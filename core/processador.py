import os
import io
import re
import shutil
import pandas as pd
from pathlib import Path

from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero
from core.migracao import mapear_rede_cache, verificar_duplicidade_em_rede, extrair_dados_migracao
from core.excel import gerar_arquivo_excel

# ---------------------------------------------------------------------------
# Constantes de caminho (lidas do regras.json — nunca hardcoded)
# ---------------------------------------------------------------------------
_RAIZ          = Path(REGRAS["diretorios"]["raiz"])          # …/PLANOS DE CORTE 2026
_PASTA_ANTIGOS = Path(REGRAS["diretorios"]["antigos"])       # X:\Egpe\ANTIGOS - NÃO USAR
_MARCADOR_2026 = "PLANOS DE CORTE 2026"                      # substring de verificação


def f_valido(f):
    c  = str(limpar(f.get(1, "")))
    a  = str(f.get(2, "")).upper()
    d  = str(f.get(3, "")).upper()
    mc = str(limpar(f.get(14, "")))
    filtros = REGRAS["filtros"]
    if any(x in a  for x in filtros["descricoes_ignoradas"]) or \
       any(x in d  for x in filtros["materiais_ignorados"])  or \
       any(x in mc for x in filtros["materiais_ignorados"]): return False
    if '*' in a and not c.startswith(tuple(filtros["prefixos_validos"])): return False
    if mc.startswith(tuple(filtros["mp_iniciais_ignoradas"])):
        is_especial = any(m in d for m in REGRAS["especiais"]["materiais_plus_5mm"]) or re.search(r'\bTS\b', d)
        return c.startswith(tuple(filtros["prefixos_validos"])) if is_especial else False
    return c.startswith(tuple(filtros["prefixos_validos"]))


def is_prensado(r):
    desc = str(r.get(3, "")).upper()
    acab = str(r.get(2, "")).upper()
    cod  = str(limpar(r.get(1, "")))
    return any(g in desc for g in REGRAS["prensados"]["descricoes_gatilho"]) or \
           cod in REGRAS["prensados"]["codigos_gatilho"]                      or \
           any(g in acab for g in REGRAS["prensados"]["acabamentos_gatilho"])


# ---------------------------------------------------------------------------
# Rotina de migração segura
# Fluxo obrigatório (nunca pular etapas):
#   1. Verificar se já existe em PLANOS DE CORTE 2026  → se sim, bloquear
#   2. Criar backup em memória do arquivo original
#   3. Processar / gerar novo arquivo
#   4. Mover original para ANTIGOS - NÃO USAR
#   5. Validação crítica — confirmar que original está em ANTIGOS
#   6. Finalizar (backup descartado automaticamente ao sair do escopo)
# Em qualquer falha → rollback + erro detalhado
# ---------------------------------------------------------------------------
def _migrar_arquivo_com_seguranca(caminho_original: Path, pasta_destino: Path,
                                   pai, blocos, id_proj, qtd_tot, molde, pai_is_prensado):
    """
    Executa a migração com segurança total contra perda de dados.
    Retorna o caminho do novo arquivo gerado.
    Lança Exception com mensagem detalhada em qualquer falha.
    """
    nome_original = caminho_original.name

    # ------------------------------------------------------------------
    # ETAPA 1 — Verificação: arquivo NÃO pode estar em PLANOS DE CORTE 2026
    # (dupla garantia — processador.py já filtra, mas defendemos aqui também)
    # ------------------------------------------------------------------
    caminho_str = str(caminho_original).replace("\\", "/")
    if _MARCADOR_2026 in caminho_str:
        raise Exception(
            f"[MIGRAÇÃO BLOQUEADA] '{nome_original}' já está em '{_MARCADOR_2026}'. "
            f"Nenhuma ação foi realizada."
        )

    # ------------------------------------------------------------------
    # ETAPA 2 — Backup em memória
    # ------------------------------------------------------------------
    try:
        with open(caminho_original, "rb") as f:
            backup_bytes = io.BytesIO(f.read())
        backup_bytes.seek(0)
        # Confirma que o backup tem conteúdo
        if backup_bytes.getbuffer().nbytes == 0:
            raise Exception("Backup resultou em 0 bytes.")
    except Exception as e:
        raise Exception(
            f"[MIGRAÇÃO ABORTADA — ETAPA 2] Falha ao criar backup de '{nome_original}': {e}\n"
            f"Arquivo original intocado."
        )

    # ------------------------------------------------------------------
    # ETAPA 3 — Processamento: gerar novo arquivo em pasta de destino
    # O arquivo original NÃO é tocado nesta etapa.
    # ------------------------------------------------------------------
    try:
        if not pasta_destino.exists():
            pasta_destino.mkdir(parents=True, exist_ok=True)
        gerar_arquivo_excel(pai, blocos, id_proj, qtd_tot, molde, str(pasta_destino), pai_is_prensado)
    except Exception as e:
        raise Exception(
            f"[MIGRAÇÃO ABORTADA — ETAPA 3] Falha ao gerar novo arquivo para '{nome_original}': {e}\n"
            f"Arquivo original intocado em: {caminho_original}"
        )

    # ------------------------------------------------------------------
    # ETAPA 4 — Mover original para ANTIGOS - NÃO USAR
    # ------------------------------------------------------------------
    try:
        _PASTA_ANTIGOS.mkdir(parents=True, exist_ok=True)
        destino_antigo = _PASTA_ANTIGOS / nome_original

        # Se já existe um arquivo com o mesmo nome em ANTIGOS, adiciona sufixo único
        if destino_antigo.exists():
            from datetime import datetime
            sufixo = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem   = caminho_original.stem
            ext    = caminho_original.suffix
            destino_antigo = _PASTA_ANTIGOS / f"{stem}_{sufixo}{ext}"

        shutil.move(str(caminho_original), str(destino_antigo))

    except Exception as e:
        # Geração foi bem-sucedida mas não conseguimos mover o original.
        # Rollback: o novo arquivo já existe mas o original também — não há perda de dados.
        # Apenas reportamos o problema sem apagar nada.
        raise Exception(
            f"[MIGRAÇÃO PARCIAL — ETAPA 4] Novo arquivo gerado com sucesso, mas falha ao mover "
            f"'{nome_original}' para ANTIGOS: {e}\n"
            f"Ação necessária: mova manualmente '{caminho_original}' para '{_PASTA_ANTIGOS}'."
        )

    # ------------------------------------------------------------------
    # ETAPA 5 — Validação crítica: confirmar que original está em ANTIGOS
    # ------------------------------------------------------------------
    if not destino_antigo.exists():
        # Arquivo original sumiu mas não chegou em ANTIGOS — situação crítica.
        # Tentamos restaurar do backup.
        try:
            backup_bytes.seek(0)
            with open(caminho_original, "wb") as f:
                f.write(backup_bytes.read())
            raise Exception(
                f"[MIGRAÇÃO FALHOU — ETAPA 5] '{nome_original}' não foi encontrado em "
                f"'{_PASTA_ANTIGOS}' após a movimentação. "
                f"Arquivo original RESTAURADO em '{caminho_original}' via backup. "
                f"Verifique permissões da pasta ANTIGOS."
            )
        except Exception as restore_err:
            raise Exception(
                f"[MIGRAÇÃO CRÍTICA — ETAPA 5] '{nome_original}' desapareceu e a restauração "
                f"também falhou: {restore_err}\n"
                f"Verifique imediatamente '{_PASTA_ANTIGOS}' e '{caminho_original}'."
            )

    # ------------------------------------------------------------------
    # ETAPA 6 — Finalização (backup descartado — sai do escopo)
    # ------------------------------------------------------------------
    return destino_antigo  # retorna onde o original foi arquivado


# ---------------------------------------------------------------------------
# Processamento principal
# ---------------------------------------------------------------------------
def processar_clipboard(is_teste=False):
    df = pd.read_clipboard(sep='\t', header=None, dtype=str).fillna('')
    dir_sistema = Path(REGRAS["diretorios"]["raiz"]) / REGRAS["diretorios"]["nome_pasta_sistema"]
    molde = dir_sistema / "planilha_molde.xlsm"

    if not molde.exists() and not is_teste: raise Exception("Molde não encontrado.")
    if df.shape[0] < 2 or df.shape[1] < 6:  raise Exception("Dados insuficientes no clipboard.")

    niveis_encontrados = [str(x).count('.') for x in df[0] if re.match(r'^\d+(\.\d+)*$', str(x).strip())]
    if not niveis_encontrados: raise Exception("Estrutura de níveis não identificada.")

    id_p_raw = str(df.iloc[1, 1]).strip().upper()

    if not re.match(r'^[A-Z]{3}\d+', id_p_raw):
        raise Exception(
            f"Não foi encontrado código pai do projeto. "
            f"Verifique se a opção \"Incluir Selecionado\" está ativa no PDM"
        )

    prefixos_validos = tuple(REGRAS["filtros"]["prefixos_validos"])
    if not any(str(limpar(c)).startswith(prefixos_validos) for c in df[1]):
        raise Exception("Nenhum código primordial válido encontrado na seleção.")

    id_p = re.sub(r'[\\/*?:"<>|]', '-', id_p_raw)
    desktop_path  = Path(os.path.join(os.path.expanduser("~"), "Desktop"))
    mapeamento    = REGRAS["diretorios"]["mapeamento_pastas"]
    pasta_marca   = next((v for k, v in mapeamento.items() if k in id_p), "Outros")
    pasta         = (desktop_path / "TESTES_GERADOR" / id_p) if is_teste \
                    else (_RAIZ / pasta_marca / id_p)

    niv_pai = min(niveis_encontrados)
    if not limpar(df.iloc[1, 1]).startswith(tuple(REGRAS["filtros"]["prefixos_validos"])):
        niv_pai += 1

    cache_rede = {} if is_teste else mapear_rede_cache()

    # Monta dicionário de pais
    cons = {}
    for _, r in df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        if cod.startswith(prefixos_validos) and nv.count('.') == niv_pai:
            if nv not in cons:
                cons[nv] = {'pai': r, 'blocos': [], 'qtd_p_total': 0,
                             'excluir_prefixos': [], 'niv_base': niv_pai}
            cons[nv]['qtd_p_total'] += float(converter_para_numero(r[5]) or 0)

    for _, r in df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        acab = limpar(r[2])
        if nv.count('.') == niv_pai + 1:
            nv_pai_str = nv.rsplit('.', 1)[0]
            if nv_pai_str in cons:
                pai_r = cons[nv_pai_str]['pai']
                if limpar(pai_r[1]).startswith('15') and not limpar(pai_r[2]) and acab:
                    if acab and acab != "YF54":
                        if nv not in cons:
                            qtd_p = cons[nv_pai_str]['qtd_p_total']
                            qtd_f = float(converter_para_numero(r[5]) or 0)
                            cons[nv] = {'pai': r, 'blocos': [], 'qtd_p_total': qtd_p * qtd_f,
                                        'excluir_prefixos': [], 'niv_base': niv_pai + 1}
                            cons[nv_pai_str]['excluir_prefixos'].append(nv)

    # ------------------------------------------------------------------
    # Classificação: duplicado (em 2026) | candidato à migração | novo
    # ------------------------------------------------------------------
    arquivos_duplicados  = []   # já estão em PLANOS DE CORTE 2026 — não processar
    arquivos_migracao    = []   # encontrados fora de 2026 — migrar
    arquivos_bloqueados  = []   # bloqueados com aviso para o usuário (legado: mantido para UI)
    processar_list       = []   # processar normalmente (sem arquivo na rede)

    for nv_p, info in cons.items():
        cod_p   = limpar(info['pai'][1])
        cam_net = None if is_teste else verificar_duplicidade_em_rede(cod_p, cache_rede)

        if cam_net:
            cam_net_str = str(cam_net).replace("\\", "/")
            if _MARCADOR_2026 in cam_net_str:
                # Já existe no formato novo — duplicado, não processar
                pasta_origem = os.path.basename(os.path.dirname(cam_net))
                arquivos_duplicados.append(
                    f"• Peça {cod_p} (Já existe em: {pasta_origem})"
                )
                arquivos_bloqueados.append(
                    f"• Peça {cod_p} (Já existe em: {pasta_origem})"
                )
            else:
                # Existe fora de 2026 — candidato à migração
                arquivos_migracao.append((nv_p, info, Path(cam_net)))
        else:
            processar_list.append((nv_p, info))

    arquivos_gerados_count = 0
    arquivos_migrados      = []   # nomes para popup informativo (RN-040)
    erros_migracao         = []   # falhas não críticas acumuladas

    def garantir_pasta():
        if not os.path.exists(pasta):
            os.makedirs(pasta)

    # ------------------------------------------------------------------
    # Processamento normal (sem arquivo na rede)
    # ------------------------------------------------------------------
    for nv_p, info in processar_list:
        mask_normal  = df[0].str.startswith(nv_p + ".")
        tem_filhos   = mask_normal.any()

        if not info['blocos']:
            _montar_blocos(df, nv_p, info, niv_pai)

        tem_validos  = any(len(b['itens']) > 0 for b in info['blocos'])

        # Gera apenas se: tem filhos válidos  OU  não tem filhos (pai isolado válido)
        if tem_validos or (not tem_filhos and f_valido(info['pai'])):
            garantir_pasta()
            if not info['blocos']:
                info['blocos'] = [{'tipo': 'normal',
                                   'itens': [{'q_unitaria_fatorada': 1.0,
                                              **info['pai'].to_dict()}]}]
            gerar_arquivo_excel(
                info['pai'], info['blocos'], id_p,
                info['qtd_p_total'], molde, pasta, is_prensado(info['pai'])
            )
            arquivos_gerados_count += 1

    # ------------------------------------------------------------------
    # Migração segura (arquivo antigo encontrado fora de PLANOS DE CORTE 2026)
    # ------------------------------------------------------------------
    for nv_p, info, caminho_antigo in arquivos_migracao:
        # Tenta extrair dados do arquivo antigo
        blocos_mig, qtd_mig = extrair_dados_migracao(str(caminho_antigo))

        if blocos_mig:
            # Usa dados migrados + quantidade da planilha antiga
            info['blocos']      = blocos_mig
            info['qtd_p_total'] = qtd_mig
        else:
            # Extração falhou ou vazia — processa do zero (RN-038)
            _montar_blocos(df, nv_p, info, niv_pai)

        # Para migração: "tem filhos" = extração retornou dados OU há filhos no df
        mask_mig        = df[0].str.startswith(nv_p + ".")
        tem_filhos_mig  = bool(blocos_mig) or mask_mig.any()
        tem_validos_mig = any(len(b['itens']) > 0 for b in info['blocos'])

        if tem_validos_mig or (not tem_filhos_mig and f_valido(info['pai'])):
            garantir_pasta()
            if not info['blocos']:
                info['blocos'] = [{'tipo': 'normal',
                                   'itens': [{'q_unitaria_fatorada': 1.0,
                                              **info['pai'].to_dict()}]}]
            try:
                _migrar_arquivo_com_seguranca(
                    caminho_original = caminho_antigo,
                    pasta_destino    = pasta,
                    pai              = info['pai'],
                    blocos           = info['blocos'],
                    id_proj          = id_p,
                    qtd_tot          = info['qtd_p_total'],
                    molde            = molde,
                    pai_is_prensado  = is_prensado(info['pai'])
                )
                arquivos_migrados.append(f"• {caminho_antigo.name}")
                arquivos_gerados_count += 1
            except Exception as e:
                erros_migracao.append(str(e))

    # Erros de migração não críticos: acumula e informa ao final
    if erros_migracao:
        raise Exception(
            "Alguns arquivos não puderam ser migrados:\n\n" +
            "\n\n".join(erros_migracao)
        )

    return {
        "pasta":           str(pasta),
        "bloqueados":      arquivos_bloqueados,   # duplicados em 2026
        "migrados":        arquivos_migrados,      # migrados com sucesso (RN-040)
        "aviso":           "Nada gerado." if arquivos_gerados_count == 0
                           and not arquivos_bloqueados else None
    }


# ---------------------------------------------------------------------------
# Helper: monta blocos de itens a partir do DataFrame (extraído do loop original)
# ---------------------------------------------------------------------------
def _montar_blocos(df, nv_p, info, niv_pai):
    c_p      = limpar(info['pai'][1])
    niv_base = info.get('niv_base', niv_pai)
    mask     = df[0].str.startswith(nv_p + ".")
    for excl in info.get('excluir_prefixos', []):
        mask = mask & ~(df[0] == excl) & ~df[0].str.startswith(excl + ".")
    desc_df = df[mask].copy()

    b_roots = {}
    for _, r in desc_df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        if (c_p.startswith('15') and cod.startswith('15') and nv.count('.') > niv_base) \
                or is_prensado(r):
            pref       = [p for p in b_roots.keys() if nv.startswith(p + ".")]
            parent_qf  = b_roots[max(pref, key=len)]['qf'] if pref else 1.0
            b_roots[nv] = {'tipo': 'prensado', 'prensado_info': r, 'itens': [],
                            'qf': float(converter_para_numero(r[5]) or 1) * parent_qf}

    bloco_a = {'tipo': 'normal', 'itens': []}
    for _, r in desc_df.iterrows():
        nv     = limpar(r[0])
        pref   = [p for p in b_roots.keys() if nv.startswith(p + ".")]
        parent = b_roots[max(pref, key=len)] if pref else None
        if not (nv in b_roots) and f_valido(r):
            ic = r.copy().to_dict()
            if parent:
                ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0) * parent['qf']
                parent['itens'].append(ic)
            elif nv.count('.') == niv_base + 1:
                ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0)
                bloco_a['itens'].append(ic)

    if bloco_a['itens']:
        info['blocos'].append(bloco_a)
    for br in b_roots.values():
        if br['itens']:
            info['blocos'].append(br)