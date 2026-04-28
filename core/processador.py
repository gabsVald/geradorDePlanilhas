"""
================================================================================
core/processador.py — Processamento Principal do Clipboard
================================================================================
Orquestra todo o fluxo de trabalho após o usuário clicar em "Colar e Gerar":

  1. Lê e valida os dados do clipboard (formato TSV do PDM)
  2. Determina a hierarquia de itens e os pais de cada planilha
  3. Classifica cada pai em: duplicado / candidato a migração / novo
  4. Para cada pai novo: monta os blocos e gera a planilha
  5. Para cada candidato a migração: executa a migração segura em 6 etapas
  6. Retorna um dicionário com resultados (pasta, bloqueados, migrados, avisos)

Módulos utilizados:
  - core/migracao.py : cache de rede, extração de dados de planilhas antigas
  - core/excel.py    : geração do arquivo .xlsm final
================================================================================
"""

# ===== IMPORTAÇÕES =====
import os
import io
import re
import shutil

import pandas as pd         # Leitura do clipboard (TSV) e manipulação de DataFrames
from pathlib import Path

from utils.config import REGRAS               # Regras globais (regras.json)
from utils.helpers import limpar, converter_para_numero
from core.migracao import (
    mapear_rede_cache,           # Varre arquivos na rede
    verificar_duplicidade_em_rede,  # Busca um código no cache da rede
    extrair_dados_migracao       # Extrai dados de planilhas antigas
)
from core.excel import gerar_arquivo_excel    # Gera o .xlsm final


# ===== CONSTANTES DE CAMINHO =====
# Lidas do regras.json para evitar caminhos hardcoded no código

_RAIZ          = Path(REGRAS["diretorios"]["raiz"])     # .../PLANOS DE CORTE 2026
_PASTA_ANTIGOS = Path(REGRAS["diretorios"]["antigos"])  # X:\Egpe\ANTIGOS - NÃO USAR
_MARCADOR_2026 = "PLANOS DE CORTE 2026"                 # Substring usada para distinguir arquivos novos dos antigos


# ===== FUNÇÕES DE VALIDAÇÃO DE ITENS =====

def f_valido(f):
    """
    Determina se um item (linha do DataFrame) deve ser incluído na planilha de corte.

    Aplica as seguintes regras em ordem:
      1. Descricoes ignoradas (ex: "CORTE") → exclui
      2. Material ignorado (fitas, tecidos, adesivos...) → exclui
      3. MP com prefixo bloqueado (9172, 93...) → exclui, exceto:
           - "LAMINADO FORM": sempre passa
           - Materiais especiais (+5mm) ou código TS: passa
      4. Código deve começar com prefixo válido ("11" ou "15") → inclui

    Parâmetros:
        f (Series|dict): Linha do DataFrame com dados do item.

    Retorna:
        bool: True se o item deve ser processado, False se deve ser ignorado.
    """
    c  = str(limpar(f.get(1, "")))   # Código do item (col 1)
    a  = str(f.get(2, "")).upper()   # Acabamento/descrição resumida (col 2)
    d  = str(f.get(3, "")).upper()   # Descrição completa (col 3)
    mc = str(limpar(f.get(14, "")))  # Material complementar / código MP (col 14)

    filtros = REGRAS["filtros"]

    # Regra 1: descrição ou material na lista de exclusão
    if any(x in a  for x in filtros["descricoes_ignoradas"]) or \
       any(x in d  for x in filtros["materiais_ignorados"])  or \
       any(x in mc for x in filtros["materiais_ignorados"]):
        return False

    # Regra 2: itens com "*" no acabamento só passam se tiverem código válido
    if '*' in a and not c.startswith(tuple(filtros["prefixos_validos"])):
        return False

    # Regra 3: MP com prefixo bloqueado (ex: 9172xx = serviços externos)
    if mc.startswith(tuple(filtros["mp_iniciais_ignoradas"])):
        # Exceção explícita: LAMINADO FORM sempre passa (código PDM especial)
        if "LAMINADO FORM" in d:
            return c.startswith(tuple(filtros["prefixos_validos"]))
        # Materiais +5mm (KRION, DURASEIN...) ou código TS também passam
        is_especial = any(m in d for m in REGRAS["especiais"]["materiais_plus_5mm"]) or re.search(r'\bTS\b', d)
        return c.startswith(tuple(filtros["prefixos_validos"])) if is_especial else False

    # Regra geral: código deve começar com "11" ou "15"
    return c.startswith(tuple(filtros["prefixos_validos"]))


def is_prensado(r):
    """
    Verifica se uma linha do DataFrame representa um item "prensado".

    Gatilhos configuráveis em regras.json (prensados):
      - descricoes_gatilho: palavras na descrição (ex: "PRENSADO")
      - codigos_gatilho: códigos específicos (ex: "1152032")
      - acabamentos_gatilho: acabamentos especiais (ex: "PRLA")

    Parâmetros:
        r (Series|dict): Linha do DataFrame.

    Retorna:
        bool: True se o item é um bloco prensado.
    """
    desc = str(r.get(3, "")).upper()
    acab = str(r.get(2, "")).upper()
    cod  = str(limpar(r.get(1, "")))

    return (
        any(g in desc for g in REGRAS["prensados"]["descricoes_gatilho"]) or
        cod in REGRAS["prensados"]["codigos_gatilho"]                      or
        any(g in acab for g in REGRAS["prensados"]["acabamentos_gatilho"])
    )


# ===== MIGRAÇÃO SEGURA EM 6 ETAPAS =====

def _migrar_arquivo_com_seguranca(
    caminho_original: Path,
    pasta_destino: Path,
    pai, blocos, id_proj, qtd_tot, molde, pai_is_prensado
):
    """
    Executa a migração de um arquivo antigo com segurança total contra perda de dados.

    Fluxo obrigatório (nunca pular etapas):
      Etapa 1 — Verificação de segurança: garante que o arquivo não está em PLANOS DE CORTE 2026
      Etapa 2 — Backup em memória do arquivo original (BytesIO)
      Etapa 3 — Geração do novo arquivo .xlsm na pasta de destino
      Etapa 4 — Movimentação do original para ANTIGOS - NÃO USAR
      Etapa 5 — Validação crítica: confirma que o original chegou em ANTIGOS
      Etapa 6 — Finalização (backup descartado ao sair do escopo)

    Em caso de falha em qualquer etapa:
      - O original nunca é apagado antes da geração ser confirmada
      - Se o arquivo sumir antes de chegar em ANTIGOS, o backup é restaurado
      - Mensagens de erro detalhadas indicam exatamente o que aconteceu

    Parâmetros:
        caminho_original (Path): Caminho completo do arquivo a migrar.
        pasta_destino (Path): Pasta onde o novo .xlsm será criado.
        pai: Linha do DataFrame com dados do pai.
        blocos (list): Blocos de itens extraídos do arquivo antigo.
        id_proj (str): Código do projeto.
        qtd_tot (float): Quantidade total (A3 do arquivo antigo).
        molde (Path): Caminho do molde .xlsm.
        pai_is_prensado (bool): Se True, todos os blocos herdam status de prensado.

    Retorna:
        Path: Caminho onde o original foi arquivado em ANTIGOS.

    Lança:
        Exception: Com mensagem detalhada em qualquer falha.
    """
    nome_original = caminho_original.name

    # ------------------------------------------------------------------
    # ETAPA 1 — VERIFICAÇÃO DE SEGURANÇA
    # Dupla garantia: processador.py já filtra por _MARCADOR_2026, mas
    # verificamos novamente aqui para proteger contra chamadas diretas.
    # ------------------------------------------------------------------
    caminho_str = str(caminho_original).replace("\\", "/")
    if _MARCADOR_2026 in caminho_str:
        raise Exception(
            f"[MIGRAÇÃO BLOQUEADA] '{nome_original}' já está em '{_MARCADOR_2026}'. "
            f"Nenhuma ação foi realizada."
        )

    # ------------------------------------------------------------------
    # ETAPA 2 — BACKUP EM MEMÓRIA
    # O backup em BytesIO permite restaurar o original se algo der errado
    # nas etapas seguintes, sem dependência de espaço em disco.
    # ------------------------------------------------------------------
    try:
        with open(caminho_original, "rb") as f:
            backup_bytes = io.BytesIO(f.read())
        backup_bytes.seek(0)
        if backup_bytes.getbuffer().nbytes == 0:
            raise Exception("Backup resultou em 0 bytes.")
    except Exception as e:
        raise Exception(
            f"[MIGRAÇÃO ABORTADA — ETAPA 2] Falha ao criar backup de '{nome_original}': {e}\n"
            f"Arquivo original intocado."
        )

    # ------------------------------------------------------------------
    # ETAPA 3 — GERAÇÃO DO NOVO ARQUIVO
    # O original NÃO é tocado nesta etapa. Se a geração falhar,
    # o original permanece intacto e o processo é abortado.
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
    # ETAPA 4 — MOVIMENTAÇÃO PARA ANTIGOS
    # Se já existe arquivo com mesmo nome em ANTIGOS, adiciona timestamp
    # para evitar sobrescrita. shutil.move em volumes de rede pode copiar
    # sem apagar a origem, então verificamos e removemos explicitamente.
    # ------------------------------------------------------------------
    try:
        _PASTA_ANTIGOS.mkdir(parents=True, exist_ok=True)
        destino_antigo = _PASTA_ANTIGOS / nome_original

        if destino_antigo.exists():
            # Adiciona sufixo timestamp para evitar colisão de nomes
            from datetime import datetime
            sufixo = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem   = caminho_original.stem
            ext    = caminho_original.suffix
            destino_antigo = _PASTA_ANTIGOS / f"{stem}_{sufixo}{ext}"

        shutil.move(str(caminho_original), str(destino_antigo))

        # Verificação extra: shutil.move entre volumes de rede pode copiar sem apagar
        if destino_antigo.exists() and caminho_original.exists():
            try:
                caminho_original.unlink()
            except Exception as del_err:
                raise Exception(
                    f"[MIGRAÇÃO PARCIAL — ETAPA 4] Arquivo copiado para ANTIGOS mas não foi "
                    f"possível remover o original: {del_err}\n"
                    f"Remova manualmente: '{caminho_original}'"
                )

    except Exception as e:
        # Novo arquivo foi gerado com sucesso mas o original não foi movido.
        # Não há perda de dados (ambos existem). Reporta sem apagar nada.
        raise Exception(
            f"[MIGRAÇÃO PARCIAL — ETAPA 4] Novo arquivo gerado com sucesso, mas falha ao mover "
            f"'{nome_original}' para ANTIGOS: {e}\n"
            f"Ação necessária: mova manualmente '{caminho_original}' para '{_PASTA_ANTIGOS}'."
        )

    # ------------------------------------------------------------------
    # ETAPA 5 — VALIDAÇÃO CRÍTICA
    # Confirma que o original chegou em ANTIGOS. Se sumiu durante a
    # movimentação, tenta restaurar do backup em memória.
    # ------------------------------------------------------------------
    if not destino_antigo.exists():
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
    # ETAPA 6 — FINALIZAÇÃO
    # backup_bytes sai do escopo e é coletado pelo GC automaticamente.
    # ------------------------------------------------------------------
    return destino_antigo  # Caminho onde o original foi arquivado


# ===== PROCESSAMENTO PRINCIPAL =====

def processar_clipboard(is_teste=False):
    """
    Ponto de entrada principal. Lê o clipboard e gera as planilhas de corte.

    Parâmetros:
        is_teste (bool): Se True, salva na pasta Desktop/TESTES_GERADOR e
                         desativa verificações de rede e migração.

    Retorna:
        dict com:
          - "pasta"      (str): Caminho da pasta onde os arquivos foram gerados.
          - "bloqueados" (list[str]): Descrições das peças já existentes em PLANOS DE CORTE 2026.
          - "migrados"   (list[str]): Nomes dos arquivos migrados com sucesso.
          - "aviso"      (str|None): Mensagem de aviso se nada foi gerado.

    Lança:
        Exception: Se os dados do clipboard forem inválidos ou o molde não for encontrado.
    """
    # ===== LEITURA E VALIDAÇÃO DO CLIPBOARD =====
    df = pd.read_clipboard(sep='\t', header=None, dtype=str).fillna('')

    dir_sistema = Path(REGRAS["diretorios"]["raiz"]) / REGRAS["diretorios"]["nome_pasta_sistema"]
    molde = dir_sistema / "planilha_molde.xlsm"

    if not molde.exists() and not is_teste:
        raise Exception("Molde não encontrado.")
    if df.shape[0] < 2 or df.shape[1] < 6:
        raise Exception("Dados insuficientes no clipboard.")

    # Detecta a estrutura de níveis hierárquicos (ex: 1.1.1 = 2 pontos = nível 2)
    niveis_encontrados = [
        str(x).count('.') for x in df[0]
        if re.match(r'^\d+(\.\d+)*$', str(x).strip())
    ]
    if not niveis_encontrados:
        raise Exception("Estrutura de níveis não identificada.")

    # Valida o ID do projeto: deve começar com 3 letras + dígitos (ex: ZAR001)
    id_p_raw = str(df.iloc[1, 1]).strip().upper()
    if not re.match(r'^[A-Z]{3}\d+', id_p_raw):
        raise Exception(
            f"Não foi encontrado código pai do projeto. "
            f"Verifique se a opção \"Incluir Selecionado\" está ativa no PDM"
        )

    prefixos_validos = tuple(REGRAS["filtros"]["prefixos_validos"])
    if not any(str(limpar(c)).startswith(prefixos_validos) for c in df[1]):
        raise Exception("Nenhum código primordial válido encontrado na seleção.")

    # Remove caracteres inválidos em nomes de arquivo do Windows
    id_p = re.sub(r'[\\/*?:"<>|]', '-', id_p_raw)

    # ===== DEFINIÇÃO DA PASTA DE DESTINO =====
    desktop_path = Path(os.path.join(os.path.expanduser("~"), "Desktop"))
    mapeamento   = REGRAS["diretorios"]["mapeamento_pastas"]
    pasta_marca  = next((v for k, v in mapeamento.items() if k in id_p), "Outros")

    # Modo teste: salva no Desktop; modo produção: salva na rede
    pasta = (
        desktop_path / "TESTES_GERADOR" / id_p
        if is_teste
        else _RAIZ / pasta_marca / id_p
    )

    # ===== DETERMINAÇÃO DO NÍVEL PAI =====
    niv_pai = min(niveis_encontrados)
    # Se o item na linha 1 não é um código válido, o nível pai está um nível abaixo
    if not limpar(df.iloc[1, 1]).startswith(tuple(REGRAS["filtros"]["prefixos_validos"])):
        niv_pai += 1

    # Cache de rede desativado em modo teste para não afetar produção
    cache_rede = {} if is_teste else mapear_rede_cache()

    # ===== MONTAGEM DO DICIONÁRIO DE PAIS =====
    cons = {}  # {nível_hierárquico: {'pai': row, 'blocos': [], 'qtd_p_total': float, ...}}

    for _, r in df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        if cod.startswith(prefixos_validos) and nv.count('.') == niv_pai:
            if nv not in cons:
                cons[nv] = {
                    'pai': r, 'blocos': [], 'qtd_p_total': 0,
                    'excluir_prefixos': [], 'niv_base': niv_pai
                }
            # Acumula quantidade total (pode haver múltiplas linhas do mesmo pai)
            cons[nv]['qtd_p_total'] += float(converter_para_numero(r[5]) or 0)

    # ===== PROMOÇÃO DE FILHOS COM ACABAMENTO (Regra 15xxx) =====
    # Pais "15xxx" sem acabamento cujos filhos têm acabamento ≠ YF54 são promovidos
    # a pais adicionais, com qtd multiplicada (pai × filho)
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
                            cons[nv] = {
                                'pai': r, 'blocos': [], 'qtd_p_total': qtd_p * qtd_f,
                                'excluir_prefixos': [], 'niv_base': niv_pai + 1
                            }
                            cons[nv_pai_str]['excluir_prefixos'].append(nv)

    # ===== CLASSIFICAÇÃO: DUPLICADO / MIGRAÇÃO / NOVO =====
    arquivos_duplicados = []   # Já existem em PLANOS DE CORTE 2026 — não processar
    arquivos_migracao   = []   # Existem fora de 2026 — migrar para o novo formato
    arquivos_bloqueados = []   # Mantido para exibição na UI (mesmo conteúdo de duplicados)
    processar_list      = []   # Sem arquivo na rede — processar normalmente

    for nv_p, info in cons.items():
        cod_p   = limpar(info['pai'][1])
        cam_net = None if is_teste else verificar_duplicidade_em_rede(cod_p, cache_rede)

        if cam_net:
            cam_net_str = str(cam_net).replace("\\", "/")
            if _MARCADOR_2026 in cam_net_str:
                # Arquivo já no formato novo — duplicata, bloqueia processamento
                pasta_origem = os.path.basename(os.path.dirname(cam_net))
                msg = f"• Peça {cod_p} (Já existe em: {pasta_origem})"
                arquivos_duplicados.append(msg)
                arquivos_bloqueados.append(msg)
            else:
                # Arquivo em pasta antiga — candidato à migração
                arquivos_migracao.append((nv_p, info, Path(cam_net)))
        else:
            processar_list.append((nv_p, info))

    arquivos_gerados_count = 0
    arquivos_migrados      = []   # Para o popup informativo (RN-040)
    erros_migracao         = []   # Erros não críticos — acumulados e exibidos no final

    def garantir_pasta():
        """Cria a pasta de destino se não existir."""
        if not os.path.exists(pasta):
            os.makedirs(pasta)

    # ===== PROCESSAMENTO NORMAL (sem arquivo na rede) =====
    for nv_p, info in processar_list:
        mask_normal = df[0].str.startswith(nv_p + ".")
        tem_filhos  = mask_normal.any()

        if not info['blocos']:
            _montar_blocos(df, nv_p, info, niv_pai)

        tem_validos = any(len(b['itens']) > 0 for b in info['blocos'])

        # Gera apenas se tem filhos válidos, OU se o pai isolado for um item válido
        if tem_validos or (not tem_filhos and f_valido(info['pai'])):
            garantir_pasta()
            if not info['blocos']:
                # Pai isolado sem filhos: gera planilha com o próprio pai como único item
                info['blocos'] = [{'tipo': 'normal', 'itens': [{'q_unitaria_fatorada': 1.0, **info['pai'].to_dict()}]}]
            gerar_arquivo_excel(
                info['pai'], info['blocos'], id_p,
                info['qtd_p_total'], molde, pasta, is_prensado(info['pai'])
            )
            arquivos_gerados_count += 1

    # ===== MIGRAÇÃO SEGURA (arquivo antigo encontrado fora de PLANOS DE CORTE 2026) =====
    for nv_p, info, caminho_antigo in arquivos_migracao:
        blocos_mig, qtd_mig = extrair_dados_migracao(str(caminho_antigo))

        if blocos_mig:
            # Usa dados do arquivo antigo (layout, materiais, dimensões originais)
            info['blocos']      = blocos_mig
            info['qtd_p_total'] = qtd_mig
        else:
            # Extração falhou: processa do zero a partir do clipboard (RN-038)
            _montar_blocos(df, nv_p, info, niv_pai)

        mask_mig        = df[0].str.startswith(nv_p + ".")
        tem_filhos_mig  = bool(blocos_mig) or mask_mig.any()
        tem_validos_mig = any(len(b['itens']) > 0 for b in info['blocos'])

        if tem_validos_mig or (not tem_filhos_mig and f_valido(info['pai'])):
            garantir_pasta()
            if not info['blocos']:
                info['blocos'] = [{'tipo': 'normal', 'itens': [{'q_unitaria_fatorada': 1.0, **info['pai'].to_dict()}]}]
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
                erros_migracao.append(str(e))  # Acumula sem interromper outros projetos

    # Erros de migração: reporta todos de uma vez ao final
    if erros_migracao:
        raise Exception(
            "Alguns arquivos não puderam ser migrados:\n\n" +
            "\n\n".join(erros_migracao)
        )

    return {
        "pasta":      str(pasta),
        "bloqueados": arquivos_bloqueados,
        "migrados":   arquivos_migrados,
        "aviso":      "Nada gerado." if arquivos_gerados_count == 0 and not arquivos_bloqueados else None
    }


# ===== HELPER: MONTAGEM DE BLOCOS DO DATAFRAME =====

def _montar_blocos(df, nv_p, info, niv_pai):
    """
    Monta os blocos de itens de um pai a partir do DataFrame do clipboard.

    Separa os filhos do pai em:
      - Blocos prensados: itens agrupados sob um cabeçalho prensado
      - Bloco normal: itens individuais sem agrupamento prensado

    Parâmetros:
        df (DataFrame): Dados completos do clipboard.
        nv_p (str): Nível hierárquico do pai (ex: "1.1").
        info (dict): Dicionário do pai com 'pai', 'blocos', 'excluir_prefixos', 'niv_base'.
        niv_pai (int): Número de pontos no nível pai.
    """
    c_p      = limpar(info['pai'][1])
    niv_base = info.get('niv_base', niv_pai)

    # Filtra os filhos diretos do pai, excluindo sub-níveis promovidos
    mask = df[0].str.startswith(nv_p + ".")
    for excl in info.get('excluir_prefixos', []):
        mask = mask & ~(df[0] == excl) & ~df[0].str.startswith(excl + ".")
    desc_df = df[mask].copy()

    # ===== IDENTIFICAÇÃO DE BLOCOS PRENSADOS =====
    b_roots = {}  # {nível_hierárquico: bloco_prensado}

    for _, r in desc_df.iterrows():
        nv, cod = limpar(r[0]), limpar(r[1])
        # Pai 15xxx com filhos 15xxx em nível mais profundo = sub-componente prensado
        if (c_p.startswith('15') and cod.startswith('15') and nv.count('.') > niv_base) \
                or is_prensado(r):
            pref      = [p for p in b_roots.keys() if nv.startswith(p + ".")]
            parent_qf = b_roots[max(pref, key=len)]['qf'] if pref else 1.0
            b_roots[nv] = {
                'tipo': 'prensado',
                'prensado_info': r,
                'itens': [],
                'qf': float(converter_para_numero(r[5]) or 1) * parent_qf  # Fator acumulado
            }

    # ===== DISTRIBUIÇÃO DOS ITENS =====
    bloco_a = {'tipo': 'normal', 'itens': []}  # Bloco para itens sem prensado pai

    for _, r in desc_df.iterrows():
        nv   = limpar(r[0])
        pref = [p for p in b_roots.keys() if nv.startswith(p + ".")]
        parent = b_roots[max(pref, key=len)] if pref else None

        if not (nv in b_roots) and f_valido(r):
            ic = r.copy().to_dict()
            if parent:
                # Item filho de um bloco prensado: fator = fator do item × fator acumulado do prensado
                ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0) * parent['qf']
                parent['itens'].append(ic)
            elif nv.count('.') == niv_base + 1:
                # Item direto do pai (nível imediatamente abaixo)
                ic['q_unitaria_fatorada'] = float(converter_para_numero(r[5]) or 0)
                bloco_a['itens'].append(ic)

    # Adiciona os blocos à lista do pai (normais primeiro, prensados depois)
    if bloco_a['itens']:
        info['blocos'].append(bloco_a)
    for br in b_roots.values():
        if br['itens']:
            info['blocos'].append(br)