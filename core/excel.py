import os
import re
from pathlib import Path
from datetime import datetime
from copy import copy

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment, PatternFill

# Importamos as nossas ferramentas e o cérebro do sistema
from utils.config import REGRAS
from utils.helpers import limpar, converter_para_numero, limpar_material_rigoroso, resource_path

def escrever_seguro(ws, coord, valor, alinhamento=None):
    """Escreve em células, tratando corretamente células mescladas."""
    try:
        cell = ws[coord]
        if cell.__class__.__name__ == 'MergedCell':
            for r in ws.merged_cells.ranges:
                if coord in r:
                    m = ws.cell(row=r.min_row, column=r.min_col)
                    m.value = valor
                    if alinhamento: 
                        m.alignment = alinhamento
                    return
        else:
            cell.value = valor
            if alinhamento: 
                cell.alignment = alinhamento
    except Exception as e: 
        print(f"[EXCEL WARN] Falha ao escrever na célula {coord}: {e}")

def tratar_cabecalho_a1(ws, id_projeto):
    """Insere logótipos ou nomes de marcas na célula A1 buscando do JSON."""
    marcas_texto = REGRAS["especiais"]["marcas_texto"]
    logos_imagem = REGRAS["especiais"]["marcas_imagem"]
    id_up = str(id_projeto).upper()
    ws['A1'].value = None

    for sigla, nome in marcas_texto.items():
        if sigla in id_up:
            escrever_seguro(ws, 'A1', nome, Alignment(horizontal='center', vertical='center'))
            ws['A1'].font = Font(size=22, bold=True)
            return

    for sigla, arq in logos_imagem.items():
        if sigla in id_up:
            path = resource_path(f"logos/{arq}.png")
            if Path(path).exists():
                img = OpenpyxlImage(path)
                img.width, img.height = 152, 42
                ws.row_dimensions[1].height = 33
                ws.add_image(img, 'A1')
                return

    # Padrão Ingecon se não encontrar marca
    path_ing = resource_path("logos/ingecon.png")
    if Path(path_ing).exists():
        img = OpenpyxlImage(path_ing)
        img.width, img.height = 152, 42
        ws.row_dimensions[1].height = 33
        ws.add_image(img, 'A1')

def ajustar_molde_elastico(ws, num_itens):
    """Redimensiona a planilha conforme a quantidade de itens."""
    for r in range(1, 50):
        ws.row_dimensions[r].hidden = False

    padrao, l_rodape, quadro = 3, 9, None
    
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= l_rodape:
            if m.min_row == l_rodape: 
                quadro = {'min_col': m.min_col, 'max_col': m.max_col, 'max_row': m.max_row}
            try: ws.unmerge_cells(str(m))
            except Exception: pass
            
    if quadro and quadro['max_row'] > l_rodape: 
        ws.delete_rows(l_rodape + 1, quadro['max_row'] - l_rodape)
        
    diff = num_itens - padrao
    if diff > 0: ws.insert_rows(l_rodape, diff)
    elif diff < 0: ws.delete_rows(l_rodape + diff, abs(diff))
    
    for r in range(6, 6 + num_itens):
        ws.row_dimensions[r].height = 25.5
        ws.cell(row=r, column=4).value = ws.cell(row=r, column=7).value = "X"
        if r > 6:
            for c in range(1, 16):
                src, tgt = ws.cell(row=6, column=c), ws.cell(row=r, column=c)
                if src.has_style: tgt._style = copy(src._style)
                
    n_inicio = 6 + num_itens
    if quadro: 
        ws.merge_cells(start_row=n_inicio, start_column=quadro['min_col'], 
                       end_row=n_inicio, end_column=quadro['max_col'])
    return n_inicio

def gerar_arquivo_excel(pai, blocos, id_proj, qtd_tot, molde, pasta, pai_is_prensado):
    """Cria o arquivo Excel formatado com os dados coletados."""
    wb = load_workbook(molde, keep_vba=True)
    ws = wb.active
    total_itens = sum(len(b['itens']) + (1 if b['tipo'] == 'prensado' else 0) for b in blocos)
    
    l_obs = ajustar_molde_elastico(ws, total_itens)
    
    # Estilos dos Botões visuais
    fill_botao = PatternFill(start_color='0078D7', end_color='0078D7', fill_type='solid')
    fill_erro = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    font_botao = Font(color='FFFFFF', bold=True, size=10)
    font_veio = Font(color='FFFFFF', bold=True, size=8)
    align_botao = Alignment(horizontal='center', vertical='center')

    # Título do Projeto
    cod_p, acab_p, desc_p = limpar(pai[1]), limpar(pai[2]), limpar(pai[3])
    acab_real = acab_p.strip(" _-")
    tit = f"{cod_p}_{acab_real} - {desc_p}" if acab_real else f"{cod_p} - {desc_p}"
    
    tratar_cabecalho_a1(ws, id_proj)
    escrever_seguro(ws, 'B3', tit)
    try: 
        ws['A3'].value = float(str(qtd_tot).replace(',', '.'))
    except Exception: 
        ws['A3'].value = qtd_tot
    
    ws['M2'].value = datetime.now().strftime('%d/%m/%Y')
    materiais_com_veio = REGRAS["especiais"]["materiais_com_veio"]
    materiais_especiais = REGRAS["especiais"]["materiais_plus_5mm"]

    row_idx = 6
    for b in sorted(blocos, key=lambda x: 0 if x['tipo'] == 'normal' else 1):
        bloco_e_prensado = (b['tipo'] == 'prensado' or pai_is_prensado)
        
        # --- PRÉ-CÁLCULO DAS DIMENSÕES PARA A DESCRIÇÃO DO PRENSADO ---
        espessura = "0"
        dim_c, dim_f = 0, 0
        
        if bloco_e_prensado:
            max_c, max_f = 0, 0
            for item in b['itens']:
                v_c_raw = limpar(item.get(15, "")) if limpar(item.get(15, "")) not in ["", "-", "="] else item.get(8, "")
                v_l_raw = limpar(item.get(16, "")) if limpar(item.get(16, "")) not in ["", "-", "="] else item.get(10, "")
                
                val_c = converter_para_numero(v_c_raw)
                val_l = converter_para_numero(v_l_raw)
                
                if isinstance(val_c, int) and val_c > max_c: max_c = val_c
                if isinstance(val_l, int) and val_l > max_f: max_f = val_l
            
            titulo_bloco = limpar(b['prensado_info'][3]) if b['tipo'] == 'prensado' else desc_p
            nums_titulo = re.findall(r'\d+', titulo_bloco)
            espessura = nums_titulo[-1] if nums_titulo else "0"
            
            dim_c = max_c - 10 if max_c > 10 else max_c
            dim_f = max_f - 10 if max_f > 10 else max_f
        # -------------------------------------------------

        # Cabeçalho Visual do Bloco Prensado
        if b['tipo'] == 'prensado':
            ws.row_dimensions[row_idx].height = 15.75
            p_d = b['prensado_info']
            ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=13)
            cell_h = ws.cell(row=row_idx, column=2)
            cell_h.value = f"{limpar(p_d[1])} - {limpar(p_d[3])}"
            cell_h.font = Font(bold=True, size=15)
            cell_h.alignment = align_botao
            row_idx += 1
        
        # Escrita dos Itens
        for item in b['itens']:
            r, is_migrado = row_idx, item.get('is_migrado', False)
            
            if is_migrado:
                d_l, m_l = item['desc_orig'], item['mat_orig']
                fita_str, veio_val = str(item['fita_orig']), item['veio_orig']
                txt_comp = f"{d_l} {m_l}".upper()
            else:
                desc_f = str(limpar(item.get(3, "")))
                txt_comp = f"{str(item.get(2, ''))} {desc_f} {str(item.get(14, ''))}".upper()
                d_l, m_l = (desc_f.split(" - ", 1) if " - " in desc_f else ("-", desc_f))
                if any(m in txt_comp for m in materiais_especiais): 
                    d_l = str(limpar(item.get(14, "")))
                fita_str = "SEC-LAM" if any(str(limpar(item.get(x, ""))) in ['-', '='] for x in [15, 16]) else "SEC"
                veio_val = None

            is_especial = any(m in txt_comp for m in materiais_especiais)
            
            # SE FOR PRENSADO, ITENS ESPECIAIS NÃO GANHAM +5
            plus = 5 if (is_especial and not is_migrado and not bloco_e_prensado) else 0

            try: 
                val_fat = float(item.get('q_unitaria_fatorada', 0))
            except Exception: 
                val_fat = 0.0
            
            # Fórmula de Quantidade
            c_formula = ws.cell(row=r, column=1)
            c_formula.value = f"={val_fat}*A3"
            if val_fat == 0: 
                c_formula.fill = fill_erro
            
            # Dimensões (Comprimento e Largura)
            val_15 = limpar(item.get(15, ""))
            v_c_raw = val_15 if val_15 not in ["", "-", "="] else item.get(8, "")
            v_c = converter_para_numero(v_c_raw)
            if isinstance(v_c, int): v_c += plus
            ws.cell(row=r, column=2).value = val_15 if val_15 in ["-", "="] else ""
            ws.cell(row=r, column=3).value = v_c
            ws.cell(row=r, column=4).value = "X"

            val_16 = limpar(item.get(16, ""))
            v_l_raw = val_16 if val_16 not in ["", "-", "="] else item.get(10, "")
            v_l = converter_para_numero(v_l_raw)
            if isinstance(v_l, int): v_l += plus
            ws.cell(row=r, column=5).value = val_16 if val_16 in ["-", "="] else ""
            ws.cell(row=r, column=6).value = v_l
            ws.cell(row=r, column=7).value = "X"

            # -------------------------------------------------------------------------
            # APLICAÇÃO DA REGRA DO PRENSADO APENAS NA DESCRIÇÃO (Coluna L / 12)
            # -------------------------------------------------------------------------
            ws.cell(row=r, column=8).value = converter_para_numero(item.get(12, ""))
            
            if bloco_e_prensado:
                ws.cell(row=r, column=12).value = f"{dim_c}X{dim_f}X{espessura}"
            else:
                ws.cell(row=r, column=12).value = converter_para_numero(d_l)
            # -------------------------------------------------------------------------

            ws.cell(row=r, column=9).value = limpar_material_rigoroso(m_l)

            if is_migrado:
                ws.cell(row=r, column=10).value = veio_val
            else:
                tem_v = any(m in txt_comp for m in materiais_com_veio)
                if "KRION" in txt_comp and str(converter_para_numero(item.get(12, ""))) == "3": 
                    tem_v = True
                if tem_v: 
                    ws.cell(row=r, column=10).value = 1
            
            ws.cell(row=r, column=11).value = fita_str
            ws.cell(row=r, column=13).value = converter_para_numero(item.get(1, ""))
            
            # Adição visual dos botões de Macro
            if ws.cell(row=r, column=10).value == 1:
                cel_v = ws.cell(row=r, column=15)
                cel_v.value = "⇄"
                cel_v.fill = fill_botao
                cel_v.font = font_veio
                cel_v.alignment = align_botao
            
            if not bloco_e_prensado and not is_especial:
                cel_n = ws.cell(row=r, column=14)
                cel_n.value = "+5"
                cel_n.fill = fill_botao
                cel_n.font = font_botao
                cel_n.alignment = align_botao
            
            row_idx += 1
    
    escrever_seguro(ws, f"A{l_obs}", f"PROJETO DE REFERÊNCIA: {id_proj}", 
                         Alignment(horizontal='left', vertical='center'))
    
    # Salva o arquivo final
    nome_f = re.sub(r'[\\/*?:\u0022<>|]', '', tit).strip()[:120]
    if not nome_f: nome_f = f"PROJETO_SEM_NOME_{cod_p}"
    
    caminho_final = os.path.join(pasta, f"{nome_f}.xlsm")
    tmp_save = caminho_final + ".tmp"
    wb.save(tmp_save)
    if os.path.exists(caminho_final): 
        os.remove(caminho_final)
    os.rename(tmp_save, caminho_final)