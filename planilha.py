import os
import sys
import shutil
import re
import threading
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from copy import copy
from openpyxl.styles import Font, Alignment

# Importações para formatação mista (Rich Text) na mesma célula
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox

# =============================
# CONFIG AUTO UPDATE
# =============================
VERSAO_ATUAL = "2.0.12"

SERVIDOR = Path(r"X:\Engenharia\GeradorPlanilhas")
ARQ_VERSAO = SERVIDOR / "version.txt"
EXE_SERVIDOR = SERVIDOR / "Gerador_Planilhas_Ingecon.exe"

COR_PRINCIPAL = "#d32732"
COR_HOVER = "#a81f28"
ctk.set_appearance_mode("Light") 

class AppIngecon(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.verificar_atualizacao()
        self.title(f"Ingecon - Gerador de Planilhas V{VERSAO_ATUAL}")
        self.geometry("480x320")
        self.configure(fg_color="#f5f5f5") 
        self.grid_columnconfigure(0, weight=1)

        self.header_frame = ctk.CTkFrame(self, fg_color=COR_PRINCIPAL, height=70, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))
        self.header_frame.grid_columnconfigure(0, weight=1)

        self.label_titulo = ctk.CTkLabel(self.header_frame, text="GERADOR DE PLANILHAS", 
                                         font=ctk.CTkFont(size=22, weight="bold"), text_color="white")
        self.label_titulo.grid(row=0, column=0, pady=20)

        self.btn_processar = ctk.CTkButton(self, text="Colar - Gerar Planilhas", command=self.iniciar_processamento,
                                           fg_color=COR_PRINCIPAL, hover_color=COR_HOVER, height=50, 
                                           corner_radius=8, font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_processar.grid(row=2, column=0, padx=40, pady=30)

        self.progress = ctk.CTkProgressBar(self, orientation="horizontal", progress_color=COR_PRINCIPAL, width=300)
        self.progress.set(0)

    def verificar_atualizacao(self):
        try:
            if ARQ_VERSAO.exists():
                with open(ARQ_VERSAO, "r") as f: v_serv = f.read().strip()
                if v_serv != VERSAO_ATUAL:
                    if messagebox.askyesno("Atualização", f"Nova versão {v_serv} disponível. Atualizar?"):
                        self.executar_patch(); os._exit(0)
        except: pass

    def executar_patch(self):
        c_exe = sys.executable
        n_exe = os.path.basename(c_exe)
        bat = f'@echo off\n:loop\ntaskkill /f /im "{n_exe}" >nul 2>&1\ndel /f /q "{c_exe}" >nul 2>&1\nif exist "{c_exe}" (timeout /t 1 >nul\ngoto loop)\ncopy /y "{EXE_SERVIDOR}" "{c_exe}"\nstart "" "{c_exe}"\nexit'
        p_bat = Path(os.environ["TEMP"]) / "patch_ingecon.bat"
        with open(p_bat, "w") as f: f.write(bat)
        os.startfile(p_bat)

    def resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'):
            p_interno = Path(sys._MEIPASS) / relative_path
            if p_interno.exists(): return str(p_interno)
        p_exe = Path(sys.executable).parent / relative_path
        if p_exe.exists(): return str(p_exe)
        return str(Path(os.getcwd()) / relative_path)

    def escrever_seguro(self, ws, coord, valor, alinhamento=None):
        try:
            cell = ws[coord]
            if cell.__class__.__name__ == 'MergedCell':
                for range_mesclado in ws.merged_cells.ranges:
                    if coord in range_mesclado:
                        master = ws.cell(row=range_mesclado.min_row, column=range_mesclado.min_col)
                        master.value = valor
                        if alinhamento: master.alignment = alinhamento
                        return
            else:
                cell.value = valor
                if alinhamento: cell.alignment = alinhamento
        except: pass

    def limpar(self, val):
        v = str(val).strip()
        if v.endswith('.0'): v = v[:-2]
        return v if v.lower() not in ['nan', ''] else ""

    def converter_para_numero(self, valor):
        limpo = self.limpar(valor)
        if not limpo: return None
        if limpo in ["-", "="]: return limpo
        try:
            v_aj = limpo.replace(',', '.')
            if '.' in v_aj: return float(v_aj)
            return int(v_aj)
        except: return limpo

    def tratar_cabecalho_a1(self, ws, id_projeto):
        logos_imagem = {
            "ARO": "amaro", "BGK": "BurguerKing", "CAM": "Camicado", "CEA": "CeA", 
            "CEN": "Centauro", "ELU": "Elubel", "FAR": "FarmaCopr", "IND": "Indian", 
            "ING": "ingecon", "MCD": "mcdonalds", "PER": "pernambucanas", "REN": "renner", 
            "TMS": "Tramontina", "TRA": "tramontinaPDV"
        }
        marcas_texto = {"ZAR": "ZARA", "ZFR": "ZAFFARI", "SEP": "SEPHORA", "PRO": "PROTÓTIPO"}
        
        id_up = str(id_projeto).upper()
        for sigla, nome_texto in marcas_texto.items():
            if sigla in id_up:
                self.escrever_seguro(ws, 'A1', nome_texto, Alignment(horizontal='center', vertical='center'))
                ws['A1'].font = Font(size=22, bold=True)
                return

        for sigla, arq_nome in logos_imagem.items():
            if sigla in id_up:
                path_logo = self.resource_path(f"logos/{arq_nome}.png")
                if Path(path_logo).exists():
                    img = OpenpyxlImage(path_logo)
                    img.width, img.height = 152, 42 
                    ws.row_dimensions[1].height = 33 
                    ws.add_image(img, 'A1')
                    return

        path_ingecon = self.resource_path("logos/ingecon.png")
        if Path(path_ingecon).exists():
            img = OpenpyxlImage(path_ingecon)
            img.width, img.height = 152, 42
            ws.row_dimensions[1].height = 33
            ws.add_image(img, 'A1')
        else:
            self.escrever_seguro(ws, 'A1', id_projeto, Alignment(horizontal='center', vertical='center'))
            ws['A1'].font = Font(size=22, bold=True)

    def limpar_material_rigoroso(self, texto):
        if not texto: return ""
        t = re.sub(r'\b(ORIG|ESS)\b', '', texto, flags=re.IGNORECASE).replace('=', '')
        t = re.sub(r'\s*\b\d+(?:[\.,]\d+)?\s*[xX].*$', '', t, flags=re.IGNORECASE)
        return re.sub(r'\s+', ' ', t).strip(' -')

    def ajustar_molde_elastico(self, ws, num_itens):
        padrao = 3
        l_rodape = 9
        quadro = None
        
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= l_rodape:
                if m.min_row == l_rodape: 
                    quadro = {'min_col': m.min_col, 'max_col': m.max_col, 'max_row': m.max_row}
                ws.unmerge_cells(str(m))
        
        if quadro and quadro['max_row'] > l_rodape:
            ws.delete_rows(l_rodape + 1, quadro['max_row'] - l_rodape)
                
        diff = num_itens - padrao
        
        if diff > 0:
            ws.insert_rows(l_rodape, diff)
        elif diff < 0:
            ws.delete_rows(l_rodape + diff, abs(diff))
            
        for r in range(6, 6 + num_itens):
            ws.row_dimensions[r].height = 25.5
            ws.cell(row=r, column=4).value = "X"
            ws.cell(row=r, column=7).value = "X"
            if r > 6:
                for c in range(1, 14):
                    src, tgt = ws.cell(row=6, column=c), ws.cell(row=r, column=c)
                    if src.has_style: tgt._style = copy(src._style)
                    
        n_inicio = 6 + num_itens
        if quadro:
            ws.merge_cells(start_row=n_inicio, start_column=quadro['min_col'], end_row=n_inicio, end_column=quadro['max_col'])
            ws.row_dimensions[n_inicio].height = 36.75
            
        return n_inicio

    def gerar_arquivo_excel(self, pai, blocos, id_proj, qtd_tot, molde, pasta):
        wb = load_workbook(molde); ws = wb.active
        total_linhas = sum(len(b['itens']) + (1 if b['tipo'] == 'prensado' else 0) for b in blocos)
        l_obs = self.ajustar_molde_elastico(ws, total_linhas)
        
        cod_p, acab_p, desc_p = self.limpar(pai[1]), self.limpar(pai[2]), self.limpar(pai[3])
        tit = f"{f'{cod_p}_{acab_p}' if acab_p else cod_p} - {desc_p}"
        
        self.tratar_cabecalho_a1(ws, id_projeto=id_proj)
        self.escrever_seguro(ws, 'B3', tit)
        
        try: ws['A3'].value = float(str(qtd_tot).replace(',', '.'))
        except: ws['A3'].value = qtd_tot

        self.escrever_seguro(ws, 'M2', datetime.now().strftime('%d/%m/%Y'))

        blocos_ordenados = sorted(blocos, key=lambda x: 0 if x['tipo'] == 'normal' else 1)

        row_idx = 6
        for bloco in blocos_ordenados:
            if bloco['tipo'] == 'prensado':
                ws.row_dimensions[row_idx].height = 15.75
                p_data = bloco['prensado_info']
                ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=13)
                cell_h = ws.cell(row=row_idx, column=2)
                cell_h.value = f"{self.limpar(p_data[1])} - {self.limpar(p_data[3])}"
                cell_h.font = Font(bold=True, size=12)
                cell_h.alignment = Alignment(horizontal='center', vertical='center')
                row_idx += 1

            for item in bloco['itens']:
                r = row_idx
                desc_full = self.limpar(item[3])
                d_l, m_l_orig = (desc_full.split(" - ", 1) if " - " in desc_full else ("-", desc_full))
                
                comp_fb, larg_fb = self.limpar(item[15]), self.limpar(item[16])
                tem_fita = any(x in [comp_fb, larg_fb] for x in ['-', '='])
                
                q_unit = item['q_unitaria_fatorada']
                cell_q = ws.cell(row=r, column=1)
                cell_q.value = f"={float(q_unit)}*A3"
                
                for col, idx in [(2,15), (3,8), (5,16), (6,10), (8,12)]:
                    v = self.converter_para_numero(item[idx])
                    c = ws.cell(row=r, column=col)
                    c.value = v
                    if isinstance(v, (int, float)):
                        c.number_format = '0' 

                ws.cell(row=r, column=9).value = self.limpar_material_rigoroso(str(m_l_orig))
                ws.cell(row=r, column=10).value = "" 
                ws.cell(row=r, column=11).value = "SEC-LAM" if tem_fita else "SEC"
                ws.cell(row=r, column=12).value = d_l
                
                cod_val = self.limpar(item[1])
                cell_cod = ws.cell(row=r, column=13)
                try:
                    if cod_val.isdigit(): cell_cod.value = int(cod_val)
                    else: cell_cod.value = float(cod_val.replace(',', '.'))
                except: cell_cod.value = cod_val
                
                row_idx += 1
        
        # Criação do texto formatado com CellRichText (Tamanho 18)
        try:
            fonte_normal = InlineFont(rFont='Arial Black', sz=18, b=False)
            fonte_negrito = InlineFont(rFont='Arial Black', sz=18, b=True)
            
            texto_quadro = CellRichText(
                TextBlock(font=fonte_normal, text="Projeto de Referência: "),
                TextBlock(font=fonte_negrito, text=id_proj)
            )
        except Exception:
            texto_quadro = f"Projeto de Referência: {id_proj}"

        self.escrever_seguro(ws, f"A{l_obs}", texto_quadro, Alignment(horizontal='left', vertical='center'))
        
        # Omitida a aplicação global de Font() nesta célula para que a CellRichText funcione sem ser sobreposta
        if isinstance(texto_quadro, str):
            try: ws[f"A{l_obs}"].font = Font(name='Arial Black', size=18, bold=True)
            except: pass
            
        wb.save(os.path.join(pasta, f"{re.sub(r'[\\/*?:\u0022<>|]', '', tit)}.xlsx"))

    def iniciar_processamento(self):
        self.btn_processar.configure(state="disabled")
        self.progress.grid(row=3, column=0, padx=20, pady=10); self.progress.start()
        threading.Thread(target=self.core_processamento, daemon=True).start()

    def core_processamento(self):
        try:
            df = pd.read_clipboard(sep='\t', header=None).fillna('')
            id_proj = "PROJETO"
            for v in df.values.flatten():
                if re.search(r'^[A-Z]{2,}\d+', str(v).strip().upper()): id_proj = str(v).strip().upper(); break
            
            MARCAS_PASTAS = {
                "ARO": "Amaro", "BGK": "BurguerKing", "CAM": "Camicado", "CEA": "CeA", 
                "CEN": "Centauro", "ELU": "Elubel", "FAR": "FarmaCopr", "IND": "Indian", 
                "ING": "Ingecon", "MCD": "McDonalds", "PER": "Pernambucanas", "REN": "Renner", 
                "TMS": "Tramontina", "TRA": "Tramontina", "ZAR": "Zara", "ZFR": "Zaffari", 
                "SEP": "Sephora", "PRO": "Prototipo"
            }
            nome_marca = "Outros"
            for sigla, nome in MARCAS_PASTAS.items():
                if sigla in id_proj.upper(): 
                    nome_marca = nome
                    break
            
            pasta = os.path.join(str(SERVIDOR), nome_marca, id_proj)
            if not os.path.exists(pasta): os.makedirs(pasta)
            molde = self.resource_path('planilha_molde.xlsx')

            def f_valido(f):
                c = self.limpar(f[1])
                return c.startswith(('11', '15')) and not any(x in self.limpar(f[14]) for x in ["92", "9172", "93"])

            consolidado = {}
            for _, row in df.iterrows():
                nv, cod = self.limpar(row[0]), self.limpar(row[1])
                if (cod.startswith(('11', '15')) or "PRENSADO" in str(row[3]).upper()) and nv.count('.') == 1:
                    if cod not in consolidado:
                        consolidado[cod] = {'pai': row, 'blocos': [], 'qtd_pai_total': 0}
                    consolidado[cod]['qtd_pai_total'] += float(self.converter_para_numero(row[5]) or 0)

            for cod_pai, info in consolidado.items():
                nv_pai = info['pai'][0]
                descendentes = df[df[0].str.startswith(nv_pai + ".")].copy()
                
                bloco_avulso = {'tipo': 'normal', 'itens': []}
                cursor = 0
                while cursor < len(descendentes):
                    row = descendentes.iloc[cursor]
                    nv_it, cod_it, desc_it = self.limpar(row[0]), self.limpar(row[1]), str(row[3]).upper()
                    
                    if "PRENSADO" in desc_it or cod_it == "1152032" or "PRLA" in str(row[2]).upper():
                        novo_bloco_prensado = {'tipo': 'prensado', 'prensado_info': row, 'itens': []}
                        q_prensado = float(self.converter_para_numero(row[5]) or 1)
                        cursor += 1
                        while cursor < len(descendentes):
                            sub_row = descendentes.iloc[cursor]
                            if not str(sub_row[0]).startswith(nv_it + "."): break
                            if f_valido(sub_row):
                                item_copy = sub_row.copy()
                                item_copy['q_unitaria_fatorada'] = float(self.converter_para_numero(sub_row[5]) or 0) * q_prensado
                                novo_bloco_prensado['itens'].append(item_copy)
                            cursor += 1
                        if novo_bloco_prensado['itens']: info['blocos'].append(novo_bloco_prensado)
                        continue
                    
                    if f_valido(row) and nv_it.count('.') == 2:
                        item_copy = row.copy()
                        item_copy['q_unitaria_fatorada'] = float(self.converter_para_numero(row[5]) or 0)
                        bloco_avulso['itens'].append(item_copy)
                    cursor += 1
                
                if bloco_avulso['itens']:
                    info['blocos'].insert(0, bloco_avulso)

            for info in consolidado.values():
                if info['blocos']: self.gerar_arquivo_excel(info['pai'], info['blocos'], id_proj, info['qtd_pai_total'], molde, pasta)

            self.after(0, self.sucesso_final, pasta)
        except Exception as e: self.after(0, self.erro_final, str(e))

    def sucesso_final(self, p): self.progress.stop(); self.progress.grid_forget(); self.btn_processar.configure(state="normal"); os.startfile(p); messagebox.showinfo("Ingecon", "Concluído!")
    def erro_final(self, m): self.progress.stop(); self.btn_processar.configure(state="normal"); messagebox.showerror("Erro", m)

if __name__ == "__main__": AppIngecon().mainloop()