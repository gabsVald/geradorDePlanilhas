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
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox

# =============================
# CONFIG AUTO UPDATE
# =============================
VERSAO_ATUAL = "1.9.3"

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
        """ Busca o caminho do arquivo de forma ultra-robusta (pathlib) """
        # 1. Tenta o caminho interno do PyInstaller (Pasta Temp)
        if hasattr(sys, '_MEIPASS'):
            p_interno = Path(sys._MEIPASS) / relative_path
            if p_interno.exists(): return str(p_interno)

        # 2. Tenta a pasta onde o executável está rodando (Drive X:)
        p_exe = Path(sys.executable).parent / relative_path
        if p_exe.exists(): return str(p_exe)

        # 3. Tenta o diretório de trabalho atual (Local)
        p_local = Path(os.getcwd()) / relative_path
        if p_local.exists(): return str(p_local)

        return str(p_exe) # Retorna o caminho do EXE como fallback

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
        logos = {"ARO": "amaro", "BGK": "BurguerKing", "CAM": "Camicado", "CEA": "CeA", "CEN": "Centauro", "ELU": "Elubel", "FAR": "FarmaCopr", "IND": "Indian", "ING": "ingecon", "MCD": "mcdonalds", "PER": "pernambucanas", "REN": "renner", "TMS": "Tramontina", "TRA": "tramontinaPDV"}
        textos_marcas = {"ZAR": "ZARA", "ZFR": "ZAFFARI", "SEP": "SEPHORA", "PRO": "Protótipo"}
        
        id_up = str(id_projeto).upper()
        
        # Procura por logos de imagem
        for sigla, arq_nome in logos.items():
            if sigla in id_up:
                # Usa Path para normalizar barras do Windows/Rede
                path_logo = self.resource_path(f"logos/{arq_nome}.png")
                if Path(path_logo).exists():
                    img = OpenpyxlImage(path_logo)
                    img.width, img.height = 152, 42 
                    ws.row_dimensions[1].height = 33 
                    ws.add_image(img, 'A1')
                    return
        
        # Fallback para textos
        texto_exibicao = id_projeto
        for sigla, nome in textos_marcas.items():
            if sigla in id_up:
                texto_exibicao = nome
                break
        self.escrever_seguro(ws, 'A1', texto_exibicao, Alignment(horizontal='center', vertical='center'))
        ws['A1'].font = Font(size=22, bold=True)

    def separar_pela_coluna_D(self, row):
        s = self.limpar(row[3])
        if " - " in s:
            p = s.split(" - ", 1)
            return p[0].strip(), self.limpar_material_rigoroso(p[1].strip())
        return "-", s

    def limpar_material_rigoroso(self, texto):
        if not texto: return ""
        t = re.sub(r'\b(ORIG|ESS)\b', '', texto, flags=re.IGNORECASE).replace('=', '')
        t = re.sub(r'\s*\b\d+(?:[\.,]\d+)?\s*[xX].*$', '', t, flags=re.IGNORECASE)
        return re.sub(r'\s+', ' ', t).strip(' -')

    def ajustar_molde_elastico(self, ws, num_itens):
        padrao, l_rodape = 3, 9
        mesclagens, quadro = [], None
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= l_rodape:
                if m.min_row == 9: quadro = {'min_col': m.min_col, 'max_col': m.max_col, 'max_row': m.max_row}
                else: mesclagens.append({'min_row': m.min_row, 'max_row': m.max_row, 'min_col': m.min_col, 'max_col': m.max_col})
                ws.unmerge_cells(str(m))
        diff = max(0, num_itens - padrao)
        if diff > 0: ws.insert_rows(l_rodape, diff)
        for r in range(6, ws.max_row + 1):
            ws.row_dimensions[r].height = 25.5
            if 6 <= r < (6 + num_itens):
                ws.cell(row=r, column=4).value = "X"; ws.cell(row=r, column=7).value = "X"
                if r > 6:
                    for c in range(1, 14):
                        src, tgt = ws.cell(row=6, column=c), ws.cell(row=r, column=c)
                        if src.has_style: tgt._style = copy(src._style)
        n_inicio = 6 + num_itens
        if quadro:
            ws.merge_cells(start_row=n_inicio, start_column=quadro['min_col'], end_row=quadro['max_row'] + diff, end_column=quadro['max_col'])
            ws.row_dimensions[n_inicio].height = 409.5
        for m in mesclagens: ws.merge_cells(start_row=m['min_row'] + diff, start_column=m['min_col'], end_row=m['max_row'] + diff, end_column=m['max_col'])
        return n_inicio

    def gerar_arquivo_excel(self, pai, itens, id_proj, qtd_tot, molde, pasta):
        wb = load_workbook(molde); ws = wb.active
        l_obs = self.ajustar_molde_elastico(ws, len(itens))
        cod, acab, desc = self.limpar(pai[1]), self.limpar(pai[2]), self.limpar(pai[3])
        tit = f"{f'{cod}_{acab}' if acab else cod} - {desc}"
        
        self.tratar_cabecalho_a1(ws, id_proj)
        self.escrever_seguro(ws, 'B3', tit)
        self.escrever_seguro(ws, 'A3', qtd_tot)
        self.escrever_seguro(ws, 'M2', datetime.now().strftime('%d/%m/%Y'))

        for i, row in enumerate(itens):
            r = 6 + i; d_l, m_l = self.separar_pela_coluna_D(row)
            q_final_val = float(row.get('qtd_final', 0))
            q_unitaria = q_final_val / float(qtd_tot or 1)
            ws.cell(row=r, column=1).value = f"={str(q_unitaria).replace(',', '.')}*A3"
            
            for col, idx in [(2,15), (3,8), (5,16), (6,10), (8,12), (13,1)]:
                raw_val = row[idx] if idx not in [8, 10, 12] else (row[idx] or row[idx+1])
                v = self.converter_para_numero(raw_val)
                c = ws.cell(row=r, column=col)
                c.value = v
                if isinstance(v, (int, float)):
                    c.number_format = '0' if isinstance(v, int) else '0.00'
                else:
                    c.number_format = '@'
                
            ws.cell(row=r, column=9).value = m_l; ws.cell(row=r, column=12).value = d_l
            ws.cell(row=r, column=11).value = "SEC-LAM" if any(x in str(row[15])+str(row[16]) for x in ["-", "="]) else "SEC"
        
        self.escrever_seguro(ws, f"A{l_obs}", id_proj, Alignment(horizontal='left', vertical='top'))
        ws.cell(row=l_obs, column=1).font = Font(size=20, bold=True)
        wb.save(os.path.join(pasta, f"{re.sub(r'[\\/*?:\u0022<>|]', '', tit)}.xlsx"))

    def iniciar_processamento(self):
        self.btn_processar.configure(state="disabled")
        self.progress.grid(row=3, column=0, padx=20, pady=10); self.progress.start()
        threading.Thread(target=self.core_processamento, daemon=True).start()

    def core_processamento(self):
        try:
            try:
                df = pd.read_clipboard(sep='\t', header=None).fillna('')
            except:
                self.after(0, self.erro_final, "Área de transferência vazia!"); return

            if df.empty or df.shape[1] < 5:
                self.after(0, self.erro_final, "Tabela inválida!"); return

            id_proj = None
            for v in df.values.flatten():
                if re.search(r'^[A-Z]{2,}\d+', str(v).strip().upper()): 
                    id_proj = str(v).strip().upper(); break
            
            if not id_proj:
                self.after(0, self.erro_final, "Projeto não identificado!"); return

            MARCAS_PASTAS = {"ARO": "Amaro", "BGK": "BurguerKing", "CAM": "Camicado", "CEA": "CeA", "CEN": "Centauro", "ELU": "Elubel", "FAR": "FarmaCopr", "IND": "Indian", "ING": "Ingecon", "MCD": "McDonalds", "PER": "Pernambucanas", "REN": "Renner", "TMS": "Tramontina", "TRA": "Tramontina", "ZAR": "Zara", "ZFR": "Zaffari", "SEP": "Sephora", "PRO": "Prototipo"}
            nome_marca = "Outros"
            for sigla, nome in MARCAS_PASTAS.items():
                if sigla in id_proj: nome_marca = nome; break
            
            pasta = os.path.join(str(SERVIDOR), nome_marca, id_proj)
            if not os.path.exists(pasta): os.makedirs(pasta)

            molde = self.resource_path('planilha_molde.xlsx')

            def f_valido(f):
                c, d, m = self.limpar(f[1]), str(f[3]).upper(), self.limpar(f[14])
                if not c: return False
                if "PRENSADO" in d or c == "1152032": return True
                return not any(x in m for x in ["92", "9172", "93"]) and c.startswith(('11', '15'))

            consolidado = {}
            grupos = {}
            for _, row in df.iterrows():
                nv, cod = self.limpar(row[0]), self.limpar(row[1])
                acab_val = str(row[2]).strip().upper()
                
                if acab_val == "*" or "CORTE" in acab_val:
                    continue

                if cod.startswith(('11', '15')) or "PRENSADO" in str(row[3]).upper():
                    p = nv.split('.'); pref = f"{p[0]}.{p[1]}" if len(p) >= 2 else p[0]
                    if pref not in grupos: grupos[pref] = []
                    grupos[pref].append(row)

            for pref, linhas in grupos.items():
                linhas_ord = sorted(linhas, key=lambda x: [int(s) if s.isdigit() else s for s in str(x[0]).split('.')])
                pai_orig = linhas_ord[0]
                k_m = self.limpar(pai_orig[1])
                if not k_m: continue

                if k_m not in consolidado:
                    consolidado[k_m] = {'pai': pai_orig.copy(), 'itens': {}, 'qtd_pai': 0}
                
                q_val_pai = self.converter_para_numero(pai_orig[5])
                q_p_atual = float(q_val_pai if isinstance(q_val_pai, (int, float)) else 0)
                consolidado[k_m]['qtd_pai'] += q_p_atual
                
                itens_encontrados = False
                cursor = 1
                while cursor < len(linhas_ord):
                    it = linhas_ord[cursor]
                    nv_it, cod_it, desc_it = self.limpar(it[0]), self.limpar(it[1]), str(it[3]).upper()
                    if "PRENSADO" in desc_it or cod_it == "1152032":
                        k_p = cod_it
                        if k_p not in consolidado: consolidado[k_p] = {'pai': it.copy(), 'itens': {}, 'qtd_pai': 0}
                        q_sub_v = self.converter_para_numero(it[5])
                        q_sub_pai = float(q_sub_v if isinstance(q_sub_v, (int, float)) else 0)
                        consolidado[k_p]['qtd_pai'] += q_sub_pai
                        sub_nv = nv_it
                        cursor += 1
                        while cursor < len(linhas_ord) and self.limpar(linhas_ord[cursor][0]).startswith(sub_nv + "."):
                            prox = linhas_ord[cursor]
                            if f_valido(prox):
                                k_f = self.limpar(prox[1]); itens_encontrados = True
                                if k_f not in consolidado[k_p]['itens']:
                                    consolidado[k_p]['itens'][k_f] = prox.copy()
                                    consolidado[k_p]['itens'][k_f]['qtd_final'] = 0
                                q_it_v = self.converter_para_numero(prox[5])
                                q_it_f = float(q_it_v if isinstance(q_it_v, (int, float)) else 0)
                                consolidado[k_p]['itens'][k_f]['qtd_final'] += (q_it_f * q_sub_pai)
                            cursor += 1
                    else:
                        if f_valido(it):
                            k_f = self.limpar(it[1]); itens_encontrados = True
                            if k_f not in consolidado[k_m]['itens']:
                                consolidado[k_m]['itens'][k_f] = it.copy()
                                consolidado[k_m]['itens'][k_f]['qtd_final'] = 0
                            q_it_v = self.converter_para_numero(it[5])
                            q_it_f = float(q_it_v if isinstance(q_it_v, (int, float)) else 0)
                            consolidado[k_m]['itens'][k_f]['qtd_final'] += (q_it_f * q_p_atual)
                        cursor += 1
                
                if not itens_encontrados and f_valido(pai_orig):
                    k_f = self.limpar(pai_orig[1])
                    if k_f not in consolidado[k_m]['itens']:
                        consolidado[k_m]['itens'][k_f] = pai_orig.copy()
                        consolidado[k_m]['itens'][k_f]['qtd_final'] = 0
                    consolidado[k_m]['itens'][k_f]['qtd_final'] += q_p_atual

            for info in consolidado.values():
                lista_filhos = list(info['itens'].values())
                if lista_filhos:
                    self.gerar_arquivo_excel(info['pai'], lista_filhos, id_proj, info['qtd_pai'], molde, pasta)

            self.after(0, self.sucesso_final, pasta)
        except Exception:
            self.after(0, self.erro_final, "Erro no processamento.")

    def sucesso_final(self, p): self.progress.stop(); self.progress.grid_forget(); self.btn_processar.configure(state="normal"); os.startfile(p); messagebox.showinfo("Ingecon", "Concluído!")
    def erro_final(self, m): self.progress.stop(); self.btn_processar.configure(state="normal"); messagebox.showerror("Erro", m)

if __name__ == "__main__":
    AppIngecon().mainloop()