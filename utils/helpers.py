"""
================================================================================
utils/helpers.py — Funções Utilitárias Globais
================================================================================
Funções puras reutilizadas em todo o sistema. Não têm dependências entre si
nem dependências de outros módulos do projeto (exceto os da stdlib).

Funções disponíveis:
  resource_path()          — Resolve caminhos dentro do executável PyInstaller
  limpar()                 — Normaliza valores vindos do Excel/Clipboard
  converter_para_numero()  — Converte string para inteiro arredondado
  limpar_material_rigoroso() — Remove sujeiras do nome do material
================================================================================
"""

import os
import sys
import re
from pathlib import Path


def resource_path(relative_path):
    """
    Resolve o caminho de um arquivo de recurso para funcionar tanto em
    modo de desenvolvimento quanto em executável compilado com PyInstaller.

    Quando empacotado como .exe, o PyInstaller extrai os recursos para um
    diretório temporário (_MEIPASS). Esta função verifica esse diretório
    primeiro e cai de volta para o diretório de trabalho atual.

    Parâmetros:
        relative_path (str): Caminho relativo ao recurso (ex: "logos/ingecon.png").

    Retorna:
        str: Caminho absoluto resolvido para o recurso.
    """
    if hasattr(sys, '_MEIPASS'):
        # Modo executável: recursos estão em um diretório temporário
        p_interno = Path(sys._MEIPASS) / relative_path
        if p_interno.exists():
            return str(p_interno)
    # Modo desenvolvimento: recursos estão no diretório de trabalho
    return str(Path(os.getcwd()) / relative_path)


def limpar(val):
    """
    Normaliza um valor vindo do Excel, Clipboard ou ODS para string limpa.

    Tratamentos aplicados:
      - None → string vazia
      - Booleanos (True/False do openpyxl) → convertidos para string antes de processar
      - ".0" no final → removido (ex: "1147128.0" → "1147128")
      - "nan", "none", "null" → string vazia

    Parâmetros:
        val: Qualquer valor (None, bool, int, float, str).

    Retorna:
        str: String limpa, ou "" se o valor for vazio/inválido.

    Exemplos:
        limpar(None)         → ""
        limpar("1147128.0")  → "1147128"
        limpar(True)         → "True"   (booleano tratado como string)
        limpar("nan")        → ""
    """
    if val is None:
        return ""

    # Booleanos do openpyxl precisam ser convertidos antes do str().strip()
    # para evitar que "True" seja confundido com um valor de célula real
    if isinstance(val, bool):
        val = str(val)

    v = str(val).strip()

    # Remove sufixo ".0" que o pandas adiciona ao ler números inteiros como float
    if v.endswith('.0'):
        v = v[:-2]

    return v if v.lower() not in ['nan', 'none', 'null', ''] else ""


def converter_para_numero(valor, retornar_marcador=False, remover_unidades=True):
    """
    Converte uma string para inteiro arredondado, com tratamento de casos especiais.

    Parâmetros:
        valor: Valor a converter (string, int, float ou None).
        retornar_marcador (bool): Se True, retorna a string original quando não é número
                                  (ex: "-" ou "="). Se False, retorna None nesses casos.
        remover_unidades (bool): Se True, remove sufixos de unidade antes de converter
                                 (pç, pc, un, und, unid, unidade). Deve ser False para
                                 códigos de peça alfanuméricos como "PÇ 1".

    Retorna:
        int | str | None:
          - int: valor convertido e arredondado
          - str: a string original se retornar_marcador=True e não for número
          - None: se não puder converter e retornar_marcador=False

    Exemplos:
        converter_para_numero("1883")        → 1883
        converter_para_numero("1883.7")      → 1884  (arredondamento)
        converter_para_numero("2 UN")        → 2     (remove unidade)
        converter_para_numero("-")           → None
        converter_para_numero("-", True)     → "-"   (retorna marcador)
        converter_para_numero("PÇ 1", remover_unidades=False) → "PÇ 1"
    """
    limpo = limpar(valor)
    if not limpo:
        return "" if retornar_marcador else None

    # Remove sufixos de unidade comuns do PDM (pode ser desativado para códigos)
    if remover_unidades:
        limpo = re.sub(r'(?i)\s*(pç|pc|un|und|unid|unidade)\.*', '', limpo)

    # Símbolos de fita de borda são preservados como string, nunca convertidos
    if limpo in ["-", "="]:
        return limpo if retornar_marcador else None

    try:
        v_aj = limpo.replace(',', '.')  # Normaliza separador decimal
        val_float = float(v_aj)
        # Arredondamento "meio acima": 1883.5 → 1884, -1.5 → -2
        return int(val_float + 0.5) if val_float >= 0 else int(val_float - 0.5)
    except Exception:
        # Não é um número: retorna a string ou None conforme o parâmetro
        return limpo if retornar_marcador else None


def limpar_material_rigoroso(texto):
    """
    Limpa o nome do material para exibição na planilha de corte.

    Remove sujeiras que o PDM insere no nome:
      1. Marcações internas: ORIG, ESS
      2. Símbolo de igualdade (=)
      3. Dimensões completas: ex. "254.7X159.2X18", "440X-260X15"
      4. Dimensões truncadas: ex. "254.7X"
      5. Sufixos de espessura: ex. "18MM", "18.5MM" (mas preserva "(3100)" em parênteses)
      6. Parênteses vazios: "()" ou "( )"
      7. Hifens residuais no início, meio ou fim

    Parâmetros:
        texto (str): Nome do material vindo do PDM (pode conter sujeiras).

    Retorna:
        str: Nome do material limpo, em maiúsculas.

    Exemplos:
        limpar_material_rigoroso("MDF CRU 254.7X159.2X18 ORIG")  → "MDF CRU"
        limpar_material_rigoroso("KRION CRYSTAL WHITE 3100MM")    → "KRION CRYSTAL WHITE"
        limpar_material_rigoroso("= MDF BP 2F BCO DIA CRI ESS")  → "MDF BP 2F BCO DIA CRI"
    """
    if not texto:
        return ""

    t = str(texto).upper()

    # 1. Remove marcações internas do PDM (ORIG = versão original; ESS = essencial)
    t = re.sub(r'\b(ORIG|ESS)\b', '', t)
    t = t.replace('=', '')

    # 2. Remove medidas completas (CxL ou CxLxE) com ponto ou vírgula como decimal
    #    Também trata hifens errôneos após "X" (ex: "440X-260X15")
    patrao_medida = r'\b\d+(?:[.,]\d+)?\s*[xX]\s*[-]?\s*\d+(?:[.,]\d+)?(?:\s*[xX]\s*[-]?\s*\d+(?:[.,]\d+)?)?\b'
    t = re.sub(patrao_medida, '', t)

    # 3. Remove fragmentos de medida truncados com decimal (ex: "254.7X" isolado)
    t = re.sub(r'\b\d+(?:[.,]\d+)?\s*[xX]\b', '', t)

    # 4. Preserva medidas em parênteses, removendo apenas o sufixo "MM"
    #    Ex: "(3100MM)" → "(3100)" — comprimento de chapa em parênteses
    t = re.sub(r'\(\s*(\d+(?:[.,]\d+)?)\s*MM\s*\)', r'(\1)', t)

    # 5. Remove sufixos de espessura/comprimento órfãos fora de parênteses
    t = re.sub(r'\b\d+(?:[.,]\d+)?\s*MM\b', '', t)

    # 6. Remove parênteses vazios que sobraram após as remoções anteriores
    t = re.sub(r'\(\s*\)', '', t)

    # 7. Remove hifens seguidos de número no final (ex: "MDF CRU - 15")
    t = re.sub(r'-\s*\d+(?:[.,]\d+)?\s*$', '', t)

    # 8. Remove hifens soltos no início ou fim da string
    t = re.sub(r'^\s*-\s*', '', t)
    t = re.sub(r'\s*-\s*$', '', t)

    # 9. Remove hifens isolados no meio (sobram quando texto dos dois lados foi removido)
    t = re.sub(r'\s+-\s+', ' ', t)

    # 10. Normaliza espaços múltiplos
    t = re.sub(r'\s+', ' ', t)

    return t.strip().strip('-').strip()