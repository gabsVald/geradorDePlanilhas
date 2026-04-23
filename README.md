# Ingecon - Gerador de Planilhas

O **Ingecon - Gerador de Planilhas** é uma solução de automação industrial desenvolvida em Python para otimizar o fluxo de trabalho de engenharia e produção. O sistema processa dados brutos (via clipboard), aplica regras de filtragem de materiais e gera planilhas Excel (.xlsm) padronizadas, além de gerenciar a migração segura de planos de corte entre diretórios de rede.

## 🚀 Principais Funcionalidades

* **Processamento via Clipboard:** Captura e trata dados copiados de listas de materiais de forma instantânea.
* **Migração Segura de Arquivos:** Gerencia a movimentação de arquivos entre pastas antigas e a nova estrutura de 2026, com verificações de duplicidade.
* **Geração de Excel Inteligente:** * Uso de `planilha_molde.xlsm` para manter macros e formatações.
    * Ajuste elástico de linhas de acordo com o volume de dados.
    * Inserção dinâmica de logotipos baseada no código do projeto.
* **Tratamento de Materiais:** Identificação automática de itens "Prensados", "Especiais" e mapeamento de códigos industriais.
* **Interface Gráfica (GUI):** Interface intuitiva construída com `customtkinter`, incluindo logs de depuração em tempo real.

## 📂 Estrutura do Projeto

* `main.py`: Inicialização do aplicativo e loop principal.
* `core/`: 
    * `excel.py`: Manipulação avançada de arquivos Excel (Openpyxl).
    * `processador.py`: Lógica de filtragem, processamento de dados e regras de negócio.
    * `migracao.py`: Utilitários para gestão de arquivos em rede.
* `ui/`: 
    * `interface.py`: Design e comportamento da janela principal.
* `utils/`: 
    * `config.py`: Gestor de configurações dos arquivos JSON.
    * `updater.py`: Sistema de verificação de atualizações.
* `regras.json` & `mapeamento.json`: Configurações de diretórios, filtros e nomes de materiais.

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3.x
* **Bibliotecas:**
    * `pandas`: Processamento e limpeza de dados.
    * `openpyxl`: Edição de planilhas Excel.
    * `customtkinter`: Interface visual moderna.
    * `pathlib`: Manipulação segura de caminhos de arquivos.

## ⚙️ Como Utilizar

1.  **Configuração Inicial:** Certifique-se de que os caminhos de rede no arquivo `regras.json` estão apontando para os servidores corretos.
2.  **Execução:** Execute o arquivo `main.py`.
3.  **Fluxo de Trabalho:**
    * Copie a tabela de dados desejada.
    * Clique no botão de processamento na interface.
    * O sistema irá validar os dados, gerar a planilha no destino correto e realizar a migração se necessário.

## ⚠️ Requisitos de Segurança

* O sistema não permite sobrescrever arquivos existentes sem validação.
* Operações críticas (leitura/escrita) utilizam backups em memória antes da execução final para evitar corrupção de dados.

---
*Documentação técnica estruturada para a equipe de Engenharia Ingecon.*
