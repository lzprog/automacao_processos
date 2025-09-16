#  Automação de Processos

Repositório de scripts em **Python** desenvolvidos para automatizar tarefas internas de uma **empresa X**.  
Os programas conectam-se a bases de dados Firebird, leem planilhas do Excel e realizam atualizações ou análises de forma totalmente automatizada.  

##  Estrutura do Projeto

| Arquivo | Descrição |
|--------|-----------|
| **alterar_valores_array.py** | Lê planilhas Excel utilizando **pandas**, gera uma lista de tuplas no formato `(id, novo_preço)` e aplica automaticamente as alterações em todos os bancos de dados das lojas. |
| **bar_code.py** | Recebe uma lista de pares **(chave de referência, código de barras)** e atualiza diretamente os registros nas bases Firebird interligadas. |
| **excel_reader_v2.py** | Analisa planilhas de notas fiscais, identifica os horários de emissão, associa aos intervalos corretos e integra essas informações de volta na planilha para facilitar o controle. |
| **nf_manager_final.py** | (Em desenvolvimento ou integração futura) Gerenciamento e processamento de notas fiscais de forma centralizada. |

---

##  Tecnologias Utilizadas

- **Python 3.10+**
- **pandas** → leitura e manipulação de planilhas
- **fdb** → conexão com banco de dados Firebird
- **openpyxl** → integração e atualização de arquivos Excel
- **pathlib**
- **ftplib**
- **os**
- **++++**

---
