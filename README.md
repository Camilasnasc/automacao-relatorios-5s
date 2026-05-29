# 📊 Sistema de Automação de Relatórios Empresariais com Python e Excel

> Projeto de Portfólio — Atividade Extra-Curricular  
> Curso: Sistemas de Informação | Período: 3° | Aluno(a): Camila Santos Nascimento

---

## 📌 Descrição

Este projeto consiste no desenvolvimento de um **MVP (Produto Mínimo Viável)** de automação empresarial utilizando Python. O sistema realiza a leitura automática de uma planilha Excel alimentada continuamente, faz o download das fotos vinculadas a cada registro, converte as imagens para o formato JPG quando necessário e gera um relatório consolidado com indicadores de desempenho.

O projeto foi desenvolvido como atividade extra-curricular com foco em automação de processos, análise de dados e Business Intelligence, aplicado a um cenário empresarial real de gestão **5S**.

---

## ✅ Funcionalidades

- 📥 Leitura automática de planilha Excel (`.xlsx` / `.xlsm`)
- 🔗 Download automático de fotos a partir de links do Google Drive
- 🖼️ Detecção do formato real das imagens (por bytes, não pela extensão)
- 🔄 Conversão automática para JPG quando necessário (PNG, WEBP, BMP, GIF)
- 📁 Organização das fotos em pastas separadas:
  - `Imagens 5S` — fotos de **antes** da atividade
  - `Imagens 5S realizadas` — fotos de **depois** da atividade
- ♻️ Controle de duplicatas — fotos já existentes são ignoradas automaticamente
- 📊 Geração automática de relatório Excel consolidado com:
  - Tabela de registros com status por cor
  - Aba separada para fotos de antes e depois
  - Indicadores automáticos (total, erros, % de cobertura)

---

## 🛠️ Tecnologias Utilizadas

| Biblioteca     | Finalidade                                        |
| -------------- | ------------------------------------------------- |
| `openpyxl`     | Leitura e escrita de arquivos Excel (.xlsx/.xlsm) |
| `requests`     | Download de imagens via HTTP (Google Drive)       |
| `Pillow (PIL)` | Detecção e conversão de formato de imagens        |
| `pathlib`      | Manipulação de caminhos e pastas                  |
| `re`           | Extração de IDs de URLs do Google Drive           |
| `datetime`     | Registro de data e hora dos downloads             |

---

## 💻 Como Instalar e Executar

### 1. Pré-requisitos

- Python 3.10 ou superior instalado
- Pip disponível

### 2. Instalar as dependências

```bash
pip install openpyxl requests Pillow
```

### 3. Configurar o script

Abra o arquivo `salvar_fotos_5s.py` e ajuste as configurações no início do arquivo:

```python
ARQUIVO_XLSX   = "GESTÃO 5S 4.0 rec.xlsm"   # nome da sua planilha
PASTA_ANTES    = Path("Imagens 5S")           # pasta para fotos de antes
PASTA_DEPOIS   = Path("Imagens 5S realizadas") # pasta para fotos de depois
RELATORIO_XLSX = "Relatorio_Fotos_5S.xlsx"    # nome do relatório gerado
```

### 4. Executar

```bash
python salvar_fotos_5s.py
```

---

## 📁 Estrutura do Projeto

```
📦 PROJETO-5S
 ┣ 📄 salvar_fotos_5s.py         # Script principal
 ┣ 📄 GESTÃO 5S 4.0 rec.xlsm    # Planilha de entrada
 ┣ 📄 Relatorio_Fotos_5S.xlsx    # Relatório gerado automaticamente
 ┣ 📁 Imagens 5S                 # Fotos de antes (gerada automaticamente)
 ┃  ┣ 🖼️ 1.jpg
 ┃  ┣ 🖼️ 2.jpg
 ┃  ┗ ...
 ┗ 📁 Imagens 5S realizadas      # Fotos de depois (gerada automaticamente)
    ┣ 🖼️ 1.jpg
    ┣ 🖼️ 2.jpg
    ┗ ...
```

---

## 📈 Relatório Gerado

O arquivo `Relatorio_Fotos_5S.xlsx` é gerado automaticamente ao final de cada execução e contém **duas abas**:

**Aba "Fotos Antes"** e **Aba "Fotos Depois"**, cada uma com:

| Coluna                | Descrição                       |
| --------------------- | ------------------------------- |
| ID                    | Identificador do registro       |
| Nome                  | Nome do responsável             |
| Área                  | Área onde foi realizado o 5S    |
| Status                | Resultado do processamento      |
| Formato Original      | Formato detectado da imagem     |
| Data/Hora do Download | Registro do momento do download |

### Legenda de cores do relatório

| Cor         | Status                   |
| ----------- | ------------------------ |
| 🟢 Verde    | Foto salva com sucesso   |
| 🟡 Amarelo  | Foto convertida para JPG |
| ⚫ Cinza    | Foto já existia, pulada  |
| 🟠 Laranja  | Registro sem foto        |
| 🔴 Vermelho | Erro no processamento    |

---

## 🗂️ Sobre a Planilha de Exemplo

A planilha disponibilizada neste repositório (`GESTÃO 5S 4.0 rec.xlsm`) é **apenas um exemplo para fins de teste e demonstração**.

- Os links de fotos **não contêm imagens reais** de nenhuma empresa ou pessoa, por questões de privacidade e prevenção ao uso não autorizado de imagens
- A planilha possui apenas **2 registros de foto de antes** e **2 registros de foto de depois**, todas retiradas aleatoriamente do Google exclusivamente para fins de teste do script
- Em um ambiente real, a planilha seria alimentada continuamente com os dados e fotos da empresa

---

## ⚠️ Observações

- As fotos no Google Drive precisam estar com permissão **"Qualquer pessoa com o link"** para que o download funcione sem autenticação
- O script pode ser executado múltiplas vezes sem risco — registros já processados são ignorados automaticamente
- Imagens muito grandes são processadas normalmente (limite de tamanho desativado)

---

## 🎯 Objetivos SMART Atendidos

- ✅ Leitura automática de planilha alimentada continuamente
- ✅ Tratamento e organização de dados sem manipulação manual
- ✅ Geração automática de relatório consolidado
- ✅ Criação de indicadores básicos de desempenho
- ✅ Exportação dos resultados em planilha formatada
- ✅ Redução significativa do tempo gasto com processos manuais repetitivos

---

## 📚 Referências

- [Documentação oficial do Python](https://docs.python.org/3/)
- [Documentação do OpenPyXL](https://openpyxl.readthedocs.io/)
- [Documentação do Requests](https://requests.readthedocs.io/)
- [Documentação do Pillow](https://pillow.readthedocs.io/)
