# SaleMachine - Sistema de Mala Direta Automatizada

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-orange)

Sistema completo de mala direta desenvolvido em Python com interface gráfica intuitiva para envio em massa de emails personalizados via Outlook.

## 🚀 Funcionalidades Principais

### ✉️ **Envio de Emails em Massa**
- Integração nativa com Microsoft Outlook
- Envio personalizado usando dados de arquivos CSV
- Delay configurável entre envios para evitar bloqueios
- Suporte a HTML no corpo dos emails

### 📊 **Gerenciamento de Contatos**
- **Importação de CSV**: Suporte a múltiplos formatos e encodings
- **Edição Interna**: Edite nomes e emails diretamente na interface
- **Visualização de Dados**: Table interativa para análise dos contatos
- **Adição/Remoção**: Gerencie contatos sem editar arquivos externos

### 📎 **Sistema de Anexos Avançado**
- **Anexos Personalizados por Contato**: Arquivos específicos para cada pessoa
- **Anexos em Lote**: Mesmo arquivo para múltiplos contatos
- **Suporte a PDF e DOCX**: Formatos mais comuns para documentos
- **Gerenciamento Visual**: Interface intuitiva para associação de arquivos

### 🎯 **Personalização Inteligente**
- **Variáveis Dinâmicas**:
  - `{primeiro_nome}` - Nome do destinatário
  - `{nome_completo}` - Nome completo
  - `{email}` - Endereço de email
- **Assunto Personalizável**
- **Template HTML** para corpo dos emails

## 🛠️ Tecnologias Utilizadas

- **Python 3.6+**
- **Pandas** - Manipulação de dados CSV
- **Tkinter** - Interface gráfica moderna
- **pywin32** - Integração com Outlook
- **Threading** - Processamento não-bloqueante

## 📦 Instalação

```bash
# Clone o repositório
git clone https://github.com/MauricioSaleMachine/MalaDiretaTK

# Entre no diretório
cd salemachine-mala-direta

# Instale as dependências
pip install pandas pywin32

