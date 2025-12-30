# 🏥 Sistema de Análise de Relatórios Hospitalares

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![Status](https://img.shields.io/badge/status-em%20desenvolvimento-yellow)
![License](https://img.shields.io/badge/license-interno-lightgrey)

Sistema desktop desenvolvido em Python para processamento, análise e geração de relatórios hospitalares de forma automatizada, com foco em procedimentos cirúrgicos.

## 📋 Índice
- [Visão Geral](#-visão-geral)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Como Usar](#-como-usar)
- [Funcionalidades](#-funcionalidades)
- [Configuração](#-configuração)
- [Desenvolvimento](#-desenvolvimento)
- [Build e Deploy](#-build-e-deploy)
- [Troubleshooting](#-troubleshooting)
- [Contribuição](#-contribuição)

## 🎯 Visão Geral

Este sistema foi desenvolvido para:
- **Automatizar** a análise de relatórios hospitalares
- **Simplificar** relatórios complexos para visualização rápida
- **Processar** dados de procedimentos cirúrgicos (NAC)
- **Gerar** relatórios organizados em Excel
- **Fornecer** interface gráfica amigável para usuários não técnicos

## 📁 Estrutura do Projeto

📦 projeto-analise-hospitalar
├── 📂 pycache/ # Caches do Python (NÃO versionar)
├── 📂 venv/ # Ambiente virtual Python
├── 📂 build/ # Arquivos temporários do PyInstaller
├── 📂 dist/ # Executável gerado
├── 📂 Prestador/ # Módulo de gestão de prestadores
├── 📂 relatorios_simplificados/ # Pasta de saída dos relatórios
├── 📂 separaRelatorio/ # Módulo de separação de relatórios
│
├── 📄 .env # Variáveis de ambiente (NÃO versionar)
├── 📄 .gitignore # Configuração do Git
├── 📄 analise.py # Lógica principal de análise
├── 📄 analise.spec # Configuração do PyInstaller
├── 📄 db.xlsx # Banco de dados em Excel
├── 📄 logo.ico # Ícone da aplicação
├── 📄 main.py # Ponto de entrada principal
├── 📄 nacCirurgico.py # Análise de procedimentos cirúrgicos
├── 📄 procedimentos.py # Gestão de procedimentos médicos
├── 📄 readme.md # Este arquivo
└── 📄 requirements.txt # Dependências do projeto


## ⚙️ Pré-requisitos

Antes de começar, você precisa ter instalado:

- **Python 3.8 ou superior**
- **pip** (gerenciador de pacotes do Python)
- **Git** (para controle de versão)
- **Ambiente Windows**

## 🔧 Instalação

### 1. Clonar o Repositório
```bash
git clone [URL_DO_SEU_REPOSITORIO]
cd [NOME_DO_PROJETO]

# Windows
python -m venv .venv
.venv\Scripts\Activate.ps1

pip install -r requirements.txt

## 📦 Build e Deploy
# Criar Executável Windows

pyinstaller --clean --noconsole --onefile .\analise.py