# Judicial Certificate Automation

Python automation tool that generates judicial certificates from Brazilian courts.

This project automates the process of accessing court websites, filling identification data (CPF/CNPJ), solving captchas, and downloading the resulting certificate in PDF format.
---
## 🇺🇸 English
### Overview

This project automates the issuance of judicial certificates from Brazilian court systems.
It simulates user interaction with court websites, fills required fields, solves captchas, and downloads the resulting certificate in PDF format.

The tool was designed to automate repetitive legal tasks and reduce manual work when retrieving public judicial certificates.
---
### Features

* Browser automation using Playwright
* Automatic captcha solving
* OCR recognition for mathematical captchas
* Automatic PDF certificate download
* Simple graphical interface
* Support for multiple Brazilian courts
---
### Technologies

* Python
* Playwright
* Tkinter
* Tesseract OCR
* SpeechRecognition
---
### Installation

Clone the repository:
```
git clone https://github.com/vitordimario/judicial-certificate-automation.git
```
Install dependencies:
```
pip install -r requirements.txt
playwright install
```
Make sure **Tesseract OCR** is installed on your system.
---
### Running

Run the application with:
```
python certidao.py
```
---
### Purpose

This project demonstrates automation of repetitive legal tasks and document retrieval from public court systems.
---

## 🇧🇷 Português

### Visão geral

Este projeto automatiza a emissão de certidões judiciais em tribunais brasileiros.

A aplicação simula a navegação de um usuário nos sites dos tribunais, preenche automaticamente os dados necessários (CPF ou CNPJ), resolve captchas e realiza o download da certidão em PDF.

A ferramenta foi desenvolvida para automatizar tarefas jurídicas repetitivas e reduzir o tempo gasto na obtenção manual de certidões públicas.

---

### Funcionalidades

* Automação de navegação utilizando Playwright
* Resolução automática de captcha
* OCR para captchas matemáticos
* Download automático da certidão em PDF
* Interface gráfica simples
* Suporte para múltiplos tribunais brasileiros

---

### Tecnologias

* Python
* Playwright
* Tkinter
* Tesseract OCR
* SpeechRecognition

---

### Instalação

Clone o repositório:

```
git clone https://github.com/vitordimario/judicial-certificate-automation.git
```

Instale as dependências:

```
pip install -r requirements.txt
playwright install
```

Certifique-se de ter o **Tesseract OCR** instalado no sistema.

---

### Execução

Execute o programa com:

```
python certidao.py
```

---

### Objetivo

Este projeto demonstra a automação de tarefas jurídicas repetitivas e a obtenção automática de documentos públicos em sistemas judiciais.

---

## Author

Vitor Di Mario

