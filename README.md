# 📧 Enviando Certificados por E-mail

Este projeto automatiza a **geração em massa de certificados personalizados** e o **envio automático por e-mail** para uma lista de alunos.

---

## ✅ Funcionalidades

- 📄 Geração de certificados com base em um modelo `.docx`
- 📊 Uso de planilha Excel com dados dos alunos
- ✉️ Envio automático dos certificados gerados por e-mail
- ⚙️ Possibilidade de customização das informações do certificado e corpo do e-mail

---

## 🛠️ Requisitos

- Python 3.8 ou superior
- Instalar as bibliotecas necessárias:

 --- Bibliotecas
pip install pandas openpyxl python-docx smtplib

## 📁 Estrutura dos Arquivos

```bash
Enviando_Email/
├── 3.Gerar Certificado_modificar_+informacoes_envio.py  # Script principal
├── Alunos.xlsx                                          # Dados dos alunos
├── Certificado1.docx                                    # Modelo de certificado
├── Gerar Certificado_Teste.py                           # Teste de geração
├── Gerar Certificado_modificar_nome_Teste.py            # Teste com modificação de nome
└── README.md                                             # Este arquivo
```

🧾 Passo a Passo para Usar
1. Preencha a planilha Alunos.xlsx
A planilha deve conter colunas como:

- Nome
- Email
- Outros campos que você queira incluir no certificado (ex: Curso, Data, etc)

2. Personalize o modelo Certificado1.docx
Use campos como {Nome}, {Curso}, {Data} etc. Esses serão substituídos pelos dados da planilha.

3. Configure o e-mail de envio no script
No arquivo 3.Gerar Certificado_modificar_+informacoes_envio.py, configure:

- email_remetente = "seu_email@gmail.com"
- senha = "sua_senha_do_app"  # Use senha de app para Gmail
- ⚠️ Use contas com autenticação de dois fatores e senhas de app para maior segurança.

4. Execute o script

- python 3.Gerar Certificado_modificar_+informacoes_envio.py
- Certificados serão gerados automaticamente e enviados por e-mail.

✏️ Personalização
- Você pode modificar o corpo do e-mail e o layout do certificado diretamente no código e no modelo .docx.

❗ Importante
- Verifique a permissão de envio de e-mails do seu provedor (ex: Gmail pode bloquear se detectar muitos envios).

- Sempre revise os dados da planilha antes de executar o script.

Feito com ❤️ para facilitar a sua automação!
