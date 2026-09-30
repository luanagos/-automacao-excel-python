import smtplib
import ssl
import os
from email.message import EmailMessage
from pathlib import Path

REMETENTE = os.environ.get("EMAIL_REMETENTE")
SENHA = os.environ.get("EMAIL_SENHA")
DESTINATARIO = os.environ.get("EMAIL_DESTINATARIO")
ARQUIVO = "relatorio_2026-09-28.xlsx"

msg = EmailMessage()
msg["Subject"] = "Relatório de vendas"
msg["From"] = REMETENTE
msg["To"] = DESTINATARIO
msg.set_content("Olá!\n\nSegue em anexo o relatório de vendas.\n\nAtt,\nLuana")

caminho = Path(ARQUIVO)
with open(caminho, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=caminho.name,
    )

contexto = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
    servidor.login(REMETENTE, SENHA)
    servidor.send_message(msg)

print("E-mail enviado!")