import pandas as pd
from win32com import client as win32


def enviar_email(parceiro, produto, email_cliente, numero_pedido):
    outlook = win32.Dispatch("outlook.application")
    email = outlook.CreateItem(0)
    email.To = email_cliente
    email.Subject = "Relatório de Vendas"
    email.Body = (
        f"Bom dia {parceiro}, Tudo bem?, \r\n"
        f"Produto: {produto} \r\n"
        f"Número do pedido: {numero_pedido}\r\n"
        f"Att, \r\n"
        f"William Ferreira \r\n"
    )
    email.Display()
    email.Send()


df = pd.read_excel(
    r"caminho do arquivo",
    sheet_name="EXERCICIO 9",
)
dados = df[
    [
        "Validade",
        "Parceiro",
        "Produto",
        "Email",
        "Número do Pedido",
    ]
]

for i, r in dados.iterrows():
    if r["Validade"] == "Válido":
        validade = r["Validade"]
        parceiro = r["Parceiro"]
        produto = r["Produto"]
        email_cliente = r["Email"]
        numero_pedido = (r["Número do Pedido"],)
        enviar_email(parceiro, produto, email_cliente, numero_pedido)
