import tkinter as tk
from tkinter import filedialog, messagebox
import ofxparse
import pandas as pd
import os


def convert_ofx_to_excel(ofx_file, output_file):
    with open(ofx_file, "rb") as f:
        ofx = ofxparse.OfxParser.parse(f)

    transactions = []

    for account in ofx.account.statement.transactions:
        transactions.append(
            {
                "ID": account.id,
                "DATA": account.date.strftime("%Y-%m-%d"),
                "MEMORANDO": account.memo,
                "VALOR": account.amount,
                "TIPO": account.type,
            }
        )

    df = pd.DataFrame(transactions)
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce")
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")

    df.to_excel(output_file, index=False, engine="openpyxl")
    messagebox.showinfo("Sucesso", f"Arquivo convertido e salvo como: {output_file}")


def selecionar_arquivo():
    input_file = filedialog.askopenfilename(filetypes=[("Arquivos OFX", "*.ofx")])
    if input_file:
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Planilhas Excel", "*.xlsx")]
        )
        if output_file:
            convert_ofx_to_excel(input_file, output_file)


def main():
    root = tk.Tk()
    root.title("Conversor OFX para Excel")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(padx=10, pady=10)

    label = tk.Label(
        frame, text="Selecione um arquivo OFX para conversão e escolha onde salvar:"
    )
    label.pack()

    botao_selecionar = tk.Button(
        frame, text="Selecionar Arquivo", command=selecionar_arquivo
    )
    botao_selecionar.pack(pady=10)

    root.mainloop()
