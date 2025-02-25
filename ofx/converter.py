import ofxparse
import pandas as pd
import argparse


def convert_ofx_to_csv_excel(ofx_file):
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
                "BENEFICIÁRIO": account.payee,
                "TIPO": account.type,
            }
        )

    df = pd.DataFrame(transactions)
    df["ID"] = pd.to_numeric(df["ID"], errors="coerce")
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
    output_file = ofx_file.rsplit(".", 1)[0] + ".xlsx"

    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Arquivo convertido e salvo como: {output_file}")


def main():
    parser = argparse.ArgumentParser(
        description="Converter arquivo OFX para CSV ou Excel"
    )
    parser.add_argument("ofx_file", help="Caminho do arquivo OFX")
    args = parser.parse_args()
    convert_ofx_to_csv_excel(args.ofx_file)
