import os
import pdfplumber
import re
import pandas as pd
import time
import math
import shutil
import sys
from datetime import datetime

def extract_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text


def main():
    print("Bem vindo ao planilhamento de boletos")

    if not os.path.isdir(f"./Boletos"):
        os.makedirs(f"./Boletos")
        print("== Coloque os boletos na pasta Boletos que foi criada e inicie o programa de novo ==")
        input("Aperte ENTER para fechar essa janela")
        sys.exit()

    if not os.path.isdir(f"./Erros"):
        os.makedirs(f"./Erros")

    boletos = os.listdir(f"./Boletos")
    errors = []
    if not boletos:
        print("Coloque todos os PDFs de boletos na plata Boletos!")
        print("== Pode fechar essa aba ==")
        sys.exit()

    output = pd.DataFrame(columns=[
        "CLIENTE", 
        "UC", 
        "LEITURA", 
        "CONSUMO COMPENSADO (kWh)", 
        "CONSUMO NÃO COMPENSADO (kWh)", 
        "GERAÇÃO ABATIDA (kWh)", 
        "VALOR CONSUMO FATURA SCEE", 
        "MÊS REFERÊNCIA", 
        "CUSTO FIO B", 
        "COSUMO SCEE (R$)", 
        "INJEÇÃO SCEE (R$)", 
        "BENEFÍCIO BRUTO SCEE", 
        "BANDEIRA TARIFÁRIA", 
        "MULTA", 
        "JUROS", 
        "VALOR FATURA", 
        "ICMS", 
        "PIS", 
        "COFINS", 
        "ILUMINAÇÃO PÚBLICA",
        "CONSUMO ENERGIA DE GERAÇÃO",
        "VENCIMENTO",
        "SALDO (kWh)",
        ])

    for boleto in boletos:
        boleto = fr".\Boletos\{boleto}"
        #tables = tabula.read_pdf(boleto, pages='all', multiple_tables=True)
        with pdfplumber.open(boleto) as pdf:
            try:
                coordenadas_usadas = set()
                print(f"=> Lendo o boleto {boleto}...")
                text = pdf.pages[0].extract_text()

                uc = text.split("Chave de Acesso em:\n")[-1].split("\n")[0]
                cliente = text.split("\nCNPJ/CPF")[0].split("\n")[-1]
                inicio_leitura = text.split("Tensão Nominal Disp: ")[-1].split("PERDAS DE TRANSFORMAÇÃO / RAMAL:")[0].split("\n")[3].split(" ")[0]
                fim_leitura = text.split("Tensão Nominal Disp: ")[-1].split("PERDAS DE TRANSFORMAÇÃO / RAMAL:")[0].split("\n")[3].split(" ")[1]
                leitura = f"{inicio_leitura} - {fim_leitura}"

                consumo_compensado = 0

                if "CONSUMO SCEE" in text:
                    consumo_compensado = text.split("CONSUMO SCEE")[-1].strip().split(" ")[1].replace(",", ".")
                
                try:
                    regex = r"(?i)consumo não compensado\s*(?:kwh\s*)?([\d,.]+)\b"
                    consumo_nao_compensado = re.search(regex, text).group(1).replace(",", ".")
                except Exception:
                    consumo_nao_compensado = 0

                rows = text.strip().split("\n")
                geracao_abatida = 0
                rows_geracao_abatida = []

                words_pdfplumber = pdf.pages[0].extract_words()
                coordenadas_usadas.clear()
                for row in rows:
                    if "INJEÇÃO SCEE" in row:
                        if "UC" in row:
                            row = row.split("UC")[-1]

                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        if x0 >= 260 and x1 <= 300:
                                            coordenadas_usadas.add(pos)
                                            rows_geracao_abatida.append(float(word.replace(".", "").replace(",", ".")))
                                            break

                            except Exception as e:
                                pass
                        geracao_abatida = math.floor(sum(rows_geracao_abatida) * 100) / 100

                valor_consumo_fatura_scee = 0

                if "CONSUMO SCEE" in text:
                    valor_consumo_fatura_scee = text.split("CONSUMO SCEE")[-1].strip().split(" ")[3].replace(".", "").replace(",", ".")

                try:
                    regex = r"\b([A-Za-z]{3})\/(\d{4})\b"
                    mes_referencia = "/".join(re.findall(regex, text)[0])
                except Exception:
                    mes_referencia = "**/****"

                rows = text.strip().split("\n")
                custo_fio_b = 0
                rows_custo_fio_b = []

                words_pdfplumber = pdf.pages[0].extract_words()
                coordenadas_usadas.clear()
                for row in rows:
                    if "PARC INJET" in row:
                        if "UC" in row:
                            row = row.split("UC")[-1]

                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        if x0 >= 360 and x1 <= 410:
                                            coordenadas_usadas.add(pos)
                                            rows_custo_fio_b.append(float(word.replace(".", "").replace(",", ".")))
                                            break

                            except Exception as e:
                                pass
                        custo_fio_b = math.floor(sum(rows_custo_fio_b) * 100) / 100
                    

                consumo_scee = 0
                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "CONSUMO SCEE" in row:
                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue

                                        coordenadas_usadas.add(pos)
                                        if x0 >= 255 and x1 <= 295:
                                            consumo_scee = word.replace(".", "").replace(",", ".")
                                            break

                            except Exception as e:
                                pass

                rows = text.strip().split("\n")
                injecao_scee = 0
                rows_injecao_scee = []

                words_pdfplumber = pdf.pages[0].extract_words()
                coordenadas_usadas.clear()
                for row in rows:
                    if "INJEÇÃO SCEE" in row:
                        if "UC" in row:
                            row = row.split("UC")[-1]

                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:
                                    
                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue

                                        coordenadas_usadas.add(pos)
                                        if x0 >= 360 and x1 <= 410:
                                            rows_injecao_scee.append(float(word.replace(".", "").replace(",", ".").replace("-", "")))
                                            break

                            except Exception as e:
                                pass

                        injecao_scee = math.floor(sum(rows_injecao_scee) * 100) / 100


                beneficio_bruto_scee = 0

                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "BENEFÍCIO TARIFÁRIO BRUTO" in row:
                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        
                                        coordenadas_usadas.add(pos)  

                                        if x0 >= 375 and x1 <= 410:
                                            beneficio_bruto_scee = word.replace(".", "").replace(",", ".")
                                        break

                            except Exception as e:
                                pass

                bandeira_tarifaria = 0

                if "BANDEIRA" in text:
                    regex = r"\b\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\b"
                    bandeira_tarifaria = re.findall(regex, text.split("BANDEIRA")[-1].strip().split("\n")[0])[2].replace(".", "").replace(",", ".")

                multa = 0
                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "MULTA" in row:
                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        if x0 >= 380 and x1 <= 410:
                                            coordenadas_usadas.add(pos)
                                            multa = word.replace(".", "").replace(",", ".")
                                            break

                            except Exception as e:
                                pass
                juros = 0

                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "JUROS" in row:
                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        if x0 >= 380 and x1 <= 410:
                                            coordenadas_usadas.add(pos)
                                            juros = word.replace(".", "").replace(",", ".")
                                            break

                            except Exception as e:
                                pass

                regex = r"R\$\*+\d+,\d{2}"

                valor_fatura = re.findall(regex, text.replace(".", ""))[0].replace("*", "").replace("R$", "").replace(",", ".")

                icms = 0

                if "ICMS" in text:
                    icms = text.split("ICMS")[-1].strip().split("\n")[0].split(" ")[-2].replace(".", "").replace(",", ".").replace("%", "")

                pis = 0

                if "PIS/PASEP" in text:
                    pis = text.split("PIS/PASEP")[-1].strip().split("\n")[0].split(" ")[-2].replace(".", "").replace(",", ".").replace("%", "")

                cofins = 0

                if "COFINS" in text:
                    cofins = text.split("COFINS")[-1].strip().split("\n")[0].split(" ")[-2].replace(".", "").replace(",", ".").replace("%", "")

                iluminacao = 0
                regex = r"\b\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\b"

                if "CONTRIB. ILUM. PÚBLICA" in text:
                    iluminacao = re.findall(regex, text.split("CONTRIB. ILUM. PÚBLICA")[-1].strip().split("\n")[0])[0].replace(".", "").replace(",", ".")

                consumo_geracao = 0
                
                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "ENERGIA GERAÇÃO - KWH" in row:
                        for word in row.split(" "):
                            try:
                                float(word.replace(".", "").replace(",", "."))

                                for word_pdfplumber in words_pdfplumber:

                                    if word_pdfplumber["text"].strip() == word.strip():
                                        x0, y0, x1, y1 = word_pdfplumber["x0"], word_pdfplumber["top"], word_pdfplumber["x1"], word_pdfplumber["bottom"]
                                        pos = (x0, y0)

                                        if pos in coordenadas_usadas:
                                            continue
                                        if x0 >= 478 and x1 <= 500:
                                            coordenadas_usadas.add(pos)
                                            consumo_geracao = word.replace(".", "").replace(",", ".")
                                            break

                            except Exception as e:
                                pass
                
                vencimento = ""

                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if mes_referencia in row:
                        vencimento = row.split(mes_referencia)[-1].strip().split(" ")[0]
                        break

                saldoKWH = 0

                rows = text.strip().split("\n")
                coordenadas_usadas.clear()
                for row in rows:
                    if "SALDO KWH" in row.upper():
                        saldoKWH = row.upper().split("SALDO KWH:")[-1].strip().split(" ")[0]
                        if saldoKWH.endswith(","):
                            saldoKWH = saldoKWH[:-1].replace(".", "").replace(",", ".")

                output.loc[len(output)] = [
                    cliente,
                    uc,
                    leitura,
                    consumo_compensado,
                    consumo_nao_compensado,
                    geracao_abatida,
                    valor_consumo_fatura_scee,
                    mes_referencia,
                    custo_fio_b,
                    consumo_scee,
                    injecao_scee,
                    beneficio_bruto_scee,
                    bandeira_tarifaria,
                    multa,
                    juros,
                    valor_fatura,
                    icms,
                    pis,
                    cofins,
                    iluminacao,
                    consumo_geracao,
                    vencimento,
                    saldoKWH,
                ]
            except Exception:
                print(f"ERROR: Não foi possível realizar o planilhamento do boleto {boleto}.")
                errors.append(boleto)

    for error in errors:
        shutil.move(error, f"./Erros")
    output.to_csv("Dados Planilhados.csv", index=False)
    output.to_excel("Dados Planilhados.xlsx", index=False)
    print("_"*100)

    input("=== PLANILHAMENTO FINALIZADO ===\n")
  
    

if __name__ == "__main__":
    main()