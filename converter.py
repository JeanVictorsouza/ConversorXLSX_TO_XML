import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xml.etree.ElementTree as ET

def selecionar_arquivos():
    root = tk.Tk()
    try:
        root.iconbitmap("icone.ico")  # Substitua pelo caminho real se necessário
    except Exception as e:
        print(f"Não foi possível carregar o ícone: {e}")
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))
    root.withdraw()

    caminho_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel", 
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    caminho_xml = filedialog.asksaveasfilename(
        defaultextension=".xml", 
        filetypes=[("Arquivos XML", "*.xml")]
    )
    root.destroy()
    return caminho_excel, caminho_xml

def excel_para_xml(caminho_excel, caminho_xml):
    # Força todos os dados como string ao ler o Excel
    xls = pd.ExcelFile(caminho_excel)
    namespace = "http://tempuri.org/FinCFOImportacao.xsd"
    root = ET.Element("FinCFOImportacao", xmlns=namespace)

    def adicionar_elementos(sheet_name):
        if sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, dtype=str)  # <- Aqui força tudo como texto
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

            for _, row in df.iterrows():
                elemento_pai = ET.SubElement(root, sheet_name)
                for col in df.columns:
                    valor = str(row[col]).strip() if pd.notna(row[col]) else ""
                    ET.SubElement(elemento_pai, col).text = valor

    for sheet in ["FCFO", "FCFOCOMPL", "FDADOSPGTO", "FDADOSPGTODEF"]:
        adicionar_elementos(sheet)

    tree = ET.ElementTree(root)
    tree.write(caminho_xml, encoding="utf-8", xml_declaration=True)
    print(f"✅ Arquivo XML gerado com sucesso em: {caminho_xml}")

def main():
    caminho_excel, caminho_xml = selecionar_arquivos()
    if caminho_excel and caminho_xml:
        print(f"📄 Arquivo Excel: {caminho_excel}")
        print(f"💾 Arquivo XML será salvo como: {caminho_xml}")
        excel_para_xml(caminho_excel, caminho_xml)
    else:
        print("❌ Erro: Nenhum arquivo selecionado.")

if __name__ == "__main__":
    main()
