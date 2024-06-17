import os
import pandas as pd
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

def check_permissions(path):
    return os.access(path, os.R_OK | os.W_OK)

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text()
    except Exception as e:
        print(f"Erro ao ler o arquivo {pdf_path}: {e}")
    return text

def get_unique_filename(directory, filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base}_{counter}{extension}"
        counter += 1
    return new_filename

def main(pdf_folder, excel_path, excel_columns):
    if not os.path.isfile(excel_path):
        print(f"Arquivo Excel não encontrado: {excel_path}")
        return

    if not check_permissions(excel_path):
        print(f"Sem permissão para ler ou escrever no arquivo Excel: {excel_path}")
        return

    try:
        df = pd.read_excel(excel_path)
        print(f"Colunas disponíveis na planilha: {df.columns.tolist()}")
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return

    for key, column in excel_columns.items():
        if column not in df.columns:
            print(f"Coluna '{column}' não encontrada na planilha Excel.")
            return

    secname = df[excel_columns["secname"]].dropna().tolist()
    prname = df[excel_columns["prname"]].dropna().tolist()

    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            if not check_permissions(pdf_path):
                print(f"Sem permissão para ler ou escrever no arquivo PDF: {pdf_path}")
                continue

            if not os.path.exists(pdf_path):
                print(f"Arquivo PDF não encontrado: {pdf_path}")
                continue

            pdf_text = extract_text_from_pdf(pdf_path)

            for sname in secname:
                for pname in prname:
                    if pname in pdf_text and sname in pdf_text:
                        new_filename = f"{pname}_{sname}.pdf"
                        new_filename = get_unique_filename(pdf_folder, new_filename)
                        new_path = os.path.join(pdf_folder, new_filename)

                        print(f"Renomeando {pdf_path} para {new_path}")

                        try:
                            os.rename(pdf_path, new_path)
                            print(f"Arquivo {filename} renomeado para {new_filename}")
                        except Exception as e:
                            print(f"Erro ao renomear o arquivo {pdf_path} para {new_path}: {e}")
                        break

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        pdf_files = [f for f in os.listdir(folder_selected) if f.endswith('.pdf')]
        if not pdf_files:
            messagebox.showerror("Erro", "Esta pasta não contém PDFs.")
        else:
            folder_path.set(folder_selected)
            list_pdfs_in_folder(folder_selected)

def list_pdfs_in_folder(folder):
    pdfs = [f for f in os.listdir(folder) if f.endswith(".pdf")]
    if pdfs:
        messagebox.showinfo("Arquivos PDF na Pasta", "Arquivos PDF encontrados:\n" + "\n".join(pdfs))
    else:
        messagebox.showinfo("Arquivos PDF na Pasta", "Nenhum arquivo PDF encontrado na pasta selecionada.")

def select_file():
    file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_selected:
        file_path.set(file_selected)

def start_process():
    if not folder_path.get() or not file_path.get():
        messagebox.showerror("Erro", "Você deve selecionar uma pasta de PDFs e uma planilha Excel.")
        return

    excel_columns = {
        "secname": "SecName",
        "prname": "PrName"
    }

    main(folder_path.get(), file_path.get(), excel_columns)
    messagebox.showinfo("Concluído", "Processo de renomeação concluído.")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Renomear PDFs")
    root.geometry("400x300")
    root.config(bg='#9ACD32')

    folder_path = tk.StringVar()
    file_path = tk.StringVar()

    tk.Label(root, text="Selecione a pasta dos PDFs:", bg='#9ACD32').pack(pady=5)
    tk.Entry(root, textvariable=folder_path, width=50).pack(pady=5)
    tk.Button(root, text="Selecionar Pasta", command=select_folder).pack(pady=5)

    tk.Label(root, text="Selecione a planilha Excel:", bg='#9ACD32').pack(pady=5)
    tk.Entry(root, textvariable=file_path, width=50).pack(pady=5)
    tk.Button(root, text="Selecionar Arquivo", command=select_file).pack(pady=5)

    tk.Button(root, text="Iniciar Processo", command=start_process).pack(pady=20)

    root.mainloop()
