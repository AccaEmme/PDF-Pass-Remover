import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pikepdf import Pdf, PasswordError
import zipfile
import fitz  # PyMuPDF

"""
================================================================================
Creato da AccaEmme, con l'ausilio predominante di ChatGPT-4.
GitHub: https://github.com/AccaEmme

Questo programma è progettato per rimuovere la password da file PDF protetti.
Può essere eseguito sia tramite un'interfaccia grafica utente (GUI) sia tramite
la riga di comando.

Come utilizzare il programma:

1. GUI:
   - Esegui il programma senza argomenti (es: `python nome_programma.py`).
   - Seleziona un file PDF protetto o un archivio ZIP contenente un PDF.
   - Specifica il percorso di destinazione per il PDF decriptato.
   - Inserisci la password del PDF e clicca su "Rimuovi Password".
   - Il PDF senza password sarà salvato nel percorso di destinazione scelto.

2. Riga di Comando:
   - Esegui il programma con i parametri necessari. Ad esempio:
     ```
     python nome_programma.py --remove-password -p "password" -i "input.pdf" -o "output.pdf"
     ```
   - Questo comando rimuoverà la password dal file PDF specificato e salverà il risultato nel percorso di destinazione.

================================================================================

Requisiti di installazione:

1. Python 3.x: Assicurati di avere installato Python 3.x sul tuo sistema.

2. Librerie Python:
   - `tkinter`: Utilizzata per creare l'interfaccia grafica. In genere è preinstallata con Python, ma su alcune distribuzioni Linux potrebbe essere necessario installarla separatamente.
   - `pikepdf`: Utilizzata per rimuovere la password dai PDF.
     Installa con: `pip install pikepdf`
   - `PyMuPDF` (fitz): Utilizzata per gestire file PDF e portfolio.
     Installa con: `pip install pymupdf`
   - `zipfile`: Modulo standard di Python per la gestione dei file ZIP.

3. Programmi di terze parti:
   - `qpdf`: Utilizzato come strumento alternativo per rimuovere la password dai PDF.
     Su Linux, installalo con: `sudo apt-get install qpdf`
     Su macOS, installalo con: `brew install qpdf`
     Su Windows, scarica e installa da: https://sourceforge.net/projects/qpdf/

================================================================================
"""

def rimuovi_password_pikepdf(input_pdf, output_pdf, password):
    """Rimuove la password da un PDF utilizzando pikepdf."""
    try:
        with Pdf.open(input_pdf, password=password) as pdf:
            pdf.save(output_pdf)
        print(f"Password rimossa con successo con pikepdf. Il nuovo file è salvato come '{output_pdf}'.")
        return True
    except PasswordError:
        print("Errore: Password non corretta o decriptazione fallita con pikepdf.")
        return False
    except Exception as e:
        print(f"Errore con pikepdf: {e}")
        return False

def rimuovi_password_qpdf(input_pdf, output_pdf, password):
    """Rimuove la password da un PDF utilizzando qpdf."""
    try:
        result = subprocess.run(
            ['qpdf', '--decrypt', '--password={}'.format(password), input_pdf, output_pdf],
            check=True,
            text=True,
            capture_output=True
        )
        print(result.stdout)
        print(f"Password rimossa con successo con qpdf. Il nuovo file è salvato come '{output_pdf}'.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Errore con qpdf: {e.stderr}")
        return False

def estrai_pdf_da_zip(zip_path, output_dir):
    """Estrai il contenuto di un archivio ZIP e restituisci il percorso del primo PDF trovato."""
    if zipfile.is_zipfile(zip_path):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
            print(f"File estratti nella cartella: {output_dir}")
            for file_name in os.listdir(output_dir):
                if file_name.lower().endswith('.pdf'):
                    pdf_path = os.path.join(output_dir, file_name)
                    print(f"PDF trovato: {pdf_path}")
                    return pdf_path
    else:
        print("Il file non è un archivio ZIP.")
    return None

def gui_mode():
    def seleziona_file_sorgente():
        file_path = filedialog.askopenfilename(
            title="Seleziona il PDF Protetto o l'Archivio ZIP",
            filetypes=[("PDF or ZIP files", "*.pdf;*.zip"), ("PDF files", "*.pdf"), ("ZIP files", "*.zip")]
        )
        if file_path:
            input_pdf_path.set(file_path)

    def seleziona_file_destinazione():
        file_path = filedialog.asksaveasfilename(
            title="Salva PDF Senza Password",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            output_pdf_path.set(file_path)

    def toggle_password_visibility():
        if password_entry.cget('show') == '':
            password_entry.config(show='*')
            toggle_button.config(text='👁')
        else:
            password_entry.config(show='')
            toggle_button.config(text='👁‍🗨')

    def rimuovi_password_gui():
        input_pdf = input_pdf_path.get()
        output_pdf = output_pdf_path.get()
        password = password_entry.get()

        if not input_pdf:
            messagebox.showerror("Errore", "Per favore, seleziona il file PDF sorgente o l'archivio ZIP.")
            return
        if not output_pdf:
            messagebox.showerror("Errore", "Per favore, specifica il percorso del file PDF di destinazione.")
            return
        if not password:
            messagebox.showerror("Errore", "Per favore, inserisci la password.")
            return

        # Se il file di input è un ZIP, estrai il PDF
        if input_pdf.lower().endswith('.zip'):
            input_pdf = estrai_pdf_da_zip(input_pdf, './estratto')
            if not input_pdf:
                messagebox.showerror("Errore", "Nessun PDF trovato nell'archivio ZIP.")
                return

        # Tentativo di rimuovere la password usando pikepdf prima
        success = rimuovi_password_pikepdf(input_pdf, output_pdf, password)
        if not success:
            # Se pikepdf non funziona, prova con qpdf
            success = rimuovi_password_qpdf(input_pdf, output_pdf, password)
        
        if success:
            messagebox.showinfo("Successo", f"Password rimossa con successo. Lo strumento utilizzato è {'pikepdf' if rimuovi_password_pikepdf else 'qpdf'}.")
        else:
            messagebox.showerror("Errore", "Errore durante la rimozione della password.")

    root = tk.Tk()
    root.title("Rimozione Password PDF")

    input_pdf_path = tk.StringVar()
    output_pdf_path = tk.StringVar()

    tk.Label(root, text="File PDF sorgente o Archivio ZIP:").pack(pady=5)
    tk.Entry(root, textvariable=input_pdf_path, width=50).pack(padx=5, pady=5)
    tk.Button(root, text="Seleziona...", command=seleziona_file_sorgente).pack(pady=5)

    tk.Label(root, text="File PDF di destinazione:").pack(pady=5)
    tk.Entry(root, textvariable=output_pdf_path, width=50).pack(padx=5, pady=5)
    tk.Button(root, text="Seleziona...", command=seleziona_file_destinazione).pack(pady=5)

    tk.Label(root, text="Password del PDF:").pack(pady=5)
    password_frame = tk.Frame(root)
    password_frame.pack(pady=5)
    password_entry = tk.Entry(password_frame, show="*")
    password_entry.pack(side=tk.LEFT, padx=5)
    toggle_button = tk.Button(password_frame, text="👁", command=toggle_password_visibility)
    toggle_button.pack(side=tk.LEFT, padx=5)

    tk.Button(root, text="Rimuovi Password", command=rimuovi_password_gui).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        import argparse
        parser = argparse.ArgumentParser(description="Rimuovi la password da un PDF.")
        parser.add_argument('--remove-password', action='store_true', help="Rimuove la password da un PDF.")
        parser.add_argument('-p', '--password', type=str, help="La password del PDF.")
        parser.add_argument('-i', '--input', type=str, help="Il percorso del file PDF sorgente.")
        parser.add_argument('-o', '--output', type=str, help="Il percorso del file PDF di destinazione.")

        args = parser.parse_args()

        if args.remove_password:
            if args.password:
                input_pdf = args.input
                output_pdf = args.output if args.output else filedialog.asksaveasfilename(title="Salva PDF Senza Password", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

                if input_pdf and output_pdf:
                    # Se il file di input è un ZIP, estrai il PDF
                    if input_pdf.lower().endswith('.zip'):
                        input_pdf = estrai_pdf_da_zip(input_pdf, './estratto')
                        if not input_pdf:
                            print("Nessun PDF trovato nell'archivio ZIP.")
                            sys.exit(1)

                    # Tentativo di rimuovere la password usando pikepdf prima
                    success = rimuovi_password_pikepdf(input_pdf, output_pdf, args.password)
                    if not success:
                        # Se pikepdf non funziona, prova con qpdf
                        success = rimuovi_password_qpdf(input_pdf, output_pdf, args.password)
                    
                    if success:
                        print(f"Password rimossa con successo. Lo strumento utilizzato è {'pikepdf' if rimuovi_password_pikepdf else 'qpdf'}.")
                    else:
                        print("Errore durante la rimozione della password.")
                else:
                    print("Errore: è necessario specificare sia il file sorgente sia il file di destinazione.")
            else:
                print("Errore: è necessario specificare la password con l'opzione -p.")
        else:
            print("Errore: è necessario utilizzare l'opzione --remove-password per rimuovere la password.")
    else:
        gui_mode()

