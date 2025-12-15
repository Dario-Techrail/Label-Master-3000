"""
Launcher principale per l'applicazione Generatore Documenti Excel.
Punto di ingresso dell'applicazione.
"""

import sys
import tkinter as tk
from tkinter import messagebox


def verifica_dipendenze():
    """Verifica che tutte le dipendenze siano installate."""
    dipendenze_mancanti = []

    try:
        import openpyxl
    except ImportError:
        dipendenze_mancanti.append("openpyxl")

    try:
        import pandas
    except ImportError:
        dipendenze_mancanti.append("pandas")

    try:
        import reportlab
    except ImportError:
        dipendenze_mancanti.append("reportlab")

    if dipendenze_mancanti:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Dipendenze Mancanti",
            f"Le seguenti librerie sono necessarie:\n\n" +
            "\n".join(dipendenze_mancanti) +
            f"\n\nInstallale con:\npip install {' '.join(dipendenze_mancanti)}"
        )
        root.destroy()
        sys.exit(1)


def main():
    """Funzione principale per avviare l'applicazione."""
    # Verifica dipendenze
    verifica_dipendenze()

    # Importa e avvia interfaccia
    try:
        from ui_modules import InterfacciaGeneratoreExcel

        root = tk.Tk()
        app = InterfacciaGeneratoreExcel(root)
        root.mainloop()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Errore di Avvio",
            f"Errore durante l'avvio dell'applicazione:\n\n{str(e)}"
        )
        root.destroy()
        sys.exit(1)


if __name__ == "__main__":
    print("=" * 70)
    print("Generatore Documenti Excel - Interfaccia Grafica")
    print("=" * 70)
    print("Avvio in corso...")
    print()
    main()
