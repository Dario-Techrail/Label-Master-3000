"""
Modulo unificato per tutte le interfacce utente dell'applicazione.
Contiene tutte le classi per le interfacce grafiche e i componenti UI.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Optional
import re
from pathlib import Path
import pandas as pd
import json
from datetime import datetime
from business_logic import (
    GeneratoreExcel, GestioneComponenti, DataProcessor,
    PDFLabelGenerator, WordLabelGenerator
)


# ============================================================================
# FUNZIONI HELPER
# ============================================================================

def bind_mousewheel_to_canvas(canvas):
    """Aggiunge il supporto per la rotella del mouse a un canvas scrollabile.

    Args:
        canvas: Il canvas a cui aggiungere il supporto per la rotella del mouse
    """
    def _on_mousewheel(event):
        # Verifica che ci sia effettivamente contenuto scrollabile
        bbox = canvas.bbox("all")
        if bbox:
            # Ottieni la posizione corrente della scrollbar
            current_view = canvas.yview()

            # Calcola lo scroll richiesto
            scroll_amount = int(-1*(event.delta/120))

            # Previeni lo scrolling oltre i limiti
            if scroll_amount < 0 and current_view[0] <= 0:
                # Già al top, non scrollare verso l'alto
                return "break"
            elif scroll_amount > 0 and current_view[1] >= 1.0:
                # Già al bottom, non scrollare verso il basso
                return "break"

            canvas.yview_scroll(scroll_amount, "units")
        return "break"

    # Binding diretto al canvas
    canvas.bind("<MouseWheel>", _on_mousewheel)

    # Funzione per propagare il binding ai widget figli
    def _bind_all_children(widget):
        widget.bind("<MouseWheel>", _on_mousewheel)
        for child in widget.winfo_children():
            _bind_all_children(child)

    # Ritarda il binding dei figli per assicurarsi che siano già creati
    canvas.after(100, lambda: _bind_all_children(canvas))


# ============================================================================
# INTERFACCIA PRINCIPALE
# ============================================================================

class InterfacciaGeneratoreExcel:
    """Interfaccia grafica principale per la generazione di documenti Excel."""

    def __init__(self, root):
        """Inizializza l'interfaccia grafica."""
        self.root = root
        self.root.title("Label Master 3000")
        self.root.geometry("1400x900")

        style = ttk.Style()
        style.theme_use('clam')

        self.gestore_componenti = GestioneComponenti()
        self.generatore = GeneratoreExcel(gestione_componenti=self.gestore_componenti)

        # Formato: {nome_componente: {'quantita': int, 'sn_iniziale_override': int|None}}
        self.componenti_selezionati: Dict[str, Dict] = {}
        self.ultimo_file_generato: Optional[Path] = None

        # Directory per i preset di componenti
        self.preset_dir = Path("DB/preset_componenti")
        self.preset_dir.mkdir(parents=True, exist_ok=True)

        self.input_file = None
        self.csv_reg_input_file = None
        self.CSV_Registrazione_file = None
        self.import_gestionale_input_file = None
        self.Import_Gestionale_file = None
        self.etichettebox_input_file = None
        self.etichettebox_output_file = None
        self.etichettepdf_image_file = None
        self.etichettepdf_output_file = None
        self.etichetteword_output_file = None

        self.csv_reg_tab = None
        self.import_gestionale_tab = None
        self.etichettebox_tab = None
        self.etichettepdf_tab = None
        self.etichetteword_tab = None

        self._crea_menu()
        self._crea_interfaccia()

    def _crea_menu(self):
        """Crea la barra dei menu."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        menu_file = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=menu_file)
        menu_file.add_command(label="Esci", command=self.root.quit)

    def _crea_interfaccia(self):
        """Crea l'interfaccia principale con notebook."""
        # Frame header con logo in alto a destra
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Spazio vuoto a sinistra per allineare il logo a destra
        ttk.Label(header_frame, text="").pack(side=tk.LEFT, expand=True)
        
        # Logo in alto a destra
        try:
            from PIL import Image, ImageTk
            logo_path = "resources/logo.png"
            if Path(logo_path).exists():
                img = Image.open(logo_path)
                img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                logo_label = ttk.Label(header_frame, image=photo)
                logo_label.image = photo  # Mantieni riferimento per evitare garbage collection
                logo_label.pack(side=tk.RIGHT, padx=5)
        except Exception as e:
            print(f"Errore nel caricamento del logo: {e}")
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Crea prima la scheda Gestione Componenti
        self.gestione_componenti_tab = GestioneComponentiTab(self.notebook, self)

        # Poi crea la scheda Generazione Documento
        main_tab = ttk.Frame(self.notebook)
        self.notebook.add(main_tab, text="Generazione Documento")

        main_container = ttk.Frame(main_tab)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Configura il grid per permettere l'espansione verticale
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        def update_scrollregion(event=None):
            # Aggiorna la scrollregion basandosi sul contenuto effettivo
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Reset alla posizione iniziale se necessario
            if canvas.yview()[0] < 0:
                canvas.yview_moveto(0)

        scrollable_frame.bind("<Configure>", update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            min_width = event.width
            canvas.itemconfig(canvas_window, width=min_width)
            # Aggiorna anche la scrollregion quando il canvas viene ridimensionato
            update_scrollregion()

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=1)
        scrollable_frame.grid_columnconfigure(2, weight=1)

        #ttk.Label(
        #    scrollable_frame,
        #    text="Generazione Documento Excel",
        #    font=("Arial", 18, "bold")
        #).grid(row=0, column=0, columnspan=3, pady=20)

        self._crea_sezione_file_input_condiviso(scrollable_frame, row_start=1)
        self._crea_sezione_dati_generali(scrollable_frame, row_start=2)
        self._crea_sezione_componenti(scrollable_frame, row_start=7)

        ttk.Button(
            scrollable_frame,
            text="Genera Documento Excel",
            command=self._genera_documento,
            style="Accent.TButton"
        ).grid(row=50, column=0, columnspan=3, pady=30, ipadx=20, ipady=10)

        # Crea le altre schede
        self.csv_reg_tab = CSVRegTab(self.notebook, self)
        self.import_gestionale_tab = ImportGestionaleTab(self.notebook, self)
        # Scheda 'Etichette Bus' rimossa dall'interfaccia principale
        self.etichettebox_tab = None
        self.etichettepdf_tab = EtichettePDFTab(self.notebook, self)
        self.etichetteword_tab = EtichetteWordTab(self.notebook, self)

        self._configura_stile()

    def _crea_sezione_file_input_condiviso(self, parent, row_start: int):
        """Crea la sezione per la selezione del file di input condiviso."""
        frame_file_input = ttk.LabelFrame(
            parent,
            text="File di Input Condiviso (per tutte le schede)",
            padding=25
        )
        frame_file_input.grid(row=row_start, column=0, columnspan=3, sticky="ew", pady=15, padx=20)

        ttk.Label(
            frame_file_input,
            text="Seleziona un file Excel da utilizzare in tutte le schede:",
            font=("Arial", 10),
            wraplength=800
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        self.shared_input_label = ttk.Label(
            frame_file_input,
            text="Nessun file selezionato",
            foreground="gray",
            font=("Arial", 10)
        )
        self.shared_input_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            frame_file_input,
            text="Scegli File Input",
            command=self._seleziona_file_input_condiviso
        ).grid(row=1, column=1, padx=10)

        frame_file_input.columnconfigure(0, weight=1)
        frame_file_input.columnconfigure(1, weight=0)

    def _seleziona_file_input_condiviso(self):
        """Seleziona il file di input condiviso."""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input condiviso",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.input_file = filename
            self.shared_input_label.config(
                text=f"{Path(filename).name}",
                foreground="green"
            )

            if self.csv_reg_tab:
                self.csv_reg_tab.load_shared_input_file(filename)
            if self.import_gestionale_tab:
                self.import_gestionale_tab.load_shared_input_file(filename)
            if self.etichettebox_tab:
                self.etichettebox_tab.load_shared_input_file(filename)
            if self.etichettepdf_tab:
                self.etichettepdf_tab.load_shared_input_file(filename)
            if self.etichetteword_tab:
                self.etichetteword_tab.load_shared_input_file(filename)

            messagebox.showinfo(
                "File Caricato",
                f"Il file '{Path(filename).name}' è stato caricato in tutte le schede."
            )

    def _crea_sezione_dati_generali(self, parent, row_start: int):
        """Crea la sezione per l'inserimento dei dati generali."""
        frame_dati = ttk.LabelFrame(parent, text="Dati Generali", padding=20)
        frame_dati.grid(row=row_start, column=0, columnspan=3, sticky="ew", pady=15, padx=20)

        # Layout a 4 colonne: label, input, label, input (due campi affiancati per riga)
        for c in range(4):
            frame_dati.columnconfigure(c, weight=1 if c % 2 == 1 else 0)

        # Row 0: Bolla Produzione | Bolla Vendita
        ttk.Label(frame_dati, text="Bolla Produzione:", font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=5, pady=6)
        self.entry_bolla_produzione = ttk.Entry(frame_dati, font=("Arial", 10))
        self.entry_bolla_produzione.grid(row=0, column=1, sticky="ew", padx=5, pady=6)

        ttk.Label(frame_dati, text="Bolla Vendita:", font=("Arial", 11)).grid(row=0, column=2, sticky="w", padx=5, pady=6)
        self.entry_bolla_vendita = ttk.Entry(frame_dati, font=("Arial", 10))
        self.entry_bolla_vendita.grid(row=0, column=3, sticky="ew", padx=5, pady=6)

        # Row 1: Numero Bus | Bus Iniziale
        ttk.Label(frame_dati, text="Numero Bus:", font=("Arial", 11)).grid(row=1, column=0, sticky="w", padx=5, pady=6)
        self.spinbox_numero_bus = ttk.Spinbox(frame_dati, from_=1, to=1000, width=12, font=("Arial", 10))
        self.spinbox_numero_bus.set(1)
        self.spinbox_numero_bus.grid(row=1, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(frame_dati, text="Bus Iniziale (opzionale):", font=("Arial", 11)).grid(row=1, column=2, sticky="w", padx=5, pady=6)
        self.spinbox_bus_iniziale = ttk.Spinbox(frame_dati, from_=1, to=999, width=12, font=("Arial", 10))
        self.spinbox_bus_iniziale.set("")
        self.spinbox_bus_iniziale.grid(row=1, column=3, sticky="w", padx=5, pady=6)

        ttk.Label(frame_dati, text="(se vuoto, inizia da 01)", font=("Arial", 8), foreground="gray").grid(row=2, column=1, sticky="w", padx=5)

        # Row 2: Fornitore | CLIENTE
        ttk.Label(frame_dati, text="Fornitore:", font=("Arial", 11)).grid(row=3, column=0, sticky="w", padx=5, pady=6)
        self.entry_fornitore = ttk.Entry(frame_dati, font=("Arial", 10))
        self.entry_fornitore.insert(0, "TECHRAIL")
        self.entry_fornitore.grid(row=3, column=1, sticky="ew", padx=5, pady=6)

        # Campi extra (2 per riga)
        self.dati_generali_extra_entries = {}

        extras = [
            ("CLIENTE", 3, 2),
            ("Modello Pullman", 4, 0),
            ("Ordine Acquisto", 4, 2),
            ("Ente_Trasporto", 5, 0),
        ]

        for name, row, col in extras:
            ttk.Label(frame_dati, text=f"{name}:", font=("Arial", 9)).grid(row=row, column=col, sticky="w", padx=5, pady=6)
            entry = ttk.Entry(frame_dati, font=("Arial", 9))
            entry.grid(row=row, column=col + 1, sticky="ew", padx=5, pady=6)
            self.dati_generali_extra_entries[name] = entry

        # Se serve uno spacing finale
        frame_dati.grid_rowconfigure(8, minsize=6)

    def _crea_sezione_componenti(self, parent, row_start: int):
        """Crea la sezione per la selezione dei componenti."""
        frame_componenti = ttk.LabelFrame(parent, text="Selezione Componenti", padding=25)
        frame_componenti.grid(row=row_start, column=0, columnspan=3, sticky="ew", pady=15, padx=20)

        frame_pulsanti_componenti = ttk.Frame(frame_componenti)
        frame_pulsanti_componenti.grid(row=0, column=0, columnspan=3, pady=10)

        ttk.Button(
            frame_pulsanti_componenti,
            text="Aggiungi Componente",
            command=self._aggiungi_componente_selezionato
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            frame_pulsanti_componenti,
            text="Aggiorna da Database",
            command=self._aggiorna_componenti_da_database
        ).pack(side=tk.LEFT, padx=5)

        ttk.Separator(frame_pulsanti_componenti, orient="vertical").pack(side=tk.LEFT, padx=15, fill=tk.Y)

        ttk.Button(
            frame_pulsanti_componenti,
            text="Salva Preset",
            command=self._salva_preset_componenti
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            frame_pulsanti_componenti,
            text="Carica Preset",
            command=self._carica_preset_componenti
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            frame_pulsanti_componenti,
            text="Gestisci Preset",
            command=self._gestisci_preset_componenti
        ).pack(side=tk.LEFT, padx=5)

        self.frame_lista_componenti = ttk.Frame(frame_componenti)
        self.frame_lista_componenti.grid(row=1, column=0, columnspan=3, sticky="ew")
        
        # Dizionario per tracciare i widget Entry per inizio indicizzazione
        self.entry_indic_widgets = {}

        headers = ["Componente", "CODE 12NC", "SN Iniziale", "Prefisso", "Inizio Indic.", "Quantità", "Azioni"]
        for idx, header in enumerate(headers):
            ttk.Label(
                self.frame_lista_componenti,
                text=header,
                font=("Arial", 10, "bold"),
                width=15 if idx < 4 else 10
            ).grid(row=0, column=idx, padx=5)

        ttk.Separator(self.frame_lista_componenti, orient="horizontal").grid(
            row=1, column=0, columnspan=7, sticky="ew", pady=5
        )

    def _aggiungi_componente_selezionato(self):
        """Apre una finestra per selezionare e aggiungere un componente."""
        finestra_selezione = tk.Toplevel(self.root)
        finestra_selezione.title("Seleziona Componente")
        finestra_selezione.geometry("600x400")

        ttk.Label(
            finestra_selezione,
            text="Seleziona un componente da aggiungere:",
            font=("Arial", 12)
        ).pack(pady=20)

        componenti_disponibili = self.gestore_componenti.ottieni_tutti_componenti()

        if not componenti_disponibili:
            ttk.Label(
                finestra_selezione,
                text="Nessun componente disponibile.\nCrea componenti dal menu Gestione.",
                font=("Arial", 10)
            ).pack(pady=20)
            ttk.Button(
                finestra_selezione,
                text="Chiudi",
                command=finestra_selezione.destroy
            ).pack(pady=10)
            return

        frame_lista = ttk.Frame(finestra_selezione)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, font=("Arial", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        for comp in componenti_disponibili:
            listbox.insert(tk.END, comp['nome'])

        frame_quantita = ttk.Frame(finestra_selezione)
        frame_quantita.pack(pady=10)

        #ttk.Label(frame_quantita, text="Quantità per bus:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        #spinbox_quantita = ttk.Spinbox(frame_quantita, from_=1, to=100, width=10)
        #spinbox_quantita.set(1)
        #spinbox_quantita.pack(side=tk.LEFT, padx=5)

        #frame_indic = ttk.Frame(finestra_selezione)
        #frame_indic.pack(pady=10)

        #ttk.Label(frame_indic, text="Inizio Indicizzazione (opzionale):", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        #entry_indic = ttk.Entry(frame_indic, width=20)
        #entry_indic.pack(side=tk.LEFT, padx=5)

        def conferma_selezione():
            selezione = listbox.curselection()
            if not selezione:
                messagebox.showwarning("Attenzione", "Seleziona un componente")
                return

            nome_componente = listbox.get(selezione[0])
            quantita = int(1)
            #indic_text = entry_indic.get().strip()

            # Inizializza con SN dal database (può essere None)
            comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_componente)
            sn_iniziale_default = comp_info.get('sn_iniziale') if comp_info else None

            # Gestisci inizio_indicizzazione_prefisso (accetta intero oppure lista separata da virgole)
            #inizio_indic = None
            #if indic_text:
            #    # supporta formato "1,3,5" oppure singolo numero "3"
            #    if ',' in indic_text:
            #        parts = [p.strip() for p in indic_text.split(',') if p.strip()]
            #        try:
            #            inizio_indic = [int(p) for p in parts]
            #        except ValueError:
            #            messagebox.showerror("Errore", "Inizio Indicizzazione deve essere una lista di numeri separati da virgola o un singolo numero")
            #            return
            #    else:
            #        try:
            #            inizio_indic = int(indic_text)
            #        except ValueError:
            #            messagebox.showerror("Errore", "Inizio Indicizzazione deve essere un numero intero")
            #            return
#
            self.componenti_selezionati[nome_componente] = {
                'quantita': quantita,
                'sn_iniziale_override': sn_iniziale_default,
                'inizio_indicizzazione_prefisso': None
            }
            self._aggiorna_lista_componenti()
            finestra_selezione.destroy()

        ttk.Button(
            finestra_selezione,
            text="Aggiungi",
            command=conferma_selezione
        ).pack(pady=10)

    def _aggiorna_lista_componenti(self):
        """Aggiorna la visualizzazione della lista componenti selezionati."""
        for widget in self.frame_lista_componenti.winfo_children()[7:]:
            widget.destroy()
        
        # Pulisci il dizionario dei widget Entry
        self.entry_indic_widgets = {}

        row = 2
        for nome_componente, comp_data in self.componenti_selezionati.items():
            comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_componente)

            if comp_info:
                ttk.Label(self.frame_lista_componenti, text=nome_componente, width=35).grid(
                    row=row, column=0, padx=5, pady=5
                )

                ttk.Label(self.frame_lista_componenti, text=comp_info.get('code_12nc', 'N/A'), width=15).grid(
                    row=row, column=1, padx=5, pady=5
                )

                # Entry modificabile per SN iniziale
                entry_sn = ttk.Entry(self.frame_lista_componenti, width=12)
                sn_value = comp_data.get('sn_iniziale_override')
                if sn_value is not None:
                    entry_sn.insert(0, str(sn_value))
                else:
                    entry_sn.insert(0, "Auto")
                entry_sn.grid(row=row, column=2, padx=5, pady=5)

                def aggiorna_sn(nome, entry):
                    sn_text = entry.get().strip()
                    if sn_text.upper() == "AUTO" or sn_text == "":
                        self.componenti_selezionati[nome]['sn_iniziale_override'] = None
                        print(f"DEBUG: SN per {nome} impostato a None (Auto)")
                    else:
                        try:
                            nuovo_sn = int(sn_text)
                            # Aggiorna prima nel dizionario locale
                            self.componenti_selezionati[nome]['sn_iniziale_override'] = nuovo_sn
                            # Poi salva nel database in modo permanente
                            self.gestore_componenti.aggiorna_sn_iniziale(nome, nuovo_sn)
                            print(f"DEBUG UI: SN per {nome} impostato a {nuovo_sn} e salvato nel DB")
                        except ValueError:
                            messagebox.showerror("Errore", f"SN iniziale non valido per {nome}")
                            entry.delete(0, tk.END)
                            entry.insert(0, "Auto")
                            self.componenti_selezionati[nome]['sn_iniziale_override'] = None
                            print(f"DEBUG: SN per {nome} non valido, reimpostato a None")

                entry_sn.bind('<FocusOut>', lambda e, nome=nome_componente, entry=entry_sn: aggiorna_sn(nome, entry))
                entry_sn.bind('<Return>', lambda e, nome=nome_componente, entry=entry_sn: aggiorna_sn(nome, entry))

                prefisso = comp_info.get('prefisso_tipo_scheda', '')
                ttk.Label(self.frame_lista_componenti, text=prefisso if prefisso else "N/A", width=10).grid(
                    row=row, column=3, padx=5, pady=5
                )

                # Entry modificabile per Inizio Indicizzazione Prefisso
                entry_indic = ttk.Entry(self.frame_lista_componenti, width=10)
                indic_value = comp_data.get('inizio_indicizzazione_prefisso')
                if indic_value is not None:
                    if isinstance(indic_value, list):
                        entry_indic.insert(0, ','.join(str(x) for x in indic_value))
                    else:
                        entry_indic.insert(0, str(indic_value))
                else:
                    entry_indic.insert(0, "")
                entry_indic.grid(row=row, column=4, padx=5, pady=5)
                
                # Memorizza il widget Entry per accedervi da _genera_documento
                self.entry_indic_widgets[nome_componente] = entry_indic

                # Traccia il valore precedente per rilevare cambiamenti
                prev_indic_value = [indic_value]  # lista per catturare in closure

                def aggiorna_indic(nome=nome_componente, entry=entry_indic):
                    indic_text = entry.get().strip()
                    # Solo se il testo è effettivamente diverso da prima
                    if indic_text == "" and prev_indic_value[0] is None:
                        # Campo vuoto e era vuoto prima: non fare nulla
                        return
                    
                    if indic_text == "":
                        self.componenti_selezionati[nome]['inizio_indicizzazione_prefisso'] = None
                        prev_indic_value[0] = None
                        print(f"DEBUG: Inizio Indic. per {nome} impostato a None")
                    else:
                        # supporta formato "1,3,5" oppure singolo numero "3"
                        if ',' in indic_text:
                            parts = [p.strip() for p in indic_text.split(',') if p.strip()]
                            try:
                                parsed = [int(p) for p in parts]
                                self.componenti_selezionati[nome]['inizio_indicizzazione_prefisso'] = parsed
                                prev_indic_value[0] = parsed
                                print(f"DEBUG: Inizio Indic. per {nome} impostato a lista {parsed}")
                            except ValueError:
                                messagebox.showerror("Errore", f"Inizio Indicizzazione deve essere una lista di numeri separati da virgola o un singolo numero per {nome}")
                                entry.delete(0, tk.END)
                                self.componenti_selezionati[nome]['inizio_indicizzazione_prefisso'] = None
                                prev_indic_value[0] = None
                                print(f"DEBUG: Inizio Indic. per {nome} non valido, reimpostato a None")
                        else:
                            try:
                                parsed = int(indic_text)
                                self.componenti_selezionati[nome]['inizio_indicizzazione_prefisso'] = parsed
                                prev_indic_value[0] = parsed
                                print(f"DEBUG: Inizio Indic. per {nome} impostato a {parsed}")
                            except ValueError:
                                messagebox.showerror("Errore", f"Inizio Indicizzazione non valido per {nome}")
                                entry.delete(0, tk.END)
                                self.componenti_selezionati[nome]['inizio_indicizzazione_prefisso'] = None
                                prev_indic_value[0] = None
                                print(f"DEBUG: Inizio Indic. per {nome} non valido, reimpostato a None")

                # Salva su Return e su FocusOut
                entry_indic.bind('<Return>', lambda e, nome=nome_componente, entry=entry_indic: aggiorna_indic())
                entry_indic.bind('<FocusOut>', lambda e, nome=nome_componente, entry=entry_indic: aggiorna_indic())

                spinbox_q = ttk.Spinbox(self.frame_lista_componenti, from_=1, to=100, width=8)
                spinbox_q.set(comp_data['quantita'])
                spinbox_q.grid(row=row, column=5, padx=5, pady=5)

                def aggiorna_quantita(nome=nome_componente, spinbox=spinbox_q):
                    try:
                        nuova_quantita = int(spinbox.get())
                        if nuova_quantita < 1:
                            spinbox.set(1)
                            nuova_quantita = 1
                        elif nuova_quantita > 100:
                            spinbox.set(100)
                            nuova_quantita = 100
                        self.componenti_selezionati[nome]['quantita'] = nuova_quantita
                        print(f"DEBUG: Quantità per {nome} aggiornata a {nuova_quantita}")
                    except ValueError:
                        spinbox.set(comp_data['quantita'])
                        print(f"DEBUG: Valore quantità non valido, reimpostato a {comp_data['quantita']}")

                spinbox_q.config(command=lambda n=nome_componente, s=spinbox_q: aggiorna_quantita(n, s))
                spinbox_q.bind('<FocusOut>', lambda e, n=nome_componente, s=spinbox_q: aggiorna_quantita(n, s))
                spinbox_q.bind('<Return>', lambda e, n=nome_componente, s=spinbox_q: aggiorna_quantita(n, s))

                ttk.Button(
                    self.frame_lista_componenti,
                    text="Rimuovi",
                    command=lambda n=nome_componente: self._rimuovi_componente(n),
                    width=10
                ).grid(row=row, column=6, padx=5, pady=5)

                row += 1

    def _aggiorna_componenti_da_database(self):
        """Aggiorna i dati dei componenti selezionati dal database."""
        componenti_aggiornati = 0
        for nome_componente in self.componenti_selezionati.keys():
            comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_componente)
            if comp_info:
                # Aggiorna sn_iniziale_override e inizio_indicizzazione_prefisso dal database
                sn_db = comp_info.get('sn_iniziale')
                self.componenti_selezionati[nome_componente]['sn_iniziale_override'] = sn_db
                self.componenti_selezionati[nome_componente]['inizio_indicizzazione_prefisso'] = comp_info.get('inizio_indicizzazione_prefisso')
                componenti_aggiornati += 1

        self._aggiorna_lista_componenti()
        messagebox.showinfo("Aggiornamento Completato",
                          f"{componenti_aggiornati} componente/i aggiornato/i dal database.")

    def _rimuovi_componente(self, nome_componente: str):
        """Rimuove un componente dalla selezione."""
        if nome_componente in self.componenti_selezionati:
            del self.componenti_selezionati[nome_componente]
            self._aggiorna_lista_componenti()

    def _salva_preset_componenti(self):
        """Salva la lista corrente di componenti come preset."""
        if not self.componenti_selezionati:
            messagebox.showwarning("Attenzione", "Nessun componente selezionato da salvare")
            return

        # Finestra per inserire il nome del preset
        finestra_salva = tk.Toplevel(self.root)
        finestra_salva.title("Salva Preset Componenti")
        finestra_salva.geometry("400x150")
        finestra_salva.transient(self.root)
        finestra_salva.grab_set()

        ttk.Label(
            finestra_salva,
            text="Inserisci un nome per questo preset:",
            font=("Arial", 11)
        ).pack(pady=20)

        entry_nome = ttk.Entry(finestra_salva, width=40, font=("Arial", 10))
        entry_nome.pack(pady=10, padx=20)
        entry_nome.focus()

        def salva():
            nome_preset = entry_nome.get().strip()
            if not nome_preset:
                messagebox.showwarning("Attenzione", "Inserisci un nome per il preset")
                return

            # Rimuovi caratteri non validi per i nomi di file
            nome_preset_safe = "".join(c for c in nome_preset if c.isalnum() or c in (' ', '-', '_')).strip()
            if not nome_preset_safe:
                messagebox.showerror("Errore", "Nome preset non valido")
                return

            preset_file = self.preset_dir / f"{nome_preset_safe}.json"

            # Verifica se esiste già
            if preset_file.exists():
                risposta = messagebox.askyesno(
                    "Preset Esistente",
                    f"Il preset '{nome_preset_safe}' esiste già. Vuoi sovrascriverlo?"
                )
                if not risposta:
                    return

            try:
                # Salva solo i nomi dei componenti, non le altre proprietà
                nomi_componenti = list(self.componenti_selezionati.keys())

                preset_data = {
                    'nome': nome_preset,
                    'data_creazione': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'componenti': nomi_componenti
                }

                with open(preset_file, 'w', encoding='utf-8') as f:
                    json.dump(preset_data, f, indent=2, ensure_ascii=False)

                messagebox.showinfo("Successo", f"Preset '{nome_preset}' salvato correttamente")
                finestra_salva.destroy()

            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel salvare il preset: {str(e)}")

        def annulla():
            finestra_salva.destroy()

        frame_pulsanti = ttk.Frame(finestra_salva)
        frame_pulsanti.pack(pady=10)

        ttk.Button(frame_pulsanti, text="Salva", command=salva).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_pulsanti, text="Annulla", command=annulla).pack(side=tk.LEFT, padx=5)

        # Permetti di salvare con Enter
        entry_nome.bind('<Return>', lambda e: salva())

    def _carica_preset_componenti(self):
        """Carica un preset di componenti salvato."""
        # Ottieni tutti i preset disponibili
        preset_files = list(self.preset_dir.glob("*.json"))

        if not preset_files:
            messagebox.showinfo("Info", "Nessun preset salvato trovato")
            return

        # Finestra per selezionare il preset
        finestra_carica = tk.Toplevel(self.root)
        finestra_carica.title("Carica Preset Componenti")
        finestra_carica.geometry("600x400")
        finestra_carica.transient(self.root)
        finestra_carica.grab_set()

        ttk.Label(
            finestra_carica,
            text="Seleziona un preset da caricare:",
            font=("Arial", 12)
        ).pack(pady=20)

        # Frame con lista e scrollbar
        frame_lista = ttk.Frame(finestra_carica)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, font=("Arial", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # Carica info preset
        preset_info_list = []
        for preset_file in sorted(preset_files, key=lambda x: x.stat().st_mtime, reverse=True):
            try:
                with open(preset_file, 'r', encoding='utf-8') as f:
                    preset_data = json.load(f)
                    nome = preset_data.get('nome', preset_file.stem)
                    data = preset_data.get('data_creazione', 'N/A')
                    componenti = preset_data.get('componenti', [])
                    num_componenti = len(componenti) if isinstance(componenti, list) else len(componenti.keys())

                    display_text = f"{nome} - {num_componenti} componenti - {data}"
                    listbox.insert(tk.END, display_text)
                    preset_info_list.append((preset_file, preset_data))
            except Exception as e:
                print(f"Errore nel caricare {preset_file}: {e}")

        def carica():
            selezione = listbox.curselection()
            if not selezione:
                messagebox.showwarning("Attenzione", "Seleziona un preset da caricare")
                return

            preset_file, preset_data = preset_info_list[selezione[0]]

            # Chiedi conferma se ci sono già componenti selezionati
            if self.componenti_selezionati:
                risposta = messagebox.askyesno(
                    "Conferma",
                    "Ci sono già componenti selezionati. Vuoi sostituirli con il preset?"
                )
                if not risposta:
                    return

            try:
                # Carica i nomi dei componenti dal preset
                nomi_componenti = preset_data.get('componenti', [])

                # Se il preset è nel vecchio formato (dict), estrai le chiavi
                if isinstance(nomi_componenti, dict):
                    nomi_componenti = list(nomi_componenti.keys())

                # Svuota i componenti selezionati
                self.componenti_selezionati.clear()

                # Carica ogni componente dal database
                componenti_non_trovati = []
                for nome_comp in nomi_componenti:
                    comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_comp)
                    if comp_info:
                        # Popola con i dati aggiornati dal database
                        self.componenti_selezionati[nome_comp] = {
                            'quantita': 1,  # Quantità default
                            'sn_iniziale_override': comp_info.get('sn_iniziale'),
                            'inizio_indicizzazione_prefisso': comp_info.get('inizio_indicizzazione_prefisso')
                        }
                    else:
                        componenti_non_trovati.append(nome_comp)

                # Aggiorna la visualizzazione
                self._aggiorna_lista_componenti()

                # Messaggio di successo con eventuali avvisi
                if componenti_non_trovati:
                    messagebox.showwarning(
                        "Preset Caricato con Avvisi",
                        f"Preset '{preset_data.get('nome')}' caricato.\n\n" +
                        f"Componenti non trovati nel database:\n" +
                        "\n".join(f"- {c}" for c in componenti_non_trovati)
                    )
                else:
                    messagebox.showinfo(
                        "Successo",
                        f"Preset '{preset_data.get('nome')}' caricato correttamente.\n" +
                        f"Tutti i dati sono stati aggiornati dal database."
                    )

                finestra_carica.destroy()

            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel caricare il preset: {str(e)}")

        def annulla():
            finestra_carica.destroy()

        frame_pulsanti = ttk.Frame(finestra_carica)
        frame_pulsanti.pack(pady=10)

        ttk.Button(frame_pulsanti, text="Carica", command=carica).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_pulsanti, text="Annulla", command=annulla).pack(side=tk.LEFT, padx=5)

        # Permetti di caricare con doppio click
        listbox.bind('<Double-Button-1>', lambda e: carica())

    def _gestisci_preset_componenti(self):
        """Finestra per gestire (visualizzare, rinominare, eliminare) i preset salvati."""
        preset_files = list(self.preset_dir.glob("*.json"))

        if not preset_files:
            messagebox.showinfo("Info", "Nessun preset salvato trovato")
            return

        # Finestra di gestione
        finestra_gestione = tk.Toplevel(self.root)
        finestra_gestione.title("Gestisci Preset Componenti")
        finestra_gestione.geometry("700x500")
        finestra_gestione.transient(self.root)
        finestra_gestione.grab_set()

        ttk.Label(
            finestra_gestione,
            text="Gestione Preset",
            font=("Arial", 14, "bold")
        ).pack(pady=15)

        # Frame con lista e dettagli
        frame_principale = ttk.Frame(finestra_gestione)
        frame_principale.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Lista preset a sinistra
        frame_lista = ttk.Frame(frame_principale)
        frame_lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        ttk.Label(frame_lista, text="Preset Salvati:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)

        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, font=("Arial", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # Dettagli preset a destra
        frame_dettagli = ttk.LabelFrame(frame_principale, text="Dettagli Preset", padding=10)
        frame_dettagli.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        text_dettagli = tk.Text(frame_dettagli, wrap=tk.WORD, font=("Arial", 9), height=15, width=30)
        text_dettagli.pack(fill=tk.BOTH, expand=True)
        text_dettagli.config(state=tk.DISABLED)

        # Funzione per caricare la lista
        def carica_lista():
            listbox.delete(0, tk.END)
            preset_info_list.clear()

            preset_files = list(self.preset_dir.glob("*.json"))
            for preset_file in sorted(preset_files, key=lambda x: x.stat().st_mtime, reverse=True):
                try:
                    with open(preset_file, 'r', encoding='utf-8') as f:
                        preset_data = json.load(f)
                        nome = preset_data.get('nome', preset_file.stem)

                        listbox.insert(tk.END, nome)
                        preset_info_list.append((preset_file, preset_data))
                except Exception as e:
                    print(f"Errore nel caricare {preset_file}: {e}")

        # Funzione per mostrare dettagli
        def mostra_dettagli(event=None):
            selezione = listbox.curselection()
            if not selezione:
                return

            preset_file, preset_data = preset_info_list[selezione[0]]

            text_dettagli.config(state=tk.NORMAL)
            text_dettagli.delete(1.0, tk.END)

            nome = preset_data.get('nome', 'N/A')
            data = preset_data.get('data_creazione', 'N/A')
            componenti = preset_data.get('componenti', [])

            # Se il preset è nel vecchio formato (dict), estrai le chiavi
            if isinstance(componenti, dict):
                componenti = list(componenti.keys())

            dettagli = f"Nome: {nome}\n\n"
            dettagli += f"Data creazione: {data}\n\n"
            dettagli += f"Numero componenti: {len(componenti)}\n\n"
            dettagli += "Componenti:\n"

            for comp_nome in componenti:
                dettagli += f"  - {comp_nome}\n"

            dettagli += "\n(I dati dei componenti verranno caricati dal database)"

            text_dettagli.insert(1.0, dettagli)
            text_dettagli.config(state=tk.DISABLED)

        preset_info_list = []
        carica_lista()
        listbox.bind('<<ListboxSelect>>', mostra_dettagli)

        # Pulsanti di azione
        frame_azioni = ttk.Frame(finestra_gestione)
        frame_azioni.pack(pady=10)

        def elimina_preset():
            selezione = listbox.curselection()
            if not selezione:
                messagebox.showwarning("Attenzione", "Seleziona un preset da eliminare")
                return

            preset_file, preset_data = preset_info_list[selezione[0]]
            nome = preset_data.get('nome', preset_file.stem)

            risposta = messagebox.askyesno(
                "Conferma Eliminazione",
                f"Sei sicuro di voler eliminare il preset '{nome}'?"
            )

            if risposta:
                try:
                    preset_file.unlink()
                    messagebox.showinfo("Successo", f"Preset '{nome}' eliminato")
                    carica_lista()
                    text_dettagli.config(state=tk.NORMAL)
                    text_dettagli.delete(1.0, tk.END)
                    text_dettagli.config(state=tk.DISABLED)
                except Exception as e:
                    messagebox.showerror("Errore", f"Errore nell'eliminare il preset: {str(e)}")

        ttk.Button(frame_azioni, text="Elimina Preset", command=elimina_preset).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_azioni, text="Chiudi", command=finestra_gestione.destroy).pack(side=tk.LEFT, padx=5)

    def _genera_documento(self):
        """Genera il documento Excel con i dati inseriti."""
        # IMPORTANTE: Leggi i valori attuali dai widget Entry di inizio indicizzazione
        # e salva in componenti_selezionati prima di procedere
        for nome_componente, entry_widget in self.entry_indic_widgets.items():
            indic_text = entry_widget.get().strip()
            
            if indic_text == "":
                self.componenti_selezionati[nome_componente]['inizio_indicizzazione_prefisso'] = None
                print(f"DEBUG _genera_documento: Inizio Indic. per {nome_componente} = None")
            else:
                # supporta formato "1,3,5" oppure singolo numero "3"
                if ',' in indic_text:
                    parts = [p.strip() for p in indic_text.split(',') if p.strip()]
                    try:
                        parsed = [int(p) for p in parts]
                        self.componenti_selezionati[nome_componente]['inizio_indicizzazione_prefisso'] = parsed
                        print(f"DEBUG _genera_documento: Inizio Indic. per {nome_componente} = lista {parsed}")
                    except ValueError:
                        messagebox.showerror("Errore", f"Inizio Indicizzazione deve essere una lista di numeri separati da virgola o un singolo numero per {nome_componente}")
                        return
                else:
                    try:
                        parsed = int(indic_text)
                        self.componenti_selezionati[nome_componente]['inizio_indicizzazione_prefisso'] = parsed
                        print(f"DEBUG _genera_documento: Inizio Indic. per {nome_componente} = {parsed}")
                    except ValueError:
                        messagebox.showerror("Errore", f"Inizio Indicizzazione non valido per {nome_componente}")
                        return
        
        bolla_produzione = self.entry_bolla_produzione.get().strip()
        bolla_vendita = self.entry_bolla_vendita.get().strip()

        if not bolla_produzione or not bolla_vendita:
            messagebox.showerror("Errore", "Bolla Produzione e Bolla Vendita sono obbligatori")
            return

        try:
            numero_bus = int(self.spinbox_numero_bus.get())
        except ValueError:
            messagebox.showerror("Errore", "Numero Bus non valido")
            return

        if numero_bus <= 0:
            messagebox.showerror("Errore", "Il numero di bus deve essere maggiore di 0")
            return

        bus_iniziale_str = self.spinbox_bus_iniziale.get().strip()
        if bus_iniziale_str:
            try:
                bus_iniziale = int(bus_iniziale_str)
                if bus_iniziale < 1:
                    messagebox.showerror("Errore", "Il bus iniziale deve essere maggiore di 0")
                    return
            except ValueError:
                messagebox.showerror("Errore", "Bus Iniziale non valido")
                return
        else:
            bus_iniziale = 1

        fornitore = self.entry_fornitore.get().strip()
        if not fornitore:
            fornitore = "TECHRAIL"

        componenti = []
        for nome_componente, comp_data in self.componenti_selezionati.items():
            comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_componente)
            if comp_info:
                # Usa sn_iniziale_override se specificato, altrimenti quello del database
                sn_override = comp_data.get('sn_iniziale_override')
                sn_finale = sn_override if sn_override is not None else comp_info.get('sn_iniziale')

                print(f"DEBUG: Componente={nome_componente}, Override={sn_override}, DB={comp_info.get('sn_iniziale')}, Finale={sn_finale}")

                componenti.append({
                    'nome': nome_componente,
                    'quantita': comp_data['quantita'],
                    'sn_iniziale': sn_finale,
                    'prefisso_tipo_scheda': comp_info.get('prefisso_tipo_scheda'),
                    'code_12nc': comp_info.get('code_12nc'),
                    'indicizzazione': comp_info.get('indicizzazione', True),
                    'inizio_indicizzazione_prefisso': comp_data.get('inizio_indicizzazione_prefisso')
                })

        nome_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="documento_componenti.xlsx"
        )

        if not nome_file:
            return

        try:
            if componenti:
                file_generato = self.generatore.crea_documento_con_componenti(
                    bolla_produzione=bolla_produzione,
                    bolla_vendita=bolla_vendita,
                    numero_bus=numero_bus,
                    componenti=componenti,
                    nome_file=nome_file,
                    bus_iniziale=bus_iniziale,
                    fornitore=fornitore
                )
            else:
                file_generato = self.generatore.crea_documento_bus(
                    bolla_produzione=bolla_produzione,
                    bolla_vendita=bolla_vendita,
                    numero_bus=numero_bus,
                    nome_file=nome_file,
                    bus_iniziale=bus_iniziale,
                    fornitore=fornitore
                )

            self.ultimo_file_generato = Path(file_generato)
            self.input_file = str(file_generato)

            # Aggiorna i valori di override con i nuovi valori dal database
            # Questo assicura che la prossima generazione parta dai valori aggiornati
            if componenti:
                for nome_componente in self.componenti_selezionati.keys():
                    comp_info = self.gestore_componenti.cerca_componente_per_nome(nome_componente)
                    if comp_info:
                        nuovo_sn = comp_info.get('sn_iniziale')
                        self.componenti_selezionati[nome_componente]['sn_iniziale_override'] = nuovo_sn
                        # Mantieni il valore di inizio_indicizzazione_prefisso che è stato persistito nel DB
                        # (è stato salvato dalla business logic durante la generazione)
                        persistito = comp_info.get('inizio_indicizzazione_prefisso')
                        self.componenti_selezionati[nome_componente]['inizio_indicizzazione_prefisso'] = persistito
                        print(f"DEBUG: Aggiornato override per {nome_componente} a {nuovo_sn}, inizio_indic={persistito}")
                # Aggiorna la visualizzazione della lista (questo ricarica anche i campi Entry)
                self._aggiorna_lista_componenti()

            self.shared_input_label.config(
                text=f"{Path(file_generato).name}",
                foreground="green"
            )

            if self.csv_reg_tab:
                self.csv_reg_tab.load_shared_input_file(file_generato)
            if self.import_gestionale_tab:
                self.import_gestionale_tab.load_shared_input_file(file_generato)
            if self.etichettebox_tab:
                self.etichettebox_tab.load_shared_input_file(file_generato)
            if self.etichettepdf_tab:
                self.etichettepdf_tab.load_shared_input_file(file_generato)
            if self.etichetteword_tab:
                self.etichetteword_tab.load_shared_input_file(file_generato)

            messagebox.showinfo(
                "Successo",
                f"Documento generato con successo!\n\n"
                f"File: {file_generato}\n\n"
                f"Il file è stato caricato automaticamente in tutte le schede."
            )

        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante la generazione:\n{str(e)}")

    def _configura_stile(self):
        """Configura gli stili dei widget."""
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 12, "bold"), padding=10)
        style.configure("TLabelframe", padding=15, relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", font=("Arial", 11, "bold"), foreground="#2C3E50")
        style.configure("TNotebook", padding=5, tabmargins=[2, 5, 2, 0])
        style.configure("TNotebook.Tab", font=("Arial", 10, "bold"), padding=[20, 10])


# ============================================================================
# TAB GESTIONE COMPONENTI
# ============================================================================

class GestioneComponentiTab:
    """Scheda dedicata alla gestione completa dei componenti (CRUD)."""

    def __init__(self, notebook, app_context):
        self.notebook = notebook
        self.app_context = app_context
        self.gestore_componenti = app_context.gestore_componenti
        self.componente_selezionato = None

        self.frame = ttk.Frame(self.notebook)
        self.notebook.add(self.frame, text="Gestione Componenti")

        self.create_widgets()
        self._carica_componenti()

    def create_widgets(self):
        """Crea il contenuto della scheda Gestione Componenti"""
        container = ttk.Frame(self.frame, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        #ttk.Label(
        #    container,
        #    text="Gestione Database Componenti",
        #    font=("Arial", 16, "bold")
        #).pack(pady=20)

        # Frame principale con layout a due colonne
        main_container = ttk.Frame(container)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Frame sinistra - Lista componenti
        frame_lista = ttk.LabelFrame(main_container, text="Componenti Esistenti", padding=15)
        frame_lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Treeview per lista componenti
        columns = ("nome", "code_12nc", "sn_iniziale", "prefisso", "indic_prefisso", "indicizzazione")
        self.tree = ttk.Treeview(frame_lista, columns=columns, show="headings", height=20)

        self.tree.heading("nome", text="Nome Componente")
        self.tree.heading("code_12nc", text="CODE 12NC")
        self.tree.heading("sn_iniziale", text="SN Iniziale")
        self.tree.heading("prefisso", text="Prefisso Tipo")
        self.tree.heading("indic_prefisso", text="Inizio Indic.")
        self.tree.heading("indicizzazione", text="Indic.")

        self.tree.column("nome", width=200)
        self.tree.column("code_12nc", width=120)
        self.tree.column("sn_iniziale", width=90)
        self.tree.column("prefisso", width=90)
        self.tree.column("indic_prefisso", width=90)
        self.tree.column("indicizzazione", width=70)

        scrollbar = ttk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_selezione_componente)

        # Frame destra - Dettagli componente
        frame_dettagli = ttk.LabelFrame(main_container, text="Dettagli Componente", padding=20)
        frame_dettagli.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Campo Nome
        ttk.Label(frame_dettagli, text="Nome Componente:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=10)
        self.entry_nome = ttk.Entry(frame_dettagli, width=40, font=("Arial", 10))
        self.entry_nome.grid(row=0, column=1, pady=10, padx=10, sticky="ew")

        # Campo CODE 12NC
        ttk.Label(frame_dettagli, text="CODE 12NC:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=10)
        self.entry_code_12nc = ttk.Entry(frame_dettagli, width=40, font=("Arial", 10))
        self.entry_code_12nc.grid(row=1, column=1, pady=10, padx=10, sticky="ew")

        # Campo SN Iniziale
        ttk.Label(frame_dettagli, text="SN Iniziale (opzionale):", font=("Arial", 10)).grid(row=2, column=0, sticky="w", pady=10)
        self.entry_sn_iniziale = ttk.Entry(frame_dettagli, width=40, font=("Arial", 10))
        self.entry_sn_iniziale.grid(row=2, column=1, pady=10, padx=10, sticky="ew")
        ttk.Label(frame_dettagli, text="(lascia vuoto per generazione automatica)", font=("Arial", 8), foreground="gray").grid(row=3, column=1, sticky="w", padx=10)

        # Campo Prefisso
        ttk.Label(frame_dettagli, text="Prefisso Tipo Scheda (opzionale):", font=("Arial", 10)).grid(row=4, column=0, sticky="w", pady=10)
        self.entry_prefisso = ttk.Entry(frame_dettagli, width=40, font=("Arial", 10))
        self.entry_prefisso.grid(row=4, column=1, pady=10, padx=10, sticky="ew")
        ttk.Label(frame_dettagli, text="(es: SU, CAM, RES - lascia vuoto se non necessario)", font=("Arial", 8), foreground="gray").grid(row=5, column=1, sticky="w", padx=10)

        # Campo Inizio Indicizzazione
        #ttk.Label(frame_dettagli, text="Inizio Indicizzazione Prefisso (opzionale):", font=("Arial", 10)).grid(row=6, column=0, sticky="w", pady=10)
        #self.entry_inizio_indic = ttk.Entry(frame_dettagli, width=40, font=("Arial", 10))
        #self.entry_inizio_indic.grid(row=6, column=1, pady=10, padx=10, sticky="ew")
        #ttk.Label(frame_dettagli, text="(numero da cui iniziare, es: 7 genera SU7, SU8, SU9...)", font=("Arial", 8), foreground="gray").grid(row=7, column=1, sticky="w", padx=10)

        # Checkbox Indicizzazione
        self.var_indicizzazione = tk.BooleanVar(value=True)
        self.chk_indicizzazione = ttk.Checkbutton(
            frame_dettagli,
            text="Indicizzazione (aggiungi numero al tipo scheda)",
            variable=self.var_indicizzazione
        )
        self.chk_indicizzazione.grid(row=8, column=1, sticky="w", padx=10, pady=5)

        # Pulsanti Salva/Annulla
        frame_pulsanti_dettagli = ttk.Frame(frame_dettagli)
        frame_pulsanti_dettagli.grid(row=9, column=0, columnspan=2, pady=30)

        self.btn_salva = ttk.Button(
            frame_pulsanti_dettagli,
            text="Salva Componente",
            command=self._salva_componente,
            state=tk.DISABLED
        )
        self.btn_salva.pack(side=tk.LEFT, padx=10)

        self.btn_annulla = ttk.Button(
            frame_pulsanti_dettagli,
            text="Annulla",
            command=self._annulla_modifica,
            state=tk.DISABLED
        )
        self.btn_annulla.pack(side=tk.LEFT, padx=10)

        frame_dettagli.grid_columnconfigure(1, weight=1)

        # Pulsanti Nuovo ed Elimina in basso, fuori dalle sezioni
        frame_pulsanti_bottom = ttk.Frame(container)
        frame_pulsanti_bottom.pack(pady=20)

        ttk.Button(frame_pulsanti_bottom, text="Nuovo Componente", command=self._nuovo_componente, width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(frame_pulsanti_bottom, text="Elimina Componente", command=self._elimina_componente, width=20).pack(side=tk.LEFT, padx=10)

    def _carica_componenti(self):
        """Carica i componenti nella Treeview."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        componenti = self.gestore_componenti.ottieni_tutti_componenti()
        for comp in componenti:
            sn_iniziale = comp.get('sn_iniziale')
            sn_text = str(sn_iniziale) if sn_iniziale is not None else "Auto"

            prefisso = comp.get('prefisso_tipo_scheda', '')
            prefisso_text = prefisso if prefisso else "N/A"

            inizio_indic = comp.get('inizio_indicizzazione_prefisso')
            inizio_indic_text = str(inizio_indic) if inizio_indic is not None else "N/A"

            indic = comp.get('indicizzazione', True)
            indic_text = "SI" if indic else "NO"

            self.tree.insert("", tk.END, values=(
                comp['nome'],
                comp.get('code_12nc', 'N/A'),
                sn_text,
                prefisso_text,
                inizio_indic_text,
                indic_text
            ))

    def _on_selezione_componente(self, event):
        """Gestisce la selezione di un componente dalla Treeview."""
        selezione = self.tree.selection()
        if not selezione:
            return

        item = self.tree.item(selezione[0])
        valori = item['values']
        nome_componente = valori[0]

        comp = self.gestore_componenti.cerca_componente_per_nome(nome_componente)
        if not comp:
            return

        self.componente_selezionato = nome_componente

        # Popola i campi
        self.entry_nome.delete(0, tk.END)
        self.entry_nome.insert(0, comp['nome'])

        self.entry_code_12nc.delete(0, tk.END)
        self.entry_code_12nc.insert(0, comp.get('code_12nc', ''))

        self.entry_sn_iniziale.delete(0, tk.END)
        sn_iniziale = comp.get('sn_iniziale')
        if sn_iniziale is not None:
            self.entry_sn_iniziale.insert(0, str(sn_iniziale))

        self.entry_prefisso.delete(0, tk.END)
        prefisso = comp.get('prefisso_tipo_scheda', '')
        if prefisso:
            self.entry_prefisso.insert(0, prefisso)

        #self.entry_inizio_indic.delete(0, tk.END)
        #inizio_indic = comp.get('inizio_indicizzazione_prefisso')
        #if inizio_indic is not None:
        #self.entry_inizio_indic.insert(0, str(inizio_indic))

        indicizzazione = comp.get('indicizzazione', True)
        self.var_indicizzazione.set(indicizzazione)

        # Abilita i pulsanti
        self.btn_salva.config(state=tk.NORMAL)
        self.btn_annulla.config(state=tk.NORMAL)

    def _nuovo_componente(self):
        """Prepara i campi per l'inserimento di un nuovo componente."""
        self.componente_selezionato = None
        self.tree.selection_remove(*self.tree.selection())

        # Pulisci tutti i campi
        self.entry_nome.delete(0, tk.END)
        self.entry_code_12nc.delete(0, tk.END)
        self.entry_sn_iniziale.delete(0, tk.END)
        self.entry_prefisso.delete(0, tk.END)
        #self.entry_inizio_indic.delete(0, tk.END)
        self.var_indicizzazione.set(True)

        # Abilita i pulsanti
        self.btn_salva.config(state=tk.NORMAL)
        self.btn_annulla.config(state=tk.NORMAL)

    def _salva_componente(self):
        """Salva il componente (nuovo o modificato)."""
        nome = self.entry_nome.get().strip()
        code_12nc = self.entry_code_12nc.get().strip()

        if not nome:
            messagebox.showwarning("Attenzione", "Inserisci il nome del componente")
            return

        if not code_12nc:
            messagebox.showwarning("Attenzione", "Inserisci il CODE 12NC")
            return

        sn_iniziale_text = self.entry_sn_iniziale.get().strip()
        sn_iniziale = None
        if sn_iniziale_text:
            try:
                sn_iniziale = int(sn_iniziale_text)
            except ValueError:
                messagebox.showerror("Errore", "SN Iniziale deve essere un numero intero")
                return

        prefisso = self.entry_prefisso.get().strip()
        prefisso = prefisso if prefisso else None

        inizio_indic_text = None#self.entry_inizio_indic.get().strip()
        inizio_indic = None
        if inizio_indic_text:
            try:
                inizio_indic = int(inizio_indic_text)
            except ValueError:
                messagebox.showerror("Errore", "Inizio Indicizzazione deve essere un numero intero")
                return

        indicizzazione = self.var_indicizzazione.get()

        if self.componente_selezionato:
            # Modifica esistente
            try:
                self.gestore_componenti.modifica_componente(
                    nome_originale=self.componente_selezionato,
                    nuovo_nome=nome,
                    code_12nc=code_12nc,
                    sn_iniziale=sn_iniziale,
                    prefisso_tipo_scheda=prefisso,
                    inizio_indicizzazione_prefisso=inizio_indic,
                    indicizzazione=indicizzazione
                )
                messagebox.showinfo("Successo", "Componente aggiornato con successo")
            except ValueError as e:
                messagebox.showerror("Errore", str(e))
                return
        else:
            # Nuovo componente
            try:
                self.gestore_componenti.aggiungi_componente(
                    nome=nome,
                    code_12nc=code_12nc,
                    sn_iniziale=sn_iniziale,
                    prefisso_tipo_scheda=prefisso,
                    inizio_indicizzazione_prefisso=inizio_indic,
                    indicizzazione=indicizzazione
                )
                messagebox.showinfo("Successo", "Componente aggiunto con successo")
            except ValueError as e:
                messagebox.showerror("Errore", str(e))
                return

        # Ricarica la lista e aggiorna anche la sezione di selezione nella prima pagina
        self._carica_componenti()
        self.app_context._aggiorna_lista_componenti()
        self._annulla_modifica()

    def _elimina_componente(self):
        """Elimina il componente selezionato."""
        selezione = self.tree.selection()
        if not selezione:
            messagebox.showwarning("Attenzione", "Seleziona un componente da eliminare")
            return

        item = self.tree.item(selezione[0])
        nome_componente = item['values'][0]

        risposta = messagebox.askyesno(
            "Conferma Eliminazione",
            f"Sei sicuro di voler eliminare il componente '{nome_componente}'?"
        )

        if risposta:
            try:
                self.gestore_componenti.elimina_componente(nome_componente)
                messagebox.showinfo("Successo", "Componente eliminato con successo")
                self._carica_componenti()
                self.app_context._aggiorna_lista_componenti()
                self._annulla_modifica()
            except ValueError as e:
                messagebox.showerror("Errore", str(e))

    def _annulla_modifica(self):
        """Annulla la modifica in corso e pulisce i campi."""
        self.componente_selezionato = None
        self.tree.selection_remove(*self.tree.selection())

        self.entry_nome.delete(0, tk.END)
        self.entry_code_12nc.delete(0, tk.END)
        self.entry_sn_iniziale.delete(0, tk.END)
        self.entry_prefisso.delete(0, tk.END)
        self.entry_inizio_indic.delete(0, tk.END)
        self.var_indicizzazione.set(True)

        self.btn_salva.config(state=tk.DISABLED)
        self.btn_annulla.config(state=tk.DISABLED)


# ============================================================================
# TAB CSV REGISTRAZIONE
# ============================================================================

class CSVRegTab:
    """Gestisce la scheda CSV Reg"""

    def __init__(self, parent_notebook, app_context):
        self.notebook = parent_notebook
        self.app_context = app_context
        self.description_checkboxes = {}

        self.frame = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame, text="CSV di Registrazione")

        self.create_widgets()

    def create_widgets(self):
        """Crea il contenuto della scheda CSV Reg"""
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configura il grid per permettere l'espansione verticale
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0, bg="white")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding=20)

        def update_scrollregion(event=None):
            # Aggiorna la scrollregion basandosi sul contenuto effettivo
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Reset alla posizione iniziale se necessario
            if canvas.yview()[0] < 0:
                canvas.yview_moveto(0)

        main_frame.bind("<Configure>", update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            # Aggiorna anche la scrollregion quando il canvas viene ridimensionato
            update_scrollregion()

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=1)

        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_rowconfigure(2, weight=0)
        main_frame.grid_rowconfigure(3, weight=0)
        main_frame.grid_rowconfigure(4, weight=0)
        main_frame.grid_rowconfigure(5, weight=0)
        main_frame.grid_rowconfigure(6, weight=1)

        #ttk.Label(
        #    main_frame,
        #    text="Generazione CSV di Registrazione",
        #    font=('Arial', 16, 'bold')
        #).grid(row=0, column=0, columnspan=3, pady=20)

        input_section = ttk.LabelFrame(main_frame, text="File di Input", padding=15)
        input_section.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            input_section,
            text="Seleziona il file Excel da processare:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.input_label = ttk.Label(
            input_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.input_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            input_section,
            text="Scegli File Input",
            command=self.select_input_file
        ).grid(row=1, column=1, padx=10)

        input_section.columnconfigure(0, weight=1)
        input_section.columnconfigure(1, weight=0)

        # La sezione "Dati Comuni" è stata spostata nella scheda principale
        # (Generazione Documento) nella sezione "Dati Generali".
        # I campi saranno letti da lì durante la generazione del CSV.

        desc_section = ttk.LabelFrame(main_frame, text="Selezione Descrizioni", padding=15)
        # Spostata su riga 2 perché i dati comuni ora sono nella scheda principale
        desc_section.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            desc_section,
            text="Seleziona quali tipologie di schede includere nel CSV Reg:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        checkbox_container = ttk.Frame(desc_section)
        checkbox_container.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=5)

        desc_section.grid_rowconfigure(1, weight=1)

        checkbox_canvas = tk.Canvas(checkbox_container, height=200, highlightthickness=0)
        checkbox_scrollbar = ttk.Scrollbar(checkbox_container, orient="vertical", command=checkbox_canvas.yview)
        self.checkbox_frame = ttk.Frame(checkbox_canvas)

        self.checkbox_frame.bind(
            "<Configure>",
            lambda _: checkbox_canvas.configure(scrollregion=checkbox_canvas.bbox("all"))
        )

        checkbox_canvas_window = checkbox_canvas.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        checkbox_canvas.configure(yscrollcommand=checkbox_scrollbar.set)

        def on_checkbox_canvas_configure(event):
            checkbox_canvas.itemconfig(checkbox_canvas_window, width=event.width)

        checkbox_canvas.bind("<Configure>", on_checkbox_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(checkbox_canvas)

        checkbox_canvas.pack(side="left", fill="both", expand=True)
        checkbox_scrollbar.pack(side="right", fill="y")

        buttons_frame = ttk.Frame(desc_section)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)

        ttk.Button(
            buttons_frame,
            text="Seleziona Tutto",
            command=self.select_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            buttons_frame,
            text="Deseleziona Tutto",
            command=self.deselect_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        desc_section.columnconfigure(0, weight=1)
        desc_section.columnconfigure(1, weight=1)

        output_section = ttk.LabelFrame(main_frame, text="File di Output", padding=15)
        output_section.grid(row=4, column=0, columnspan=3, sticky="ew", pady=10, padx=10)

        ttk.Label(
            output_section,
            text="File Output CSV Reg:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.output_label = ttk.Label(
            output_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.output_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            output_section,
            text="Scegli Percorso Output",
            command=self.select_output_file
        ).grid(row=1, column=1, padx=10)

        output_section.columnconfigure(0, weight=1)
        output_section.columnconfigure(1, weight=0)

        generation_frame = ttk.Frame(main_frame)
        generation_frame.grid(row=5, column=0, columnspan=3, pady=20)

        self.generate_button = ttk.Button(
            generation_frame,
            text="GENERA CSV DI REGISTRAZIONE",
            command=self.generate_csvreg,
            style="Accent.TButton"
        )
        self.generate_button.pack(pady=10, ipadx=20, ipady=10)
        self.generate_button.state(['disabled'])

        self.progress = ttk.Progressbar(generation_frame, mode='indeterminate', length=400)
        self.progress.pack(pady=5)

        self.status_label = ttk.Label(generation_frame, text="", foreground="blue", font=('Arial', 10))
        self.status_label.pack(pady=5)

        # Aggiungi spacer vuoto per push del contenuto in alto
        spacer = ttk.Frame(main_frame)
        spacer.grid(row=6, column=0, columnspan=3, sticky="nsew")

    # I campi dati comuni per CSV Reg sono stati spostati nella scheda principale
    # sotto la sezione 'Dati Generali' per centralizzare l'inserimento.

    def select_input_file(self):
        """Seleziona il file Excel di input"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.csv_reg_input_file = Path(filename)
            self.input_label.config(
                text=f"{self.app_context.csv_reg_input_file.name}",
                foreground="green"
            )

            try:
                descriptions = DataProcessor.extract_unique_descriptions(
                    self.app_context.csv_reg_input_file
                )
                self.load_descriptions(descriptions)
                self.status_label.config(
                    text=f"Trovate {len(descriptions)} descrizioni",
                    foreground="green"
                )
            except Exception as e:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )
                self.status_label.config(
                    text="Errore nel caricamento delle descrizioni",
                    foreground="red"
                )

            self.update_button_state()

    def select_output_file(self):
        """Seleziona il percorso di output per CSV Reg"""
        if hasattr(self.app_context, 'csv_reg_input_file') and self.app_context.csv_reg_input_file:
            initialdir = Path(self.app_context.csv_reg_input_file).parent
            initialfile = "CSV_Registrazione.xlsx"
        elif hasattr(self.app_context, 'CSV_Registrazione_file') and self.app_context.CSV_Registrazione_file:
            initialdir = Path(self.app_context.CSV_Registrazione_file).parent
            initialfile = Path(self.app_context.CSV_Registrazione_file).name
        else:
            initialdir = Path.home()
            initialfile = "CSV_Registrazione.xlsx"

        filename = filedialog.asksaveasfilename(
            title="Scegli percorso e nome file CSV Reg",
            initialdir=initialdir,
            initialfile=initialfile,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.CSV_Registrazione_file = Path(filename)
            self.output_label.config(
                text=f"{self.app_context.CSV_Registrazione_file.name}",
                foreground="green"
            )
            self.update_button_state()

    def select_all_descriptions(self):
        """Seleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(True)

    def deselect_all_descriptions(self):
        """Deseleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(False)

    def load_descriptions(self, descriptions):
        """Carica le descrizioni disponibili e crea le checkbox"""
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()
        self.description_checkboxes.clear()

        for idx, desc in enumerate(descriptions):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.checkbox_frame, text=desc, variable=var)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            self.description_checkboxes[desc] = var

    def update_button_state(self):
        """Abilita il bottone GENERA CSV REG solo se input e output sono impostati"""
        has_input = (
            hasattr(self.app_context, 'csv_reg_input_file') and
            self.app_context.csv_reg_input_file is not None
        )
        has_output = (
            hasattr(self.app_context, 'CSV_Registrazione_file') and
            self.app_context.CSV_Registrazione_file is not None
        )

        if has_input and has_output:
            self.generate_button.state(['!disabled'])
        else:
            self.generate_button.state(['disabled'])

    def load_shared_input_file(self, file_path):
        """Carica il file di input condiviso."""
        self.app_context.csv_reg_input_file = Path(file_path)
        self.input_label.config(
            text=f"{self.app_context.csv_reg_input_file.name} (File Condiviso)",
            foreground="blue"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.csv_reg_input_file
            )
            self.load_descriptions(descriptions)
            self.status_label.config(
                text=f"File condiviso caricato - {len(descriptions)} descrizioni",
                foreground="blue"
            )
        except Exception as e:
            self.status_label.config(
                text="Errore nel caricamento del file condiviso",
                foreground="red"
            )

        self.update_button_state()

    def load_from_main_tab(self, generated_file_path, silent=False):
        """Carica automaticamente il file generato dalla scheda principale."""
        self.app_context.csv_reg_input_file = Path(generated_file_path)
        self.input_label.config(
            text=f"{self.app_context.csv_reg_input_file.name}",
            foreground="blue" if not silent else "green"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.csv_reg_input_file
            )
            self.load_descriptions(descriptions)
            if not silent:
                self.status_label.config(
                    text=f"File caricato automaticamente - {len(descriptions)} descrizioni",
                    foreground="blue"
                )
            else:
                self.status_label.config(
                    text=f"File pronto - {len(descriptions)} descrizioni",
                    foreground="green"
                )
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )

        self.update_button_state()

    def generate_csvreg(self):
        """Genera il file CSV_Reg filtrando per le descrizioni selezionate"""
        try:
            selected_descriptions = [
                desc for desc, var in self.description_checkboxes.items() if var.get()
            ]

            if not selected_descriptions:
                messagebox.showwarning("Attenzione", "Seleziona almeno una descrizione!")
                return

            extra_fields = {}
            # Preleva i campi comuni dalla scheda principale (Generazione Documento)
            if hasattr(self.app_context, 'dati_generali_extra_entries'):
                for field_name, entry in self.app_context.dati_generali_extra_entries.items():
                    value = entry.get().strip()
                    extra_fields[field_name] = value
            else:
                # Fallback: se per qualche motivo la vecchia struttura è ancora presente
                if hasattr(self, 'extra_fields_entries'):
                    for field_name, entry in self.extra_fields_entries.items():
                        value = entry.get().strip()
                        extra_fields[field_name] = value

            # Aggiungi esplicitamente le Bolla (i nomi sono quelli attesi da DataProcessor)
            try:
                bolla_prod = self.app_context.entry_bolla_produzione.get().strip()
            except Exception:
                bolla_prod = ""
            try:
                bolla_vend = self.app_context.entry_bolla_vendita.get().strip()
            except Exception:
                bolla_vend = ""

            extra_fields["Bolla Produzione"] = bolla_prod
            extra_fields["Bolla Vendita Techrail"] = bolla_vend

            self.status_label.config(text="Generazione CSV Reg in corso...", foreground="blue")
            self.progress.start()
            self.frame.update()

            rows_count = DataProcessor.generate_csv_reg(
                self.app_context.csv_reg_input_file,
                self.app_context.CSV_Registrazione_file,
                selected_descriptions,
                extra_fields
            )

            self.progress.stop()
            self.status_label.config(
                text=f"CSV Reg generato! Righe: {rows_count}",
                foreground="green"
            )

            messagebox.showinfo(
                "Successo",
                f"File CSV Reg generato con successo!\n\n"
                f"File output:\n{self.app_context.CSV_Registrazione_file}\n\n"
                f"Righe generate: {rows_count}\n"
                f"Descrizioni incluse: {len(selected_descriptions)}"
            )

        except ValueError as ve:
            self.progress.stop()
            self.status_label.config(text="Nessuna riga da generare", foreground="red")
            messagebox.showwarning("Attenzione", str(ve))
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Errore durante la generazione", foreground="red")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n\n{str(e)}")


# ============================================================================
# TAB IMPORT GESTIONALE
# ============================================================================

class ImportGestionaleTab:
    """Gestisce la scheda Import Gestionale"""

    def __init__(self, parent_notebook, app_context):
        self.notebook = parent_notebook
        self.app_context = app_context
        self.description_checkboxes = {}

        self.frame = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame, text="Import Gestionale")

        self.create_widgets()

    def create_widgets(self):
        """Crea il contenuto della scheda Import Gestionale"""
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configura il grid per permettere l'espansione verticale
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0, bg="white")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding=20)

        def update_scrollregion(event=None):
            # Aggiorna la scrollregion basandosi sul contenuto effettivo
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Reset alla posizione iniziale se necessario
            if canvas.yview()[0] < 0:
                canvas.yview_moveto(0)

        main_frame.bind("<Configure>", update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            # Aggiorna anche la scrollregion quando il canvas viene ridimensionato
            update_scrollregion()

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=1)

        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_rowconfigure(2, weight=0)
        main_frame.grid_rowconfigure(3, weight=0)
        main_frame.grid_rowconfigure(4, weight=0)
        main_frame.grid_rowconfigure(5, weight=0)
        main_frame.grid_rowconfigure(6, weight=1)

        #ttk.Label(
        #    main_frame,
        #    text="Import Gestionale",
        #    font=('Arial', 16, 'bold')
        #).grid(row=0, column=0, columnspan=3, pady=20)

        input_section = ttk.LabelFrame(main_frame, text="File di Input", padding=15)
        input_section.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            input_section,
            text="Seleziona il file Excel da processare:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.input_label = ttk.Label(
            input_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.input_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            input_section,
            text="Scegli File Input",
            command=self.select_input_file
        ).grid(row=1, column=1, padx=10)

        input_section.columnconfigure(0, weight=1)
        input_section.columnconfigure(1, weight=0)

        # La sezione "Dati Comuni" è stata spostata nella scheda principale
        # (Generazione Documento) nella sezione "Dati Generali".
        # I campi saranno letti da lì durante la generazione del file.

        desc_section = ttk.LabelFrame(main_frame, text="Selezione Descrizioni", padding=15)
        # Spostata su riga 2 perché i dati comuni ora sono nella scheda principale
        desc_section.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            desc_section,
            text="Seleziona quali tipologie di schede includere:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        checkbox_container = ttk.Frame(desc_section)
        checkbox_container.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=5)

        desc_section.grid_rowconfigure(1, weight=1)

        checkbox_canvas = tk.Canvas(checkbox_container, height=200, highlightthickness=0)
        checkbox_scrollbar = ttk.Scrollbar(checkbox_container, orient="vertical", command=checkbox_canvas.yview)
        self.checkbox_frame = ttk.Frame(checkbox_canvas)

        self.checkbox_frame.bind(
            "<Configure>",
            lambda _: checkbox_canvas.configure(scrollregion=checkbox_canvas.bbox("all"))
        )

        checkbox_canvas_window = checkbox_canvas.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        checkbox_canvas.configure(yscrollcommand=checkbox_scrollbar.set)

        def on_checkbox_canvas_configure(event):
            checkbox_canvas.itemconfig(checkbox_canvas_window, width=event.width)

        checkbox_canvas.bind("<Configure>", on_checkbox_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(checkbox_canvas)

        checkbox_canvas.pack(side="left", fill="both", expand=True)
        checkbox_scrollbar.pack(side="right", fill="y")

        buttons_frame = ttk.Frame(desc_section)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)

        ttk.Button(
            buttons_frame,
            text="Seleziona Tutto",
            command=self.select_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            buttons_frame,
            text="Deseleziona Tutto",
            command=self.deselect_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        desc_section.columnconfigure(0, weight=1)
        desc_section.columnconfigure(1, weight=1)

        output_section = ttk.LabelFrame(main_frame, text="File di Output", padding=15)
        output_section.grid(row=4, column=0, columnspan=3, sticky="ew", pady=10, padx=10)

        ttk.Label(
            output_section,
            text="File Output:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.output_label = ttk.Label(
            output_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.output_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            output_section,
            text="Scegli Percorso Output",
            command=self.select_output_file
        ).grid(row=1, column=1, padx=10)

        output_section.columnconfigure(0, weight=1)
        output_section.columnconfigure(1, weight=0)

        generation_frame = ttk.Frame(main_frame)
        generation_frame.grid(row=5, column=0, columnspan=3, pady=20)

        self.generate_button = ttk.Button(
            generation_frame,
            text="GENERA IMPORT GESTIONALE",
            command=self.generate_import_gestionale,
            style="Accent.TButton"
        )
        self.generate_button.pack(pady=10, ipadx=20, ipady=10)
        self.generate_button.state(['disabled'])

        self.progress = ttk.Progressbar(generation_frame, mode='indeterminate', length=400)
        self.progress.pack(pady=5)

        self.status_label = ttk.Label(generation_frame, text="", foreground="blue", font=('Arial', 10))
        self.status_label.pack(pady=5)

        # Aggiungi spacer vuoto per push del contenuto in alto
        spacer = ttk.Frame(main_frame)
        spacer.grid(row=6, column=0, columnspan=3, sticky="nsew")

    def select_input_file(self):
        """Seleziona il file Excel di input"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.import_gestionale_input_file = Path(filename)
            self.input_label.config(
                text=f"{self.app_context.import_gestionale_input_file.name}",
                foreground="green"
            )

            try:
                descriptions = DataProcessor.extract_unique_descriptions(
                    self.app_context.import_gestionale_input_file
                )
                self.load_descriptions(descriptions)
                self.status_label.config(
                    text=f"Trovate {len(descriptions)} descrizioni",
                    foreground="green"
                )
            except Exception as e:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )
                self.status_label.config(
                    text="Errore nel caricamento delle descrizioni",
                    foreground="red"
                )

            self.update_button_state()

    def select_output_file(self):
        """Seleziona il percorso di output"""
        if hasattr(self.app_context, 'import_gestionale_input_file') and self.app_context.import_gestionale_input_file:
            initialdir = Path(self.app_context.import_gestionale_input_file).parent
            initialfile = "Import_Gestionale.xlsx"
        elif hasattr(self.app_context, 'Import_Gestionale_file') and self.app_context.Import_Gestionale_file:
            initialdir = Path(self.app_context.Import_Gestionale_file).parent
            initialfile = Path(self.app_context.Import_Gestionale_file).name
        else:
            initialdir = Path.home()
            initialfile = "Import_Gestionale.xlsx"

        filename = filedialog.asksaveasfilename(
            title="Scegli percorso e nome file",
            initialdir=initialdir,
            initialfile=initialfile,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.Import_Gestionale_file = Path(filename)
            self.output_label.config(
                text=f"{self.app_context.Import_Gestionale_file.name}",
                foreground="green"
            )
            self.update_button_state()

    def select_all_descriptions(self):
        """Seleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(True)

    def deselect_all_descriptions(self):
        """Deseleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(False)

    def load_descriptions(self, descriptions):
        """Carica le descrizioni disponibili e crea le checkbox"""
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()
        self.description_checkboxes.clear()

        for idx, desc in enumerate(descriptions):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.checkbox_frame, text=desc, variable=var)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            self.description_checkboxes[desc] = var

    def update_button_state(self):
        """Abilita il bottone GENERA solo se input e output sono impostati"""
        has_input = (
            hasattr(self.app_context, 'import_gestionale_input_file') and
            self.app_context.import_gestionale_input_file is not None
        )
        has_output = (
            hasattr(self.app_context, 'Import_Gestionale_file') and
            self.app_context.Import_Gestionale_file is not None
        )

        if has_input and has_output:
            self.generate_button.state(['!disabled'])
        else:
            self.generate_button.state(['disabled'])

    def load_shared_input_file(self, file_path):
        """Carica il file di input condiviso."""
        self.app_context.import_gestionale_input_file = Path(file_path)
        self.input_label.config(
            text=f"{self.app_context.import_gestionale_input_file.name} (File Condiviso)",
            foreground="blue"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.import_gestionale_input_file
            )
            self.load_descriptions(descriptions)
            self.status_label.config(
                text=f"File condiviso caricato - {len(descriptions)} descrizioni",
                foreground="blue"
            )
        except Exception as e:
            self.status_label.config(
                text="Errore nel caricamento del file condiviso",
                foreground="red"
            )

        self.update_button_state()

    def load_from_main_tab(self, generated_file_path, silent=False):
        """Carica automaticamente il file generato dalla scheda principale."""
        self.app_context.import_gestionale_input_file = Path(generated_file_path)
        self.input_label.config(
            text=f"{self.app_context.import_gestionale_input_file.name}",
            foreground="blue" if not silent else "green"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.import_gestionale_input_file
            )
            self.load_descriptions(descriptions)
            if not silent:
                self.status_label.config(
                    text=f"File caricato automaticamente - {len(descriptions)} descrizioni",
                    foreground="blue"
                )
            else:
                self.status_label.config(
                    text=f"File pronto - {len(descriptions)} descrizioni",
                    foreground="green"
                )
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )

        self.update_button_state()

    def generate_import_gestionale(self):
        """Genera il file Import Gestionale filtrando per le descrizioni selezionate"""
        try:
            selected_descriptions = [
                desc for desc, var in self.description_checkboxes.items() if var.get()
            ]

            if not selected_descriptions:
                messagebox.showwarning("Attenzione", "Seleziona almeno una descrizione!")
                return

            extra_fields = {}
            # Preleva i campi comuni dalla scheda principale (Generazione Documento)
            if hasattr(self.app_context, 'dati_generali_extra_entries'):
                for field_name, entry in self.app_context.dati_generali_extra_entries.items():
                    value = entry.get().strip()
                    extra_fields[field_name] = value
            else:
                # Fallback: se per qualche motivo la vecchia struttura è ancora presente
                if hasattr(self, 'extra_fields_entries'):
                    for field_name, entry in self.extra_fields_entries.items():
                        value = entry.get().strip()
                        extra_fields[field_name] = value

            # Aggiungi esplicitamente le Bolla (i nomi sono quelli attesi da DataProcessor)
            try:
                bolla_prod = self.app_context.entry_bolla_produzione.get().strip()
            except Exception:
                bolla_prod = ""
            try:
                bolla_vend = self.app_context.entry_bolla_vendita.get().strip()
            except Exception:
                bolla_vend = ""

            extra_fields["Bolla Produzione"] = bolla_prod
            extra_fields["Bolla Vendita Techrail"] = bolla_vend

            self.status_label.config(text="Generazione Import Gestionale in corso...", foreground="blue")
            self.progress.start()
            self.frame.update()

            rows_count = DataProcessor.generate_import_gestionale(
                self.app_context.import_gestionale_input_file,
                self.app_context.Import_Gestionale_file,
                selected_descriptions,
                extra_fields
            )

            self.progress.stop()
            self.status_label.config(
                text=f"Import Gestionale generato! Righe: {rows_count}",
                foreground="green"
            )

            messagebox.showinfo(
                "Successo",
                f"File Import Gestionale generato con successo!\n\n"
                f"File output:\n{self.app_context.Import_Gestionale_file}\n\n"
                f"Righe generate: {rows_count}\n"
                f"Descrizioni incluse: {len(selected_descriptions)}"
            )

        except ValueError as ve:
            self.progress.stop()
            self.status_label.config(text="Nessuna riga da generare", foreground="red")
            messagebox.showwarning("Attenzione", str(ve))
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Errore durante la generazione", foreground="red")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n\n{str(e)}")


class EtichetteBoxTab:
    """Gestisce la scheda Etichette Bus"""

    def __init__(self, parent_notebook, app_context):
        self.notebook = parent_notebook
        self.app_context = app_context
        self.description_checkboxes = {}

        self.frame = ttk.Frame(parent_notebook)
        # Nota: la scheda 'Etichette Bus' non viene aggiunta automaticamente al notebook
        # per rimuoverla dall'interfaccia principale. Se si desidera ripristinarla,
        # decommentare la riga seguente.
        # parent_notebook.add(self.frame, text="Etichette Bus")

        self.create_widgets()

    def create_widgets(self):
        """Crea il contenuto della scheda Etichette Bus"""
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configura il grid per permettere l'espansione verticale
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding=20)

        main_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=1)

        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(3, weight=0)
        main_frame.grid_rowconfigure(4, weight=0)

        ttk.Label(
            main_frame,
            text="Etichette Bus - Filtro per Descrizione",
            font=('Arial', 16, 'bold')
        ).grid(row=0, column=0, columnspan=3, pady=20)

        input_section = ttk.LabelFrame(main_frame, text="File di Input", padding=15)
        input_section.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            input_section,
            text="Seleziona il file Excel da processare:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.input_label = ttk.Label(
            input_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.input_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            input_section,
            text="Scegli File Input",
            command=self.select_input_file
        ).grid(row=1, column=1, padx=10)

        input_section.columnconfigure(0, weight=1)
        input_section.columnconfigure(1, weight=0)

        desc_section = ttk.LabelFrame(main_frame, text="Selezione Descrizioni", padding=15)
        desc_section.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            desc_section,
            text="Seleziona quali tipologie di schede includere nelle etichette:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        checkbox_container = ttk.Frame(desc_section)
        checkbox_container.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=5)

        desc_section.grid_rowconfigure(1, weight=1)

        checkbox_canvas = tk.Canvas(checkbox_container, height=200, highlightthickness=0)
        checkbox_scrollbar = ttk.Scrollbar(checkbox_container, orient="vertical", command=checkbox_canvas.yview)
        self.checkbox_frame = ttk.Frame(checkbox_canvas)

        self.checkbox_frame.bind(
            "<Configure>",
            lambda _: checkbox_canvas.configure(scrollregion=checkbox_canvas.bbox("all"))
        )

        checkbox_canvas_window = checkbox_canvas.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        checkbox_canvas.configure(yscrollcommand=checkbox_scrollbar.set)

        def on_checkbox_canvas_configure(event):
            checkbox_canvas.itemconfig(checkbox_canvas_window, width=event.width)

        checkbox_canvas.bind("<Configure>", on_checkbox_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(checkbox_canvas)

        checkbox_canvas.pack(side="left", fill="both", expand=True)
        checkbox_scrollbar.pack(side="right", fill="y")

        buttons_frame = ttk.Frame(desc_section)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)

        ttk.Button(
            buttons_frame,
            text="Seleziona Tutto",
            command=self.select_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            buttons_frame,
            text="Deseleziona Tutto",
            command=self.deselect_all_descriptions
        ).pack(side=tk.LEFT, padx=5)

        desc_section.columnconfigure(0, weight=1)
        desc_section.columnconfigure(1, weight=1)
        desc_section.columnconfigure(2, weight=1)
        desc_section.rowconfigure(1, weight=1)

        output_section = ttk.LabelFrame(main_frame, text="File di Output", padding=15)
        output_section.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=10, padx=10)

        ttk.Label(
            output_section,
            text="File Output Etichette Bus:",
            font=('Arial', 10)
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        self.output_label = ttk.Label(
            output_section,
            text="Nessun file selezionato",
            foreground="gray",
            font=('Arial', 10)
        )
        self.output_label.grid(row=1, column=0, sticky="ew", padx=5)

        ttk.Button(
            output_section,
            text="Scegli Percorso Output",
            command=self.select_output_file
        ).grid(row=1, column=1, padx=10)

        output_section.columnconfigure(0, weight=1)
        output_section.columnconfigure(1, weight=0)

        generation_frame = ttk.Frame(main_frame)
        generation_frame.grid(row=4, column=0, columnspan=3, pady=20)

        self.generate_button = ttk.Button(
            generation_frame,
            text="GENERA ETICHETTE BUS",
            command=self.generate_etichettebox,
            style="Accent.TButton"
        )
        self.generate_button.pack(pady=10, ipadx=20, ipady=10)
        self.generate_button.state(['disabled'])

        self.progress = ttk.Progressbar(generation_frame, mode='indeterminate', length=400)
        self.progress.pack(pady=5)

        self.status_label = ttk.Label(generation_frame, text="", foreground="blue", font=('Arial', 10))
        self.status_label.pack(pady=5)

        main_frame.rowconfigure(2, weight=1)

    def select_input_file(self):
        """Seleziona il file Excel di input"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.etichettebox_input_file = Path(filename)
            self.input_label.config(
                text=f"{self.app_context.etichettebox_input_file.name}",
                foreground="green"
            )

            try:
                descriptions = DataProcessor.extract_unique_descriptions(
                    self.app_context.etichettebox_input_file
                )
                self.load_descriptions(descriptions)
                self.status_label.config(
                    text=f"Trovate {len(descriptions)} descrizioni",
                    foreground="green"
                )
            except Exception as e:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )
                self.status_label.config(
                    text="Errore nel caricamento delle descrizioni",
                    foreground="red"
                )

            self.update_button_state()

    def select_output_file(self):
        """Seleziona il percorso di output per Etichette Bus"""
        if hasattr(self.app_context, 'etichettebox_input_file') and self.app_context.etichettebox_input_file:
            initialdir = Path(self.app_context.etichettebox_input_file).parent
            initialfile = "EtichetteBOX_Output.xlsx"
        elif hasattr(self.app_context, 'etichettebox_output_file') and self.app_context.etichettebox_output_file:
            initialdir = Path(self.app_context.etichettebox_output_file).parent
            initialfile = Path(self.app_context.etichettebox_output_file).name
        else:
            initialdir = Path.home()
            initialfile = "EtichetteBOX_Output.xlsx"

        filename = filedialog.asksaveasfilename(
            title="Scegli percorso e nome file Etichette Bus",
            initialdir=initialdir,
            initialfile=initialfile,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if filename:
            self.app_context.etichettebox_output_file = Path(filename)
            self.output_label.config(
                text=f"{self.app_context.etichettebox_output_file.name}",
                foreground="green"
            )
            self.update_button_state()

    def select_all_descriptions(self):
        """Seleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(True)

    def deselect_all_descriptions(self):
        """Deseleziona tutte le descrizioni"""
        for var in self.description_checkboxes.values():
            var.set(False)

    def load_descriptions(self, descriptions):
        """Carica le descrizioni disponibili e crea le checkbox"""
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()
        self.description_checkboxes.clear()

        for idx, desc in enumerate(descriptions):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.checkbox_frame, text=desc, variable=var)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            self.description_checkboxes[desc] = var

    def update_button_state(self):
        """Abilita il bottone GENERA solo se input e output sono impostati"""
        has_input = (
            hasattr(self.app_context, 'etichettebox_input_file') and
            self.app_context.etichettebox_input_file is not None
        )
        has_output = (
            hasattr(self.app_context, 'etichettebox_output_file') and
            self.app_context.etichettebox_output_file is not None
        )

        if has_input and has_output:
            self.generate_button.state(['!disabled'])
        else:
            self.generate_button.state(['disabled'])

    def load_shared_input_file(self, file_path):
        """Carica il file di input condiviso."""
        self.app_context.etichettebox_input_file = Path(file_path)
        self.input_label.config(
            text=f"{self.app_context.etichettebox_input_file.name} (File Condiviso)",
            foreground="blue"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.etichettebox_input_file
            )
            self.load_descriptions(descriptions)
            self.status_label.config(
                text=f"File condiviso caricato - {len(descriptions)} descrizioni",
                foreground="blue"
            )
        except Exception as e:
            self.status_label.config(
                text="Errore nel caricamento del file condiviso",
                foreground="red"
            )

        self.update_button_state()

    def load_from_main_tab(self, generated_file_path, silent=False):
        """Carica automaticamente il file generato dalla scheda principale."""
        self.app_context.etichettebox_input_file = Path(generated_file_path)
        self.input_label.config(
            text=f"{self.app_context.etichettebox_input_file.name}",
            foreground="blue" if not silent else "green"
        )

        try:
            descriptions = DataProcessor.extract_unique_descriptions(
                self.app_context.etichettebox_input_file
            )
            self.load_descriptions(descriptions)
            if not silent:
                self.status_label.config(
                    text=f"File caricato automaticamente - {len(descriptions)} descrizioni",
                    foreground="blue"
                )
            else:
                self.status_label.config(
                    text=f"File pronto - {len(descriptions)} descrizioni",
                    foreground="green"
                )
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "Errore",
                    f"Errore nel caricamento delle descrizioni:\n{str(e)}"
                )

        self.update_button_state()

    def generate_etichettebox(self):
        """Genera il file EtichetteBOX.xlsx"""
        try:
            selected_descriptions = [
                desc for desc, var in self.description_checkboxes.items() if var.get()
            ]

            if not selected_descriptions:
                messagebox.showwarning("Attenzione", "Seleziona almeno una descrizione!")
                return

            self.status_label.config(text="Generazione Etichette Bus in corso...", foreground="blue")
            self.progress.start()
            self.frame.update()

            rows_count = DataProcessor.generate_etichettebox_excel(
                self.app_context.etichettebox_input_file,
                self.app_context.etichettebox_output_file,
                selected_descriptions
            )

            self.progress.stop()
            self.status_label.config(
                text=f"Etichette Bus generate! Righe: {rows_count}",
                foreground="green"
            )

            messagebox.showinfo(
                "Successo",
                f"File Etichette Bus generato con successo!\n\n"
                f"File output:\n{self.app_context.etichettebox_output_file}\n\n"
                f"Righe generate: {rows_count}\n"
                f"Descrizioni incluse: {len(selected_descriptions)}"
            )

        except ValueError as ve:
            self.progress.stop()
            self.status_label.config(text="Nessun bus da generare", foreground="red")
            messagebox.showwarning("Attenzione", str(ve))
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Errore durante la generazione", foreground="red")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n\n{str(e)}")


# ============================================================================
# TAB Etichette Naz
# ============================================================================

class EtichettePDFTab:
    """Gestisce la scheda EtichettePDF"""

    def __init__(self, notebook, app_context):
        self.notebook = notebook
        self.app_context = app_context
        self.tipo_scheda_checkboxes = {}
        # mappa checkbox_label -> (descrizione, prefisso_tipo) usata dalla scheda Word
        self._word_label_mapping: Dict[str, tuple] = {}
        self.filter_enabled_var = tk.BooleanVar(value=True)

        self.frame = ttk.Frame(self.notebook)
        self.notebook.add(self.frame, text="Etichette Naz")

        self.create_widgets()

    def create_widgets(self):
        """Crea il contenuto della scheda EtichettePDF"""
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configura il grid per permettere l'espansione verticale
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0, bg="white")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        def update_scrollregion(event=None):
            # Aggiorna la scrollregion basandosi sul contenuto effettivo
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Reset alla posizione iniziale se necessario
            if canvas.yview()[0] < 0:
                canvas.yview_moveto(0)

        self.scrollable_frame.bind("<Configure>", update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            # Aggiorna anche la scrollregion quando il canvas viene ridimensionato
            update_scrollregion()

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        self.scrollable_frame.grid_columnconfigure(2, weight=1)

        self.scrollable_frame.grid_rowconfigure(1, weight=0)
        self.scrollable_frame.grid_rowconfigure(2, weight=0)
        self.scrollable_frame.grid_rowconfigure(3, weight=0)
        self.scrollable_frame.grid_rowconfigure(4, weight=0)
        self.scrollable_frame.grid_rowconfigure(5, weight=0)
        self.scrollable_frame.grid_rowconfigure(17, weight=1)

        #ttk.Label(self.scrollable_frame, text="Generazione Etichette Naz",
        #          font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=3, pady=10)

        self._crea_sezione_input_file(row_start=1)

        ttk.Label(self.scrollable_frame, text="Seleziona quali tipi di schede includere:",
                  font=('Arial', 10)).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5, padx=10)

        checkbox_container = ttk.Frame(self.scrollable_frame)
        checkbox_container.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=5, padx=10)

        canvas_cb = tk.Canvas(checkbox_container, height=150, highlightthickness=0)
        scrollbar_cb = ttk.Scrollbar(checkbox_container, orient="vertical", command=canvas_cb.yview)
        self.checkbox_frame = ttk.Frame(canvas_cb)

        self.checkbox_frame.bind(
            "<Configure>",
            lambda _: canvas_cb.configure(scrollregion=canvas_cb.bbox("all"))
        )

        checkbox_canvas_window = canvas_cb.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        canvas_cb.configure(yscrollcommand=scrollbar_cb.set)

        def on_checkbox_canvas_configure(event):
            canvas_cb.itemconfig(checkbox_canvas_window, width=event.width)

        canvas_cb.bind("<Configure>", on_checkbox_canvas_configure)

        canvas_cb.pack(side="left", fill="both", expand=True)
        scrollbar_cb.pack(side="right", fill="y")

        buttons_frame = ttk.Frame(self.scrollable_frame)
        buttons_frame.grid(row=5, column=0, columnspan=3, pady=5)
        ttk.Button(buttons_frame, text="Seleziona Tutto",
                   command=self.select_all_tipo_scheda).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Deseleziona Tutto",
                   command=self.deselect_all_tipo_scheda).pack(side=tk.LEFT, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(self.scrollable_frame, text="Ripetizioni:",
                  font=('Arial', 10, 'bold')).grid(row=7, column=0, sticky=tk.W, pady=5, padx=10)

        self.repetitions_frame = ttk.Frame(self.scrollable_frame)
        self.repetitions_frame.grid(row=7, column=1, columnspan=2, sticky=tk.W, padx=10, pady=5)

        self.repetitions_var = tk.IntVar(value=1)
        self.repetitions_spinbox = ttk.Spinbox(
            self.repetitions_frame,
            from_=1,
            to=100,
            textvariable=self.repetitions_var,
            width=10
        )
        self.repetitions_spinbox.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.repetitions_frame, text="(numero di volte che tutti i dati vengono ripetuti)",
                  font=('Arial', 9), foreground="gray").pack(side=tk.LEFT, padx=5)

        ttk.Label(self.scrollable_frame, text="Coordinate prima entrata:",
                  font=('Arial', 10, 'bold')).grid(row=8, column=0, sticky=tk.W, pady=5, padx=10)

        self.coordinates_frame = ttk.Frame(self.scrollable_frame)
        self.coordinates_frame.grid(row=8, column=1, columnspan=2, sticky=tk.W, padx=10, pady=5)

        ttk.Label(self.coordinates_frame, text="Colonna:").pack(side=tk.LEFT, padx=5)
        self.start_column_var = tk.IntVar(value=1)
        ttk.Spinbox(
            self.coordinates_frame,
            from_=1,
            to=4,
            textvariable=self.start_column_var,
            width=5
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(self.coordinates_frame, text="Riga:").pack(side=tk.LEFT, padx=5)
        self.start_row_var = tk.IntVar(value=1)
        ttk.Spinbox(
            self.coordinates_frame,
            from_=1,
            to=21,
            textvariable=self.start_row_var,
            width=5
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(self.coordinates_frame, text="(foglio: 4 colonne x 21 righe)",
                  font=('Arial', 9), foreground="gray").pack(side=tk.LEFT, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(self.scrollable_frame, text="File Immagine Logo:",
                  font=('Arial', 10, 'bold')).grid(row=10, column=0, sticky=tk.W, pady=5, padx=10)

        self.image_label = ttk.Label(self.scrollable_frame, text=r"Resources\image.png (default)", foreground="gray")
        self.image_label.grid(row=10, column=0, columnspan=2, sticky=tk.W, padx=15)

        ttk.Button(self.scrollable_frame, text="Seleziona Immagine",
                   command=self.select_image_file).grid(row=10, column=2, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(self.scrollable_frame, text="File Output PDF:",
                  font=('Arial', 10, 'bold')).grid(row=12, column=0, sticky=tk.W, pady=5, padx=10)

        self.output_label = ttk.Label(self.scrollable_frame, text="Nessun file selezionato", foreground="gray")
        self.output_label.grid(row=13, column=0, columnspan=2, sticky=tk.W, padx=15)

        ttk.Button(self.scrollable_frame, text="Scegli Percorso",
                   command=self.select_output_file).grid(row=13, column=2, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.generate_button = ttk.Button(self.scrollable_frame, text="GENERA PDF ETICHETTE", command=self.generate_pdf)
        self.generate_button.grid(row=15, column=0, columnspan=3, pady=20)
        self.generate_button.state(['disabled'])

        self.progress = ttk.Progressbar(self.scrollable_frame, mode='indeterminate')
        self.progress.grid(row=16, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=10)

        self.status_label = ttk.Label(self.scrollable_frame, text="", foreground="blue")
        self.status_label.grid(row=17, column=0, columnspan=3, pady=5)

    def _crea_sezione_input_file(self, row_start: int):
        """Crea la sezione per la selezione del file di input"""
        frame = ttk.LabelFrame(
            self.scrollable_frame,
            text="File di Input",
            padding=10
        )
        frame.grid(row=row_start, column=0, columnspan=3, sticky="ew", pady=10, padx=10)

        ttk.Label(frame, text="File Excel di Input:", font=("Arial", 10)).grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.input_label = ttk.Label(frame, text="Nessun file selezionato", foreground="gray")
        self.input_label.grid(row=0, column=1, sticky="w", padx=10)

        ttk.Button(frame, text="Seleziona File",
                   command=self.select_input_file).grid(row=0, column=2, padx=5)

    def select_input_file(self):
        """Seleziona il file Excel di input"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input",
            filetypes=[
                ("File Excel", "*.xlsx *.xls"),
                ("XLSX files", "*.xlsx"),
                ("XLS files", "*.xls"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.app_context.input_file = filename
            self.input_label.config(text=Path(filename).name, foreground="green")
            self.update_button_state()
            self.load_tipo_scheda_from_file()

    def load_tipo_scheda_from_file(self):
        """Carica i CODE 12NC dal file di input"""
        try:
            if not self.app_context.input_file:
                return

            df = pd.read_excel(self.app_context.input_file)

            if "CODE 12NC" not in df.columns:
                messagebox.showwarning("Attenzione", "La colonna 'CODE 12NC' non trovata")
                return

            # Crea un dizionario con CODE 12NC come chiave e descrizione come valore
            code_desc_map = {}

            # Prova diversi nomi per la colonna descrizione
            desc_col = None
            for col_name in ["DESCRIZIONE", "Descrizione", "descrizione"]:
                if col_name in df.columns:
                    desc_col = col_name
                    break

            if desc_col:
                for _, row in df.iterrows():
                    code = str(row["CODE 12NC"]) if pd.notna(row["CODE 12NC"]) else None
                    desc = str(row[desc_col]) if pd.notna(row[desc_col]) else ""
                    if code and code not in code_desc_map:
                        code_desc_map[code] = desc
            else:
                # Se la colonna descrizione non esiste, usa solo i codici
                code_12nc = [str(t) for t in df["CODE 12NC"].unique() if pd.notna(t)]
                code_desc_map = {code: "" for code in code_12nc}

            # Ordina per CODE 12NC
            sorted_codes = sorted(code_desc_map.keys())
            sorted_code_desc_map = {code: code_desc_map[code] for code in sorted_codes}

            self.load_tipo_scheda(sorted_code_desc_map)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nel caricamento dei CODE 12NC:\n{str(e)}")

    def select_all_tipo_scheda(self):
        """Seleziona tutti i CODE 12NC"""
        for var in self.tipo_scheda_checkboxes.values():
            var.set(True)

    def deselect_all_tipo_scheda(self):
        """Deseleziona tutti i CODE 12NC"""
        for var in self.tipo_scheda_checkboxes.values():
            var.set(False)

    def load_tipo_scheda(self, code_desc_map):
        """Carica i CODE 12NC disponibili e crea le checkbox

        Args:
            code_desc_map: Dizionario {CODE_12NC: DESCRIZIONE} oppure lista di CODE_12NC
        """
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()
        self.tipo_scheda_checkboxes.clear()

        # Se è una lista (retrocompatibilità), converti in dizionario
        if isinstance(code_desc_map, list):
            code_desc_map = {code: "" for code in code_desc_map}

        for idx, (code, desc) in enumerate(code_desc_map.items()):
            var = tk.BooleanVar(value=False)
            # Mostra CODE 12NC - DESCRIZIONE se la descrizione esiste
            label_text = f"{code} - {desc}" if desc else code
            cb = ttk.Checkbutton(self.checkbox_frame, text=label_text, variable=var)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            self.tipo_scheda_checkboxes[code] = var
            cb.configure(state='normal')

    def select_image_file(self):
        """Seleziona il file immagine per il logo"""
        initialdir = Path.cwd()
        if self.app_context.input_file:
            initialdir = Path(self.app_context.input_file).parent

        filename = filedialog.askopenfilename(
            title="Seleziona immagine logo",
            initialdir=initialdir,
            filetypes=[
                ("Immagini", "*.png *.jpg *.jpeg"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.app_context.etichettepdf_image_file = filename
            self.image_label.config(text=Path(filename).name, foreground="green")
            self.update_button_state()

    def select_output_file(self):
        """Seleziona il percorso di output per il PDF"""
        if self.app_context.input_file:
            initialdir = Path(self.app_context.input_file).parent
            initialfile = "Etichette Naz.pdf"
        else:
            initialdir = Path.home()
            initialfile = "Etichette Naz.pdf"

        filename = filedialog.asksaveasfilename(
            title="Scegli percorso e nome file PDF",
            initialdir=initialdir,
            initialfile=initialfile,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.app_context.etichettepdf_output_file = Path(filename)
            self.output_label.config(text=f"{self.app_context.etichettepdf_output_file.name}", foreground="green")
            self.update_button_state()

    def update_button_state(self):
        """Abilita il bottone GENERA PDF solo se input e output sono impostati"""
        has_input = hasattr(self.app_context, 'input_file') and self.app_context.input_file is not None
        has_output = hasattr(self.app_context, 'etichettepdf_output_file') and self.app_context.etichettepdf_output_file is not None

        if has_input and has_output:
            self.generate_button.state(['!disabled'])
        else:
            self.generate_button.state(['disabled'])

    def load_shared_input_file(self, file_path):
        """Carica il file di input condiviso."""
        self.app_context.input_file = str(file_path)
        self.input_label.config(
            text=f"{Path(file_path).name} (File Condiviso)",
            foreground="blue"
        )
        self.update_button_state()
        self.load_tipo_scheda_from_file()
        self.status_label.config(
            text="File condiviso caricato",
            foreground="blue"
        )

    def load_from_main_tab(self, generated_file_path):
        """Carica automaticamente il file generato dalla scheda principale."""
        self.app_context.input_file = str(generated_file_path)
        self.input_label.config(
            text=Path(generated_file_path).name,
            foreground="green"
        )
        self.update_button_state()
        self.load_tipo_scheda_from_file()
        self.status_label.config(
            text="File pronto per la generazione",
            foreground="green"
        )

    def generate_pdf(self):
        """Genera il file PDF con le etichette"""
        try:
            image_path = self.app_context.etichettepdf_image_file if hasattr(self.app_context, 'etichettepdf_image_file') and self.app_context.etichettepdf_image_file else r"Resources\image.png"

            if not Path(image_path).exists():
                messagebox.showwarning(
                    "Attenzione",
                    f"File immagine non trovato: {image_path}\n\n"
                    "Seleziona un'immagine valida."
                )
                return

            self.status_label.config(text="Generazione PDF in corso...", foreground="blue")
            self.progress.start()
            self.app_context.root.update()

            filter_enabled = self.filter_enabled_var.get()
            selected_tipo_scheda = None

            if filter_enabled:
                selected_tipo_scheda = [tipo for tipo, var in self.tipo_scheda_checkboxes.items() if var.get()]

                if not selected_tipo_scheda:
                    messagebox.showwarning("Attenzione", "Seleziona almeno un CODE 12NC!")
                    self.progress.stop()
                    return

            repetitions = self.repetitions_var.get()
            start_column = self.start_column_var.get()
            start_row = self.start_row_var.get()

            label_count = PDFLabelGenerator.generate_pdf_labels(
                self.app_context.input_file,
                self.app_context.etichettepdf_output_file,
                image_path,
                filter_enabled,
                selected_tipo_scheda,
                repetitions,
                start_column,
                start_row
            )

            self.progress.stop()
            self.status_label.config(
                text=f"PDF generato! Etichette create: {label_count}",
                foreground="green"
            )

            if filter_enabled:
                msg = (f"File PDF generato con successo!\n\n"
                       f"File output:\n{self.app_context.etichettepdf_output_file}\n\n"
                       f"Etichette generate: {label_count}\n"
                       f"CODE 12NC inclusi: {len(selected_tipo_scheda)}")
            else:
                msg = (f"File PDF generato con successo!\n\n"
                       f"File output:\n{self.app_context.etichettepdf_output_file}\n\n"
                       f"Etichette generate: {label_count}")

            messagebox.showinfo("Successo", msg)

        except ValueError as ve:
            self.progress.stop()
            self.status_label.config(text="Nessuna etichetta da generare", foreground="red")
            messagebox.showwarning("Attenzione", str(ve))
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Errore durante la generazione", foreground="red")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n\n{str(e)}")


# ============================================================================
# TAB Etichette Interne
# ============================================================================

class EtichetteWordTab:
    """Gestisce la scheda Etichette Interne"""

    def __init__(self, notebook, app_context):
        self.notebook = notebook
        self.app_context = app_context
        self.tipo_scheda_checkboxes = {}
        # mappa checkbox_label -> (descrizione, prefisso_tipo) usata dalla scheda Word
        self._word_label_mapping: Dict[str, tuple] = {}
        self.filter_enabled_var = tk.BooleanVar(value=True)

        self.frame = ttk.Frame(self.notebook)
        self.notebook.add(self.frame, text="Etichette Interne")

        self.create_widgets()

    def create_widgets(self):
        """Crea il contenuto della scheda Etichette Interne"""
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configura il grid per permettere l'espansione verticale
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0, bg="white")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        def update_scrollregion(event=None):
            # Aggiorna la scrollregion basandosi sul contenuto effettivo
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Reset alla posizione iniziale se necessario
            if canvas.yview()[0] < 0:
                canvas.yview_moveto(0)

        self.scrollable_frame.bind("<Configure>", update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            # Aggiorna anche la scrollregion quando il canvas viene ridimensionato
            update_scrollregion()

        canvas.bind("<Configure>", on_canvas_configure)

        # Aggiungi supporto rotella del mouse
        bind_mousewheel_to_canvas(canvas)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        self.scrollable_frame.grid_columnconfigure(2, weight=1)

        self.scrollable_frame.grid_rowconfigure(1, weight=0)
        self.scrollable_frame.grid_rowconfigure(2, weight=0)
        self.scrollable_frame.grid_rowconfigure(3, weight=0)
        self.scrollable_frame.grid_rowconfigure(4, weight=0)
        self.scrollable_frame.grid_rowconfigure(5, weight=0)
        self.scrollable_frame.grid_rowconfigure(15, weight=1)

        #ttk.Label(self.scrollable_frame, text="Generazione Etichette Interne (A4 Portrate)",
        #          font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=3, pady=10)

        self._crea_sezione_input_file(row_start=1)

        ttk.Label(self.scrollable_frame, text="Seleziona quali tipi di schede includere:",
                  font=('Arial', 10)).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5, padx=10)

        checkbox_container = ttk.Frame(self.scrollable_frame)
        checkbox_container.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=5, padx=10)

        canvas_cb = tk.Canvas(checkbox_container, height=150, highlightthickness=0)
        scrollbar_cb = ttk.Scrollbar(checkbox_container, orient="vertical", command=canvas_cb.yview)
        self.checkbox_frame = ttk.Frame(canvas_cb)

        self.checkbox_frame.bind(
            "<Configure>",
            lambda _: canvas_cb.configure(scrollregion=canvas_cb.bbox("all"))
        )

        checkbox_canvas_window = canvas_cb.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        canvas_cb.configure(yscrollcommand=scrollbar_cb.set)

        def on_checkbox_canvas_configure(event):
            canvas_cb.itemconfig(checkbox_canvas_window, width=event.width)

        canvas_cb.bind("<Configure>", on_checkbox_canvas_configure)

        canvas_cb.pack(side="left", fill="both", expand=True)
        scrollbar_cb.pack(side="right", fill="y")

        buttons_frame = ttk.Frame(self.scrollable_frame)
        buttons_frame.grid(row=5, column=0, columnspan=3, pady=5)
        ttk.Button(buttons_frame, text="Seleziona Tutto",
                   command=self.select_all_tipo_scheda).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Deseleziona Tutto",
                   command=self.deselect_all_tipo_scheda).pack(side=tk.LEFT, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(self.scrollable_frame, text="Ripetizioni:",
              font=('Arial', 10, 'bold')).grid(row=7, column=0, sticky=tk.W, pady=5, padx=10)
        self.repetitions_var = tk.IntVar(value=1)
        self.repetitions_spinbox = ttk.Spinbox(self.scrollable_frame, from_=1, to=100, textvariable=self.repetitions_var, width=8)
        self.repetitions_spinbox.grid(row=7, column=1, sticky=tk.W, padx=5)
        ttk.Label(self.scrollable_frame, text="(numero di volte che tutti i dati vengono ripetuti)", font=('Arial', 9), foreground="gray").grid(row=7, column=2, sticky=tk.W)

        ttk.Label(self.scrollable_frame, text="Coordinate prima entrata:",
              font=('Arial', 10, 'bold')).grid(row=8, column=0, sticky=tk.W, pady=5, padx=10)
        coords_frame = ttk.Frame(self.scrollable_frame)
        coords_frame.grid(row=8, column=1, columnspan=2, sticky=tk.W, padx=5)
        ttk.Label(coords_frame, text="Colonna:").pack(side=tk.LEFT, padx=5)
        self.start_column_var = tk.IntVar(value=1)
        ttk.Spinbox(coords_frame, from_=1, to=8, textvariable=self.start_column_var, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(coords_frame, text="Riga:").pack(side=tk.LEFT, padx=5)
        self.start_row_var = tk.IntVar(value=1)
        ttk.Spinbox(coords_frame, from_=1, to=20, textvariable=self.start_row_var, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(coords_frame, text="(foglio: 8 colonne x 20 righe)", font=('Arial', 9), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Checkbox per etichette con sfondo nero
        self.add_black_labels_var = tk.BooleanVar(value=True)
        add_black_checkbox = ttk.Checkbutton(self.scrollable_frame,
                                            text="Aggiungi etichette con sfondo nero",
                                            variable=self.add_black_labels_var)
        add_black_checkbox.grid(row=8, column=2, sticky=tk.W, pady=5, padx=10)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(self.scrollable_frame, text="File Output Etichette Interne:",
              font=('Arial', 10, 'bold')).grid(row=10, column=0, sticky=tk.W, pady=5, padx=10)

        self.output_label = ttk.Label(self.scrollable_frame, text="Nessun file selezionato", foreground="gray")
        self.output_label.grid(row=11, column=0, columnspan=2, sticky=tk.W, padx=15)

        ttk.Button(self.scrollable_frame, text="Scegli Percorso",
               command=self.select_output_file).grid(row=11, column=2, padx=5)

        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=12, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.generate_button = ttk.Button(self.scrollable_frame, text="GENERA Etichette Interne", command=self.generate_etichetteword)
        self.generate_button.grid(row=13, column=0, columnspan=3, pady=20)
        self.generate_button.state(['disabled'])

        self.progress = ttk.Progressbar(self.scrollable_frame, mode='indeterminate')
        self.progress.grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=10)

        self.status_label = ttk.Label(self.scrollable_frame, text="", foreground="blue")
        self.status_label.grid(row=15, column=0, columnspan=3, pady=5)

    def _crea_sezione_input_file(self, row_start: int):
        """Crea la sezione per la selezione del file di input"""
        frame = ttk.LabelFrame(
            self.scrollable_frame,
            text="File di Input",
            padding=10
        )
        frame.grid(row=row_start, column=0, columnspan=3, sticky="ew", pady=10, padx=10)

        ttk.Label(frame, text="File Excel di Input:", font=("Arial", 10)).grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.input_label = ttk.Label(frame, text="Nessun file selezionato", foreground="gray")
        self.input_label.grid(row=0, column=1, sticky="w", padx=10)

        ttk.Button(frame, text="Seleziona File",
                   command=self.select_input_file).grid(row=0, column=2, padx=5)

    def select_input_file(self):
        """Seleziona il file Excel di input"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel di input",
            filetypes=[
                ("File Excel", "*.xlsx *.xls"),
                ("XLSX files", "*.xlsx"),
                ("XLS files", "*.xls"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.app_context.input_file = filename
            self.input_label.config(text=Path(filename).name, foreground="green")
            self.update_button_state()
            self.load_tipo_scheda_from_file()

    def load_tipo_scheda_from_file(self):
        """Carica la lista di componenti mostrata nella selezione.

        Mostriamo per ogni componente la descrizione e il prefisso del tipo scheda
        (es. "NOME COMPONENTE - SU"). Se il prefisso non è disponibile viene mostrata
        solo la descrizione.
        """
        try:
            if not self.app_context.input_file:
                return

            df = pd.read_excel(self.app_context.input_file)

            if "Descrizione" not in df.columns:
                messagebox.showwarning("Attenzione", "La colonna 'Descrizione' non trovata")
                return

            # Tipo Scheda può mancare per alcune righe; usiamo la prima occorrenza per descrizione
            descrizioni = [d for d in df["Descrizione"].unique() if pd.notna(d)]

            items = []
            self._word_label_mapping.clear()

            for descr in sorted(descrizioni):
                # Trova la prima riga con questa descrizione che abbia un Tipo Scheda
                sub = df[df["Descrizione"] == descr]
                tipo_val = None
                if "Tipo Scheda" in df.columns:
                    non_null = sub["Tipo Scheda"].dropna()
                    if not non_null.empty:
                        tipo_val = str(non_null.iloc[0]).strip()

                # Estrai prefisso rimuovendo cifre finali
                prefisso = ""
                if tipo_val:
                    prefisso = re.sub(r"\d+$", "", tipo_val).strip()

                if prefisso:
                    label = f"{descr} - {prefisso}"
                else:
                    label = f"{descr}"

                items.append(label)
                self._word_label_mapping[label] = (descr, prefisso)

            self.load_tipo_scheda(items)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nel caricamento dei dati:\n{str(e)}")

    def select_all_tipo_scheda(self):
        """Seleziona tutti i tipi scheda"""
        for var in self.tipo_scheda_checkboxes.values():
            var.set(True)

    def deselect_all_tipo_scheda(self):
        """Deseleziona tutti i tipi scheda"""
        for var in self.tipo_scheda_checkboxes.values():
            var.set(False)

    def select_output_file(self):
        """Seleziona il percorso di output"""
        if self.app_context.input_file:
            initialdir = Path(self.app_context.input_file).parent
            initialfile = "Etichette Interne.pdf"
        else:
            initialdir = Path.home()
            initialfile = "Etichette Interne.pdf"

        filename = filedialog.asksaveasfilename(
            title="Scegli percorso e nome file Etichette Naz",
            initialdir=initialdir,
            initialfile=initialfile,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.app_context.etichetteword_output_file = Path(filename)
            self.output_label.config(text=f"{self.app_context.etichetteword_output_file.name}", foreground="green")
            self.update_button_state()

    def update_button_state(self):
        """Abilita il bottone GENERA"""
        has_input = hasattr(self.app_context, 'input_file') and self.app_context.input_file is not None
        has_output = hasattr(self.app_context, 'etichetteword_output_file') and self.app_context.etichetteword_output_file is not None

        if has_input and has_output:
            self.generate_button.state(['!disabled'])
        else:
            self.generate_button.state(['disabled'])

    def load_shared_input_file(self, file_path):
        """Carica il file di input condiviso."""
        self.app_context.input_file = str(file_path)
        self.input_label.config(
            text=f"{Path(file_path).name} (File Condiviso)",
            foreground="blue"
        )
        self.update_button_state()
        self.load_tipo_scheda_from_file()
        self.status_label.config(
            text="File condiviso caricato",
            foreground="blue"
        )

    def load_from_main_tab(self, generated_file_path):
        """Carica automaticamente il file generato."""
        self.app_context.input_file = str(generated_file_path)
        self.input_label.config(
            text=Path(generated_file_path).name,
            foreground="green"
        )
        self.update_button_state()
        self.load_tipo_scheda_from_file()
        self.status_label.config(
            text="File pronto per la generazione",
            foreground="green"
        )

    def load_tipo_scheda(self, tipi_scheda):
        """Carica i tipi scheda disponibili"""
        for widget in self.checkbox_frame.winfo_children():
            widget.destroy()
        self.tipo_scheda_checkboxes.clear()

        for idx, tipo in enumerate(tipi_scheda):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.checkbox_frame, text=tipo, variable=var)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            self.tipo_scheda_checkboxes[tipo] = var
            cb.configure(state='normal')

    def generate_etichetteword(self):
        """Genera il documento Word con le etichette"""
        try:
            self.status_label.config(text="Generazione Etichette Interne in corso...", foreground="blue")
            self.progress.start()
            self.app_context.root.update()

            filter_enabled = self.filter_enabled_var.get()
            selected_tipo_scheda = None

            if filter_enabled:
                # Le checkbox ora contengono etichette del tipo "DESCRIZIONE - PREFISSO"
                selected_labels = [lbl for lbl, var in self.tipo_scheda_checkboxes.items() if var.get()]

                if not selected_labels:
                    messagebox.showwarning("Attenzione", "Seleziona almeno un elemento dalla lista!")
                    self.progress.stop()
                    return

                # Traduci le etichette selezionate in tipi scheda reali leggendo il file di input
                try:
                    df = pd.read_excel(self.app_context.input_file)
                except Exception as e:
                    self.progress.stop()
                    messagebox.showerror("Errore", f"Impossibile leggere il file di input:\n{e}")
                    return

                selected_tipo_scheda = []
                for lbl in selected_labels:
                    mapping = self._word_label_mapping.get(lbl)
                    if not mapping:
                        continue
                    descr, prefisso = mapping
                    # Filtra le righe corrispondenti alla descrizione
                    sub = df[df['Descrizione'] == descr]
                    if 'Tipo Scheda' in df.columns:
                        if prefisso:
                            matches = sub[sub['Tipo Scheda'].astype(str).str.startswith(prefisso, na=False)]['Tipo Scheda'].unique()
                        else:
                            matches = sub['Tipo Scheda'].dropna().unique()
                        for m in matches:
                            selected_tipo_scheda.append(str(m))

                # rimuovi duplicati
                selected_tipo_scheda = sorted(list(dict.fromkeys(selected_tipo_scheda)))

                if not selected_tipo_scheda:
                    messagebox.showwarning("Attenzione", "Nessun Tipo Scheda trovato per le selezioni effettuate")
                    self.progress.stop()
                    return

            repetitions = getattr(self, 'repetitions_var', tk.IntVar(value=1)).get()
            start_column = getattr(self, 'start_column_var', tk.IntVar(value=1)).get()
            start_row = getattr(self, 'start_row_var', tk.IntVar(value=1)).get()
            add_black_labels = getattr(self, 'add_black_labels_var', tk.BooleanVar(value=True)).get()

            label_count = WordLabelGenerator.generate_word_labels(
                self.app_context.input_file,
                self.app_context.etichetteword_output_file,
                filter_enabled,
                selected_tipo_scheda,
                repetitions,
                start_column,
                start_row,
                add_black_labels=add_black_labels
            )

            self.progress.stop()
            self.status_label.config(
                text=f"Etichette Interne generate! Etichette totali: {label_count}",
                foreground="green"
            )

            if filter_enabled:
                msg = (f"File Etichette Interne generato con successo!\n\n"
                       f"File output:\n{self.app_context.etichetteword_output_file}\n\n"
                       f"Etichette generate: {label_count}\n"
                       f"Tipi scheda inclusi: {len(selected_tipo_scheda)}\n"
                       f"Formato: A4 Portrate con righe alternate bianco/nero")
            else:
                msg = (f"File Etichette Interne generato con successo!\n\n"
                       f"File output:\n{self.app_context.etichetteword_output_file}\n\n"
                       f"Etichette generate: {label_count}\n"
                       f"Formato: A4 Portrate con righe alternate bianco/nero")

            messagebox.showinfo("Successo", msg)

        except ValueError as ve:
            self.progress.stop()
            self.status_label.config(text="Nessuna etichetta da generare", foreground="red")
            messagebox.showwarning("Attenzione", str(ve))
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Errore durante la generazione", foreground="red")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n\n{str(e)}")
