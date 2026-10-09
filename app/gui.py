import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import threading

import fitz
from PIL import Image, ImageTk

from .constants import CLASSES
from .database import WorkDatabase
from .importer import import_folder
from .exporter import export_project
from .structure import structure_parts
from .register_matcher import (
    load_register,
    normalize_personnummer,
    find_best_register_record,
)

from .multi_pass_ocr import valid_luhn

class ReviewApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sjöfartsbevisbearbetare v0.5")
        self.root.geometry("1600x950")
        self.root.minsize(1200, 720)

        self.db = None
        self.register_path = None
        self.register_by_personnummer = None
        # Samtliga sidor från arbetsdatabasen.
        self.all_rows = []

        # De sidor som för tillfället visas i gränssnittet.
        # Vid normal visning innehåller listan alla sidor.
        # Vid filtrering innehåller den endast matchande sidor.
        self.rows = []

        self.index = 0
        self.photo = None

        # Aktuellt visningsläge:
        # "alla", "osakra" eller "utan_personnummer".
        self.view_mode = "alla"

        self.class_var = tk.StringVar()
        self.pnr_var = tk.StringVar()

        self.title_var = tk.StringVar(
            value="Ingen databas öppnad"
        )

        self.status = tk.StringVar(
            value="Redo"
        )
        self.view_status = tk.StringVar(
            value="Visning: Alla sidor"
        )
        # Aktuell position i materialet.
        self.position_progress = [
            tk.DoubleVar()
            for _ in range(3)
        ]

        self.position_progress_text = [
            tk.StringVar(value=text)
            for text in (
                "Aktuell PDF: –",
                "Aktuell volym: –",
                "Aktuell serie: –",
            )
        ]

        # Faktiskt granskat innehåll.
        self.review_progress = [
            tk.DoubleVar()
            for _ in range(5)
        ]

        self.review_progress_text = [
            tk.StringVar(value=text)
            for text in (
                "Granskade sidor: –",
                "Granskade dokument: –",
                "Färdiga volymer: –",
                "Personnummer från skript: –",
                "Personnummer korrigerade av människa: –",
            )
        ]

        self.zoom = 1.8
        self.fit_scale = 1.0
        self.render_scale = 1.0

        self.image_origin = (0, 0)
        self.buttons = {}

        self.build()
        self.bind_keys()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close
        )

    def build(self):
        self.root.rowconfigure(
            2,
            weight=1
        )

        self.root.columnconfigure(
            0,
            weight=1
        )

        # -------------------------------------------------
        # Huvudverktygsfält
        # -------------------------------------------------

        toolbar = ttk.Frame(
            self.root,
            padding=(10, 8)
        )

        toolbar.grid(
            row=0,
            column=0,
            sticky="ew"
        )

        toolbar_buttons = [
            (
                "Ny arbetsdatabas",
                self.new_db
            ),
            (
                "Öppna arbetsdatabas",
                self.open_db
            ),
            (
                "Läs in seriemapp",
                self.load_folder
            ),
            (
                "Exportera resultat",
                self.do_export
            ),
            (
                "Navigera",
                self.open_navigation_dialog
            ),
        ]

        for text, command in toolbar_buttons:
            ttk.Button(
                toolbar,
                text=text,
                command=command
            ).pack(
                side="left",
                padx=(0, 6)
            )

        ttk.Label(
            toolbar,
            textvariable=self.view_status,
            foreground="#005a9e"
        ).pack(
            side="left",
            padx=(12, 0)
        )
        ttk.Label(
            toolbar,
            text=(
                "F1–F6: klassificera    "
                "←/→: byt sida    "
                "Ctrl+S: spara"
            ),
            foreground="#555"
        ).pack(
            side="right"
        )

        # -------------------------------------------------
        # Progress
        # -------------------------------------------------

        progress_container = ttk.Frame(
            self.root
        )

        progress_container.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=10
        )

        progress_container.columnconfigure(
            0,
            weight=1
        )

        progress_container.columnconfigure(
            1,
            weight=1
        )

        # Aktuell position.
        position_frame = ttk.LabelFrame(
            progress_container,
            text="Aktuell position",
            padding=7
        )

        position_frame.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 5)
        )

        position_frame.columnconfigure(
            1,
            weight=1
        )

        for i in range(3):
            ttk.Label(
                position_frame,
                textvariable=self.position_progress_text[i],
                width=43
            ).grid(
                row=i,
                column=0,
                sticky="w",
                padx=(0, 5)
            )

            ttk.Progressbar(
                position_frame,
                variable=self.position_progress[i],
                maximum=100
            ).grid(
                row=i,
                column=1,
                sticky="ew",
                pady=1
            )

        # Granskningsprogress.
        review_frame = ttk.LabelFrame(
            progress_container,
            text="Granskningsprogress",
            padding=7
        )

        review_frame.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(5, 0)
        )

        review_frame.columnconfigure(
            1,
            weight=1
        )

        for i in range(5):
            ttk.Label(
                review_frame,
                textvariable=self.review_progress_text[i],
                width=43
            ).grid(
                row=i,
                column=0,
                sticky="w",
                padx=(0, 5)
            )

            ttk.Progressbar(
                review_frame,
                variable=self.review_progress[i],
                maximum=100
            ).grid(
                row=i,
                column=1,
                sticky="ew",
                pady=1
            )

        # -------------------------------------------------
        # Huvudinnehåll
        # -------------------------------------------------

        content = ttk.Panedwindow(
            self.root,
            orient="horizontal"
        )

        content.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=10,
            pady=8
        )

        pdf_frame = ttk.LabelFrame(
            content,
            text="PDF-sida",
            padding=4
        )

        info = ttk.Frame(
            content,
            padding=(8, 0, 0, 0)
        )

        # PDF-fönstret får större delen av utrymmet.
        content.add(
            pdf_frame,
            weight=5
        )

        content.add(
            info,
            weight=2
        )

        # -------------------------------------------------
        # PDF-panel
        # -------------------------------------------------

        pdf_frame.rowconfigure(
            2,
            weight=1
        )

        pdf_frame.columnconfigure(
            0,
            weight=1
        )

        ttk.Label(
            pdf_frame,
            textvariable=self.title_var,
            font=("Segoe UI", 12, "bold")
        ).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 4)
        )

        # Zoomkontroller direkt ovanför PDF-rutan.
        pdf_toolbar = ttk.Frame(
            pdf_frame
        )

        pdf_toolbar.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(0, 4)
        )

        ttk.Button(
            pdf_toolbar,
            text="−",
            width=3,
            command=self.zoom_out
        ).pack(
            side="left"
        )

        self.zoom_label = ttk.Label(
            pdf_toolbar,
            text="100%",
            width=7,
            anchor="center"
        )

        self.zoom_label.pack(
            side="left"
        )

        ttk.Button(
            pdf_toolbar,
            text="+",
            width=3,
            command=self.zoom_in
        ).pack(
            side="left"
        )

        ttk.Button(
            pdf_toolbar,
            text="Anpassa",
            command=self.fit_page
        ).pack(
            side="left",
            padx=(6, 0)
        )

        ttk.Label(
            pdf_toolbar,
            text=(
                "Ctrl + mushjul: zooma    "
                "Mittenknapp: flytta sidan"
            ),
            foreground="#555"
        ).pack(
            side="left",
            padx=(15, 0)
        )

        self.canvas = tk.Canvas(
            pdf_frame,
            bg="#555",
            highlightthickness=0,
            xscrollincrement=1,
            yscrollincrement=1
        )

        xbar = ttk.Scrollbar(
            pdf_frame,
            orient="horizontal",
            command=self.canvas.xview
        )

        ybar = ttk.Scrollbar(
            pdf_frame,
            orient="vertical",
            command=self.canvas.yview
        )

        self.canvas.configure(
            xscrollcommand=xbar.set,
            yscrollcommand=ybar.set
        )

        self.canvas.grid(
            row=2,
            column=0,
            sticky="nsew"
        )

        ybar.grid(
            row=2,
            column=1,
            sticky="ns"
        )

        xbar.grid(
            row=3,
            column=0,
            sticky="ew"
        )

        self.canvas.bind(
            "<Configure>",
            self.on_canvas_resize
        )

        # Zoom på Windows.
        self.canvas.bind(
            "<Control-MouseWheel>",
            self.mouse_zoom
        )

        # Zoom på Linux.
        self.canvas.bind(
            "<Control-Button-4>",
            lambda event: self.mouse_zoom_linux(1)
        )

        self.canvas.bind(
            "<Control-Button-5>",
            lambda event: self.mouse_zoom_linux(-1)
        )

        # Panorering med musens mittenknapp.
        self.canvas.bind(
            "<ButtonPress-2>",
            self.pan_start
        )

        self.canvas.bind(
            "<B2-Motion>",
            self.pan_move
        )

        # -------------------------------------------------
        # Informations- och granskningspanel
        # -------------------------------------------------

        info.columnconfigure(
            0,
            weight=1
        )

        info.columnconfigure(
            1,
            weight=1
        )

        # Poängorsaker och OCR-text får expanderbart utrymme.
        info.rowconfigure(
            8,
            weight=1
        )

        info.rowconfigure(
            10,
            weight=2
        )

        ttk.Label(
            info,
            text="Klassificering",
            font=("Segoe UI", 11, "bold")
        ).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 5)
        )

        for i, (key, label, hotkey) in enumerate(CLASSES):
            button = ttk.Button(
                info,
                text=f"{hotkey} – {label}",
                command=lambda k=key: self.set_class(k)
            )

            button.grid(
                row=1 + i // 2,
                column=i % 2,
                sticky="ew",
                padx=2,
                pady=2
            )

            self.buttons[key] = button

        ttk.Label(
            info,
            text="Personnummer"
        ).grid(
            row=4,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(12, 0)
        )

        personnummer_frame = ttk.Frame(
            info
        )

        personnummer_frame.grid(
            row=5,
            column=0,
            columnspan=2,
            sticky="ew"
        )

        personnummer_frame.columnconfigure(
            0,
            weight=1
        )

        self.pnr_entry = ttk.Entry(
            personnummer_frame,
            textvariable=self.pnr_var
        )

        self.pnr_entry.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 6)
        )

        self.pnr_entry.bind(
            "<FocusOut>",
            lambda event: self.save()
        )

        ttk.Button(
            personnummer_frame,
            text="Sök efter korrigerat personnummer",
            command=self.search_corrected_personnummer
        ).grid(
            row=0,
            column=1,
            sticky="e"
        )

        self.score_label = ttk.Label(
            info,
            text="Poäng: –",
            font=("Segoe UI", 10, "bold")
        )

        self.score_label.grid(
            row=6,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(12, 3)
        )

        # -------------------------------------------------
        # Poängorsaker
        # -------------------------------------------------

        ttk.Label(
            info,
            text="Poängorsaker"
        ).grid(
            row=7,
            column=0,
            columnspan=2,
            sticky="w"
        )

        reasons_frame = ttk.Frame(
            info
        )

        reasons_frame.grid(
            row=8,
            column=0,
            columnspan=2,
            sticky="nsew"
        )

        reasons_frame.rowconfigure(
            0,
            weight=1
        )

        reasons_frame.columnconfigure(
            0,
            weight=1
        )

        self.reasons = tk.Listbox(
            reasons_frame,
            exportselection=False
        )

        reasons_scrollbar = ttk.Scrollbar(
            reasons_frame,
            orient="vertical",
            command=self.reasons.yview
        )

        self.reasons.configure(
            yscrollcommand=reasons_scrollbar.set
        )

        self.reasons.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        reasons_scrollbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        # -------------------------------------------------
        # OCR-resultat
        # -------------------------------------------------

        ttk.Label(
            info,
            text="OCR-resultat"
        ).grid(
            row=9,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(10, 4)
        )

        ocr_frame = ttk.Frame(
            info
        )

        ocr_frame.grid(
            row=10,
            column=0,
            columnspan=2,
            sticky="nsew"
        )

        ocr_frame.rowconfigure(
            0,
            weight=1
        )

        ocr_frame.columnconfigure(
            0,
            weight=1
        )

        self.ocr = tk.Text(
            ocr_frame,
            height=12,
            wrap="word",
            font=("Consolas", 9),
            undo=False
        )

        ocr_scrollbar = ttk.Scrollbar(
            ocr_frame,
            orient="vertical",
            command=self.ocr.yview
        )

        self.ocr.configure(
            yscrollcommand=ocr_scrollbar.set
        )

        self.ocr.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        ocr_scrollbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        # Texten ska kunna markeras och kopieras,
        # men inte ändras av användaren.
        self.ocr.bind(
            "<Key>",
            self.block_ocr_edit
        )

        # -------------------------------------------------
        # Navigeringsknappar
        # -------------------------------------------------

        navigation = ttk.Frame(
            info
        )

        navigation.grid(
            row=11,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(8, 0)
        )

        navigation.columnconfigure(
            0,
            weight=1
        )

        navigation.columnconfigure(
            1,
            weight=1
        )

        ttk.Button(
            navigation,
            text="← Föregående",
            command=lambda: self.move(-1)
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 2)
        )

        ttk.Button(
            navigation,
            text="Nästa →",
            command=lambda: self.move(1)
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(2, 0)
        )

        # -------------------------------------------------
        # Statusrad
        # -------------------------------------------------

        ttk.Label(
            self.root,
            textvariable=self.status,
            anchor="w"
        ).grid(
            row=3,
            column=0,
            sticky="ew",
            padx=10,
            pady=(0, 6)
        )

    def bind_keys(self):
        for key, _, function_key in CLASSES:
            self.root.bind(
                f"<{function_key}>",
                lambda event, k=key: self.set_class(k)
            )

        self.root.bind(
            "<Left>",
            lambda event: self.move(-1)
        )

        self.root.bind(
            "<Right>",
            lambda event: self.move(1)
        )

        self.root.bind(
            "<Control-s>",
            lambda event: self.save()
        )

    def block_ocr_edit(self, event):
        """
        Hindrar användaren från att ändra OCR-texten men tillåter
        navigering, markering och kopiering.
        """

        allowed_keys = {
            "Left",
            "Right",
            "Up",
            "Down",
            "Home",
            "End",
            "Prior",
            "Next",
        }

        if event.keysym in allowed_keys:
            return None

        # Tillåt Ctrl+C och Ctrl+A.
        if event.state & 0x0004:
            if event.keysym.lower() in {
                "c",
                "a",
            }:
                return None

        return "break"

    def select_register(self):
        register_path = filedialog.askopenfilename(
            title="Välj personnummerregister",
            filetypes=[
                (
                    "Excel-arbetsbok",
                    "*.xlsx",
                ),
                (
                    "Alla filer",
                    "*.*",
                ),
            ],
        )

        if not register_path:
            self.register_path = None
            self.register_by_personnummer = None
            return False

        try:
            register_data = load_register(
                register_path,
            )

        except Exception as exception:
            messagebox.showerror(
                "Registerfel",
                str(exception),
            )

            self.register_path = None
            self.register_by_personnummer = None
            return False

        self.register_path = register_path
        self.register_by_personnummer = register_data

        number_of_records = sum(
            len(records)
            for records in register_data.values()
        )

        self.status.set(
            "Register inläst: "
            f"{number_of_records} registerrader, "
            f"{len(register_data)} unika personnummer"
        )

        return True

    def new_db(self):
        path = filedialog.asksaveasfilename(
            title="Skapa arbetsdatabas",
            defaultextension=".sqlite",
            filetypes=[
                (
                    "SQLite",
                    "*.sqlite"
                ),
                (
                    "Alla filer",
                    "*.*"
                ),
            ]
        )

        if path:
            if not self.select_register():
                messagebox.showwarning(
                    "Inget register valt",
                    (
                        "Arbetsdatabasen öppnas inte eftersom "
                        "inget giltigt register valdes."
                    ),
                )
                return

            self.connect(path)

    def open_db(self):
        path = filedialog.askopenfilename(
            title="Öppna arbetsdatabas",
            filetypes=[
                (
                    "SQLite",
                    "*.sqlite *.db"
                ),
                (
                    "Alla filer",
                    "*.*"
                ),
            ]
        )

        if path:
            if not self.select_register():
                messagebox.showwarning(
                    "Inget register valt",
                    (
                        "Arbetsdatabasen öppnas inte eftersom "
                        "inget giltigt register valdes."
                    ),
                )
                return

            self.connect(path)

    def connect(self, path):
        if self.db:
            self.db.close()

        self.db = WorkDatabase(path)

        self.all_rows = list(
            self.db.pages()
        )

        self.rows = list(
            self.all_rows
        )

        self.view_mode = "alla"

        self.view_status.set(
            f"Visning: Alla sidor "
            f"({len(self.rows)})"
        )

        self.index = 0

        last_page_id = (
            self.db.get_current_position()
        )

        position_restored = False

        if last_page_id is not None:
            for row_index, row in enumerate(
                    self.rows
            ):
                if row["id"] == last_page_id:
                    self.index = row_index
                    position_restored = True
                    break

        # Om ingen sparad position finns börjar vi på första
        # sida som ännu inte är granskad.
        if (
                not position_restored
                and self.rows
        ):
            for row_index, row in enumerate(
                    self.rows
            ):
                if not row["reviewed"]:
                    self.index = row_index
                    break

        self.show()

        if (
                position_restored
                and self.rows
        ):
            current_row = self.rows[
                self.index
            ]

            self.status.set(
                "Föregående session återställd: "
                f'{current_row["relative_pdf"]}, '
                f'sida {current_row["page_number"]} '
                f'av {current_row["page_count"]}.'
            )

        elif self.rows:
            current_row = self.rows[
                self.index
            ]

            self.status.set(
                "Arbetsdatabasen öppnades vid "
                "första ogranskade sidan: "
                f'{current_row["relative_pdf"]}, '
                f'sida {current_row["page_number"]}.'
            )

        else:
            self.status.set(
                f"Databas öppnad: {path}"
            )

    def load_folder(self):
        if not self.db:
            messagebox.showwarning(
                "Ingen databas",
                (
                    "Skapa eller öppna en "
                    "arbetsdatabas först."
                )
            )
            return

        if not self.register_by_personnummer:
            messagebox.showwarning(
                "Inget register inläst",
                (
                    "Läs in ett personnummerregister innan "
                    "seriemappen importeras."
                ),
            )
            return

        folder = filedialog.askdirectory(
            title="Välj seriemapp"
        )

        if not folder:
            return

        self.status.set(
            "Importen startar..."
        )

        def update_import_progress(
            current,
            total,
            filename
        ):
            self.root.after(
                0,
                lambda: self.status.set(
                    f"Importerar {current}/{total}: "
                    f"{filename}"
                )
            )

        def work():
            count, errors = import_folder(
                self.db,
                folder,
                update_import_progress,
                register_by_personnummer=(
                    self.register_by_personnummer
                ),
            )

            self.root.after(
                0,
                lambda: self.import_done(
                    count,
                    errors
                )
            )

        threading.Thread(
            target=work,
            daemon=True
        ).start()

    def import_done(self, count, errors):
        self.refresh()

        self.status.set(
            f"Import klar: {count} PDF-filer, "
            f"{len(errors)} fel"
        )

        if errors:
            messagebox.showwarning(
                "Importavvikelser",
                "\n".join(errors[:20])
            )

    def refresh(self, preserve_page_id=None):
        """
        Läser om samtliga databasposter och återskapar
        det aktiva visningsläget.

        preserve_page_id används för att försöka behålla
        den sida som användaren befinner sig på.
        """

        if preserve_page_id is None and self.rows:
            preserve_page_id = self.rows[self.index]["id"]

        if self.db:
            self.all_rows = list(
                self.db.pages()
            )
        else:
            self.all_rows = []

        self.apply_view_filter(
            preserve_page_id=preserve_page_id
        )

    def apply_view_filter(
            self,
            preserve_page_id=None,
    ):
        """
        Skapar self.rows utifrån aktuellt visningsläge.

        Lägen:
        alla:
            Visar samtliga sidor.

        osakra:
            Visar sidor med poäng under 100.

        utan_personnummer:
            Visar sidor där inget personnummer finns.
        """

        if self.view_mode == "osakra":
            self.rows = [
                row
                for row in self.all_rows
                if float(row["score"] or 0) < 100
            ]

            self.view_status.set(
                f"Visning: Osäkra sidor "
                f"({len(self.rows)})"
            )

        elif self.view_mode == "utan_personnummer":
            self.rows = [
                row
                for row in self.all_rows
                if not str(
                    row["personnummer"] or ""
                ).strip()
            ]

            self.view_status.set(
                f"Visning: Sidor utan personnummer "
                f"({len(self.rows)})"
            )

        else:
            self.view_mode = "alla"
            self.rows = list(
                self.all_rows
            )

            self.view_status.set(
                f"Visning: Alla sidor "
                f"({len(self.rows)})"
            )

        if not self.rows:
            self.index = 0

            self.title_var.set(
                "Inga sidor motsvarar det valda filtret"
            )

            self.class_var.set("")
            self.pnr_var.set("")

            self.score_label.config(
                text="Poäng: –"
            )

            self.reasons.delete(
                0,
                "end"
            )

            self.ocr.delete(
                "1.0",
                "end"
            )

            self.canvas.delete(
                "all"
            )

            self.clear_progress()
            return

        self.index = 0

        if preserve_page_id is not None:
            for row_index, row in enumerate(
                    self.rows
            ):
                if row["id"] == preserve_page_id:
                    self.index = row_index
                    break

        self.show()
    def show(self):
        if not self.rows:
            self.title_var.set(
                "Inga importerade PDF-sidor"
            )

            self.class_var.set("")
            self.pnr_var.set("")

            self.score_label.config(
                text="Poäng: –"
            )

            self.reasons.delete(
                0,
                "end"
            )

            self.ocr.delete(
                "1.0",
                "end"
            )

            self.canvas.delete(
                "all"
            )

            self.clear_progress()
            return

        row = self.rows[self.index]
        # Spara den aktuella positionen varje gång en sida visas.
        if self.db:
            self.db.save_current_position(
                row["id"]
            )

        self.title_var.set(
            f'{row["relative_pdf"]} | '
            f'sida {row["page_number"]}/'
            f'{row["page_count"]}'
        )

        self.class_var.set(
            row["classification"]
        )

        self.pnr_var.set(
            row["personnummer"]
        )

        self.score_label.config(
            text=f'Poäng: {row["score"]}/100'
        )

        self.reasons.delete(
            0,
            "end"
        )

        try:
            reasons = json.loads(
                row["score_reasons"] or "[]"
            )
        except (
            json.JSONDecodeError,
            TypeError
        ):
            reasons = []

        for reason in reasons:
            self.reasons.insert(
                "end",
                "• " + str(reason)
            )

        # Visa OCR-resultatet.
        self.ocr.delete(
            "1.0",
            "end"
        )

        self.ocr.insert(
            "1.0",
            row["ocr_text"] or ""
        )

        # Flytta OCR-rutan till början varje gång sidan byts.
        self.ocr.see(
            "1.0"
        )

        self.zoom = 1.8

        self.update_progress()
        self.update_class_buttons()
        self.render()

    def clear_progress(self):
        position_labels = [
            "Aktuell PDF: –",
            "Aktuell volym: –",
            "Aktuell serie: –",
        ]

        review_labels = [
            "Granskade sidor: –",
            "Granskade dokument: –",
            "Färdiga volymer: –",
            "Personnummer från skript: –",
            "Personnummer korrigerade av människa: –",
        ]

        for index in range(3):
            self.position_progress[index].set(0)
            self.position_progress_text[index].set(
                position_labels[index]
            )

        for index in range(5):
            self.review_progress[index].set(0)
            self.review_progress_text[index].set(
                review_labels[index]
            )
    def update_progress(self):

        if not self.rows:
            self.clear_progress()
            return
        all_progress_rows = (
            self.all_rows
            if self.all_rows
            else self.rows
        )
        current_row = self.rows[
            self.index
        ]

        current_series, current_volume = (
            structure_parts(
                current_row["relative_pdf"]
            )
        )

        # -------------------------------------------------
        # Bygg datastruktur för dokument och volymer
        # -------------------------------------------------

        documents = {}

        for row in all_progress_rows:
            relative_pdf = row[
                "relative_pdf"
            ]

            if relative_pdf not in documents:
                series, volume = (
                    structure_parts(
                        relative_pdf
                    )
                )

                documents[relative_pdf] = {
                    "series": series,
                    "volume": volume,
                    "pages": [],
                }

            documents[
                relative_pdf
            ]["pages"].append(row)

        volume_documents = {}

        for relative_pdf, document in (
                documents.items()
        ):
            volume_key = (
                document["series"],
                document["volume"],
            )

            if volume_key not in volume_documents:
                volume_documents[
                    volume_key
                ] = []

            volume_documents[
                volume_key
            ].append(
                relative_pdf
            )

        # -------------------------------------------------
        # Aktuell position
        # -------------------------------------------------

        current_document_pages = documents[
            current_row["relative_pdf"]
        ]["pages"]

        current_page_position = next(
            position
            for position, page
            in enumerate(
                current_document_pages,
                start=1
            )
            if page["id"] == current_row["id"]
        )

        current_volume_key = (
            current_series,
            current_volume,
        )

        current_volume_documents = (
            volume_documents[
                current_volume_key
            ]
        )

        current_document_position = (
                current_volume_documents.index(
                    current_row["relative_pdf"]
                )
                + 1
        )

        all_volumes = list(
            volume_documents.keys()
        )

        current_volume_position = (
                all_volumes.index(
                    current_volume_key
                )
                + 1
        )

        position_values = [
            (
                current_page_position,
                len(current_document_pages)
            ),
            (
                current_document_position,
                len(current_volume_documents)
            ),
            (
                current_volume_position,
                len(all_volumes)
            ),
        ]

        position_labels = [
            (
                f"Aktuell PDF: sida "
                f"{current_page_position}/"
                f"{len(current_document_pages)} • "
                f"{len(current_document_pages) - current_page_position} "
                f"sidor kvar"
            ),
            (
                f"Aktuell volym "
                f"{current_volume or '–'}: "
                f"dokument "
                f"{current_document_position}/"
                f"{len(current_volume_documents)}"
            ),
            (
                f"Aktuell serie "
                f"{current_series or '–'}: "
                f"volym "
                f"{current_volume_position}/"
                f"{len(all_volumes)}"
            ),
        ]

        for index, (
                completed,
                total
        ) in enumerate(position_values):
            self.position_progress[
                index
            ].set(
                100
                * completed
                / max(
                    1,
                    total
                )
            )

            self.position_progress_text[
                index
            ].set(
                position_labels[index]
            )

        # -------------------------------------------------
        # Faktiskt granskat innehåll
        # -------------------------------------------------
        script_personnummer = sum(
            1
            for row in all_progress_rows
            if row["personnummer_source"] == "skript"
        )

        human_personnummer = sum(
            1
            for row in all_progress_rows
            if row["personnummer_source"] == "människa"
        )
        total_pages = len(
            all_progress_rows
        )

        reviewed_pages = sum(
            1
            for row in all_progress_rows
            if bool(row["reviewed"])
        )

        total_documents = len(
            documents
        )

        reviewed_documents = 0

        for document in documents.values():
            pages = document["pages"]

            # Dokumentet är färdigt först när alla dess
            # sidor har markerats som granskade.
            if (
                    pages
                    and all(
                bool(page["reviewed"])
                for page in pages
            )
            ):
                reviewed_documents += 1

        total_volumes = len(
            volume_documents
        )

        reviewed_volumes = 0

        for document_paths in (
                volume_documents.values()
        ):
            volume_is_complete = True

            for document_path in (
                    document_paths
            ):
                pages = documents[
                    document_path
                ]["pages"]

                if not (
                        pages
                        and all(
                    bool(page["reviewed"])
                    for page in pages
                )
                ):
                    volume_is_complete = False
                    break

            if volume_is_complete:
                reviewed_volumes += 1

        review_values = [
            (
                reviewed_pages,
                total_pages
            ),
            (
                reviewed_documents,
                total_documents
            ),
            (
                reviewed_volumes,
                total_volumes
            ),
            (
                script_personnummer,
                total_pages
            ),
            (
                human_personnummer,
                total_pages
            )
        ]

        review_labels = [
            (
                f"Granskade sidor: "
                f"{reviewed_pages}/{total_pages}"
            ),
            (
                f"Granskade dokument: "
                f"{reviewed_documents}/{total_documents}"
            ),
            (
                f"Färdiga volymer: "
                f"{reviewed_volumes}/{total_volumes}"
            ),
            (
                f"Personnummer från skript: "
                f"{script_personnummer}/{total_pages}"
            ),
            (
                f"Personnummer korrigerade av människa: "
                f"{human_personnummer}/{total_pages}"
            ),
        ]

        for index, (
                completed,
                total
        ) in enumerate(review_values):
            self.review_progress[
                index
            ].set(
                100
                * completed
                / max(
                    1,
                    total
                )
            )

            self.review_progress_text[
                index
            ].set(
                review_labels[index]
            )
    def update_class_buttons(self):
        selected_class = (
            self.class_var.get()
        )

        for key, button in (
            self.buttons.items()
        ):
            if key == selected_class:
                button.state(
                    ["pressed"]
                )
            else:
                button.state(
                    ["!pressed"]
                )

    def search_corrected_personnummer(self):
        """
        Söker efter det personnummer som användaren har skrivit
        i personnummerfältet.

        Om personnumret finns i registret jämförs tillhörande
        förnamn och efternamn med den aktuella sidans OCR-text.
        Resultatet sparas i arbetsdatabasen.
        """

        if not self.db:
            messagebox.showwarning(
                "Ingen arbetsdatabas",
                "Öppna en arbetsdatabas först."
            )
            return

        if not self.rows:
            messagebox.showwarning(
                "Ingen sida vald",
                "Det finns ingen aktuell sida att kontrollera."
            )
            return

        if not self.register_by_personnummer:
            messagebox.showwarning(
                "Inget register inläst",
                (
                    "Det finns inget inläst personnummerregister. "
                    "Öppna arbetsdatabasen tillsammans med ett register."
                )
            )
            return

        entered_personnummer = self.pnr_var.get()

        normalized_personnummer = normalize_personnummer(
            entered_personnummer
        )

        if not normalized_personnummer:
            messagebox.showwarning(
                "Ogiltigt personnummer",
                (
                    "Det inskrivna värdet kunde inte tolkas som ett "
                    "personnummer.\n\n"
                    "Använd exempelvis formatet 620505-5137."
                )
            )
            return

        # Uppdatera personnummerfältet med normaliserat format.
        self.pnr_var.set(
            normalized_personnummer
        )

        current_row = self.rows[
            self.index
        ]

        ocr_text = str(
            current_row["ocr_text"] or ""
        )

        score = 0
        reasons = [
            (
                f"Manuellt korrigerat personnummer: "
                f"{normalized_personnummer}"
            )
        ]

        # Kontrollera svensk kontrollsiffra.
        checksum_valid = valid_luhn(
            normalized_personnummer
        )

        if checksum_valid:
            score += 10

            reasons.append(
                "Personnumrets kontrollsiffra är giltig: +10 poäng"
            )
        else:
            score -= 10

            reasons.append(
                "Personnumrets kontrollsiffra är inte giltig: -10 poäng"
            )

        matching_records = self.register_by_personnummer.get(
            normalized_personnummer,
            []
        )

        if not matching_records:
            score -= 20

            reasons.append(
                "Det korrigerade personnumret hittades inte i registret: "
                "-20 poäng"
            )

            score = max(
                0,
                min(
                    100,
                    score
                )
            )

            self.save_corrected_validation_result(
                personnummer=normalized_personnummer,
                score=score,
                reasons=reasons,
            )

            messagebox.showwarning(
                "Ingen registerträff",
                (
                    f"Personnumret {normalized_personnummer} "
                    "hittades inte i registret.\n\n"
                    "Kontrollera personnumret en gång till."
                )
            )

            return

        score += 60

        reasons.append(
            "Det korrigerade personnumret hittades i registret: "
            "+60 poäng"
        )

        best_record, name_match = find_best_register_record(
            matching_records,
            ocr_text,
        )

        if best_record is not None and name_match is not None:
            score += name_match.points

            reasons.extend(
                name_match.reasons
            )

            reasons.append(
                (
                    f"Vald registerpost: "
                    f"{best_record.fornamn} "
                    f"{best_record.efternamn}"
                )
            )

            register_name = (
                f"{best_record.fornamn} "
                f"{best_record.efternamn}"
            ).strip()

            fornamn_percent = round(
                name_match.fornamn_similarity * 100
            )

            efternamn_percent = round(
                name_match.efternamn_similarity * 100
            )

        else:
            register_name = "Namnet kunde inte läsas från registret"
            fornamn_percent = 0
            efternamn_percent = 0

            reasons.append(
                "Ingen namnpost kunde jämföras med OCR-resultatet"
            )

        score = max(
            0,
            min(
                100,
                score
            )
        )

        self.save_corrected_validation_result(
            personnummer=normalized_personnummer,
            score=score,
            reasons=reasons,
        )

        messagebox.showinfo(
            "Registerkontroll klar",
            (
                f"Personnummer: {normalized_personnummer}\n"
                f"Registerpost: {register_name}\n\n"
                f"Förnamnsmatchning: {fornamn_percent} %\n"
                f"Efternamnsmatchning: {efternamn_percent} %\n\n"
                f"Ny poäng: {score}/100"
            )
        )

    def save_corrected_validation_result(
            self,
            personnummer,
            score,
            reasons,
    ):
        """
        Sparar resultatet från kontrollen av ett manuellt
        korrigerat personnummer.
        """

        if not self.rows:
            return

        current_page_id = self.rows[
            self.index
        ]["id"]

        score_reasons_json = json.dumps(
            reasons,
            ensure_ascii=False,
        )

        self.db.update_validation_result(
            page_id=current_page_id,
            personnummer=personnummer,
            score=score,
            score_reasons=score_reasons_json,
        )

        self.db.save_current_position(
            current_page_id
        )

        # Hämta den uppdaterade databasen.
        self.all_rows = list(
            self.db.pages()
        )

        # Återskapa aktuellt filter utan att byta sida i onödan.
        self.apply_view_filter_without_show(
            preserve_page_id=current_page_id
        )

        if self.rows:
            matching_index = None

            for row_index, row in enumerate(
                    self.rows
            ):
                if row["id"] == current_page_id:
                    matching_index = row_index
                    break

            if matching_index is not None:
                self.index = matching_index
                self.show()
            else:
                # Sidan kan ha försvunnit ur filtret
                # "Sidor utan personnummer".
                self.index = min(
                    self.index,
                    len(self.rows) - 1
                )

                self.show()

        else:
            self.title_var.set(
                "Inga sidor motsvarar det valda filtret"
            )

            self.class_var.set("")
            self.pnr_var.set("")

            self.score_label.config(
                text="Poäng: –"
            )

            self.reasons.delete(
                0,
                "end"
            )

            self.ocr.delete(
                "1.0",
                "end"
            )

            self.canvas.delete(
                "all"
            )

            self.clear_progress()

        self.status.set(
            (
                f"Personnummer {personnummer} kontrollerades "
                f"mot registret."
            )
        )
    def save(self):
        if not self.db:
            return

        if not self.rows:
            return

        current_page_id = self.rows[
            self.index
        ]["id"]

        current_index = self.index

        self.db.save(
            current_page_id,
            self.class_var.get(),
            self.pnr_var.get()
        )

        self.db.save_current_position(
            current_page_id
        )

        self.all_rows = list(
            self.db.pages()
        )

        self.apply_view_filter_without_show(
            preserve_page_id=current_page_id
        )

        page_still_visible = any(
            row["id"] == current_page_id
            for row in self.rows
        )

        if not page_still_visible and self.rows:
            self.index = min(
                current_index,
                len(self.rows) - 1
            )

            self.show()
            return

        if not self.rows:
            self.title_var.set(
                "Inga sidor motsvarar det valda filtret"
            )

            self.class_var.set("")
            self.pnr_var.set("")

            self.score_label.config(
                text="Poäng: –"
            )

            self.reasons.delete(
                0,
                "end"
            )

            self.ocr.delete(
                "1.0",
                "end"
            )

            self.canvas.delete(
                "all"
            )

            self.clear_progress()

            self.status.set(
                "Det finns inga fler sidor i den valda visningen."
            )
            return

        self.update_progress()

        self.status.set(
            "Ändringarna har sparats."
        )

    def apply_view_filter_without_show(
            self,
            preserve_page_id=None,
    ):
        """
        Samma filtrering som apply_view_filter, men utan
        att återge PDF-sidan på nytt.

        Används efter sparning så att gränssnittet inte
        blinkar eller återrenderar PDF-filen i onödan.
        """

        if self.view_mode == "osakra":
            self.rows = [
                row
                for row in self.all_rows
                if float(row["score"] or 0) < 100
            ]

            self.view_status.set(
                f"Visning: Osäkra sidor "
                f"({len(self.rows)})"
            )

        elif self.view_mode == "utan_personnummer":
            self.rows = [
                row
                for row in self.all_rows
                if not str(
                    row["personnummer"] or ""
                ).strip()
            ]

            self.view_status.set(
                f"Visning: Sidor utan personnummer "
                f"({len(self.rows)})"
            )


        else:

            self.view_mode = "alla"

            self.rows = list(

                self.all_rows

            )

            self.view_status.set(

                f"Visning: Alla sidor "

                f"({len(self.rows)})"

            )

        if not self.rows:
            self.index = 0
            return

        self.index = min(
            self.index,
            len(self.rows) - 1
        )

        if preserve_page_id is not None:
            for row_index, row in enumerate(
                    self.rows
            ):
                if row["id"] == preserve_page_id:
                    self.index = row_index
                    break

    def set_class(self, value):
        if not self.rows:
            return

        current_page_id = self.rows[
            self.index
        ]["id"]

        self.class_var.set(
            value
        )

        self.update_class_buttons()
        self.save()

        if not self.rows:
            return

        current_page_still_visible = any(
            row["id"] == current_page_id
            for row in self.rows
        )

        if current_page_still_visible:
            self.move(
                1,
                save=False
            )


    def move(
        self,
        amount,
        save=True
    ):
        if not self.rows:
            return

        if save:
            self.save()

        self.index = max(
            0,
            min(
                len(self.rows) - 1,
                self.index + amount
            )
        )

        self.show()

    def fit_page(self):
        self.zoom = 1.0
        self.render()

    def zoom_in(self):
        if not self.rows:
            return

        self.zoom = min(
            6.0,
            self.zoom * 1.25
        )

        self.render()

    def zoom_out(self):
        if not self.rows:
            return

        self.zoom = max(
            0.25,
            self.zoom / 1.25
        )

        self.render()

    def mouse_zoom(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

        return "break"

    def mouse_zoom_linux(
        self,
        direction
    ):
        if direction > 0:
            self.zoom_in()
        else:
            self.zoom_out()

        return "break"

    def on_canvas_resize(self, event):
        if (
            self.rows
            and self.zoom == 1.0
        ):
            self.render()

    def render(self):
        if not self.rows:
            return

        row = self.rows[self.index]

        try:
            document = fitz.open(
                row["source_pdf"]
            )

            page = document[
                row["page_number"] - 1
            ]

            page_rect = page.rect

            available_width = max(
                100,
                self.canvas.winfo_width()
                - 30
            )

            available_height = max(
                100,
                self.canvas.winfo_height()
                - 30
            )

            self.fit_scale = min(
                available_width
                / page_rect.width,
                available_height
                / page_rect.height
            )

            self.render_scale = (
                self.fit_scale
                * self.zoom
            )

            pixmap = page.get_pixmap(
                matrix=fitz.Matrix(
                    self.render_scale,
                    self.render_scale
                ),
                alpha=False
            )

            image = Image.frombytes(
                "RGB",
                [
                    pixmap.width,
                    pixmap.height
                ],
                pixmap.samples
            )

            document.close()

            self.photo = ImageTk.PhotoImage(
                image
            )

            self.canvas.delete(
                "all"
            )

            margin = 15

            self.image_origin = (
                margin,
                margin
            )

            self.canvas.create_image(
                margin,
                margin,
                image=self.photo,
                anchor="nw",
                tags="page"
            )

            self.canvas.configure(
                scrollregion=(
                    0,
                    0,
                    pixmap.width
                    + 2 * margin,
                    pixmap.height
                    + 2 * margin
                )
            )

            self.zoom_label.config(
                text=(
                    f"{self.zoom * 100:.0f}%"
                )
            )

            # Börja högst upp och längst till vänster
            # när en ny sida visas.
            self.canvas.xview_moveto(0)
            self.canvas.yview_moveto(0)

        except Exception as exc:
            self.status.set(
                "Förhandsvisningsfel: "
                f"{exc}"
            )

    def pan_start(self, event):
        self.canvas.scan_mark(
            event.x,
            event.y
        )

    def pan_move(self, event):
        self.canvas.scan_dragto(
            event.x,
            event.y,
            gain=1
        )

    def get_navigation_structure(self):
        """
        Bygger en struktur med volymer och dokument.

        Resultat:

        {
            "Volym 1": {
                "fil1.pdf": första_sidans_id,
                "fil2.pdf": första_sidans_id,
            }
        }
        """

        navigation_structure = {}

        for row in self.all_rows:
            series, volume = structure_parts(
                row["relative_pdf"]
            )

            if not volume:
                volume = "Okänd volym"

            document_name = row[
                "relative_pdf"
            ]

            volume_documents = (
                navigation_structure.setdefault(
                    volume,
                    {}
                )
            )

            if document_name not in volume_documents:
                volume_documents[
                    document_name
                ] = row["id"]

        return navigation_structure

    def go_to_page_id(self, page_id):
        """
        Avslutar eventuell filtrering och går direkt
        till en sida i den fullständiga listan.
        """

        self.view_mode = "alla"

        self.rows = list(
            self.all_rows
        )

        for row_index, row in enumerate(
                self.rows
        ):
            if row["id"] == page_id:
                self.index = row_index
                self.view_status.set(
                    f"Visning: Alla sidor "
                    f"({len(self.rows)})"
                )
                self.show()
                return

        messagebox.showwarning(
            "Sidan hittades inte",
            "Den valda sidan kunde inte hittas."
        )

    def open_navigation_dialog(self):
        if not self.db or not self.all_rows:
            messagebox.showwarning(
                "Inget material",
                (
                    "Öppna en arbetsdatabas med inläst "
                    "material innan du navigerar."
                ),
            )
            return

        navigation_structure = self.get_navigation_structure()

        dialog = tk.Toplevel(self.root)
        dialog.title("Navigera")
        dialog.geometry("720x540")
        dialog.minsize(620, 450)
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.columnconfigure(
            0,
            weight=1,
        )

        dialog.rowconfigure(
            2,
            weight=1,
        )

        # -------------------------------------------------
        # Visningsalternativ
        # -------------------------------------------------

        filters_frame = ttk.LabelFrame(
            dialog,
            text="Visningsalternativ",
            padding=10,
        )

        filters_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=10,
            pady=(10, 5),
        )

        for column in range(3):
            filters_frame.columnconfigure(
                column,
                weight=1,
            )

        ttk.Button(
            filters_frame,
            text="Visa alla sidor",
            command=lambda: self.activate_view_mode(
                "alla",
                dialog,
            ),
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 4),
        )

        ttk.Button(
            filters_frame,
            text="Visa osäkra sidor",
            command=lambda: self.activate_view_mode(
                "osakra",
                dialog,
            ),
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=4,
        )

        ttk.Button(
            filters_frame,
            text="Visa sidor utan personnummer",
            command=lambda: self.activate_view_mode(
                "utan_personnummer",
                dialog,
            ),
        ).grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(4, 0),
        )

        # -------------------------------------------------
        # Volymval
        # -------------------------------------------------

        selection_frame = ttk.Frame(
            dialog,
            padding=(10, 5),
        )

        selection_frame.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        selection_frame.columnconfigure(
            1,
            weight=1,
        )

        ttk.Label(
            selection_frame,
            text="Volym:",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
        )

        volume_var = tk.StringVar()

        volume_box = ttk.Combobox(
            selection_frame,
            textvariable=volume_var,
            state="readonly",
            values=sorted(
                navigation_structure.keys()
            ),
        )

        volume_box.grid(
            row=0,
            column=1,
            sticky="ew",
        )

        # -------------------------------------------------
        # Dokumentlista
        # -------------------------------------------------

        documents_frame = ttk.LabelFrame(
            dialog,
            text="Dokument i vald volym",
            padding=8,
        )

        documents_frame.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=10,
            pady=5,
        )

        documents_frame.columnconfigure(
            0,
            weight=1,
        )

        documents_frame.rowconfigure(
            0,
            weight=1,
        )

        document_list = tk.Listbox(
            documents_frame,
            exportselection=False,
        )

        document_scrollbar = ttk.Scrollbar(
            documents_frame,
            orient="vertical",
            command=document_list.yview,
        )

        document_list.configure(
            yscrollcommand=document_scrollbar.set,
        )

        document_list.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        document_scrollbar.grid(
            row=0,
            column=1,
            sticky="ns",
        )

        # Listan innehåller de relativa PDF-sökvägar som
        # motsvarar posterna som visas i Listbox.
        displayed_documents = []

        def update_document_list(event=None):
            selected_volume = volume_var.get()

            document_list.delete(
                0,
                "end",
            )

            displayed_documents.clear()

            if not selected_volume:
                return

            documents = navigation_structure.get(
                selected_volume,
                {},
            )

            for relative_pdf in sorted(
                    documents.keys()
            ):
                displayed_documents.append(
                    relative_pdf
                )

                document_list.insert(
                    "end",
                    relative_pdf,
                )

            if displayed_documents:
                document_list.selection_set(0)
                document_list.activate(0)
                document_list.see(0)

        def go_to_selected_document():
            selected_volume = volume_var.get()
            selection = document_list.curselection()

            if not selected_volume:
                messagebox.showwarning(
                    "Ingen volym vald",
                    "Välj först en volym.",
                    parent=dialog,
                )
                return

            if not selection:
                messagebox.showwarning(
                    "Inget dokument valt",
                    "Välj ett dokument i listan.",
                    parent=dialog,
                )
                return

            document_position = selection[0]

            relative_pdf = displayed_documents[
                document_position
            ]

            page_id = navigation_structure[
                selected_volume
            ][relative_pdf]

            dialog.destroy()

            self.go_to_page_id(
                page_id
            )

        volume_box.bind(
            "<<ComboboxSelected>>",
            update_document_list,
        )

        document_list.bind(
            "<Double-Button-1>",
            lambda event: go_to_selected_document(),
        )

        document_list.bind(
            "<Return>",
            lambda event: go_to_selected_document(),
        )

        # -------------------------------------------------
        # Dialogknappar
        # -------------------------------------------------

        buttons_frame = ttk.Frame(
            dialog,
            padding=10,
        )

        buttons_frame.grid(
            row=3,
            column=0,
            sticky="ew",
        )

        ttk.Button(
            buttons_frame,
            text="Gå till dokument",
            command=go_to_selected_document,
        ).pack(
            side="right",
            padx=(5, 0),
        )

        ttk.Button(
            buttons_frame,
            text="Stäng",
            command=dialog.destroy,
        ).pack(
            side="right",
        )

        volumes = sorted(
            navigation_structure.keys()
        )

        if volumes:
            volume_var.set(
                volumes[0]
            )

            update_document_list()

            volume_box.focus_set()

    def activate_view_mode(
            self,
            view_mode,
            dialog=None,
    ):
        """
        Aktiverar ett visningsläge och går till den första
        sidan som motsvarar filtret.
        """

        current_page_id = None

        if self.rows:
            current_page_id = self.rows[
                self.index
            ]["id"]

        self.view_mode = view_mode

        if dialog is not None:
            dialog.destroy()

        self.apply_view_filter(
            preserve_page_id=current_page_id
        )

        if not self.rows:
            if view_mode == "osakra":
                messagebox.showinfo(
                    "Inga osäkra sidor",
                    (
                        "Det finns inga sidor med "
                        "poäng under 100."
                    ),
                )

            elif view_mode == "utan_personnummer":
                messagebox.showinfo(
                    "Inga sidor utan personnummer",
                    (
                        "Alla sidor har ett registrerat "
                        "personnummer."
                    ),
                )

            return

        if view_mode == "osakra":
            self.status.set(
                f"Visar {len(self.rows)} osäkra sidor."
            )

        elif view_mode == "utan_personnummer":
            self.status.set(
                f"Visar {len(self.rows)} sidor "
                f"utan personnummer."
            )

        else:
            self.status.set(
                f"Visar samtliga "
                f"{len(self.rows)} sidor."
            )
    def do_export(self):
        if not self.db:
            messagebox.showwarning(
                "Ingen databas",
                (
                    "Skapa eller öppna en "
                    "arbetsdatabas först."
                )
            )
            return

        self.save()

        folder = filedialog.askdirectory(
            title="Välj exportmapp"
        )

        if not folder:
            return

        try:
            exported_documents = (
                export_project(
                    self.db,
                    folder
                )
            )

            messagebox.showinfo(
                "Export klar",
                (
                    f"{exported_documents} "
                    f"dokument exporterades."
                )
            )

            self.status.set(
                f"Export klar: {folder}"
            )

        except Exception as exc:
            messagebox.showerror(
                "Exportfel",
                str(exc)
            )

            self.status.set(
                f"Exporten misslyckades: {exc}"
            )

    def close(self):
        if self.db:
            try:
                self.save()
            except Exception:
                pass

            self.db.close()

        self.root.destroy()

    def run(self):
        self.root.mainloop()