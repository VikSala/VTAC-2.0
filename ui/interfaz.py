import json
import os
import threading
import tkinter as tk
from enum import Enum, auto
from tkinter import *  # Considera importar selectivamente
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import webbrowser

# --- INICIO: Modificación para resolver ImportError de core.exceptions ---
import sys

_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)


# --- FIN: Modificación para resolver ImportError ---


class Interface:
    class CloseType(Enum):
        NORMAL = auto()
        PROHIBITED = auto()
        EXCEPTION = auto()

    def __init__(self, controller):
        self.window = None
        self.controller = controller
        self.WINDOW_WIDTH = 1200
        self.WINDOW_HEIGHT = 750
        self.BACKGROUND_COLOR = '#060032'
        self.logo_img_main_menu = None  # Para evitar GC

        self.BACK_BUTTON_STYLE = {
            'font': ('Helvetica', 14, 'bold'), 'padx': 20, 'text': 'Menú Principal',
            'pady': 10, 'bg': "#FC9C94", 'fg': 'black', 'relief': "raised",
            'borderwidth': 3, 'cursor': 'hand2'
        }
        self.button_style = {
            'font': ('Helvetica', 11, 'bold'), 'relief': 'raised', 'borderwidth': 2,
            'padx': 15, 'pady': 8, 'cursor': 'hand2', 'bg': '#4A90E2', 'fg': 'black'
        }
        self.checkbutton_style = {
            'font': ('Helvetica', 11, 'bold'), 'padx': 15, 'pady': 8, 'fg': 'white',
            'bg': self.BACKGROUND_COLOR, 'selectcolor': '#3A7BD5',
            'activebackground': '#1A4480', 'activeforeground': 'white'
        }
        self.label_style = {
            'font': ('Helvetica', 12, 'bold'), 'fg': 'white', 'bg': self.BACKGROUND_COLOR
        }
        self.title_style = {
            'font': ('Helvetica', 16, 'bold'), 'fg': '#FFD700', 'bg': self.BACKGROUND_COLOR
        }
        self.path_selector_failed = False
        self.progress_window = None
        self.progress_text = None
        self.status_label = None
        self._setup_event_listeners()

    def _setup_event_listeners(self):
        self.controller.event_manager.subscribe('status_update', self._on_status_update)
        self.controller.event_manager.subscribe('progress_update', self._on_progress_update)
        self.controller.event_manager.subscribe('scraping_completed', self._on_scraping_completed)
        self.controller.event_manager.subscribe('scraping_error', self._on_scraping_error)
        self.controller.event_manager.subscribe('operation_completed', self._on_operation_completed)
        self.controller.event_manager.subscribe('operation_error', self._on_operation_error)

    def _on_status_update(self, status, message):
        if self.status_label and self.status_label.winfo_exists():
            self.status_label.after(0, lambda: self.status_label.config(text=f"⚡ Estado: {message}"))

    def _on_progress_update(self, message):
        if self.progress_text and self.progress_window and self.progress_window.winfo_exists():
            self.progress_window.after(0, lambda: self._safe_update_progress(message))

    def _safe_update_progress(self, message):
        try:
            if self.progress_text and self.progress_text.winfo_exists():
                self.progress_text.insert(tk.END, message + "\n")
                self.progress_text.see(tk.END)
        except tk.TclError:
            pass

    def _on_scraping_completed(self, message):
        self._close_progress_window()
        self.show_info_message(message, "Scraping Completado")

    def _on_scraping_error(self, error_message):
        self._close_progress_window()
        messagebox.showerror("Error en Scraping", f"Error durante el scraping:\n{error_message}")

    def _on_operation_completed(self, message):
        if "generado:" in message.lower() or "completado" in message.lower() or "finalizado" in message.lower() or "cerrada" in message.lower():
            if self.progress_window and self.progress_window.winfo_exists():
                self._close_progress_window()
            self.show_info_message(message, "Éxito")

    def _on_operation_error(self, error_message):
        if self.progress_window and self.progress_window.winfo_exists():
            self._close_progress_window()
        messagebox.showerror("Error en Operación", f"Error durante la operación:\n{error_message}")

    def _close_progress_window(self):
        if self.progress_window and self.progress_window.winfo_exists():
            self.progress_window.after(0, self.progress_window.destroy)
        self.progress_window = None
        self.progress_text = None
        self.status_label = None

    @classmethod
    def hover(cls, button, on_enter, on_leave):
        button.bind("<Enter>", lambda e, h=on_enter: button.config(background=h))
        button.bind("<Leave>", lambda e, h=on_leave: button.config(background=h))

    def configure_window(self):
        if not self.window:
            self.window = self.controller.master
        self.window.title('🚀 Optima Scraper Pro')
        self.window.configure(background=self.BACKGROUND_COLOR)
        self.window.geometry(f'{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}')
        self.window.minsize(800, 600)
        self.window.protocol("WM_DELETE_WINDOW", self._on_main_window_close)
        # Configurar grid para que la fila 0 y columna 0 se expandan
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

    def _draw_main_menu(self, canvas):
        canvas.delete("all")
        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        if canvas_width <= 1 or canvas_height <= 1:
            canvas.after(50, lambda: self._draw_main_menu(canvas))
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_dir, 'logo.png')
        try:
            temp_logo_img = PhotoImage(master=self.window, file=logo_path)
            original_width = temp_logo_img.width()
            original_height = temp_logo_img.height()

            # Escalar logo para que no ocupe más del 30% del alto o 40% del ancho
            max_h = canvas_height * 0.30
            max_w = canvas_width * 0.40

            scale_h = max_h / original_height if original_height > 0 else 1
            scale_w = max_w / original_width if original_width > 0 else 1
            scale = min(scale_h, scale_w, 1.0)  # No agrandar más que el original

            subsample_factor = max(1,
                                   int(1 / scale)) if scale > 0 else 10  # Evitar división por cero o subsample muy grande

            self.logo_img_main_menu = temp_logo_img.subsample(subsample_factor, subsample_factor)
        except tk.TclError:
            self.logo_img_main_menu = None

        center_x = canvas_width / 2
        if self.logo_img_main_menu:
            logo_y_pos = canvas_height * 0.25  # Un poco más arriba
            canvas.create_image(center_x, logo_y_pos, image=self.logo_img_main_menu, anchor='center', tags="logo")
            button_y_start = logo_y_pos + (self.logo_img_main_menu.height() / 2) + (
                        canvas_height * 0.08)  # Espacio dinámico
        else:
            button_y_start = canvas_height * 0.35

        colors = {'button1': "#FF5733", 'button2': "#21A35A", 'button3': "#3B80EA", 'button4': "#8E44AD"}
        button_style_main = {'font': ('Helvetica', 14, 'bold'), 'relief': "raised", 'borderwidth': 4, 'cursor': 'hand2'}

        button_width, button_height = 140, 60
        # Ajustar button_width y button_height si la pantalla es muy pequeña
        button_width = max(80, min(button_width, int(canvas_width * 0.15)))
        button_height = max(40, min(button_height, int(canvas_height * 0.1)))

        button_spacing_x = button_width * 0.20

        total_width_row1 = 3 * button_width + 2 * button_spacing_x
        start_x_row1 = center_x - total_width_row1 / 2 + button_width / 2

        btn_widgets = []  # Para aplicar hover

        b1 = Button(canvas, text="Scrapeo", command=lambda: self.scraper_interface(), **button_style_main,
                    bg=colors['button1'], fg="white")
        canvas.create_window(start_x_row1, button_y_start, window=b1, anchor="center", width=button_width,
                             height=button_height, tags="btn_scrapeo")
        btn_widgets.append((b1, colors['button1'], "#FF8566"))

        b2 = Button(canvas, text="Mega Excel", command=lambda: self.excel_interface(), **button_style_main,
                    bg=colors['button2'], fg="white")
        canvas.create_window(start_x_row1 + button_width + button_spacing_x, button_y_start, window=b2, anchor="center",
                             width=button_width, height=button_height, tags="btn_excel")
        btn_widgets.append((b2, colors['button2'], "#49D888"))

        b3 = Button(canvas, text="Merge", command=lambda: self.merge_interface(), **button_style_main,
                    bg=colors['button3'], fg="white")
        canvas.create_window(start_x_row1 + 2 * (button_width + button_spacing_x), button_y_start, window=b3,
                             anchor="center", width=button_width, height=button_height, tags="btn_merge")
        btn_widgets.append((b3, colors['button3'], "#8CB1E8"))

        button_y_row2 = button_y_start + button_height + (canvas_height * 0.05)  # Espacio dinámico
        b4 = Button(canvas, text="Import Odoo", command=lambda: self.import_interface(), **button_style_main,
                    bg=colors['button4'], fg="white")
        canvas.create_window(center_x, button_y_row2, window=b4, anchor="center", width=button_width,
                             height=button_height, tags="btn_odoo")
        btn_widgets.append((b4, colors['button4'], "#B19CD9"))

        for btn, base_color, hover_color in btn_widgets:
            self.hover(btn, hover_color, base_color)

    def start(self, current_view_frame=None):
        # Renombrado argumento para claridad
        first_start = False
        if not self.window:
            self.configure_window()
            first_start = True

        self.window.deiconify()

        if first_start:
            self.window.update_idletasks()
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            win_width = self.WINDOW_WIDTH
            win_height = self.WINDOW_HEIGHT
            x = (screen_width // 2) - (win_width // 2)
            y = (screen_height // 2) - (win_height // 2)
            self.window.geometry(f'{win_width}x{win_height}+{x}+{y}')

        if current_view_frame and current_view_frame.winfo_exists():
            current_view_frame.destroy()

        # Limpiar widgets anteriores de la ventana principal si no son el canvas
        for widget in self.window.winfo_children():
            if isinstance(widget, tk.Canvas) and hasattr(self, 'main_canvas_menu') and widget == self.main_canvas_menu:
                continue  # No destruir el canvas del menú si ya existe y lo vamos a usar
            widget.destroy()

        if not (hasattr(self, 'main_canvas_menu') and self.main_canvas_menu and self.main_canvas_menu.winfo_exists()):
            self.main_canvas_menu = self._create_main_menu_canvas(self.window)

        self.main_canvas_menu.grid(row=0, column=0, sticky="nsew")  # Asegurar que el canvas está en el grid
        self._draw_main_menu(self.main_canvas_menu)

        if first_start:
            self.window.mainloop()

    def _create_main_menu_canvas(self, window):
        canvas = Canvas(window, bg=self.BACKGROUND_COLOR, highlightthickness=0)
        # El canvas usará grid para expandirse, configurado en configure_window
        # El bind se hace aquí para asegurar que solo el canvas del menú tiene este comportamiento
        canvas.bind("<Configure>", lambda event: self._draw_main_menu(event.widget))
        return canvas

    def _clear_window_content(self):
        # Destruye todos los widgets hijos directos de self.window
        for widget in self.window.winfo_children():
            widget.destroy()

    def _setup_common_interface_structure(self, title_text):
        self._clear_window_content()  # Limpiar todo el contenido anterior (incluyendo canvas del menú)

        base_frame = Frame(self.window, bg=self.BACKGROUND_COLOR)
        base_frame.grid(row=0, column=0, sticky="nsew")  # Usa grid para expandir

        header_frame = Frame(base_frame, bg=self.BACKGROUND_COLOR, height=80)
        header_frame.pack(fill='x', padx=20, pady=10)
        header_frame.pack_propagate(False)

        Button(header_frame, command=lambda: self.start(current_view_frame=base_frame),
               **self.BACK_BUTTON_STYLE).pack(side='left', anchor='nw')

        content_outer_frame = Frame(base_frame, bg=self.BACKGROUND_COLOR)
        content_outer_frame.pack(fill='both', expand=True, padx=10, pady=5)  # Contenedor para centrar

        content_frame = Frame(content_outer_frame, bg=self.BACKGROUND_COLOR)
        content_frame.pack(expand=True)  # Este frame se centrará dentro de content_outer_frame

        title_frame = Frame(content_frame, bg=self.BACKGROUND_COLOR, height=50)
        title_frame.pack(fill='x', pady=(0, 20))  # Dentro del content_frame centrado
        title_frame.pack_propagate(False)
        Label(title_frame, text=title_text, **self.title_style).pack(expand=True)

        return content_frame  # Devolver el frame donde se pondrá el contenido específico de la vista

    def scraper_interface(self):
        content_frame = self._setup_common_interface_structure("CONFIGURACIÓN SCRAPER")

        regions_frame = LabelFrame(content_frame, text="Regiones Objetivo", font=('Helvetica', 12, 'bold'), fg='white',
                                   bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        regions_frame.pack(pady=(0, 20))

        region_center_frame = Frame(regions_frame, bg=self.BACKGROUND_COLOR)
        region_center_frame.pack(pady=15)
        region_options = [("🇪🇸 ESPAÑA", self.controller.scraper_controller.SPAIN),
                          ("🇬🇧 UK", self.controller.scraper_controller.UK),
                          ("🇮🇹 ITALIA", self.controller.scraper_controller.ITALIA),
                          ("💡 BUYLED", self.controller.scraper_controller.BUYLED)]
        for i, (text, var) in enumerate(region_options): Checkbutton(region_center_frame, text=text, variable=var,
                                                                     **self.checkbutton_style).grid(row=0, column=i,
                                                                                                    padx=20)

        options_frame = LabelFrame(content_frame, text="Acciones de Scraping", font=('Helvetica', 12, 'bold'),
                                   fg='white', bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        options_frame.pack(pady=(0, 20))

        parent_var, child_var = self.controller.scraper_controller.IF_EXTRACT_ITEM_INFO, self.controller.scraper_controller.IF_ONLY_NEW_PRODUCTS
        child_check = Checkbutton(options_frame, text="Solo nuevos productos", variable=child_var,
                                  **self.checkbutton_style)

        def parent_changed():
            child_check.config(state='disabled' if not parent_var.get() else 'normal')
            if not parent_var.get(): child_var.set(0)

        def child_changed():
            if child_var.get(): parent_var.set(1)

        parent_check = Checkbutton(options_frame, text="Scrapear información productos", variable=parent_var,
                                   command=parent_changed, **self.checkbutton_style)
        parent_check.grid(row=0, column=0, sticky="w", padx=20, pady=8)
        child_check.config(command=child_changed)
        child_check.grid(row=1, column=0, sticky="w", padx=50, pady=8)
        parent_changed()

        execution_frame = Frame(content_frame, bg=self.BACKGROUND_COLOR)
        execution_frame.pack(pady=30)
        execute_btn = Button(execution_frame, text="🚀 Ejecutar Scraping", command=self._execute_scraper_with_validation,
                             font=('Helvetica', 14, 'bold'), bg='#DC3545', fg='white', relief='raised', borderwidth=3,
                             padx=30, pady=10, cursor='hand2')
        execute_btn.pack()
        self.hover(execute_btn, "#C82333", "#DC3545")

    def excel_interface(self):
        content_frame = self._setup_common_interface_structure("CONFIGURACIÓN EXCEL")

        options_frame = LabelFrame(content_frame, text="Acciones Disponibles", font=('Helvetica', 12, 'bold'),
                                   fg='white', bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        options_frame.pack(pady=(0, 10), padx=10)

        options = [
            ("Restaurar de copia seguridad", self.controller.excel_controller.RESTAURAR_BACKUP),
            ("Generar bloque columnas", self.controller.excel_controller.GENERAR_NUEVO_BLOQUE),
            ("Añadir / Actualizar Productos", self.controller.excel_controller.OBTENER_DATOS_PRINCIPAL_HOJAS),
            ("Generar Excel nuevas unidades", self.controller.excel_controller.GENERAR_EXCEL_NUEVAS_UNIDADES),
            ("Rellenar nuevas unidades", self.controller.excel_controller.RELLENAR_NUEVAS_UNIDADES),
            ("Rellenar stocks", self.controller.excel_controller.RELLENAR_STOCKS),
            ("Calcular número de ventas", self.controller.excel_controller.CALCULAR_NUM_VENTAS),
            ("Calcular valor de ventas", self.controller.excel_controller.CALCULAR_VALOR_VENTAS),
            ("Calcular valor de stock", self.controller.excel_controller.CALCULAR_VALOR_STOCK),
            ("Calcular total valores monetarios", self.controller.excel_controller.CALCULAR_TOTAL_VALORES_MONETARIOS),
            ("Rellenar hoja principal", self.controller.excel_controller.RELLENAR_PRINCIPAL)
        ]
        cols = 2
        for i, (text, var) in enumerate(options):
            Checkbutton(options_frame, text=text, variable=var, **self.checkbutton_style).grid(row=i // cols,
                                                                                               column=i % cols,
                                                                                               sticky="w", padx=10,
                                                                                               pady=2)

        sheets_frame = LabelFrame(content_frame, text="Hojas de Trabajo", font=('Helvetica', 12, 'bold'), fg='white',
                                  bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        sheets_frame.pack(pady=(10, 10), padx=10)
        sheet_options = [("BULGARIA", self.controller.excel_controller.BULGARIA),
                         ("UK", self.controller.excel_controller.UK),
                         ("ITALIA", self.controller.excel_controller.ITALIA),
                         ("BUYLED", self.controller.excel_controller.BUYLED),
                         ("MADRID", self.controller.excel_controller.MADRID),
                         ("POLONIA", self.controller.excel_controller.POLONIA)]

        sheet_buttons_frame = Frame(sheets_frame, bg=self.BACKGROUND_COLOR)  # Frame para centrar los botones
        sheet_buttons_frame.pack(pady=5)
        for i, (text, var) in enumerate(sheet_options):
            check = Checkbutton(sheet_buttons_frame, text=text, variable=var, **self.checkbutton_style)
            check.pack(side='left', padx=5, expand=True)

        file_frame = LabelFrame(content_frame, text="Archivo Excel", font=('Helvetica', 12, 'bold'), fg='white',
                                bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        file_frame.pack(fill='x', pady=(10, 10), padx=10)
        Button(file_frame, text="📁 Seleccionar Mega Excel",
               command=lambda: self.select_file(self.controller.excel_controller.MEGA_EXCEL_PATH),
               **self.button_style).pack(side='left', padx=10, pady=5)
        Entry(file_frame, textvariable=self.controller.excel_controller.MEGA_EXCEL_PATH, font=('Helvetica', 11),
              state='disabled', bg='#E8E8E8').pack(side='left', fill='x', expand=True, padx=(0, 10), pady=5, ipady=3)

        execute_frame = Frame(content_frame, bg=self.BACKGROUND_COLOR)
        execute_frame.pack(pady=15)
        execute_btn = Button(execute_frame, text="🚀 Ejecutar Acciones Excel",
                             command=self._execute_excel_with_validation,
                             font=('Helvetica', 14, 'bold'), bg='#28A745', fg='white', relief='raised', borderwidth=3,
                             padx=30, pady=10, cursor='hand2')
        execute_btn.pack()
        self.hover(execute_btn, "#218838", "#28A745")

    def merge_interface(self):
        content_frame = self._setup_common_interface_structure("COMPARACIÓN Y FUSIÓN DE DATOS (MERGE)")

        original_merge_frame = LabelFrame(content_frame, text="Fusión desde Excel de Scrapping",
                                          font=('Helvetica', 12, 'bold'), fg='white',
                                          bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        original_merge_frame.pack(pady=(5, 10), padx=10, fill='x')
        excel_input_frame = Frame(original_merge_frame, bg=self.BACKGROUND_COLOR)
        excel_input_frame.pack(fill='x', padx=5, pady=(10, 5))
        Button(excel_input_frame, text="📄 Seleccionar Excel Scrapping",  # Texto más corto
               command=lambda: self.select_file(self.controller.merge_controller.SCRAPPED_DATA_EXCEL_PATH),
               **self.button_style).pack(side='left', padx=(0, 5))
        Entry(excel_input_frame, textvariable=self.controller.merge_controller.SCRAPPED_DATA_EXCEL_PATH,
              font=('Helvetica', 10), state="disabled", bg='#E8E8E8').pack(side='left', fill='x', expand=True, ipady=3)

        actions_frame_original = Frame(original_merge_frame, bg=self.BACKGROUND_COLOR)
        actions_frame_original.pack(pady=(10, 15))
        compare_btn_style = {**self.button_style, 'font': ('Helvetica', 11, 'bold'), 'bg': '#17A2B8', 'fg': 'white',
                             'padx': 15, 'pady': 6}  # Ligeramente más pequeño
        compare_btn = Button(actions_frame_original, text="📊 Iniciar Comparación",
                             command=self._execute_data_comparison, **compare_btn_style)
        compare_btn.pack(side='left', padx=(0, 10))
        self.hover(compare_btn, "#138496", "#17A2B8")
        generate_merged_btn_style = {**self.button_style, 'font': ('Helvetica', 11, 'bold'), 'bg': '#28A745',
                                     'fg': 'white', 'padx': 15, 'pady': 6}
        generate_merged_btn = Button(actions_frame_original, text="💾 Generar Fusionado",
                                     command=self._execute_final_merge_generation, **generate_merged_btn_style)
        generate_merged_btn.pack(side='left')
        self.hover(generate_merged_btn, "#218838", "#28A745")

        excel_diff_frame = LabelFrame(content_frame, text="Comparación entre Excels Fusionados",
                                      font=('Helvetica', 12, 'bold'), fg='white',
                                      bg=self.BACKGROUND_COLOR, labelanchor='n', padx=15, pady=10)
        excel_diff_frame.pack(pady=(15, 15), padx=10, fill='x')
        file_selection_diff_outer_frame = Frame(excel_diff_frame, bg=self.BACKGROUND_COLOR)
        file_selection_diff_outer_frame.pack(fill='x', padx=5, pady=(10, 5))

        file_selection_diff_frame_old = Frame(file_selection_diff_outer_frame, bg=self.BACKGROUND_COLOR)
        file_selection_diff_frame_old.pack(fill='x', pady=(0, 5))
        Button(file_selection_diff_frame_old, text="📁 Excel Antiguo", **{**self.button_style, 'padx': 10, 'pady': 5},
               # Más pequeño
               command=lambda: self.select_file(self.controller.merge_controller.MERGED_EXCEL_PATH_OLD)).pack(
            side='left', padx=(0, 5))
        Entry(file_selection_diff_frame_old, textvariable=self.controller.merge_controller.MERGED_EXCEL_PATH_OLD,
              font=('Helvetica', 10), state="disabled", bg='#E8E8E8').pack(side='left', fill='x', expand=True, ipady=3)

        file_selection_diff_frame_new = Frame(file_selection_diff_outer_frame, bg=self.BACKGROUND_COLOR)
        file_selection_diff_frame_new.pack(fill='x', pady=(5, 5))
        Button(file_selection_diff_frame_new, text="📁 Excel Nuevo", **{**self.button_style, 'padx': 10, 'pady': 5},
               command=lambda: self.select_file(self.controller.merge_controller.MERGED_EXCEL_PATH_NEW)).pack(
            side='left', padx=(0, 5))
        Entry(file_selection_diff_frame_new, textvariable=self.controller.merge_controller.MERGED_EXCEL_PATH_NEW,
              font=('Helvetica', 10), state="disabled", bg='#E8E8E8').pack(side='left', fill='x', expand=True, ipady=3)

        actions_frame_diff = Frame(excel_diff_frame, bg=self.BACKGROUND_COLOR)
        actions_frame_diff.pack(pady=(10, 15))
        excel_diff_btn_style = {**self.button_style, 'font': ('Helvetica', 12, 'bold'), 'bg': '#FF8C00', 'fg': 'white',
                                'padx': 20, 'pady': 8}
        excel_diff_btn = Button(actions_frame_diff, text="🔍 Comparar y Reportar",
                                command=self._execute_excel_merge_comparison, **excel_diff_btn_style)
        excel_diff_btn.pack()
        self.hover(excel_diff_btn, "#E07B00", "#FF8C00")

    def import_interface(self):
        content_frame = self._setup_common_interface_structure("📥 CONFIGURACIÓN IMPORT ODOO")

        sections = [
            ("📦 Productos", [
                ("Importar productos", self.controller.import_controller.IF_IMPORT_PRODUCTS, 0),
                ("Publicar productos", self.controller.import_controller.IF_PUBLISH_PRODUCTS, 0)]),
            ("💰 Información Comercial", [
                ("Importar Informacion Comercial", self.controller.import_controller.IF_IMPORT_COMERCIAL_INFO, 0)])
        ]

        for section_title, options in sections:
            section_frame = LabelFrame(content_frame, text=section_title, font=('Helvetica', 12, 'bold'),
                                       fg='#FFD700', bg=self.BACKGROUND_COLOR, labelanchor='n', relief='ridge',
                                       borderwidth=2, padx=25, pady=15)
            section_frame.pack(pady=(10, 15), padx=20, ipadx=10, ipady=10)

            for i, (text, var, indent_level) in enumerate(options):
                check = Checkbutton(section_frame, text=text, variable=var, **self.checkbutton_style)
                check.pack(anchor="w", padx=10 + (indent_level * 25), pady=5)

        action_frame = Frame(content_frame, bg=self.BACKGROUND_COLOR)
        action_frame.pack(pady=25)
        import_btn = Button(action_frame, text="🚀 Ejecutar Importación", command=self._execute_import_with_validation,
                            font=('Helvetica', 14, 'bold'), bg='#6F42C1', fg='white', relief='raised', borderwidth=3,
                            padx=30, pady=12, cursor='hand2')
        import_btn.pack()
        self.hover(import_btn, "#5A379E", "#6F42C1")

    def _execute_excel_with_validation(self):
        try:
            if not self.controller.excel_controller.MEGA_EXCEL_PATH.get():
                self.show_info_message("Por favor, selecciona primero un archivo Mega Excel.", "Archivo Requerido")
                return
            if not any(var.get() for var_name, var in self.controller.excel_controller.__dict__.items() if
                       isinstance(var, tk.BooleanVar) and var_name != 'RESTAURAR_BACKUP'):
                self.show_info_message("Por favor, selecciona al menos una acción a realizar.", "Acción Requerida")
                return
            self._create_generic_progress_window("Procesando Excel", "📊 PROCESANDO ARCHIVO EXCEL...")
            self.controller.execute_excel_controller_actions()
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error Inesperado", f"Error: {str(e)}")

    def _execute_scraper_with_validation(self):
        if self.controller.is_critical_operation_active():
            self.show_info_message("Ya hay otra operación crítica en ejecución.", "Operación en Curso")
            return
        if not self._validate_scraper_options():
            self.show_info_message("Seleccione al menos una acción y una región para el scraping.",
                                   "Configuración Incompleta")
            return
        self._create_scraper_progress_window()
        try:
            self.controller.execute_scraper_controller_actions()
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error Inesperado", f"Error: {str(e)}")

    def _validate_scraper_options(self):
        actions = self.controller.scraper_controller.IF_EXTRACT_ITEM_INFO.get()
        regions = any([self.controller.scraper_controller.SPAIN.get(), self.controller.scraper_controller.UK.get(),
                       self.controller.scraper_controller.ITALIA.get(),
                       self.controller.scraper_controller.BUYLED.get()])
        return actions and regions

    def _execute_data_comparison(self):
        from core import App_Merge
        try:
            excel_path = self.controller.merge_controller.SCRAPPED_DATA_EXCEL_PATH.get()
            if not excel_path or not os.path.exists(excel_path):
                self.show_info_message("Selecciona un Excel de scrapping válido.", "Archivo Requerido");
                return
            self._create_generic_progress_window("Comparando Datos", "📊 COMPARANDO DATOS Y GENERANDO HTML...")
            selections_path = self.controller.merge_controller.USER_SELECTIONS_PATH.get()
            App_Merge.execute_comparison_and_launch_html(self.controller.event_manager, excel_path, selections_path)
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error en Comparación", f"Error: {str(e)}")

    def _execute_final_merge_generation(self):
        from core import App_Merge
        try:
            excel_path = self.controller.merge_controller.SCRAPPED_DATA_EXCEL_PATH.get()
            selections_path = self.controller.merge_controller.USER_SELECTIONS_PATH.get()
            if not excel_path or not os.path.exists(excel_path):
                self.show_info_message("Selecciona el Excel de scrapping.", "Archivo Requerido");
                return
            if not selections_path or not os.path.exists(selections_path):
                self.show_info_message(
                    f"Archivo de selecciones no encontrado:\n{selections_path}\nRealiza la comparación primero.",
                    "Archivo Requerido")
                return

            output_filename = filedialog.asksaveasfilename(
                title="Guardar Archivo Fusionado Como", defaultextension=".xlsx",
                initialdir=os.path.join(os.getcwd(), "data", "vtac_merged"),
                filetypes=[("Excel files", "*.xlsx"), ("JSON files", "*.json")])
            if not output_filename: self.show_info_message("Guardado cancelado.", "Cancelado"); return

            self._create_generic_progress_window("Generando Archivo Fusionado", "💾 GENERANDO ARCHIVO FUSIONADO...")
            App_Merge.execute_final_merge_file_generation(self.controller.event_manager, excel_path, selections_path, output_filename)
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error en Generación", f"Error: {str(e)}")

    def _execute_excel_merge_comparison(self):
        from core import App_Merge
        try:
            excel_old = self.controller.merge_controller.MERGED_EXCEL_PATH_OLD.get()
            excel_new = self.controller.merge_controller.MERGED_EXCEL_PATH_NEW.get()
            if not excel_old or not os.path.exists(excel_old):
                self.show_info_message("Selecciona un Excel 'Antiguo (Base)' válido.", "Archivo Requerido")
                return
            if not excel_new or not os.path.exists(excel_new):
                self.show_info_message("Selecciona un Excel 'Nuevo (A comparar)' válido.", "Archivo Requerido")
                return

            output_report_filename = filedialog.asksaveasfilename(
                title="Guardar Reporte de Comparación Como", defaultextension=".xlsx",
                initialdir=os.path.join(os.getcwd(), "data", "vtac_merged", "reports"),
                filetypes=[("Excel files", "*.xlsx")])
            if not output_report_filename: self.show_info_message("Guardado de reporte cancelado.", "Cancelado"); return
            os.makedirs(os.path.dirname(output_report_filename), exist_ok=True)

            self._create_generic_progress_window("Comparando Excels Merge", "🔍 COMPARANDO EXCEL MERGES...")
            selections_diff_path = os.path.join('data', 'common', 'json', 'user_selections_excel_diff.json')
            App_Merge.execute_excel_merge_comparison_and_launch_html(self.controller.event_manager, excel_old, excel_new, selections_diff_path, output_report_filename)
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error en Comparación Excels", f"Error: {str(e)}")

    def _execute_import_with_validation(self):
        try:
            if not self._validate_import_options():
                self.show_info_message("Seleccione al menos una acción para importar.", "Acción Requerida")
                return
            self._create_generic_progress_window("Importando a Odoo", "📥 IMPORTANDO DATOS A ODOO...")
            self.controller.execute_import_controller_action()
        except Exception as e:
            self._close_progress_window(); messagebox.showerror("Error Inesperado", f"Error: {str(e)}")

    def _validate_import_options(self):
        return (self.controller.import_controller.IF_IMPORT_PRODUCTS.get() or
                self.controller.import_controller.IF_PUBLISH_PRODUCTS.get() or
                self.controller.import_controller.IF_IMPORT_COMERCIAL_INFO.get())

    def _create_generic_progress_window(self, title, banner_text):
        if self.progress_window and self.progress_window.winfo_exists(): self.progress_window.destroy()
        self.progress_window = tk.Toplevel(self.window)
        self.progress_window.title(title)
        self.progress_window.geometry("800x500")
        self.progress_window.configure(bg='white')
        self.progress_window.transient(self.window)
        self.progress_window.grab_set()
        self.window.update_idletasks()
        main_x, main_y, main_w, main_h = self.window.winfo_x(), self.window.winfo_y(), self.window.winfo_width(), self.window.winfo_height()
        self.progress_window.update_idletasks()
        prog_w, prog_h = self.progress_window.winfo_width(), self.progress_window.winfo_height()
        pos_x, pos_y = main_x + (main_w // 2) - (prog_w // 2), main_y + (main_h // 2) - (prog_h // 2)
        self.progress_window.geometry(f"+{pos_x}+{pos_y}")

        header_bar = tk.Frame(self.progress_window, bg='#007BFF', height=60)
        header_bar.pack(fill='x')
        header_bar.pack_propagate(False)
        tk.Label(header_bar, text=banner_text, font=('Helvetica', 16, 'bold'), bg='#007BFF', fg='white').pack(expand=True)

        info_bar = tk.Frame(self.progress_window, bg='#E6F2FF', height=50)
        info_bar.pack(fill='x', padx=10, pady=(10, 0))
        info_bar.pack_propagate(False)
        tk.Label(info_bar, text="💡 Procesando... por favor, espera.", font=('Helvetica', 11, 'italic'), bg='#E6F2FF', fg='#0056b3').pack(expand=True)

        text_frame = Frame(self.progress_window, bg='white')
        text_frame.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
        self.progress_text = ScrolledText(text_frame, wrap=tk.WORD, font=('Consolas', 10), bg='#FAFAFA', fg='#333333',
                                          relief='solid', borderwidth=1)
        self.progress_text.pack(fill=tk.BOTH, expand=True)
        status_frame = Frame(self.progress_window, bg='#F5F5F5', height=40)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)
        self.status_label = Label(status_frame, text="⚡ Estado: Iniciando...", font=('Helvetica', 10, 'bold'),
                                  bg='#F5F5F5', fg='#007BFF')
        self.status_label.pack(expand=True)
        self.progress_window.protocol("WM_DELETE_WINDOW", self._on_generic_progress_close)

    def _on_generic_progress_close(self):
        if self.controller.is_critical_operation_active():
            if not messagebox.askyesno("Operación en curso", "Operación en curso. ¿Cancelar y cerrar?",
                                       parent=self.progress_window): return
            self.controller.cancel_current_operation()
        self._close_progress_window()

    def _create_scraper_progress_window(self):  # Similar a generic, pero personalizada
        if self.progress_window and self.progress_window.winfo_exists(): self.progress_window.destroy()
        self.progress_window = tk.Toplevel(self.window)
        self.progress_window.title("🚀 Progreso del Scraping")
        self.progress_window.geometry("800x500")
        self.progress_window.configure(bg='white')
        self.progress_window.transient(self.window)
        self.progress_window.grab_set()
        # (Centrado igual que generic)
        self.window.update_idletasks()
        main_x, main_y, main_w, main_h = self.window.winfo_x(), self.window.winfo_y(), self.window.winfo_width(), self.window.winfo_height()
        self.progress_window.update_idletasks()
        prog_w, prog_h = self.progress_window.winfo_width(), self.progress_window.winfo_height()
        self.progress_window.geometry(
            f"+{main_x + (main_w // 2) - (prog_w // 2)}+{main_y + (main_h // 2) - (prog_h // 2)}")
        header_bar = tk.Frame(self.progress_window, bg='#4CAF50', height=60)
        header_bar.pack(fill='x')
        header_bar.pack_propagate(False)
        tk.Label(
            header_bar,
            text="🚀 PROCESO DE SCRAPING EN CURSO",
            font=('Helvetica', 16, 'bold'),
            bg='#4CAF50',
            fg='white'
        ).pack(expand=True)
        info_bar = tk.Frame(self.progress_window, bg='#E8F5E8', height=50)
        info_bar.pack(fill='x', padx=10, pady=(10, 0))
        info_bar.pack_propagate(False)
        tk.Label(
            info_bar,
            text="💡 Mantenga esta ventana abierta.",
            font=('Helvetica', 11, 'italic'),
            bg='#E8F5E8',
            fg='#2E7D32'
        ).pack(expand=True)
        text_frame = Frame(self.progress_window, bg='white')
        text_frame.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
        self.progress_text = ScrolledText(text_frame, wrap=tk.WORD, font=('Consolas', 10), bg='#FAFAFA', fg='#333333',
                                          relief='solid', bd=1)
        self.progress_text.pack(fill=tk.BOTH, expand=True)
        status_frame = Frame(self.progress_window, bg='#F5F5F5', height=40)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)
        self.status_label = Label(status_frame, text="⚡ Estado: Iniciando...", font=('Helvetica', 10, 'bold'),
                                  bg='#F5F5F5', fg='#1976D2')
        self.status_label.pack(expand=True)
        self.progress_window.protocol("WM_DELETE_WINDOW", self._on_scraper_progress_window_close)

    def _on_scraper_progress_window_close(self):
        if self.controller.is_scraping_active():
            if messagebox.askyesno("Scraping en curso", "Scraping en progreso. ¿Cancelar?",
                                   parent=self.progress_window):
                self.controller.cancel_current_operation()
        else:
            self._close_progress_window()

    @classmethod
    def select_file(cls, file_path_var):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("JSON files", "*.json"), ("All files", "*.*")])
        if filename: file_path_var.set(filename)

    @classmethod
    def select_dir(cls, dir_path_var):
        dirname = filedialog.askdirectory()
        if dirname: dir_path_var.set(dirname)

    def _on_main_window_close(self):
        if self.controller.is_critical_operation_active():
            if messagebox.askyesno("Operación en Curso", "Operación crítica en curso. ¿Salir y cancelar?",
                                   parent=self.window):
                self.controller.cancel_current_operation()
                if self.window and self.window.winfo_exists(): self.window.destroy()
            else:
                return
        else:
            if messagebox.askyesno("Salir", "¿Seguro que desea salir?", parent=self.window):
                if self.window and self.window.winfo_exists(): self.window.destroy()

    def show_info_message(self, text: str = '',
                          title: str = 'Info'):  # Movido aquí para que esté definido antes de usarlo
        messagebox.showinfo(title, text)

    # El resto de métodos como show_modal_message, show_selection_message, etc., no necesitan cambios inmediatos para el responsive
    # pero deben asegurar que usan self.window como parent si crean Toplevels.


if __name__ == "__main__":
    os.makedirs(os.path.join('data', 'common', 'json'), exist_ok=True)
    os.makedirs(os.path.join('data', 'common', 'html_comparison'), exist_ok=True)
    os.makedirs(os.path.join('data', 'vtac_merged'), exist_ok=True)
    os.makedirs(os.path.join('data', 'vtac_merged', 'reports'), exist_ok=True)

    root = tk.Tk()
    root.withdraw()

    from controller import Controller

    main_controller = Controller(master=root)
    app_interface = Interface(main_controller)

    app_interface.start()
