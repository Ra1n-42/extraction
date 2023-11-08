import customtkinter
import tkinter.filedialog as filedialog
from functions import *
from CTkToolTip import *


class ScrollableLabelButtonFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)
        self.tooltip = CTkToolTip
        self.command = command
        self.label_list = []
        self.tooltip_list = []

    def add_item(self, item, image=None):
        label = customtkinter.CTkLabel(self, text=item, image=image, compound="left", padx=5, anchor="w")
        label.grid(row=len(self.label_list), column=0, pady=(0, 5), sticky="w")
        self.label_list.append(label)


    def remove_item(self, item):
        for label in self.label_list:
            if item == label.cget("text"):
                label.destroy()
                self.label_list.remove(label)
                return

    def coloring_item(self, index, color="green"):
        for key, label in enumerate(self.label_list):
            if key == index:
                label.configure(text_color=color)

    def tooltip_item(self, index, text):
        for key, label in enumerate(self.label_list):
            if key == index:
                tool = self.tooltip(label, message=text)
                self.tooltip_list.append(tool)
                tool.hide()

    def tooltip_change_text(self, index, value, color=None):
        for key, tooltip in enumerate(self.tooltip_list):
            if key == index:
                tooltip.configure(message=value, text_color=color or "red")
                tooltip.show()


class ToplevelWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, text=None, posx=None, posy=None, **kwargs):
        super().__init__(*args, **kwargs)
        if posx is not None and posy is not None:
            self.geometry(f"420x250+{posx + 105}+{posy + 10}")
        else:
            self.geometry(f"420x250")
        self.title("Info")
        self.label = customtkinter.CTkLabel(self, text=text, font=customtkinter.CTkFont(size=12, weight="bold"))
        self.label.grid(row=0, column=0, padx=(20, 0), pady=(20, 0))
        self.label.focus_set()


class App(customtkinter.CTk):

    def __init__(self):
        super().__init__()

        # Größe des Hauptfensters festlegen
        window_width = 645
        window_height = 290

        # Position berechnen, um das Hauptfenster mittig zu platzieren
        self.x = (self.winfo_screenwidth() - window_width) // 2
        self.y = (self.winfo_screenheight() - window_height) // 2
        self.title("Extraction")
        self.geometry(f"{window_width}x{window_height}+{self.x}+{self.y}")

        self.Load_variables()
        self.Load_Frames()

    def Load_variables(self):
        self.toplevel_window = None
        self.resizable(False, False)
        self.current_progress: float = 0.0
        self.max_progressbar = float

    def Load_Frames(self):

    
        self.sidebar_frame = customtkinter.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Einstellungen",
                                                 font=customtkinter.CTkFont(size=20, weight="bold")).grid(row=1,
                                                                                                          column=0,
                                                                                                          padx=20,
                                                                                                          pady=(20, 10))
        self.Ordner = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_find_source_folder,
                                              text="Source-Ordner")
        self.Ordner.grid(row=2, column=0, padx=10, pady=(0, 10))

        self.Excel_Datei = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_find_file,
                                                   text="Zieldatei", fg_color="grey", state="disabled")
        self.Excel_Datei.grid(row=3, column=0, padx=0, pady=(0, 10))

        self.Speicherort_Fehlermeldungen = customtkinter.CTkButton(self.sidebar_frame, command=self.button_error_save_folder,
                                                   text="ERROR-Ordner")
        self.Speicherort_Fehlermeldungen.grid(row=4, column=0, padx=0, pady=(0, 0))

        self.Start = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_start_event,
                                             text="Start", fg_color="grey", state="disabled")
        self.Start.grid(row=5, column=0, padx=0, pady=(0, 20))


        self.scrollable_label_button_frame = ScrollableLabelButtonFrame(master=self, width=400, corner_radius=10)
        self.scrollable_label_button_frame.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")

        self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.slider_progressbar_frame.grid(row=2, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
        self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
        self.progressbar_1.grid(row=0, column=0, padx=(20, 0), pady=(0, 10), sticky="ew")
        self.progressbar_1.set(self.current_progress)
        self.attribute_value = "von Yoe \n\nfür RITTER Starkstromtechnik Magdeburg GmbH & Co. KG\n\nVersion 1.5 \n\nFunktion: Lesen und schreiben von Datensätze auf Zieldatei\nAutomatische Ausgabe von Fehlern via Erstellung \n.txt Dateien.\n\nNach Start:\nMit der Maus über die Dateien um fehlerhafte \nTabellen zu sehen"
        self.Help_btn = customtkinter.CTkButton(self.sidebar_frame, command=self.open_toplevel,
                                                text="Info", height=20, width=50)
        self.Help_btn.grid(row=5, column=0, padx=0, pady=(90, 0))
        self.Help_btn.configure(textvariable=self.attribute_value)
        self.activate_buttons()

        self.Status_Label = customtkinter.CTkLabel(self, text="Status: Standby")
        self.Status_Label.grid(row=1, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")


    def open_toplevel(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ToplevelWindow(self, text=self.Help_btn.cget("textvariable"), posx=self.x,
                                                  posy=self.y)
            self.toplevel_window.lift()
        else:
            self.toplevel_window.lift()

    def sidebar_button_find_file(self):
        selected_file = filedialog.askopenfilename()
        if selected_file:
            self.write_in_config("file_path", selected_file)
        self.activate_buttons()

    def write_in_config(self, key, value):
        with open('config.json', 'w', encoding='utf-8') as tmp_file:
            config["data"] = []
            config["meldungen"] = []
            config["notexisting"] = []
            config[key] = value
            json.dump(config, tmp_file, indent=4, ensure_ascii=False)

    def button_error_save_folder(self):
        selected_folder = filedialog.askdirectory()
        if selected_folder:
            self.write_in_config("error_path", selected_folder)
        self.activate_buttons()

    def sidebar_button_find_source_folder(self):
        selected_folder = filedialog.askdirectory()
        if selected_folder:
            self.write_in_config("source_path", selected_folder)
        self.activate_buttons()

    def activate_buttons(self):
        if config.source_path != "":
            self.Ordner.configure(fg_color="green", hover_color="#004e00")
            self.Excel_Datei.configure(fg_color="#1f6aa5", state="normal")
            self.load_files()

        if config.file_path != "":
            self.Excel_Datei.configure(fg_color="green", hover_color="#004e00")
            self.Start.configure(state="normal")

        if config.file_path != "" and config.source_path != "":
            self.Start.configure(fg_color="#1f6aa5", hover_color="#133e61")

        if config.error_path != "":
            self.Speicherort_Fehlermeldungen.configure(fg_color="green", hover_color="#004e00")

    def Refresh(self):
        for label in self.scrollable_label_button_frame.label_list:
            label.destroy()
        self.scrollable_label_button_frame.label_list.clear()

    def Update_Progress_Bar(self, max):
        self.current_progress += 1.0
        self.progressbar_1.set(self.current_progress / max)

    def verbinde_pfad_mit_datei(self, error_file):
        error_path = config.error_path
        if error_path:  # Überprüfe, ob error_path nicht leer ist
            file_path = os.path.join(error_path, error_file)
        else:
            file_path = error_file

        return file_path

    def oefne_Zieldatei_und_schreibe_aus_Datenbank(self):
        __update_count = 0
        workbook = load_workbook(EXTRACTED_FILE, data_only=True)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            __update_count += update_to_file(workbook.sheetnames, sheet, sheet_name)
        workbook.save(EXTRACTED_FILE)
        workbook.close()
        if len(config.notexisting) > 0:
            with open(self.verbinde_pfad_mit_datei("Fehlerhafte Stundenzettel.txt"), "w") as file:
                for message in config.notexisting:
                    error_messages = re.sub(r'.*?m', '', message)
                    file.write(error_messages + "\n")
        return __update_count

    def Upate_Progressbar(self):
        pass

    def Lese_Daten_aus_Source_path(self):
        self.Status_Label.configure(text="Status: Dateien werden gelesen.")
        self.Start.configure(state="disabled")
        _temp = 0
        if config.file_path != "" and config.source_path != "":  # Sicherheitshalber
            excel_file_list = [filename for filename in os.listdir(config.source_path) if
                                              filename[-1] != '#' and '.xlsx' in filename]
            enum = enumerate(excel_file_list)
            for index, filename in enum:
                start = time.time()
                path = config.source_path + '\\' + filename
                workbook = load_workbook(filename=path, data_only=True, read_only=True)
                temp = workbook.sheetnames
                for sheet_name in workbook.sheetnames:
                    Kostenstelle_in_Datei = getDataFromWeekWorkTime(workbook[sheet_name], filename)
                    if Kostenstelle_in_Datei or Kostenstelle_in_Datei is None:
                        temp.remove(sheet_name)
                print(time.time() - start)
                if len(temp) == 0:
                    self.scrollable_label_button_frame.tooltip_change_text(index, "Hier ist alles Okay", color="green")
                else:
                    self.scrollable_label_button_frame.tooltip_change_text(index, str(temp))
                workbook.close()
                self.scrollable_label_button_frame.coloring_item(index, color="#93c4cc")
                self.Update_Progress_Bar(len(excel_file_list))
            self.Status_Label.configure(text="Status: Dateien werden in Zieldatei überschrieben.")

            _count = self.oefne_Zieldatei_und_schreibe_aus_Datenbank()
            self.current_progress = 0.0
            for i in range(_count):
                self.Update_Progress_Bar(_count)
            _unique_set = list(set(config.meldungen))
            if len(_unique_set) > 0:
                with open(self.verbinde_pfad_mit_datei("fehlermeldungen.txt"), 'w') as file:
                    for _ in _unique_set:
                        error_messages = re.sub(r'.*?m', '', _)
                        file.write(error_messages + "\n")
            self.Status_Label.configure(text=""+"{} Fehlermeldung/en ; {}/{} geupdatet".format(len(_unique_set), _count, len(config.data))+" > Fertig.")

        else:
            self.Status_Label.configure(text="Fehlercode: 221")
        self.Start.configure(state="normal")

    def sidebar_button_start_event(self):
        threading.Thread(target=self.Lese_Daten_aus_Source_path, daemon=True).start()

    def load_files(self):
        self.Refresh()
        for index, filename in enumerate([filename for filename in os.listdir(config.source_path) if
                                          filename[-1] != '#' and '.xlsx' in filename]):
            self.scrollable_label_button_frame.add_item(filename)
            self.scrollable_label_button_frame.tooltip_item(index, "")
        if len(self.scrollable_label_button_frame.label_list) < 1:
            self.scrollable_label_button_frame.add_item("Keine Daten gefunden, wähle ein anderen Ordner.")

    def label_button_frame_event(self, item):
        print(f"label button frame clicked: {item}")


if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    app = App()
    app.mainloop()
