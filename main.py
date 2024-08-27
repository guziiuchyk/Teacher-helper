from typing import Tuple
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import PatternFill, Border, Side, Alignment
import json
import customtkinter as CTk

class FileManager:
    def __init__(self, config_file = "config.json"):
        self.config_file = config_file
        self.config = self._load_config()

    def _load_config(self):
        try: 
            with open(self.config_file, 'r') as file: #Try to read config file
                return json.loads(file.read()) 
        except FileNotFoundError:
            print("W: Configuration file not found, the default configuration set will be used.")
            input() 
            exit()
        except json.JSONDecodeError:
            print("Error: An error occurred while reading the configuration file.")
            input()
            exit()

class ConfigManager:
    def __init__(self, config):
            self.is_select_all = config["is_select_all"]
            self.selected_fields = config["selected_fields"]
            self.displayed_fields = config["displayed_fields"]
            self.templates = config["templates"]

class HtmlParcer:
    def __init__(self, html):
        self.soup = BeautifulSoup(html, "html.parser")
        self.fields_list = self._parse_fields()
        self.progress_list = self._parse_progress()

    def _parse_fields(self):
        list = [field.text for field in self.soup.thead.find_all("span")]
        list.insert(0, "Opiskelijan nimi")
        return(list)

    def _parse_progress(self):
        return self.soup.tbody.find_all("tr")

class DataManager:
    def __init__(self, parser, config):
        self.parser = parser
        self.config = config
        self.table = self._initialize_table()
        self.indexes_list = self._initialize_indexes()
        self._process_data()

    def _initialize_table(self):
        if self.config.is_select_all:
            return {field: [] for field in ["Opiskelijan nimi"] + self.parser.fields_list}
        else:
            return {self.config.displayed_fields[i]: [] for i in range(len(self.config.selected_fields))}

    def _initialize_indexes(self):
        print(self.parser.fields_list)
        if not self.config.is_select_all:
            return [self.parser.fields_list.index(field) for field in self.config.selected_fields]
        return []

    def _delete_spaces(self, text):
        return ''.join(char for char in text if char.isalnum())

    def _process_data(self):
        for n, row in enumerate(self.parser.progress_list):
            elements = row.find_all("td")
            if n == 0:
                continue
            for n2, element in enumerate(elements):
                if not self.config.is_select_all and n2 not in self.indexes_list:
                    continue

                field_name = self._get_field_name(n2)
                value = element.a.text if element.find("a") else self._delete_spaces(element.text)
                value = "X" if value.lower() == "o" else value
                self.table[field_name].append(value)

    def _get_field_name(self, n2):
        if self.config.is_select_all:
            return self.parser.fields_list[n2]
        return self.config.displayed_fields[self.indexes_list.index(n2)]

class ExcelWriter:
    def __init__(self, table):
        self.table = table
        self.df = pd.DataFrame(table)
        self._write_to_excel()

    def _write_to_excel(self):
        with pd.ExcelWriter('students.xlsx', engine='openpyxl') as writer:
            self.df.to_excel(writer, sheet_name='Table', index=False)
            work_sheet = writer.sheets['Table']
            self._adjust_columns(work_sheet)
            self._apply_styles(work_sheet)

    def _adjust_columns(self, sheet):
        max_length = 0
        for col in sheet.columns:
            if col[0].column_letter == 'A':
                for cell in col:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[col[0].column_letter].width = adjusted_width

    def _apply_styles(self, sheet):
        fill = PatternFill(start_color="828181", end_color="828181", fill_type='solid')
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')

        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = center_alignment
                if cell.row % 2 == 0:
                    cell.fill = fill

class Gui(CTk.CTk): 
    def __init__(self, app, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._app = app

        self._FONT = "Inter"
        self._WHITE_COLOR = "#EBEBEB"
        self._BLUE_COLOR = "#1256D3"
        self._PURPLE_COLOR = "#5A4AEA"
        self._HOVER_PURPLE_COLOR = "#4533b5"

        self.geometry("800x500")
        self.title("Teacher helper")
        self._set_appearance_mode("light")
        self.resizable(height=False, width=False)

        

        self.load_menu() 

    def load_menu(self):
        self._clear()
        self._load_header()
        self.menu_text = CTk.CTkLabel(master=self,fg_color=self._WHITE_COLOR,text_color="black",text="Copy code, and click the button", font=(self._FONT,32))
        self.menu_text.grid(row=1, column=0, pady=(90, 0))
        self.menu_button = CTk.CTkButton(master=self, command=self._app.menu_button_handle,hover_color=self._HOVER_PURPLE_COLOR,width=170,height=45, text_color="black",corner_radius=11,border_width=1,border_color="black",text="Paste",font=(self._FONT,30), fg_color=self._PURPLE_COLOR, bg_color=self._WHITE_COLOR)
        self.menu_button.grid(row=2, column=0, pady=(10,0))

    def load_main(self):
        self._clear()
        self._load_header()

        BUTTON_WIDTH = 80
        BUTTON_PADDING = 5
        BUTTON_FONT_SIZE = 18

        hero_text = CTk.CTkLabel(master=self, fg_color=self._WHITE_COLOR, text_color="black", text="Choose template", font=(self._FONT, 24))
        hero_text.grid(row=1,column=0, sticky="w", padx=(7,0), pady=(7,0))
        template_clear = CTk.CTkButton(master=self,command=lambda: self._app.select_all_checkboxes(0),hover_color=self._HOVER_PURPLE_COLOR,width=80,height=20, text_color="black",corner_radius=11,border_width=1,border_color="black",text="Clear",font=(self._FONT,BUTTON_FONT_SIZE), fg_color=self._PURPLE_COLOR, bg_color=self._WHITE_COLOR)
        template_clear.grid(row=2, column=0, sticky="w", padx=(BUTTON_PADDING,0),pady=(BUTTON_PADDING, 0))
        template_all = CTk.CTkButton(master=self, command=lambda: self._app.select_all_checkboxes(1),hover_color=self._HOVER_PURPLE_COLOR,width=80,height=20, text_color="black",corner_radius=11,border_width=1,border_color="black",text="All",font=(self._FONT,BUTTON_FONT_SIZE), fg_color=self._PURPLE_COLOR, bg_color=self._WHITE_COLOR)
        template_all.grid(row=2, column=0,sticky="w", padx=(BUTTON_WIDTH+BUTTON_PADDING*2,0),pady=(BUTTON_PADDING, 0))

        for n,i in enumerate(self._app.config_manager.templates):
            button = CTk.CTkButton(master=self,command=lambda i=i: self._app.select_checkboxes_by_template(i),hover_color=self._HOVER_PURPLE_COLOR,width=80,height=20, text_color="black",corner_radius=11,border_width=1,border_color="black",text=i,font=(self._FONT,BUTTON_FONT_SIZE), fg_color=self._PURPLE_COLOR, bg_color=self._WHITE_COLOR)
            button.grid(row=2, column=0,sticky="w", padx=(BUTTON_WIDTH*(n+2)+BUTTON_PADDING*(n+3),0),pady=(BUTTON_PADDING, 0))
        
        template_add = CTk.CTkButton(master=self,hover_color=self._HOVER_PURPLE_COLOR,width=30,height=20, text_color="black",corner_radius=11,border_width=1,border_color="black",text="+",font=(self._FONT,BUTTON_FONT_SIZE), fg_color=self._PURPLE_COLOR, bg_color=self._WHITE_COLOR)
        template_add.grid(row=2, column=0,sticky="w",padx=(345,0),pady=(BUTTON_PADDING,0))

        checkbox_frame = CTk.CTkScrollableFrame(master=self,fg_color="#D5D5D5",bg_color=self._WHITE_COLOR, width=380, height=300)
        checkbox_frame.grid(row=3, column=0, sticky="w", padx=(10,0), pady=(10,0))

        self.checkboxes_list = []
        self.entry_list = []

        for n,i in enumerate(self._app.html_parser.fields_list):
            if n == 0:
                continue

            state = "disabled"

            checkbox_var = CTk.IntVar(value=0)
            checkbox_var.trace_add("write", self._app.on_select_checkbox)
            checkbox = CTk.CTkCheckBox(master=checkbox_frame,variable=checkbox_var, text=i, font=(self._FONT,16), text_color="black", checkbox_width=20, checkbox_height=20)
            checkbox.grid(row=n, column=0, sticky="w")
            self.checkboxes_list.append(checkbox)
            #self.checkboxes_list.append(checkbox_var)
            
            entry = CTk.CTkEntry(master=checkbox_frame, state=state, fg_color="#D5D5D5",text_color="black")
            entry.grid(row=n, column=1, sticky="w", padx=(10,0))
            entry.bind("<Return>", self._app.focus_next_entry)
            self.entry_list.append(entry)
        
    def _clear(self):
        for e in self.winfo_children():
            e.destroy()

    def _load_header(self):
        header_frame = CTk.CTkFrame(master=self,border_width=0, fg_color=self._BLUE_COLOR, width=800, height=73, corner_radius=0)
        header_frame__text = CTk.CTkLabel(master=self, text="Teacher helper", bg_color=self._BLUE_COLOR,font=(self._FONT,40))
        header_frame__text.grid(row=0, column=0)
        header_frame.grid(row=0, column=0)
class App:
    def __init__(self):
        print("Starting application...")
        self.config_manager = None
        self.html_parser = None
        self.gui = Gui(self)
        self._start()
        self.gui.mainloop()

    def _start(self):
        self._file_manager = FileManager()
        self.config_manager = ConfigManager(self._file_manager.config)
    
    def _parse_html(self):
        self.html_parser = HtmlParcer(self._html)
        self._data_manager = DataManager(self.html_parser, self.config_manager)
    
    def _write_to_excel(self):
        self.excel_writer = ExcelWriter(self._data_manager.table)

    def on_select_checkbox(self, *args):
        index = int(args[0][6:])
        element = self.gui.checkboxes_list[index]
        value = element.get()
        if value == 1:
            self.gui.entry_list[index].configure(state="normal", fg_color=self.gui._WHITE_COLOR)
        else:
            self.gui.entry_list[index].delete(0, CTk.END)
            self.gui.entry_list[index].configure(state="disabled", fg_color="#D5D5D5")
    def select_all_checkboxes(self, value):
        state = "normal"
        if value != 1:
            state = "disabled"
        for n in range(0, len(self.gui.checkboxes_list)):
            #self.gui.checkboxes_list[n].set(value)
            if value == 1:
                self.gui.checkboxes_list[n].select()
            else:
                self.gui.checkboxes_list[n].deselect()
            if value == 1:
                self.gui.entry_list[n].configure(state=state, fg_color=self.gui._WHITE_COLOR)
            else:
                self.gui.entry_list[n].delete(0, CTk.END)
                self.gui.entry_list[n].configure(state=state, fg_color="#D5D5D5")
                
    def menu_button_handle(self):
        try:
            self._html = self.gui.clipboard_get()
        except:
            print("Cant get data from clipboard")
            return
        try:
            self._parse_html()
        except:
            print("Wrong html code")
            return
        self.gui.load_main()
        #print(self._data_manager.table)
    
    def select_checkboxes_by_template(self, name):
        print(name)
        template = self.config_manager.templates[name]
        self.select_all_checkboxes(0)
        for i in template:
            for n2,j in enumerate(self.gui.checkboxes_list):
                text = j.cget("text")
                if text == i[0]:
                    j.select()
                    self.gui.entry_list[n2].configure(state="normal", fg_color=self.gui._WHITE_COLOR)
                    if len(i) == 2:
                        self.gui.entry_list[n2].delete(0, CTk.END)
                        self.gui.entry_list[n2].insert(0, i[1])

    def focus_next_entry(self, event):
        current_widget = event.widget

        for i, ctk_entry in enumerate(self.entry_list):
            if ctk_entry._entry == current_widget:
                next_index = i + 1
                while next_index < len(self.entry_list):
                    next_widget = self.entry_list[next_index]
                    if next_widget.cget("state") == "normal":
                        next_widget.focus_set()
                        break
                    next_index += 1
                break
if __name__ == "__main__":
    app = App()