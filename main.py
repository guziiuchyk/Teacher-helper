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
            with open(self.config_file, 'r') as file:
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
    def __init__(self, parser, config, app):
        self._app = app
        self.parser = parser
        self.config = config
        self.selected_fields = self._initialize_selected_fields()
        self.table = self._initialize_table()
        self.indexes_list = self._initialize_indexes()
        self._process_data()

    def _initialize_selected_fields(self):
        fields_list = [["Opiskelijan nimi"]]

        for n,i in enumerate(self._app.gui.checkboxes_list):
            if i.get() == 1:
                element = [i.cget("text")]
                if len(self._app.gui.entry_list[n].get()) != 0:
                    element.append(self._app.gui.entry_list[n].get())
                fields_list.append(element)
        return fields_list

    def _initialize_table(self):
        table = {}
        table["Opiskelijan nimi"] = [] 
        for i in self.selected_fields:
            print(i)
            table[i[-1]] = []
        return table

    def _initialize_indexes(self):
        indexes_list = []
        for i in self.selected_fields:
            if i[0] in self.parser.fields_list:
                indexes_list.append(self.parser.fields_list.index(i[0]))
        return indexes_list

    def _delete_spaces(self, text):
        return ''.join(char for char in text if char.isalnum())

    def _process_data(self):
        print(self.indexes_list)
        print(self.selected_fields)
        for n,row in enumerate(self.parser.progress_list):
            if n == 0:
                continue
            elements = row.find_all("td")
            for idx, selected_field in enumerate(self.indexes_list):
                field_name = self.selected_fields[idx][-1]
                element = elements[selected_field]  
                value = element.a.text if element.find("a") else self._delete_spaces(element.text)
                value = "X" if value.lower() == "o" else value
                self.table[field_name].append(value)

class ExcelWriter:
    def __init__(self, table, total_lines, filename="students"):
        self.table = table
        self.df = pd.DataFrame(table)
        self.total_lines = total_lines
        if  len(self.df) < self.total_lines:
            empty_rows = self.total_lines - len(self.df)
            empty_data = pd.DataFrame([[""] * len(self.df.columns)] * empty_rows, columns=self.df.columns)
            self.df = pd.concat([self.df, empty_data], ignore_index=True)
        self.filename = filename
        self._write_to_excel()

    def _write_to_excel(self):
        with pd.ExcelWriter(self.filename+".xlsx", engine='openpyxl') as writer:
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
            
            entry = CTk.CTkEntry(master=checkbox_frame, state=state,font=(self._FONT, 18), fg_color="#D5D5D5",text_color="black")
            entry.grid(row=n, column=1, sticky="w", padx=(10,0))
            entry.bind("<Return>", self._app.focus_next_entry)
            self.entry_list.append(entry)
        
        self.file_name_frame = CTk.CTkFrame(master=self,fg_color=self._WHITE_COLOR,bg_color=self._WHITE_COLOR)
        self.file_name_frame.grid(row=3, column=0,sticky="ne",padx=(0,100))
        self.file_name_frame_text = CTk.CTkLabel(master=self.file_name_frame, text="Write the file name", font=(self._FONT, 24),text_color="black")
        self.file_name_frame_text.grid(row=0,column=0)
        self.file_name_frame_entry = CTk.CTkEntry(master=self.file_name_frame, text_color="black",fg_color=self._WHITE_COLOR, font=(self._FONT,18))
        self.file_name_frame_entry.grid(row=1,column=0, pady=(5,0))
        
        self.column_count_frame = CTk.CTkFrame(master=self,fg_color=self._WHITE_COLOR,bg_color=self._WHITE_COLOR)
        self.column_count_frame.grid(row=3, column=0,sticky="se",padx=(0,66), pady=(0,150))
        self.column_count_frame_text = CTk.CTkLabel(master=self.column_count_frame, text="Write the number of lines", font=(self._FONT, 24),text_color="black")
        self.column_count_frame_text.grid(row=0,column=0)
        self.column_count_frame_entry = CTk.CTkEntry(master=self.column_count_frame, text_color="black",fg_color=self._WHITE_COLOR, font=(self._FONT,18))
        self.column_count_frame_entry.grid(row=1,column=0, pady=(5,0))
        
        self.back_button = CTk.CTkButton(master=self,command=lambda: self._app.change_window(0),hover_color=self._HOVER_PURPLE_COLOR, text="Back", fg_color=self._PURPLE_COLOR, font=(self._FONT, 18), bg_color=self._WHITE_COLOR, width=70, border_width=1, border_color="black", text_color="black")
        self.back_button.grid(row=4, column=0, sticky="ws", padx=(10,0), pady=(5,0))
        self.back_button = CTk.CTkButton(master=self,command=self._app.compilate_data,hover_color=self._HOVER_PURPLE_COLOR, text="Next", fg_color=self._PURPLE_COLOR, font=(self._FONT, 18), bg_color=self._WHITE_COLOR, width=70, border_width=1, border_color="black", text_color="black")
        self.back_button.grid(row=4, column=0, sticky="se", padx=(0,10), pady=(5,0))

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
    
    def _write_to_excel(self):
        filename = self.gui.file_name_frame_entry.get()
        total_lines = self.gui.column_count_frame_entry.get()
        try:
            total_lines = int(total_lines)
        except:
            total_lines = 0
        if len(filename) == 0:
            filename = "students"
        self.excel_writer = ExcelWriter(self._data_manager.table,total_lines , filename)

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
            if value == 1:
                self.gui.checkboxes_list[n].select()
            else:
                self.gui.checkboxes_list[n].deselect()
            if value == 1:
                self.gui.entry_list[n].configure(state=state, fg_color=self.gui._WHITE_COLOR)
            else:
                self.gui.entry_list[n].delete(0, CTk.END)
                self.gui.entry_list[n].configure(state=state, fg_color="#D5D5D5")

    def compilate_data(self):
        self._data_manager = DataManager(self.html_parser, self.config_manager,self)
        print(self._data_manager.table)
        self._write_to_excel()

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
    
    def change_window(self, index):
        if index==0:
            self.gui.load_menu()
if __name__ == "__main__":
    app = App()