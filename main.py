from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import PatternFill, Border, Side, Alignment
import json

print("""
██╗    ██╗███████╗██╗      ██████╗ ██████╗ ███╗   ███╗███████╗    
██║    ██║██╔════╝██║     ██╔════╝██╔═══██╗████╗ ████║██╔════╝    
██║ █╗ ██║█████╗  ██║     ██║     ██║   ██║██╔████╔██║█████╗    
██║███╗██║██╔══╝  ██║     ██║     ██║   ██║██║╚██╔╝██║██╔══╝      
╚███╔███╔╝███████╗███████╗╚██████╗╚██████╔╝██║ ╚═╝ ██║███████╗    
 ╚══╝╚══╝ ╚══════╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝                                        
""")
print("Version: 1.2")
print("Create a file named html.txt and paste the HTML code there.")

class FileManager:
    def __init__(self, html_file="html.txt", config_file = "config.json"):
        self.html_file = html_file
        self.config_file = config_file
        self.html = self._load_html()
        self.config = self._load_config()

    def _load_html(self):
        try:
            with open(self.html_file, "r", encoding="utf-8") as file:
                return file.read()
        except FileNotFoundError:
            print("Html file not found")
            input()
            exit()

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
        if config is None:
            self.is_select_all = True
            self.selected_fields = []
            self.displayed_fields = []
        else:
            self.is_select_all = config["is_select_all"]
            self.selected_fields = config["selected_fields"]
            self.displayed_fields = config["displayed_fields"]

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

class Gui: 
    pass

class App:
    def __init__(self):
        print("Starting application...")
        self.file_manager = FileManager()
        self.config_manager = ConfigManager(self.file_manager.config)
        self.html_parser = HtmlParcer(self.file_manager.html)
        self.data_processor = DataManager(self.html_parser, self.config_manager)
        self.excel_writer = ExcelWriter(self.data_processor.table)

if __name__ == "__main__":
    app = App()
    print("If successful, check the file students.xlsx in the same folder.")
    print("Powered by guziiuchyk@gmail.com <3")
    input()