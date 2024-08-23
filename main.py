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

config = None
html = None

try:
    with open("html.txt", "r", encoding="utf-8") as file: #Open and read html.txt file
        html = file.read()
except FileNotFoundError:
    print("Error: html.txt file not found. Please create it.")
    input()
    exit()

filename = "config.json" #Name of config file

try: 
    with open(filename, 'r') as file: #Try to read config file
        config = json.loads(file.read()) 
except FileNotFoundError:
    print("Warning: Configuration file not found, the default configuration set will be used.")
except json.JSONDecodeError:
    print("Warning: An error occurred while reading the configuration file.")

print("Press Enter to continue if you are ready.")
input() #Empty input so that the program waits until Enter is pressed

#Fields that should be in the Excel table
#base_selected_fields = ["YHTmaluo18.7","YHTmaluo18.8","YHTmaluo18.9","YHTmaluo18.10",
#                   "YHTmaluo18.1","YHTmaluo18.11","YHTmaluo18.12","YHTmaluo18.3"]
selected_fields = []

#These are the names that the fields will be replaced with later
#(The name and desired name should be at the same index)
#base_displayed_fields = ["I","II","III","IV","Ma","Fy","Ke","FyKe",]
displayed_fields = []

#Will program select all categries
is_select_all = False
 
try:
    if config is None: #If config file is not found
        is_select_all = True
    else: #Get values from config file
        selected_fields = config["selected_fields"]
        displayed_fields = config["displayed_fields"]
        is_select_all = config["is_select_all"]
except:
    print("Error: An error occurred while reading the configuration.")
    input()
    exit()
 
 
#A simple function that removes all unnecessary characters from a string, such as spaces or tabs
def deleteSpaces(text):
    return ''.join(char for char in text if char.isalnum())
 
soup = BeautifulSoup (html, 'html.parser') #Parse html
 
clear_fields_list = soup.thead.find_all("span") #Find all fields at the page
progress_list = soup.tbody.find_all("tr") # Find all students progres

clear_progress_list = ["Opiskelijan nimi"] #After i will write here "progress_list" with out HTML elements

for i in clear_fields_list: #write text with out HTML
    clear_progress_list.append(i.text)



#clear_progress_list.insert(0, "Opiskelijan nimi")

#This array is needed to link the current order of fields with the desired order of fields
indexes_list = []
 
#In this object we will store the data that we will later populate into Excel
table = {}
 
#The student list is always required, so we create an object for it in advance
 
#The loop creates objects in the table where we will later write the data
#The data can vary depending on the "selected_fields" array

if is_select_all:
    for i in clear_progress_list:
        table[i] = []
else:
    for i in selected_fields:
        table[displayed_fields[selected_fields.index(i)]] = []
        indexes_list.append(clear_progress_list.index(i))

#The main loop of the program that searches for data, separates the required data
#and writes it to the table.
for n,i in enumerate(progress_list): #The loop iterates through all the <tr> tags
    elements = i.find_all("td")
    if n == 0: #This check skips the first element, which is not needed
        continue
    #table["Opiskelijan nimi"].append(elements[0].a.text)
    for n2, j in enumerate(elements): #The loop iterates through the <td> tags that were found inside the <tr> tag
        #print(f"{n2}: {deleteSpaces(j.text)}")
        #print(clear_progress_list)

        if is_select_all == False and not n2 in indexes_list:
            continue
        field_name = ""
        if(is_select_all):
            field_name = clear_progress_list[n2]
        else:
            field_name = displayed_fields[indexes_list.index(n2)]
        if j.find("a"): #Name blocks also packed in <a> tag
            table[field_name].append(j.a.text)
            continue
        clear_str = deleteSpaces(j.text) # Delete spaces from text
        if(clear_str == "O" or clear_str == "o"): # Replace O to X
            table[field_name].append("X") # Add X to table
            continue
        table[field_name].append(deleteSpaces(clear_str)) # Add text to table
        continue
        
        #if is_select_all:
        #    field_name = clear_progress_list[n2]
        #    #print(n2)
        #    #print(field_name)
        #    if j.find("a"):
        #        table[field_name].append(j.a.text)
        #        continue
        #    clear_str = deleteSpaces(j.text)
        #    if(clear_str == "O" or clear_str == "o"): # Replace O to X
        #        table[field_name].append("X") # Add X to table
        #        continue
        #    table[field_name].append(deleteSpaces(clear_str)) # Add text to table
        #    continue
        #
        #if n2 in indexes_list: #Check whether the category is relevant to us
        #    field_name = displayed_fields[indexes_list.index(n2)] #Find key
        #    if j.find("a"):
        #       table[field_name].append(j.a.text)
        #        continue
        #    clear_str = deleteSpaces(j.text) # Delete spaces from text
        #    if(clear_str == "O" or clear_str == "o"): # Replace O to X
        #        table[field_name].append("X") # Add X to table
        #        continue
        #    table[field_name].append(deleteSpaces(clear_str)) # Add text to table

#print(table)
df = pd.DataFrame(table) #Add table to data frame
 
#The next line writes the DataFrame to Excel, if you don't need styles, just use this code
#df.to_excel(writer, sheet_name='Table', index=False)
 
#the addition of styles for the Excel table
with pd.ExcelWriter('students.xlsx', engine='openpyxl') as writer:

    df.to_excel(writer, sheet_name='Table', index=False) #Write to excel

    work_book = writer.book
    work_sheet = writer.sheets['Table'] # Get a sheet from excel
 
    #The loop that iterates through the table columns is needed to adjust the first column (A)
    #to fit the length of the longest name
    max_length = 0 #The variable for longest lenght
    for col in work_sheet.columns:
        if col[0].column_letter == 'A' or is_select_all: #Check is this A column  
            #The loop iterates through the cells in the column and looks for the longest element
            for cell in col:
                #Checks if the length of the iterated cell is longest than the maximum detected length.
                #if so it updates the variable with the max_length.
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            adjusted_width = (max_length + 2) #Add +2 so that the text does not touch the cells border.
            work_sheet.column_dimensions[col[0].column_letter].width = adjusted_width #Change (A) length
    #for col in work_sheet.columns:
    #    max_length = 0
    #    column = col[0].column_letter  # Получаем букву колонки (A, B, C, ...)
    #    for cell in col:
    #        try:  # Найти максимальную длину контента в ячейке
    #            if len(str(cell.value)) > max_length:
    #                max_length = len(cell.value)
    #        except:
    #            pass
    #    adjusted_width = (max_length + 2)
    #    work_sheet.column_dimensions[column].width = adjusted_width
    fill_color = "828181" #Gray color

    #Fill object
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
 
    #Border object
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    #Centers all elements
    center_alignment = Alignment(horizontal='center', vertical='center')

    #Add a black-and-white background for the rows.
    for row in work_sheet.iter_rows(min_row=2, max_row=work_sheet.max_row):
        for cell in row: #Adds styles such as: border, background and centers
            cell.border = border
            cell.alignment = center_alignment
            if cell.row % 2 == 0:
                cell.fill = fill  
                #cell.font = openpyxl.styles.Font(color='FFFFFF') #Font size

print("If successful, check the file students.xlsx in the same folder.")
print("Powered by guziiuchyk@gmail.com <3")
input()