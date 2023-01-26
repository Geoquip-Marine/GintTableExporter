import pyodbc
import os
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import *
import pandas as pd
from tkinter import messagebox
import warnings
warnings.filterwarnings("ignore")
import customtkinter as ct
import time
import datetime

def main():
    class Application(ct.CTkFrame):
        ct.set_appearance_mode("light")

        def __init__(self):
            super(Application, self).__init__(window)
            window.title("gINT table exporter")
            window.resizable(False,False)
            window.wm_iconbitmap(bitmap='common/assets/gint.ico', default='common/assets/gint.ico')
            window.geometry("125+125")

            self.file_location: str 
            self.gintpath: str
            self.export_dir: str
            self.gint_name: str
            self.tables = ()
            self.bh_list = ()
            self.selected_bh_list = ()

            #create all widgets, assign functions to click events, place in window and formatting

            self.button_open = ct.CTkButton(window, text="Select a gINT project...", command=self.get_file_location, fg_color="#7FB3D5", 
            corner_radius=10, hover_color="#58D68D", text_color="#000000", text_color_disabled="#CECECE", font=("Tahoma",12))
            self.button_open.grid(row=1, column=1, pady=(30,20), columnspan=3)

            self.menulabel = ct.CTkLabel(window, text="Select table(s) below to export from gINT.\nSelect multiple tables with CTRL.", font=("Tahoma",12), text_color="#000000")
            self.menulabel.grid(row=2, column=1, columnspan=3, pady=(5,15))

            self.button_export = ct.CTkButton(window, text="Export Table(s)", command=self.export_table, fg_color="#7FB3D5", 
            corner_radius=10, hover_color="#58D68D", text_color="#000000", text_color_disabled="#CECECE", font=("Tahoma",12), width=150)
            self.button_export.grid(row=4, column=1, pady=10, columnspan=3)
            self.button_export.configure(state=tk.DISABLED)

            self.n = StringVar(value=self.tables)
            self.chosentable = tk.Listbox(window, height=10, width=45, listvariable=self.n, selectmode='extended', justify="center")
            self.chosentable.grid(row=5, column=1, pady=10, padx=(60, 0), columnspan=2)
            self.chosentable.configure(state=tk.DISABLED)

            self.scrollbar = ct.CTkScrollbar(window, command=self.chosentable.yview, button_color="#7FB3D5", button_hover_color="#58D68D", minimum_pixel_length=40, border_spacing=3, height=0)
            self.scrollbar.grid(row=5, column=3, padx=(0, 40), pady=10, sticky="ns")
            self.chosentable.config(yscrollcommand=self.scrollbar.set)

            self.menulabel2 = ct.CTkLabel(window, text="Select borehole(s) below.\nSelecting none will export all by default.", font=("Tahoma",12), text_color="#000000")
            self.menulabel2.grid(row=6, column=1, columnspan=3,  pady=(12,12))

            self.y = StringVar(value=self.bh_list)
            self.pointtable = tk.Listbox(window, height = 10, width=45, selectmode='extended', justify="center", listvariable=self.y)
            self.pointtable.grid(row=7, column=1, pady=(10, 20), padx=(60, 0), columnspan=2)

            self.pointscroll = ct.CTkScrollbar(window, command=self.pointtable.yview, button_color="#7FB3D5", button_hover_color="#58D68D", minimum_pixel_length=40, border_spacing=3, height=0)
            self.pointscroll.grid(row=7, column=3, padx=(0, 40), pady=(10, 20), sticky="ns")
            self.pointtable.config(yscrollcommand=self.pointscroll.set)
            self.pointtable.configure(state=tk.DISABLED)

            self.borehole_button = ct.CTkButton(window, text="Add selected Borehole(s) to range", command=self.get_bhs, fg_color="#7FB3D5", 
            corner_radius=10, hover_color="#58D68D", text_color="#000000", text_color_disabled="#CECECE", font=("Tahoma",12))
            self.borehole_button.grid(row=8, column=1, pady=(0,25), columnspan=3)
            self.borehole_button.configure(state=tk.DISABLED)

            self.progress_bar = ct.CTkProgressBar(window, mode="determinate")
            self.progress_bar.set(0)
            self.progress_bar.configure(progress_color="#58D68D")
            self.progress_bar.grid(row=9, column=1, columnspan=3, padx=100, pady=(0, 20))


        def disable_buttons(self):
            self.button_open.configure(state=tk.DISABLED)
            self.button_export.configure(state=tk.DISABLED)
            self.borehole_button.configure(state=tk.DISABLED)
            self.chosentable.configure(state=tk.DISABLED)
            self.pointtable.configure(state=tk.DISABLED)

        def enable_buttons(self):
            self.button_open.configure(state=tk.NORMAL)
            self.button_export.configure(state=tk.NORMAL)
            self.borehole_button.configure(state=tk.NORMAL)
            self.chosentable.configure(state=tk.NORMAL)
            self.pointtable.configure(state=tk.NORMAL)

        def get_bhs(self):
            #reset the selected borehole list to avoid conflict after selecting another gint, get the selected boreholes from bottom list box
            self.selected_bh_list = ()
            list_bhs = ""
            bh_selection = self.pointtable.curselection()
            list_bhs = ",".join([self.pointtable.get(i) for i in bh_selection])
            if not list_bhs == "":
                self.selected_bh_list = list_bhs
            print("Boreholes added: " + str(list_bhs))

        def get_file_location(self):
            #get gint location
            self.file_location = filedialog.askopenfilename(filetypes=[('gINT Project', '*.gpj')])
            if self.file_location == '':
                messagebox.showwarning(title="No gINT selected", message="You must select a gINT")
                return
            print(f"Opening {self.file_location}...")

            #check the dir exists
            if os.path.exists(os.path.dirname((self.file_location))):    
                os.chdir(os.path.dirname((self.file_location)))
                self.gintpath = os.path.dirname((self.file_location))
            else:
                raise ValueError("Please check the directory is correct.")
            
            #get the project name for filename
            self.gint_name = self.file_location.rsplit("/", 1)[1]
            self.gint_name = self.gint_name.split(" ", 1)[0]

            #establish connection to sql database (gint)
            try:
                conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+self.file_location+';')
            except Exception as e:
                print(f"Couldn't establish connection with gINT. Please ensure you have Access Driver 64-bit installed. {e}")
                return

            self.enable_buttons()

            #load bottom listbox with borehole ids from the loaded gint
            bh_query = "SELECT PointID FROM POINT"
            point_list = pd.read_sql(bh_query, conn)
            point_id = point_list['PointID'].tolist()
            self.bh_list = point_id
            self.pointtable.delete(0, END)
            point_id = sorted(point_id)
            for x in range (0, len(point_id)):
                self.pointtable.insert(END, point_id[x])

            #get all the tables from gint and put them into a list
            cursor = conn.cursor()
            tableNames = [x[2] for x in cursor.tables() if x[3] == 'TABLE']

            INVALIDCHARS = '<>:"/\|?* �'

            #remove the access database system tables, load table names to top listbox
            self.tables = [x for x in tableNames if not str(x).startswith(u'\x7f') and not "DATGEL_SETTINGS" in str(x) and not " " in str(x) and not "GINT" in str(x)]
        
            self.chosentable.delete(0, END)
            self.tables = sorted(self.tables)
            for x in range (0, len(self.tables)):
                self.chosentable.insert(END, self.tables[x])

            #reset the selected boreholes from the bottom listbox if another gint is loaded
            self.selected_bh_list = ()
            

        def export_table(self):
            self.disable_buttons()
            #get dir of exported file of dataframes from sql query
            self.path_msg = messagebox.showinfo('Select destination folder:','Select directory to save export.')
            self.path_directory = filedialog.askdirectory(title="Select directory to save export: ")
            if self.path_directory == "":
                if messagebox.askokcancel("You must select a directory!", "You have not selected a directory to save export."):
                    self.path_directory = filedialog.askdirectory(title="Select directory to save export: ")
                    if self.path_directory == "":
                        self.enable_buttons()
                        return
                else:
                    self.enable_buttons()
                    return

            #get tables selected from top listbox
            self.export_dir = self.path_directory + "/"
            print(f"Export directory: {self.export_dir}")
            self.progress_bar.set(0)
            tableselection = self.chosentable.curselection()
            tableselect = ",".join([self.chosentable.get(i) for i in tableselection])

            if tableselect == "":
                messagebox.showwarning(title="No table selected.", message="You must select at least one table!")

            #double check we can talk to gint before calling sql query
            try:
                conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+self.file_location+';')
            except Exception as e:
                print(f"Couldn't establish connection with gINT. Please ensure you have Access Driver 64-bit installed. {e}")
                self.enable_buttons()
                return

            list_all_tables = tableselect.split(",")
            print("Tables selected: " + str(list_all_tables))

            #we need to not only get the selected boreholes into a list if its more than 1 table,
            #but also to add commas and quotation marks so it passes the syntax for the SQL query when it's passed as a string
            if len(self.selected_bh_list) < 1:
                bh_select = ""
                num_bh = len(self.bh_list)
                for i in range (0, num_bh):
                    bh_select += str("'" + self.bh_list[i] + "',")
                bh_select = bh_select[:-1]
                print(f"Boreholes selected: {str(bh_select)}")
            else:
                bh_selected = self.selected_bh_list
                bh_selected = bh_selected.split(",")
                num_bh = len(bh_selected)
                bh_select = ""
                for i in range (0, num_bh):
                    bh_select += "'" + bh_selected[i] + "',"
                bh_select = bh_select[:-1]  
                print(f"Boreholes selected: {str(bh_select)}")

            #after all data from gint is extracted, create some variables to use later
            cur_time = time.time()
            timestamp = str(datetime.datetime.fromtimestamp(cur_time).strftime('%Y-%m-%d (%H_%M)'))
            num_tables = len(list_all_tables)
            df_dict = {}

            #looping over tables select to query data and save to a new sheet in the workbook
            for df_iteration in range(0, num_tables):
                self.update_idletasks()
                time.sleep(0.01)
                sheetname = list_all_tables[df_iteration]
                #save stcn_data to a seperate csv to save time for large dataframe
                if sheetname == "STCN_DATA":  
                    print(sheetname + " extracting from gINT...")
                    query = f"SELECT * FROM {str(list_all_tables[df_iteration])} WHERE PointID IN ({str(bh_select)})"
                    df_iteration = pd.read_sql(query, conn)
                    df_iteration.sort_values("PointID")
                    df_iteration.drop(['GintRecID'], axis=1, inplace=True)
                    df_iteration.to_csv(f"{self.export_dir}{self.gint_name} - {sheetname}_{timestamp}.csv", index=False)
                    #messagebox.showwarning(title="STCN_DATA exported seperately (filesize too big)", message=f"{sheetname} has been exported to a seperate .csv file.")
                    continue
                print(sheetname + " extracing from gINT...")

                #the progress bar is updated up to 50% from the number of tables queried from gint
                progress_update = ((100 / num_tables) / 100) / 2
                value = self.progress_bar.get()
                value += progress_update
                self.progress_bar.set(value)
                window.update()

                #query the data from gint
                try:
                    #sleep so it doesn't max out cpu load in single thread from all the sql queries
                    time.sleep(0.01)
                    query = f"SELECT * FROM {str(list_all_tables[df_iteration])} WHERE PointID IN ({str(bh_select)})"
                    df_iteration = pd.read_sql(query, conn)
                    #amke sure to sort the values from location name alphabetically, all other columns follow suit it seems
                    df_iteration.sort_values("PointID")
                except:
                    #another query if the table doesn't have any pointid field (like 'dict' or 'project')
                    time.sleep(0.01)
                    query = (f"SELECT * FROM {list_all_tables[df_iteration]}")
                    df_iteration = pd.read_sql(query, conn)

                #get rid of that pesky gintrecid, save the table selected as the dataframe dict's key and the sheetname for excel
                df_iteration.drop(['GintRecID'], axis=1, inplace=True)    
                df_dict[sheetname] = df_iteration

            #put all the dataframes from sql quert into a dict to use to loop through for excel writing, check if they're empty and save empty list
            final_dataframes = [(k,v) for (k,v) in df_dict.items() if not v.empty]
            final_dataframes = dict(final_dataframes)
            empty_dataframes = [k for (k,v) in df_dict.items() if v.empty]

            #make the filename from the number of tables selected, gint filename (project) and timestamp
            table_names = list(final_dataframes.keys())
            if len(table_names) < 5:
                table_names = table_names
            else:
                table_names = table_names[0:5]

            filename_tables = str(table_names).replace(",","_").replace("'","").replace("[","").replace("]","")

            self.filename = (f"{self.gint_name} - {filename_tables}_{timestamp}")

            print(f"""------------------------------------------------------
    Saving tables to excel file, this may take a while...
    ------------------------------------------------------""")

            #create the excel file with the first dataframe from dict, so pd.excelwriter can be called (can only be used on existing excel workbook to append more sheets)
            if not len(final_dataframes.keys()) < 1:
                next(iter(final_dataframes.values())).to_excel(f"{self.export_dir}{self.filename}.xlsx", sheet_name=(f"{next(iter(final_dataframes))}"), index=None, index_label=None)
                final_writer = pd.ExcelWriter(f"{self.export_dir}{self.filename}.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace")
            else:
                print(f"All selected tables are empty! Please select others. Tables selected: {empty_dataframes}")
                self.enable_buttons()
                return

            #for every key (table name) and value (table data) in the sql query dict, append to excel sheet and update progress bar, saving only at the end for performance
            for (k,v) in final_dataframes.items():
                print(f"Writing {k} to excel...")
                v.to_excel(final_writer, sheet_name=(f"{str(k)}"), index=None, index_label=None)
                self.update_idletasks()
                time.sleep(0.01)
                #update the progress bar for the remaining 50% based on how many non-empty dataframes are being written to excel
                try:
                    progress_update = ((100 / (len(final_dataframes.keys()) -1 )) / 100) / 2
                except ZeroDivisionError:
                    progress_update = (((100 / (1) )) / 100) / 2
                value = self.progress_bar.get()
                value += progress_update
                self.progress_bar.set(value)
                window.update()
            final_writer.save()
            
            if empty_dataframes != []:
                print(f"""----------------------------------------------------
    The following tables were empty, and were skipped...
    ----------------------------------------------------
    {empty_dataframes}""")

            print(f"""---------------------
    EXCEL EXPORT COMPLETE               
    ---------------------
    {self.export_dir}{self.filename}.xlsx""")
            
            self.enable_buttons()

    window = ct.CTk()
    app = Application()
    window.mainloop()

if __name__ == '__main__':
    main()
