# nuitka-project: --disable-console
# nuitka-project: --standalone
# nuitka-project: --windows-icon-from-ico="\Icon.ico"
# nuitka-project: --enable-plugin=tk-inter
import sys
import os
import shutil
import pandas as pd
import time 
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

copy_file_src = "" # Update source path here
copy_file_dst = "" # Update destination path here
delete_file_src = "" # Update approved path here
icon_path = r"\Icon.ico"
event_log_file = "\\event_log.txt"
# Global variables
file_src = ""
file_dst = ""
current_item_checking = ""

main_button_option = ""
    
# List for found/not found items
item_not_found_list = []
item_in_pending_list = []
item_not_in_approved_list = []

copy_option_msg = ""
hidden_option = "Unhide"
# Main function
def main():
    
    global runtime_label, root, header_font, item_read_label, file_paths_menu, menu_bar, excel_file

    excel_file = "\\move_file.xlsx" # Update excel path here
    header_font = ("Arial", 12, "bold")
    
    # Main window
    root = tk.Tk()
    root.title("File Management")
    root.geometry("500x250")
    root.focus_force()
    root.iconbitmap(icon_path)
    # Function to display file paths 
    def open_file_location(selection):

        # To-do Add option to change file path location
        if selection == "set":
            messagebox.showinfo("In Progress", "In Progress")
        if selection == "excel file input":
            file_path = filedialog.askopenfilename(filetypes = [("Csv Files", "*.csv"), ("Excel Files", "*.xlsx")])
            if file_path:
                global excel_file
                excel_file = file_path
                messagebox.showinfo("Excel Location", f"File Path: {file_path}")
        elif selection == excel_file:
            if os.path.exists(excel_file):
                excel_directory = os.path.dirname(excel_file)
                os.startfile(excel_directory)
                messagebox.showinfo("Excel Location", f"File Path: {selection}")
            else:
                messagebox.showerror("Error", f"File Path Not Found - Current File Path:\n{excel_file}")

        elif selection != excel_file:
            if os.path.exists(selection):
                os.startfile(selection)
            else:
                messagebox.showerror("Error", f"File Path Not Found - Current File Path:\n{selection}")
        else:
            messagebox.showerror("Error", "Error in open_file_location function")

    # Function to toggle hidden menu
    hidden_menu_var = tk.IntVar()
    hidden_menu_var.set(0)

    def toggle_hidden_menu_option(event):

        if hidden_menu_var.get() == 0:
            hidden_menu_var.set(1)
            hidden_menu = tk.Menu(menu_bar, tearoff = 0)
            menu_bar.add_cascade(label = "Hidden", menu = hidden_menu)
            hidden_menu.add_command(label = "Set Paths", command = lambda: open_file_location("set")) # To-do Add more menu if needed
            hidden_menu.add_command(label = "Excel File", command = lambda: open_file_location("excel file input"))
        else:
            hidden_menu_var.set(0)
            global excel_file
            excel_file = "\\move_file.xlsx"
            messagebox.showinfo("Excel File", "Currently Using Default Excel File:\n" + excel_file)
            menu_bar.delete("Hidden")
            
    # Bind function to toggle hidden menu
    root.bind_all("<Control-Shift-A>", toggle_hidden_menu_option)

    # Widget menu
    menu_bar = tk.Menu(root)
    root.config(menu = menu_bar)
    
    file = tk.Menu(menu_bar, tearoff = 0)
    menu_bar.add_cascade(label = "File", menu = file)
    file.add_command(label = "Exit", command = sys.exit)
    
    file_paths_menu = tk.Menu(menu_bar, tearoff = 0)
    menu_bar.add_cascade(label = "File Locations", menu = file_paths_menu)

    file_paths_menu.add_command(label = "Excel Location",foreground = "blue", 
            font = ("Arial", 8, "bold"), command = lambda:  open_file_location(excel_file))
    file_paths_menu.add_command(label = "Copy Source", command = lambda: open_file_location(copy_file_src))
    file_paths_menu.add_command(label = "Copy Destination",command = lambda: open_file_location(copy_file_dst))
    file_paths_menu.add_command(label = "Delete Location", command = lambda: open_file_location(delete_file_src))
   
    calculation_menu = tk.Menu(menu_bar, tearoff = 0)
    menu_bar.add_cascade(label = "Calculation Tools", menu = calculation_menu)
    calculation_menu.add_command(label = "Scaling Tool", command = on_scaling_tool_select)

    calc_img = tk.PhotoImage(file = r"\Calc.png")
    calc_img = calc_img.subsample(2, 2)
    calculation_menu.add_command(label = "Calculator", image = calc_img, compound = "left", command = on_calulator_select)
    
    style = ttk.Style() 
    style.configure("M.TButton", width = 25, font = ("Arial", 9))
    style.map("M.TButton", background = [('pressed','black')])
    
    lbl = ttk.Label(root, text="Select Options Below", font=header_font)
    lbl.grid(column = 0, row = 0, columnspan = 2)

    delete_btn = ttk.Button(root, text = "Approved Folder Delete Files", command = on_delete_button_click, style = "M.TButton")
    delete_btn.grid(column = 0, row = 1, padx = 2,pady = 2, ipadx = 2.5, ipady = 2.5, sticky = tk.N)

    copy_btn = ttk.Button(root, text="Autoprint Output Copy Files", command = on_copy_button_click, style = "M.TButton")
    copy_btn.grid(column = 0, row = 2, padx = 2, pady = 2, ipadx = 2.5, ipady = 2.5)

    ttk.Button(root, text = "Pending", command = on_pending_button_click).grid(column = 1, row = 1, padx = 2, ipadx = 2.5, ipady = 2.5)
    
    ttk.Button(root, text = "Exit", command = sys.exit).grid(column = 1, row = 2, padx = 2, ipadx = 2.5, ipady = 2.5)

    item_read_label = ttk.Label(root, text = "", font = ("Arial", 9, "underline"))
    item_read_label.place(relx = 1.0, rely = 1.0, anchor = tk.SE, x = -5, y = -2)

    runtime_label = ttk.Label(root, text = "Program Runtime: 00:00",font = ("Arial", 9, "bold"))
    runtime_label.place(relx = 0, rely = 1, anchor = tk.SW, x = 5, y = -2)
    # Call to center window
    center_window(root)
    root.mainloop()
# Global Variables for user input validation
accepted_symbols = "xX+-÷*()^/"
accepted_numbers = "0123456789."

 # Function for validate number input
def validate_number(value):
    if value.startswith('.'):
        value = '0' + value
    
    for char in value:
        if char not in accepted_symbols and char not in accepted_numbers:
            return False 
    return True

# Function for Scaling tool
def on_scaling_tool_select():
    
    scaling_tool_frame = tk.Toplevel()
    scaling_tool_frame.title("Scaling Tool")
    scaling_tool_frame.geometry("300x150")
    scaling_tool_frame.iconbitmap(icon_path)
    scaling_tool_frame.resizable(0,0)
    center_window(scaling_tool_frame)

    ttk.Label(scaling_tool_frame, text = "Diameter 1:").grid(row = 0, column = 0, padx = 5, pady = 5)
    dia1_entry = ttk.Entry(scaling_tool_frame, width = 10, justify = 'right')
    dia1_entry.config(validate = 'key', validatecommand = (scaling_tool_frame.register(validate_number), '%P'))
    dia1_entry.grid(row = 0, column = 1, padx = 5, pady = 5)
    dia1_entry.focus_set()

    ttk.Label(scaling_tool_frame, text = "Scaling Dim:").grid(row = 1, column = 0, padx = 5, pady = 5)
    x_entry = ttk.Entry(scaling_tool_frame, width = 10, justify = 'right')
    x_entry.config(validate = 'key', validatecommand = (scaling_tool_frame.register(validate_number), '%P'))
    x_entry.grid(row = 1, column = 1, padx = 5, pady = 5)

    ttk.Label(scaling_tool_frame, text = "Diameter 2:").grid(row = 2, column = 0, padx = 5, pady = 5)
    dia2_entry = ttk.Entry(scaling_tool_frame, width = 10, justify = 'right')
    dia2_entry.config(validate = 'key', validatecommand = (scaling_tool_frame.register(validate_number), '%P'))
    dia2_entry.grid(row = 2, column = 1, padx = 5, pady = 5)

    result_label = ttk.Label(scaling_tool_frame, text = "", font = ("Arial", 9, "bold"))
    result_label.grid(row = 1, column = 2, padx = 5, pady = 5, ipady = 10, ipadx = 10)
    # Function for continue button
    def on_continue_button_click():

        if dia1_entry.get() == '' or dia2_entry.get() == '' or x_entry.get() == '':
            messagebox.showerror("Error", "Please enter all values")
            on_scaling_tool_select()
        else:
            result_label.config(text = "")

            dia1 = float(dia1_entry.get())
            dia2 = float(dia2_entry.get())
            x = float(x_entry.get())

            result = round((x/dia1) * dia2, 8)

            ttk.Label(scaling_tool_frame, text = "Result New Dim:", font = ("Arial", 9, "bold")).grid(row = 0, column = 2, padx = 5, pady = 5)
            result_label.config(text = result, background = "white", relief = "solid", border = 1, justify = 'center', font = ("Arial", 9, "bold"))

    ttk.Button(scaling_tool_frame, text = "Back", command = scaling_tool_frame.destroy).grid(row = 3, column = 0, padx = 5, pady = 5)
    ttk.Button(scaling_tool_frame, text = "Calculate", command = on_continue_button_click).grid(row = 3, column = 1, padx = 5, pady = 5)
    scaling_tool_frame.mainloop()

# Function for calculator
def on_calulator_select():
    # Made Global to be used in other functions
    global show_tooltip, hide_tooltip
    calculator = tk.Toplevel()
    calculator.title("Calculator")
    calculator.geometry("392x400")
    calculator.iconbitmap(icon_path)
    calculator.resizable(0,0)

    # Function to make window topmost
    def make_top(event):
        if calculator.attributes('-topmost'):
            
            calculator.attributes('-topmost', False)
            top_label = tk.Label(calculator, text = "Topmost OFF", font = ("Arial", 9, "bold"))
        else:
            calculator.attributes('-topmost', True)
            top_label = tk.Label(calculator, text = "Topmost ON", font = ("Arial", 9, "bold"))
    
    calculator.bind("<KeyPress-t>", make_top)
    calculator.bind("<KeyPress-T>", make_top)
    # Function to calculate user input
    def equal(event):
        expression = entry.get()

        if expression == "":
            value_error()
        else:
            # Convert the symbols to standard symbols
            expression = expression.replace("x", "*")
            expression = expression.replace("÷", "/")
            expression = expression.replace("^", "**")
            # Catch errors with eval() method
            try:
                result = str(eval(expression))
                label_result = f"{expression} = "
                display_result(result, label_result)
                entry.focus_set()
            except (ZeroDivisionError, SyntaxError, NameError, TypeError, KeyError, ValueError) as e:
                messagebox.showerror("Error", f"Error: {e}")
            
    # Function for displaying error when value is not entered
    def value_error():
        messagebox.showerror("Error", "Please enter a value")
        entry.focus_set()
    # Function for clearing user input
    def clear_entry(event):
        entry.delete(0, tk.END)
        entry_label.config(text = "")
        entry.insert(0, "0")
        entry.focus_set()
        entry.select_range(0, tk.END)
    # Function to display the result
    def display_result(result, label_result):
        entry_label.config(text = label_result)
        entry.delete(0, tk.END)
        entry.insert(0, result)
    # Function for converting mm and inch
    def mm_and_inch_convert(option):
        symbol_found = False
        # Converts input to mm from inch
        #if option == "mm":
        if entry.get() == "":
            value_error()
            return
        # Checks for symbols before continuing
        for i in entry.get():
            if i in accepted_symbols:
                messagebox.showerror("Error", f"Please enter numeric value only")
                symbol_found = True
                break
        if symbol_found == False:
            try:
                if option == "mm":
                    result = round(eval(entry.get())*25.4, 6)
                    label_result = f"{entry.get()}in x 25.4 = {result}mm"
                    display_result(result, label_result)

                if option == "inch":
                    result = round(eval(entry.get())/25.4, 6)
                    label_result = f"{entry.get()}mm ÷ 25.4 = {result}in"
                    display_result(result, label_result)
            except (ZeroDivisionError, SyntaxError, NameError, TypeError, KeyError, ValueError) as e:
                messagebox.showerror("Error", f"Error: {e}")
       
    # Function for math symbol button on calculator
    def math_btn_click(symbol):
        entry.insert(tk.END, symbol)
    
    # Purpose is to display calculation breakdown 
    entry_label = tk.Label(calculator, borderwidth = 0, foreground = "grey", background = "#595954", relief = "groove", font = ("Arial", 12))
    entry_label.grid(row = 0, column = 0, columnspan = 4, sticky = "nswe")

    entry = tk.Entry(calculator, width = 26, borderwidth = 0, justify="center", highlightcolor = "white",
                     foreground = "white", background = "#595954", relief = "groove", font = ("Arial", 20))
    entry.grid(row = 1, column = 0, columnspan = 4, sticky = "nswe")
    # Validate user input per key press
    entry.config(validate = 'key', validatecommand = (calculator.register(validate_number), '%P'))
    entry.focus_set()
    entry.insert(0, "0") # Default value
    entry.select_range(0, tk.END)
    
    mm_btn = tk.Button(calculator, text = "mm", command = lambda: mm_and_inch_convert("mm"),
                       foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    mm_btn.grid(row = 2, column = 0, sticky = "ew")

    inch_btn = tk.Button(calculator, text = "inch", command = lambda: mm_and_inch_convert("inch"),
                       foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    inch_btn.grid(row = 2, column = 1, sticky = "ew")

    equal_btn = tk.Button(calculator, text = "=", command = lambda: equal(event = None),
                          foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    equal_btn.grid(row = 2, column = 2, sticky = "ew")

    multiply_btn = tk.Button(calculator, text = "x", command = lambda: math_btn_click("x"),
                            foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    multiply_btn.grid(row = 3, column = 2, sticky = "ew")

    divide_btn = tk.Button(calculator, text = "/", command = lambda: math_btn_click("/"),
                            foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    divide_btn.grid(row = 4, column = 2, sticky = "ew")

    add_btn = tk.Button(calculator, text = "+", command = lambda: math_btn_click("+"),
                        foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    add_btn.grid(row = 5, column = 2, sticky = "ew")

    subtract_btn = tk.Button(calculator, text = "-", command = lambda: math_btn_click("-"),
                             foreground = "white", background = "#2E2E2B", relief = "groove", font = ("Arial", 15))
    subtract_btn.grid(row = 6, column = 2, sticky = "ew")
    
    clear_btn = tk.Button(calculator, text = "clear", command = lambda: clear_entry(event = None), font = ("Arial", 15))
    clear_btn.grid(row = 2, column = 3, sticky = "ew")
    
    # -----Keybinds and Tooltips-----
    # Function to show tooltip
    def show_tooltip(text):
        global tooltip_window
        tooltip_window = tk.Toplevel()
        tooltip_label = tk.Label(tooltip_window, text = text, border = 1, relief = "solid")
        tooltip_label.pack()
        # Remove window header
        tooltip_window.wm_overrideredirect(1)
        # Position the tooltip
        x = calculator.winfo_pointerx() + 20
        y = calculator.winfo_pointery() + 20    
        tooltip_window.geometry(f"+{x}+{y}")

    # Function to hide tooltip
    def hide_tooltip(event):
        try:
            global tooltip_window
        
            tooltip_window.destroy()
            tooltip_window = None
        except AttributeError:
            pass 
            
    # For mm Button
    def on_m_keypress(event):
        mm_and_inch_convert("mm")
    calculator.bind("<KeyPress-m>", on_m_keypress)
    calculator.bind("<KeyPress-M>", on_m_keypress)
    # mm Tooltip
    def mm_btn_tooltip(event):
        text = "Converts MM to Inch (M)"
        show_tooltip(text)
    # Mouse hover and leave will display tool tips
    mm_btn.bind("<Enter>", mm_btn_tooltip)
    mm_btn.bind("<Leave>", hide_tooltip)

    # For Inch Button
    def on_i_keypress(event):
        mm_and_inch_convert("inch")
    calculator.bind("<KeyPress-i>", on_i_keypress)
    calculator.bind("<KeyPress-I>", on_i_keypress)
    # Inch Tooltip
    def inch_btn_tooltip(event):
        text = "Converts Inch to MM (i)"
        show_tooltip(text)
    inch_btn.bind("<Enter>", inch_btn_tooltip)
    inch_btn.bind("<Leave>", hide_tooltip)

    # For equal Button
    calculator.bind("<Return>", equal)
    # Equal Tooltip
    def equal_btn_tooltip(event):
        text = "Calculate Input (Enter)"
        show_tooltip(text)
    equal_btn.bind("<Enter>", equal_btn_tooltip)
    equal_btn.bind("<Leave>", hide_tooltip)

    # For clear Button
    calculator.bind("<KeyPress-C>", clear_entry)
    calculator.bind("<KeyPress-c>", clear_entry)
    # Clear Tooltip
    def clear_btn_tooltip(event):
        text = "Clear Input (C)"
        show_tooltip(text)
    clear_btn.bind("<Enter>", clear_btn_tooltip)
    clear_btn.bind("<Leave>", hide_tooltip)
    
    calculator.mainloop()

# Function for delete button 
def on_delete_button_click():
    global main_button_option
    main_button_option = "Delete"
    clear_list()
    select_file_type()

# Function for copy button
def on_copy_button_click(): 
    global main_button_option
    main_button_option = "Copy"

    clear_list()
    copy_checkbutton_option()

# Function for pending button
def on_pending_button_click():
    copy_file_src = "W:\\Technical\\PRINTS\\PENDING APPROVAL" 
    copy_file_dst = "W:\\Technical\\PRINTS\\APPROVED PRINTS" 
    global pending_frame, main_button_option

    main_button_option = "Pending"

    pending_frame = tk.Toplevel()
    pending_frame.title("Pending Copy")
    pending_frame.geometry("275x100")
    pending_frame.iconbitmap(icon_path)
    pending_frame.resizable(0,0)
    center_window(pending_frame)
    # Function for continue button inside pending frame
    def on_continue_button_click():
        item = item_entry.get()
        pending_option = "continue"
        if item == "":
            messagebox.showerror("Error", "Please enter part number")
            pending_frame.destroy()
            on_pending_button_click()

        else:
            start_time = time.time()
            pending_copy_file(item,start_time, pending_option)
            messagebox.showinfo(pending_copy_status,message)
            pending_frame.destroy()
            on_pending_button_click()
        
    # Function for copying files from pending to approved folder
    def pending_copy_file(item, start_time, pending_option):
        global message, pending_copy_status
        file_src = (f"{copy_file_src}\\{item}")
        file_dst = (f"{copy_file_dst}\\{item}")
        revision_folder_src = (f"{file_src}\\OLD")
        revision_folder_dst = (f"{file_dst}\\OLD")

        if os.path.exists(file_src):
            if os.listdir(file_src) != []:
                #Creates folder if destination folder does not exist
                if not os.path.exists(file_dst):
                    os.mkdir(file_dst)

                # Handles revision folder
                if os.path.exists(revision_folder_src):
                    if not os.path.exists(revision_folder_dst):
                        os.mkdir(revision_folder_dst)

                    for file in os.listdir(revision_folder_src):
                        shutil.copy2(os.path.join(revision_folder_src, file), revision_folder_dst)
                    
                    for file in os.listdir(file_src):
                        if file != "OLD": # Exclude OLD folder otherwise throw permission error when copying
                            shutil.copy2(os.path.join(file_src, file), file_dst)

                    if pending_option == "continue":
                        run_time = runtime(start_time)
                        runtime_label.config(text = run_time)
                        message = (f"{item} folder copied from pending to Approval folder\n")
                        pending_copy_status = "Success"
                        event_log()     
                else:
                    for file in os.listdir(file_src):
                        shutil.copy2(os.path.join(file_src, file), file_dst)
                        run_time = runtime(start_time)
                        runtime_label.config(text = run_time)
                        message = (f"{item} folder copied from pending to Approval folder\n")
                        pending_copy_status = "Success"
                        event_log()     
            else:
                messagebox.showerror("Error", f"{item} Folder is empty")
        else:
            if not os.path.exists(file_src):
                messagebox.showerror("Error", f"{item} Folder not found in Pending Folder")
            
    # Function to use excel to read multiple entries
    def on_excel_button_click():
        pending_option = "excel"
        index = 0
        start_time = time.time()
        if read_excel_file() == "Error":
            return
        else:
            for item in standard_item_list:
                pending_copy_file(item, start_time, pending_option)
                index = index + 1
            
            run_time = runtime(start_time)
            runtime_label.config(text = run_time)
            message = (f"{index} Item(s) copied from pending to Approval folder")
            messagebox.showinfo("", f"Finished:\n{message}")
            event_log()
            pending_frame.destroy()
            on_pending_button_click()
                
    # Bind enter key to continue button
    pending_frame.bind("<Return>", lambda event: on_continue_button_click())

    style = ttk.Style()
    style.configure("CP.TButton", font = ("Arial", 9))

    ttk.Label(pending_frame, text = "Enter Part Number: ", font = ("Arial",12,"bold")).grid(row = 0, column = 0, columnspan = 3, pady = 5)
    item_entry = ttk.Entry(pending_frame, width = 20,justify = "center", font = ("Arial",12))
    item_entry.grid(row = 1, column = 0, columnspan = 3)
    item_entry.focus_set()

    ttk.Button(pending_frame, text = "Continue", command = on_continue_button_click, style = "CP.TButton").grid(row = 2, column = 1, pady = 10, padx = 2, sticky = "se")
    ttk.Button(pending_frame, text = "Exit", command = pending_frame.destroy, style = "CP.TButton").grid(row = 2, column = 2, pady = 10, padx = 2, sticky = "w")
    ttk.Button(pending_frame, text = "Use Excel", command = on_excel_button_click, style = "CP.TButton").grid(row = 2, column = 0, pady = 10, padx = 2, sticky = "e")
    pending_frame.mainloop()

# Clear list being used in message box
def clear_list():
    item_not_found_list.clear()
    item_in_pending_list.clear()
    item_not_in_approved_list.clear()

# Function for center window
def center_window(window):
    window.update_idletasks()
    width, height = window.winfo_width(), window.winfo_height()
    x = window.winfo_screenwidth() // 2 - width // 2
    y = window.winfo_screenheight() // 2 - height // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# Function for file select radio frame
def select_file_type():
    global select_file_frame, selected_file_type
    
    select_file_frame = tk.Toplevel() # Important to use top level window
    select_file_frame.title(" ")
    center_window(select_file_frame)
    select_file_frame.geometry("175x175")
    select_file_frame.iconbitmap(icon_path)
    select_file_frame.resizable(0,0)
    
    if main_button_option == "Copy":
        ttk.Label(select_file_frame, text = f"{copy_option_msg}", font = ("Arial",9,"bold")).pack()
    selected_file_type = tk.StringVar()
    selected_file_type.set(".pdf")
    rf_label = ttk.Label(select_file_frame, text = "Select File Type", font=header_font)
    rf_label.pack(anchor = tk.N)

    files = [".pdf", ".txt"] # To-do Add more file types if needed
    for file in files:
        ttk.Radiobutton(select_file_frame, text = file, variable = selected_file_type, value = file).pack(padx=5)

    # Function for continue button inside radio frame
    def on_continue_button_click():
        global start_time, file_type
        file_type = selected_file_type.get()
        start_time = time.time()
        if main_button_option == "Delete":
            delete_file()
        elif main_button_option == "Copy":
            read_excel_file()
        else:
            messagebox.showwarning("Error", "Please select an option")   
      
    continue_button = ttk.Button(select_file_frame, text = "Continue", command = on_continue_button_click)
    continue_button.pack(fill = 'x', padx = 20, pady = 10, side = tk.BOTTOM)
    select_file_frame.mainloop()

# Function for copy button options
def copy_checkbutton_option():

    global copy_checkbutton, overwrite_checkbutton, muraki_checkbutton
    
    copy_checkbutton = tk.Toplevel()
    copy_checkbutton.geometry("175x150")
    copy_checkbutton.title(" ")
    copy_checkbutton.iconbitmap(icon_path)
    copy_checkbutton.resizable(False,False)
    center_window(copy_checkbutton)
    # Function for continue button after copy option
    def on_continue_button_click():
        
        # Global variables to be used in user input function
        global overwrite_option, muraki_column_option, copy_option_msg

        if overwrite_checkbutton.get() == 1 and muraki_checkbutton.get() == 1:
            overwrite_option = "yes"
            muraki_column_option = "yes"
            overwrite_msg = "\nOverwriting existing files"
            muraki_option_msg = "\nUsing Muraki column"
            
        elif overwrite_checkbutton.get() == 1 or muraki_checkbutton.get() == 1:
            
            if overwrite_checkbutton.get() == 1:
                overwrite_option = "yes"
                overwrite_msg = "\nOverwriting existing files"
            else:
                overwrite_option = "no"
                overwrite_msg = "\nNot overwriting existing files"

            if muraki_checkbutton.get() == 1:
                muraki_column_option = "yes"
                muraki_option_msg = "\nUsing Muraki column"
            else:
                muraki_column_option = "no"
                muraki_option_msg = "\nNot using Muraki column"
        # When user chooses no options
        else:  
            overwrite_option = "no"
            muraki_column_option = "no"
            overwrite_msg = "\nNot overwriting existing files"
            muraki_option_msg = "\nNot using Muraki column"
        
        copy_option_msg = (f"{overwrite_msg + muraki_option_msg}\n")
        # Call to select_file_type function
        select_file_type()
       
    # To-do Will need to consolidate into one variable if adding more options
    overwrite_checkbutton = tk.IntVar()
    muraki_checkbutton = tk.IntVar()

    ttk.Label(copy_checkbutton, text = "Optional:",font = header_font).pack(side = tk.TOP, anchor = tk.N)
    ttk.Label(copy_checkbutton, text = "Select Options Below",font = header_font).pack(side = tk.TOP, anchor = tk.N)
    
    ttk.Checkbutton(copy_checkbutton, text = "Overwrite File", variable = overwrite_checkbutton).pack(side = tk.TOP, anchor = tk.N)
    ttk.Checkbutton(copy_checkbutton, text = "Use Muraki Column", variable = muraki_checkbutton).pack(side = tk.TOP, anchor = tk.N)

    continue_button = ttk.Button(copy_checkbutton, text = "Continue", command = on_continue_button_click)
    continue_button.pack(fill='x', padx = 20, pady = 10, side = tk.BOTTOM)
    copy_checkbutton.mainloop()
    
# Function to delete files
def delete_file():
    item_deleted = []
    
    if read_excel_file() == "Error":
        return
    else:
        if os.path.exists(delete_file_src):
            for standard_item in standard_item_list:
                file_dst = f"{delete_file_src}\\{standard_item}\\{standard_item}{file_type}" 
                if os.path.exists(file_dst):
                    os.remove(file_dst)
                    item_deleted.append(f"{standard_item}{file_type}")
                else:
                    item_not_in_approved_list.append(f"{standard_item}{file_type}") 
                    continue
                    
            if item_not_in_approved_list != []:

                not_approved_list = "\n".join(f"{item}" for item in item_not_in_approved_list)
                missing_message = (f"{len(item_not_in_approved_list)} print(s) not found in" 
                                    f" Approved Prints folder:\n{not_approved_list}\n")
            else:
                missing_message = ""

            if item_deleted != []:
                deleted_list = "\n".join(f"{item}" for item in item_deleted)
                deleted_message = (f"\n{len(item_deleted)} print(s) deleted:\n{deleted_list}\n")
                deleted_message_shorten = (f"\n{len(item_deleted)} print(s) deleted\n\n")
            else:
                deleted_message = ""
                deleted_message_shorten = ""

            global message
            message = (f"Delete Option:\n{missing_message}{deleted_message}")

            run_time = runtime(start_time)
            runtime_label.config(text = run_time)
            select_file_frame.destroy()
            messagebox.showinfo(f"({file_type}) Deleted", f"{deleted_message_shorten}{missing_message}")
            # Call to event_log function
            event_log()
        # Handles file path error
        else:
            messagebox.showerror("Approved Print",f"{delete_file_src} does not exist")
        return

# Function for displaying runtime
def runtime(start_time):
    global run_time
    end_time = time.time()
    program_runtime = end_time - start_time
    if program_runtime < 60:
       run_time = (f"Program Runtime: {round(program_runtime,5)} seconds")
    elif 60 <= program_runtime < 3600:
        run_time = (f"Program Runtime: {round((program_runtime/60),2)} minutes")
    elif 3600 <= program_runtime < 86400:
        run_time = (f"Program Runtime: {round((program_runtime/3600),2)} hours")
    return run_time

# Retriving data from excel file
def read_excel_file():
    global standard_item_list, muraki_item_list
    standard_item_list = []
    muraki_item_list = []
    excel_error = False

    # Closes copy_checkbutton if it exists
    if main_button_option == "Copy":
        copy_checkbutton.destroy()
        select_file_frame.destroy()

    # Reading from .xlsx or .csv file
    if os.path.exists(excel_file):
        if excel_file.endswith(".csv"):
            excel_item = pd.read_csv(excel_file)
            excel_folder_column = "Item"
            excel_file_column = "FileName"

        elif excel_file.endswith(".xlsx"):
            excel_item = pd.read_excel(excel_file)
            excel_folder_column = "standard_name"
            excel_file_column = "muraki_name"

        # Handles missing columns
        if excel_folder_column not in excel_item.columns or excel_file_column not in excel_item.columns:
                messagebox.showerror("Excel File",f"{excel_file} does not contain required columns")
                excel_error = "Column Error"
        else:       
            for index, row in excel_item.iterrows():
                if main_button_option == "Delete" or main_button_option == "Pending":
                    standard_item_list.append((row[excel_folder_column]))

                if main_button_option == "Copy":
                    if pd.isnull(row[excel_file_column]):
                        muraki_item_list = []
                    else:
                        muraki_item_list.append((row[excel_file_column], index))
                    standard_item_list.append((row[excel_folder_column], index))
            # Handles empty excel file
            if standard_item_list == []:
                excel_error = True
    else:
        messagebox.showerror("Excel File",f"{excel_file} does not exist")
        excel_error = True

    # Displaying number of items read
    if excel_error == False:
        excel_msg = (f"{len(standard_item_list)} Item(s) read From Excel")
        item_read_label.config(text = excel_msg)

        if main_button_option == "Copy":
            user_input(standard_item_list, muraki_item_list,current_item_checking)

    # Handles empty and not found excel file path
    if excel_error == True or excel_error == "Column Error":
        if excel_error != "Column Error":
            error_message = (f"{excel_file} is empty or not found")
            messagebox.showerror("Error", error_message)
        if main_button_option == "Copy":
            copy_checkbutton_option()

        if main_button_option == "Delete":
            select_file_frame.destroy()

        if main_button_option == "Pending":
            if excel_error != "Column Error":
                pending_frame.destroy()
            on_pending_button_click()
        return "Error"
# User input choices
def user_input(standard_item_list, muraki_item_list,current_item_checking):

    if overwrite_option in ["yes", "no"] and muraki_column_option in ["yes", "no"]:
        # Checks if the Muraki column is not empty and user selection
        if muraki_column_option == "yes" and muraki_item_list != []:

            for standard_info in standard_item_list:
                standard_item, standard_id = standard_info  
                for muraki_info in muraki_item_list:
                    muraki_item, muraki_id = muraki_info  
                    # Checking if id matches
                    if standard_id == muraki_id:
                        file_dst = f"{copy_file_dst}\\{standard_item}" # File destination path - change if needed

                        standard_item_file = f"{standard_item}{file_type}"
                        muraki_item_file = f"{muraki_item}{file_type}"
                        file_names = [standard_item_file, muraki_item_file]
                        # Call to check_file function to check if files/folders exist
                        for current_item_checking in file_names:
                            file_src = f"{copy_file_src}\\{current_item_checking}"
                            check_file(standard_item, file_src, file_dst, current_item_checking, overwrite_option)
            # Call to print_list function to display output
            print_list()
        
        # Checks if the Muraki column is empty but user selected yes to using column
        elif muraki_column_option == "yes" and muraki_item_list == []:

            last_text = os.path.basename(excel_file)

            messagebox.showinfo("Error",f"Please fill in the 'Muraki column in' -- {last_text} --"
                                "or do not select 'Use Muraki Name' if you do not want the Muraki column")
            # call to on_copy_button_click function again
            on_copy_button_click()

        # If not using the Muraki column in excel sheet               
        elif muraki_column_option == "no":

            for standard_info in standard_item_list:
                standard_item, standard_id = standard_info 

                file_dst = f"{copy_file_dst}\\{standard_item}" # File destination path - change if needed
                standard_item_file = f"{standard_item}{file_type}"
                file_names = [standard_item_file]
                # Call to check_file function to check if files/folders exist
                for current_item_checking in file_names:
                    file_src = f"{copy_file_src}\\{current_item_checking}"
                    check_file(standard_item, file_src, file_dst, current_item_checking, overwrite_option)
            # Call to print_list function
            print_list()
        # Handles unknown error
        else:
            messagebox.showerror("Error - User Input","Error in user selection") 
    else:
        messagebox.showerror("Error", "Error in User_Input function")

# Check if source file and destination folders exist
def check_file(standard_item, file_src, file_dst, current_item_checking, overwrite_option):
    
    # Creates folder at destination if it does not exist
    if not os.path.exists(file_dst):
        os.mkdir(file_dst)

    # Prevents overwriting 
    if overwrite_option == "no":
        if os.path.isfile(os.path.join(file_dst, current_item_checking)):
            if current_item_checking not in item_in_pending_list:
                item_in_pending_list.append(current_item_checking)
        elif not os.path.isfile(os.path.join(file_dst, current_item_checking)):
            if os.path.isfile(file_src):
                shutil.copy(file_src, file_dst)
            else:
                item_not_found_list.append(current_item_checking) 
        else:
            messagebox.showerror("Error", f"Error in check_file function overwrite = NO for {current_item_checking}")
    
    # Copies file to destination folder
    if overwrite_option == "yes":
        if os.path.isfile(file_src):
            shutil.copy(file_src, file_dst)
        elif not os.path.isfile(file_src):
            item_not_found_list.append(current_item_checking)
        else:
            messagebox.showerror("Error", f"Error in check_file function overwrite = YES for {current_item_checking}")
       
# Prints list of items not found in output folder and list of items that have already been found in pending approval
def print_list():
    global message
    pending_list = []
    not_found_list = []
    message = ""
    run_time = runtime(start_time)
    runtime_label.config(text = run_time)

    if item_in_pending_list != [] or item_not_found_list != []:

        if item_in_pending_list != []:
            pending_list = "".join(f"\n{file}" for file in item_in_pending_list)
            pending_message = (f"\n{len(item_in_pending_list)} print(s) already exist in Pending Approval (Destination) folder:{pending_list}\n")
            pending_message_shorten = (f"\n{len(item_in_pending_list)} print(s) already exist in Pending Approval (Destination) folder.\n")
        else:
            pending_message = "" 
            pending_message_shorten = ""

        if item_not_found_list != []:
            not_found_list = "".join(f"\n{item}" for item in item_not_found_list)
            not_found_message = (f"\n{len(item_not_found_list)} print(s) not found in AutoPrint Output (Source) folder:{not_found_list}\n")
            not_found_message_shorten = (f"\n{len(item_not_found_list)} print(s) not found in AutoPrint Output (Source) folder.\n")
        else:
            not_found_message = ""
            not_found_message_shorten = ""

        message = (f"Copy Option: {pending_message}{not_found_message}")
        message_shorten = (pending_message_shorten + not_found_message_shorten)
        messagebox.showinfo("Finished", message_shorten)

        # Call to event_log function
        event_log()
    else:
        messagebox.showinfo("Finished", "All Done")
    select_file_frame.destroy()
# Store event log of run time and output
def event_log():
        timestamp = datetime.now().strftime("%m-%d-%Y %H:%M:%S")

        with open(event_log_file, "r") as f:
            event_log_content = f.read()
        
        with open(event_log_file, "w") as f:
            f.write(f"\n--------Program Executed at: {timestamp}--------\n")
            f.write(f"{run_time}\n")
            f.write(f"\n--Output--\n{message}\n")
            f.write(event_log_content)
# Main    
if __name__ == '__main__':
    try:
        # Call to main function
        main()
    except Exception as e:
        messagebox.showinfo("Error", f"An error occurred: {e}")
else:
    print("AutoPrintMoveFile.py is being imported")