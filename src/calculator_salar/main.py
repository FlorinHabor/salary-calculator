import customtkinter
from core import parse_excel_file as parse_excel_file
from core import get_data_for_employee as get_data_for_employee
import os
from customtkinter import CTkToplevel 
from tkinter import ttk
from tkinter import filedialog
import openpyxl

app = customtkinter.CTk()
app.title("Calculator Salarii")
style = ttk.Style()

style.configure(
    "Treeview",
    font=("Segoe UI", 10),      # change size here
    rowheight=30                # increase row spacing here
)

style.configure(
    "Treeview.Heading",
    font=("Segoe UI", 12, "bold")
)

global file_path
global file_name_label
global results_table
global month_working_days

def import_button_callback():
    global file_path
    global file_name_label
    file_path = filedialog.askopenfilename(title="Selectează un fișier", filetypes=[("Excel files", ["*.xlsx" , "*.xls"])])
    if file_path:
        file_name_label.configure(text=os.path.basename(file_path))
    
def hours_to_hhmm(hours: float) -> str:
    h = int(hours)
    m = round((hours - h) * 60)
    return f"{h:02d}:{m:02d}"

def calculate_missing_hoours(worked_hours): 
    global month_working_days
    if month_working_days.get().isdigit():
        total_required_hours = int(month_working_days.get()) * 8
        missing_hours = max(0, total_required_hours - worked_hours)
        if missing_hours > 8:
            days = int(missing_hours / 8) 
            hours = missing_hours % 8
            if hours > 0.25:
                return f"{days} zile si {hours_to_hhmm(hours)}"
            else:
                return f"{days} zile"
        return hours_to_hhmm(missing_hours)
    else:
        return "N/A"

def calculate_button_callback():
    global results_table
    if file_path:
        employee_data = parse_excel_file(file_path)

        results_table.delete(*results_table.get_children())

        for name, data in employee_data.items():
            results_table.insert(
                "", "end",
                values=(
                    name,
                    data.get("worked_days", 0),
                    hours_to_hhmm(data.get("normal_hours", 0)),
                    hours_to_hhmm(data.get("overtime", 0)),
                    calculate_missing_hoours(data.get("normal_hours", 0)) if month_working_days.get().isdigit() else "N/A"
                )
            )
        
    else:
        print("Nu a fost selectat niciun fișier.")

def open_employee_modal(employee_name):
    global file_path
    employee_data = get_data_for_employee(employee_name, file_path) 

    modal = customtkinter.CTkToplevel(master=app)
    modal.geometry("1100x600")
    modal.title(f"Detalii pentru {employee_name}")
    
    details_table = ttk.Treeview(modal, columns=("Data", "Ore in program", "Ore totale", "Ore normale", "Ore suplimentare", "Ore lipsite"), show="headings")

    details_table.heading("Data", text="Data")
    details_table.heading("Ore in program", text="Ore in program")
    details_table.heading("Ore totale", text="Ore totale")
    details_table.heading("Ore normale", text="Ore normale")
    details_table.heading("Ore suplimentare", text="Ore suplimentare")
    details_table.heading("Ore lipsite", text="Ore lipsite")

    details_table.column("Data", anchor="center")
    details_table.column("Ore in program", anchor="center")
    details_table.column("Ore totale", anchor="center")
    details_table.column("Ore normale", anchor="center")
    details_table.column("Ore suplimentare", anchor="center")
    details_table.column("Ore lipsite", anchor="center")
    details_table.tag_configure("overtime_row", background="#9cfa9c")   
    details_table.tag_configure("missing_row", background="#f8a74a") 
    details_table.pack(padx=20, pady=20, fill="both", expand=True)
    details_table.style = style
    details_table.delete(*details_table.get_children())
    
    for row_data in employee_data:
        total_hours = row_data.get("total_hours", 0)

        row_tag = None

        if total_hours >= 8:
            row_tag = "overtime_row"
        elif total_hours < 8:
            row_tag = "missing_row" 
        details_table.insert(
            "", "end",
            values=(
                row_data.get("date", "Err"),
                hours_to_hhmm(row_data.get("required_hours", 8)),
                hours_to_hhmm(row_data.get("total_hours", 0)),
                hours_to_hhmm(row_data.get("normal_hours", 0)),
                hours_to_hhmm(row_data.get("overtime", 0)) if row_data.get("overtime", 0) != 0 else "N/A",
                hours_to_hhmm(row_data.get("missing_hours", 0)) if row_data.get("missing_hours", 0) > 0 else "N/A",
            ),
            tags=(row_tag,) if row_tag else ()
        )
        


def on_row_select(event):
    global results_table
    selected_item = results_table.focus()
    if not selected_item:
        return

    values = results_table.item(selected_item, "values")
    employee_name = values[0]

    open_employee_modal(employee_name) 

def create_gui(): 
    global file_name_label
    global results_table
    global month_working_days

    app.geometry("1200x700")
    customtkinter.set_appearance_mode("dark")
    
    app_title = customtkinter.CTkLabel(app, text="Calculator ore", font=customtkinter.CTkFont(size=30, weight="bold", family="Comic Sans MS"))
    app_title.configure(text_color="#F89161")
    app_title.pack(side = "top", padx=20, pady=20)

    inputs_frame = customtkinter.CTkFrame(app, height=100)
    inputs_frame.pack(side = "top", padx=20, pady=20, fill= "x")

    month_working_days = customtkinter.CTkEntry(inputs_frame, height=30, width=150, placeholder_text="Zile lucrătoare")
    month_working_days.pack(side = "left", padx=10, pady=10)

    import_button = customtkinter.CTkButton(inputs_frame, text="Încarcă un fișier", command=import_button_callback)
    import_button.pack(side = "left", padx=10, pady=10)
    import_button.configure(height=30, width=200)
    import_button.configure(fg_color="#C5511C", hover_color="#FC7638", text_color="white")
    
    file_name_label = customtkinter.CTkLabel(inputs_frame, text="")
    file_name_label.pack(side="left", padx=10, pady=10)

    calculate_button = customtkinter.CTkButton(inputs_frame, text="Calculează", command=calculate_button_callback)
    calculate_button.pack(side = "right", padx=10, pady=10)
    calculate_button.configure(height=30, width=200)
    calculate_button.configure(fg_color="#0EAD0E", hover_color="#73DF1B", text_color="white")

    table_frame = customtkinter.CTkFrame(app, height=500)
    table_frame.pack(side = "bottom", padx=20, pady=20, fill= "both", expand=True)

    results_table = ttk.Treeview(table_frame, columns=("Angajat", "Zile lucrate", "Ore normale", "Ore suplimentare", "Ore absentate"), show="headings")
    results_table.heading("Angajat", text="Angajat")
    results_table.heading("Zile lucrate", text="Zile lucrate")
    results_table.heading("Ore normale", text="Ore normale")
    results_table.heading("Ore suplimentare", text="Ore suplimentare")
    results_table.heading("Ore absentate", text="Ore absentate")
    results_table.column("Zile lucrate", anchor="center")
    results_table.column("Ore normale", anchor="center")
    results_table.column("Ore suplimentare", anchor="center")
    results_table.column("Ore absentate", anchor="center")
    results_table.bind("<<TreeviewSelect>>", on_row_select)
    results_table.pack(fill="both", expand=True, padx=10, pady=10, side= "top")
    
    app.mainloop()

    

def main() -> None:
    create_gui()


if __name__ == "__main__":
    main()