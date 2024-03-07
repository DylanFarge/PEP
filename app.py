import os
import webbrowser
import tkinter as tk
from src.mark_ratings import mark
from tkinter.filedialog import askopenfilename, askdirectory
import sys
from src.mark_ratings import run_cmd
from src.self_learn_groups import learn_groups

def select_file(excel_type):
    if excel_type == "xlsx":
        gf = askopenfilename(title="Select file", filetypes=(("Excel files", "*.xlsx"),))
        if gf == None or gf == "" or gf == ' ' or gf == ():
            Application.group_file = "no file selected..."
            Application.label.config(text="no file selected...")
        else:
            Application.group_file = gf
            Application.label.config(text=gf)
            Application.label.config(fg="black")
    elif excel_type == "csv":
        pf = askopenfilename(title="Select file", filetypes=(("CSV files", "*.csv"),))
        if pf == None or pf == "" or pf == ' ' or pf == ():
            Application.ratings_file = "no file selected..."
            Application.label2.config(text="no file selected...")
        else:
            Application.ratings_file = pf
            Application.label2.config(text=pf)
            Application.label2.config(fg="black")
    else:
        raise ValueError(f"Unknown excel type: {excel_type}. Only accepts 'xlsx' or 'csv'")

def select_directory():
    od = askdirectory(title="Select directory")
    if od != None and od != "" and od != ' ' and od != ():
        Application.output_dir = od
        Application.label3.config(text=od)

def execute():
    
    try:
        float(Application.cap_entry.get())
        Application.cap_entry.config(highlightbackground='black', highlightcolor='black', highlightthickness=1)
    except ValueError:
        Application.cap_entry.config(highlightbackground='red', highlightcolor='red', highlightthickness=2)
        return
    
    Application.btn4.config(state="disabled")
    if Application.group_file == "no file selected..." or Application.group_file == "":
        Application.label.config(fg="red")
        Application.btn4.config(state="normal")
        raise FileNotFoundError("No group allocation file selected. Please select a file.")
    elif Application.ratings_file == "no file selected..." or Application.ratings_file == "":
        Application.label2.config(fg="red")
        Application.btn4.config(state="normal")
        raise FileNotFoundError("No peer ratings file selected. Please select a file.")
    else:
        try:
            print("Students can rate themselves:", Application.check_var.get())
            mark([
                    Application.group_file, 
                    Application.ratings_file,
                    Application.output_dir,
                ], 
                self_rate=Application.check_var.get(),
                settings= {
                    "capped_mark": float(Application.cap_entry.get()),
                    "correct_ids":Application.check_var2.get(),
                    "create_file":True
                }
            )
            Application.btn4.config(bg="green")
            Application.errorLabel["text"] = ""
        except AttributeError as e:
            print(e)
        except Exception as e:
            print(e)
            Application.btn4.config(bg="red")
            Application.errorLabel = tk.Label(master=Application.frameMain, text="Something went wrong with processing the group allocation file and/or the peer review file. Please make sure the files are correct or that the column names are what is expected.", bg="white", fg="red", wraplength=500)
            Application.errorLabel.grid(row=5, column=0, columnspan=2, padx=10, pady=10)
        finally:
            Application.btn4.config(state="normal")
            

def get_fonts(texttype: str):
    return dict(
        heading=("Arial", 23, "bold", "underline"),  
        subheading=("Arial", 12, "bold"),
    )[texttype]

class Application():

    group_file = "no file selected..."
    ratings_file = "no file selected..."
    output_dir = os.getcwd()+"/output"
    
    app = tk.Tk()
    app.columnconfigure(0, weight=1)
    app.columnconfigure(1, weight=3)

    app.title("PEP: Peer Evaluation Program")
    app.configure(bg="white")

    frameREADME = tk.Frame(relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frameREADME.columnconfigure(0, weight=1)
    frameREADME.columnconfigure(1, weight=1)
    frameREADME.columnconfigure(2, weight=1)
    frameREADME.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    label = tk.Label(master=frameREADME, text="How To", bg="white", font=get_fonts("heading"))
    label.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

    btn_PDF = tk.Button(master=frameREADME, text="README.pdf", width=21, command=lambda: os.system("open README.pdf"), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_PDF.grid(row=1, column=0, padx=10, pady=10)

    btn_MARKDOWN = tk.Button(master=frameREADME, text="README.md", width=21, command=lambda: os.system("open README.md"), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_MARKDOWN.grid(row=1, column=1, padx=10, pady=10)

    btn_GITHUB = tk.Button(master=frameREADME, text="GitHub", width=21, command=lambda: webbrowser.open('https://github.com/DylanFarge/PRP'), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_GITHUB.grid(row=1, column=2, padx=10, pady=10)

    frameMain = tk.Frame(relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frameMain.columnconfigure(0, weight=1)
    frameMain.columnconfigure(1, weight=1)
    frameMain.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    label = tk.Label(master=frameMain, text="PEP: Peer Evaluation Program", bg="white", font=get_fonts("heading"))
    label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

    frame1 = tk.Frame(master=frameMain, relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frame1.columnconfigure(0, weight=1)
    frame1.columnconfigure(1, weight=3)
    frame1.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

    btn = tk.Button(master=frame1, text="Select Group Allocation File", width=25, command=lambda: select_file("xlsx"), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label = tk.Label(master=frame1, text=group_file, bg="white")
    btn.grid(row=0, column=0, padx=10, pady=10)
    label.grid(row=0, column=1, padx=10, pady=10)

    btn2 = tk.Button(master=frame1, text="Select Peer Ratings File", width=25, command=lambda: select_file("csv"), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label2 = tk.Label(master=frame1, text=ratings_file, bg="white")
    btn2.grid(row=1, column=0, padx=10, pady=10)
    label2.grid(row=1, column=1, padx=10, pady=10)

    frameSettings = tk.Frame(master=frameMain, bg="white")
    frameSettings.columnconfigure(0, weight=1)
    frameSettings.columnconfigure(1, weight=1)
    frameSettings.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    check_var = tk.IntVar(value=1)
    check = tk.Checkbutton(master=frameSettings, text="Students can rate themselves", variable=check_var, bg="white", highlightthickness=0, activeforeground="#08F", activebackground="white", font=get_fonts("subheading"))
    check.grid(row=0, column=0, padx=10, pady=10)

    check_var2 = tk.IntVar(value=1)
    check2 = tk.Checkbutton(master=frameSettings, text="Auto correct IDs", variable=check_var2, bg="white", highlightthickness=0, activeforeground="#08F", activebackground="white", font=get_fonts("subheading"))
    check2.grid(row=0, column=1, padx=10, pady=10)

    label4 = tk.Label(master=frameSettings, text="Capped Scale:", bg="white", font=get_fonts("subheading"), justify='left')
    label4.grid(row=1, column=0, padx=10, pady=10, sticky="E")

    cap_entry = tk.Entry(master=frameSettings, bg="white", font=get_fonts("subheading"), justify='left', textvariable=tk.DoubleVar(value=1.1))
    cap_entry.grid(row=1, column=1, padx=10, pady=10)

    # This is in attempt to deselect the Capped Mark entry.
    def change_focus(event):
        try: event.widget.focus_set() 
        except: pass
    frameSettings.bind_all("<Button>", change_focus)

    btn3 = tk.Button(master=frameMain, text="Select Output File", width=25, command=lambda: select_directory(), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label3 = tk.Label(master=frameMain, text=output_dir, bg="white")

    btn3.grid(row=3, column=0, padx=10, pady=10, sticky="E")
    label3.grid(row=3, column=1, padx=10, pady=10, sticky="W")

    btn4 = tk.Button(master=frameMain, text="Run", width=25, command=execute, bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn4.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

    def run():
        Application.app.mainloop()

if __name__ == "__main__":

    if len(sys.argv) > 1:

        if sys.argv[1] == "cmd":
            run_cmd()

        elif sys.argv[1] == "run":
            Application.run()

        elif sys.argv[1] == "self-learn":
            if len(sys.argv) > 2:
                path = sys.argv[2]
            else:
                path = "~/PEP/files/p1/2024-MT00548-Project 1 Chat peer rating-responses.csv"
            
            learn_groups(path, to_file=True)

        else:
            print("Unknown command:", sys.argv[1])
    else:
        spacing = 30
        print("----------------------------------------------<<< PEP OPTIONS >>>----------------------------------------------")
        print(f"{'run':>{spacing}} : Run the GUI.")
        print(f"{'cmd':>{spacing}} : Run via the command line with hard-coded files and output directory.")
        print(f"{'cmd <group> <ratings> <output>':>{spacing}} : Run the command line interface with the given files and output directory.")
        print(f"{'self-learn':>{spacing}} : Run the self-learning groups function with hard-coded file and output directory.")
        print(f"{'self-learn <ratings file>':>{spacing}} : Generates groups based off of peer evaluations' cross referencing. Outputs to 'output/'")
        print("---------------------------------------------------------------------------------------------------------------")
