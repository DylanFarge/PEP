import os
import webbrowser
import tkinter as tk
from src.mark_ratings import mark
from tkinter.filedialog import askopenfilename, askdirectory

def select_file(excel_type):
    if excel_type == "xlsx":
        Application.group_file = askopenfilename(title="Select file", filetypes=(("Excel files", "*.xlsx"),))
        Application.label.config(text=Application.group_file if Application.group_file else "no file selected...") 
        Application.label.config(fg="black") if Application.group_file else None
    elif excel_type == "csv":
        Application.ratings_file = askopenfilename(title="Select file", filetypes=(("CSV files", "*.csv"),))
        Application.label2.config(text=Application.ratings_file if Application.ratings_file else "no file selected...") 
        Application.label2.config(fg="black") if Application.group_file else None
    else:
        raise ValueError(f"Unknown excel type: {excel_type}. Only accepts 'xlsx' or 'csv'")

def select_directory():
    Application.output_dir = askdirectory(title="Select directory")
    Application.label3.config(text=Application.output_dir) if Application.output_dir else None

def execute():
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
            ], self_rate=Application.check_var.get())
            Application.btn4.config(bg="green")
            Application.errorLabel["text"] = ""
        except AttributeError as e:
            print(e)
        except Exception as e:
            print(e)
            Application.btn4.config(bg="red")
            Application.errorLabel = tk.Label(master=Application.frameMain, text="Something went wrong with processing the group allocation file and/or the peer review file. Please make sure the files are correct and try again.", bg="white", fg="red", wraplength=500)
            Application.errorLabel.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
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

    app.title("PRP: Peer Review Processor")
    app.configure(bg="white")

    frameREADME = tk.Frame(relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frameREADME.columnconfigure(0, weight=1)
    frameREADME.columnconfigure(1, weight=1)
    frameREADME.columnconfigure(2, weight=1)
    frameREADME.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    label = tk.Label(master=frameREADME, text="How To", bg="white", font=get_fonts("heading"))
    label.grid(row=0, column=0, columnspan=3, padx=5, pady=5)

    btn_PDF = tk.Button(master=frameREADME, text="README.pdf", width=21, command=lambda: os.system("open README.pdf"), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_PDF.grid(row=1, column=0, padx=5, pady=5)

    btn_MARKDOWN = tk.Button(master=frameREADME, text="README.md", width=21, command=lambda: os.system("open README.md"), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_MARKDOWN.grid(row=1, column=1, padx=5, pady=5)

    btn_GITHUB = tk.Button(master=frameREADME, text="GitHub", width=21, command=lambda: webbrowser.open('https://github.com/DylanFarge/PRP'), bg="#0BF", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn_GITHUB.grid(row=1, column=2, padx=5, pady=5)

    frameMain = tk.Frame(relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frameMain.columnconfigure(0, weight=1)
    frameMain.columnconfigure(1, weight=1)
    frameMain.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ns")

    label = tk.Label(master=frameMain, text="PRP: Peer Review Processor", bg="white", font=get_fonts("heading"))
    label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    frame1 = tk.Frame(master=frameMain, relief=tk.RAISED, borderwidth=1, bg="white", highlightbackground="#0AF", highlightthickness=2)
    frame1.columnconfigure(0, weight=1)
    frame1.columnconfigure(1, weight=3)
    frame1.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

    btn = tk.Button(master=frame1, text="Select Group Allocation File", width=25, command=lambda: select_file("xlsx"), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label = tk.Label(master=frame1, text=group_file, bg="white")
    btn.grid(row=0, column=0, padx=5, pady=5)
    label.grid(row=0, column=1, padx=5, pady=5)

    btn2 = tk.Button(master=frame1, text="Select Peer Ratings File", width=25, command=lambda: select_file("csv"), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label2 = tk.Label(master=frame1, text=ratings_file, bg="white")
    btn2.grid(row=1, column=0, padx=5, pady=5)
    label2.grid(row=1, column=1, padx=5, pady=5)

    check_var = tk.IntVar(value=1)
    check = tk.Checkbutton(master=frameMain, text="Students can rate themselves", variable=check_var, bg="white", highlightthickness=0, activeforeground="#08F", activebackground="white", font=get_fonts("subheading"))
    check.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

    btn3 = tk.Button(master=frameMain, text="Select Output File", width=25, command=lambda: select_directory(), bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    label3 = tk.Label(master=frameMain, text=output_dir, bg="white")

    btn3.grid(row=3, column=0, padx=5, pady=5)
    label3.grid(row=3, column=1, padx=5, pady=5)

    btn4 = tk.Button(master=frameMain, text="Run", width=25, command=execute, bg="#08F", fg="white", activebackground="#0FF", font=get_fonts("subheading"))
    btn4.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

    def run():
        Application.app.mainloop()

if __name__ == "__main__":
    Application.run()