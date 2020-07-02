from tkinter import *
from tkinter import scrolledtext, filedialog, messagebox
from GUI.texts import *
from core.parser import parse_excel


exc_path = "./data.xlsx"


def on_out_mult_pressed():
    messagebox.showinfo(DIALOG_ERROR_TITLE, ERROR_WIP)


def on_parse_btn_pressed():
    tmp = template_box.get(1.0, END)
    try:
        parsed = parse_excel(exc_path, tmp)
        preview_box.delete(1.0, END)
        preview_box.insert(INSERT, parsed)
    except:
        messagebox.showinfo(DIALOG_ERROR_TITLE, ERROR_CANNOT_PARSE)


def on_xlsx_btn_pressed():
    global exc_path
    exc_path = filedialog.askopenfilename(
        title=XLSX_IMPORT_DIALOG_TITLE,
        filetypes=(
            ("Excel 2007+", "*.xlsx"),
            ("Excel 2003-", "*.xls"),
            ("All files", "*.*"),
        ),
    )


def on_temp_btn_pressed():
    temp_path = filedialog.askopenfilename(
        title=TEMP_IMPORT_DIALOG_TITLE,
        filetypes=(("Text", "*.txt"), ("All files", "*.*")),
    )
    with open(temp_path, encoding="utf8") as f:
        temp = f.read()
    template_box.delete(1.0, END)
    template_box.insert(INSERT, temp)


def on_out_sing_pressed():
    f_path = filedialog.asksaveasfilename(
        title=EXPORT_DIALOG_TITLE, filetypes=(("Text", "*.txt"), ("All files", "*.*")),
    )
    with open(f_path, "w") as f:
        f.write(preview_box.get(1.0, END))


# initialization
def main():
    global template_box, preview_box

    root = Tk()
    root.title(WIN_TITLE)

    output_fm = Frame(root)
    output_fm.grid(column=0, row=0, padx=10, pady=2)

    control_fm = Frame(root)
    control_fm.grid(column=1, row=0, padx=10, pady=5)

    # preview box
    preview_lb = Label(output_fm, text=OUTPUT_LB)
    preview_lb.grid(column=0, row=0)

    preview_box = scrolledtext.ScrolledText(output_fm, width=50, height=10)
    preview_box.grid(column=0, row=1)
    preview_box.insert(INSERT, DEFAULT_OUTPUT)

    # output botton
    output_btn_fm = Frame(output_fm)
    output_btn_fm.grid(column=0, row=2)

    phase_btn = Button(
        output_btn_fm, text=PHASE_BTN, bg="green", command=on_parse_btn_pressed
    )
    phase_btn.grid(column=0, row=0)

    output_btn_lb = Label(output_btn_fm, text=OUTPUT_BTN_LB)
    output_btn_lb.grid(column=1, row=0, padx=10)

    output_sing_btn = Button(
        output_btn_fm, text=OUTPUT_SING_BT, command=on_out_sing_pressed
    )
    output_sing_btn.grid(column=2, row=0)

    output_sing_btn = Button(
        output_btn_fm, text=OUTPUT_MULT_BT, command=on_out_mult_pressed
    )
    output_sing_btn.grid(column=3, row=0)

    # texplate text box
    template_lb = Label(control_fm, text=TEMP_LB)
    template_lb.grid(column=0, row=0, padx=10)
    template_box = scrolledtext.ScrolledText(control_fm, width=25, height=10)
    template_box.grid(column=0, row=1)

    # import buttons
    import_btn_fm = Frame(control_fm)
    import_btn_fm.grid(column=0, row=3)

    import_lb = Label(import_btn_fm, text=IMPORT_LB)
    import_lb.grid(column=0, row=0, padx=10)
    import_xlsx_btn = Button(
        import_btn_fm, text=IMPORT_XLSX_BTN, command=on_xlsx_btn_pressed
    )
    import_xlsx_btn.grid(column=1, row=0)
    import_temp_btn = Button(
        import_btn_fm, text=IMPORT_TEMP_BTN, command=on_temp_btn_pressed
    )
    import_temp_btn.grid(column=2, row=0)

    root.mainloop()
