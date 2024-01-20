import tkinter as tk
from tkinter import filedialog, messagebox
import os
import PyPDF2
from pdf2docx import Converter
import fitz
from tkinter import IntVar

file_names = ''
file_list = ''


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f"{width}x{height}+{int(x)}+{int(y)}")


def choose_files(*quanity):
    global file_list, file_names

    if len(quanity) == 1:
        file = filedialog.askopenfilename(title="Choose Files", filetypes=[("PDF Files", "*.pdf")])
        file_list = file

        file_names = os.path.basename(file)

        selected_files_label.config(text="".join(file_names))
    else:
        files = filedialog.askopenfilenames(title="Choose Files", filetypes=[("PDF Files", "*.pdf")])
        file_list = list(files)

        file_names = [os.path.basename(file) for file in file_list]

        selected_files_label.config(text="\n".join(file_names))


def merge():
    global selected_files_label

    def merge_pdfs():

        print(file_list)
        if not file_list:
            print("You didnt choose files to merge.")
            return

        merged_pdf = PyPDF2.PdfMerger()

        for file in file_list:
            merged_pdf.append(file)

        output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
                                                        initialfile='Name of a file to save')
        if output_file_path:
            with open(output_file_path, "wb") as output_file:
                merged_pdf.write(output_file)

            print("Pliki zostały połączone w:", output_file_path)
        else:
            print("Operacja scalania została anulowana.")

    window.withdraw()
    window_merge = tk.Tk()
    window_merge.title("Merge")
    window_merge.configure(bg="black")
    window_width = 900
    window_height = 500
    center_window(window_merge, window_width, window_height)

    label2 = tk.Label(window_merge, text="Choose Files To Merge", font=('Arial', 40, 'bold'), fg='red', bg='black')
    label2.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

    choose_button = tk.Button(window_merge, text="Choose Files", font=('Arial', 15, 'bold'), fg='red', bg='black',
                              width=22,
                              activeforeground='white', activebackground='black', command=lambda: choose_files(2,1))
    choose_button.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    merge_files_button = tk.Button(window_merge, text="Merge Files", font=('Arial', 15, 'bold'), fg='red', bg='black',
                                   width=22,
                                   activeforeground='white', activebackground='black', command=merge_pdfs)
    merge_files_button.place(relx=0.5, rely=0.4, anchor=tk.CENTER)

    # Label do wyświetlenia nazw wybranych plików
    selected_files_label = tk.Label(window_merge, text=file_names, font=('Arial', 12), fg='white', bg='black',
                                    justify=tk.LEFT)
    selected_files_label.place(relx=0.5, rely=0.55, anchor=tk.CENTER)

    window_merge.mainloop()


def delete_pages():
    global selected_files_label, entry_var

    def remove_pages(entry, entry2=None):
        try:
            pages = entry.get()

            pdf_writer = PyPDF2.PdfWriter()
            outputfile = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
                                                      initialfile='Name_of_a_file_to_save.pdf')
            pdf_reader = PyPDF2.PdfReader(file_list)

            if entry2 == None:
                for page_number in range(len(pdf_reader.pages)):
                    if str(page_number + 1) not in pages:
                        strona = pdf_reader.pages[page_number]
                        pdf_writer.add_page(strona)
            else:
                max_sites = len(pdf_reader.pages)

                star_page = int(entry.get())
                end_page = int(entry2.get())

                if star_page > end_page:
                    star_page, end_page = end_page, star_page

                if max_sites < end_page:
                    result = messagebox.askyesno("You want do dele more sites than is ", "You want do dele more sites than is "
                                                                                         "in document . Do you want to continue")
                    if result:
                        pass
                    else:
                        return

                for page_num in range(max_sites):
                     if page_num < star_page - 1 or page_num > end_page - 1:
                        pdf_writer.add_page(pdf_reader.pages[page_num])


            with open(outputfile, 'wb') as outputfile:
                pdf_writer.write(outputfile)
            messagebox.showinfo("Sukces")
        except ValueError:
            messagebox.showerror("Błąd", "Wprowadź prawidłowe numery stron.")

    def radio_button_selected(selected_option):
        print(selected_option)
        if not file_list:
            messagebox.showinfo("info", "You didnt choose files to delete.")
            return
        else:
            # Usuń poprzednie elementy interfejsu z wyjątkiem RadioButtons
            for widget in window_delete.winfo_children():
                if not isinstance(widget, (tk.Radiobutton, tk.Label)):
                    widget.destroy()

            if selected_option == 1:
                print('Opcja 1')
                pages_label = tk.Label(window_delete, text="Podaj numery stron do usunięcia (oddzielone przecinkami):")
                pages_label.place(relx=0.2, rely=0.4, anchor=tk.CENTER)

                entry_var = tk.StringVar()

                pages_entry = tk.Entry(window_delete, textvariable=entry_var)
                pages_entry.place(relx=0.2, rely=0.5, anchor=tk.CENTER)

                delete_button = tk.Button(window_delete, text="Delete", font=('Arial', 15, 'bold'), fg='red',
                                          bg='black',
                                          width=22,
                                          activeforeground='white', activebackground='black',
                                          command=lambda: remove_pages(pages_entry))
                delete_button.place(relx=0.2, rely=0.6, anchor=tk.CENTER)

            elif selected_option == 2:
                print('Opcja 2')
                pages_label = tk.Label(window_delete, text="Podaj zakres stron do usunięcia:")
                pages_label.place(relx=0.8, rely=0.4, anchor=tk.CENTER)

                entry_var = tk.StringVar()
                entry_var2 = tk.StringVar()

                pages_entry_range = tk.Entry(window_delete, textvariable=entry_var)
                pages_entry_range.place(relx=0.8, rely=0.5, anchor=tk.CENTER)

                pages_entry_range2 = tk.Entry(window_delete, textvariable=entry_var2)
                pages_entry_range2.place(relx=0.8, rely=0.6, anchor=tk.CENTER)

                delete_button = tk.Button(window_delete, text="Delete", font=('Arial', 15, 'bold'), fg='red',
                                          bg='black',
                                          width=22,
                                          activeforeground='white', activebackground='black',
                                          command=lambda: remove_pages(pages_entry_range, pages_entry_range2))
                delete_button.place(relx=0.8, rely=0.7, anchor=tk.CENTER)

    window.withdraw()
    radio_var = tk.IntVar()
    radio_var.set(2)

    window_delete = tk.Tk()
    window_delete.title("Delete Pages")
    window_delete.configure(bg="black")
    window_width = 900
    window_height = 500
    center_window(window_delete, window_width, window_height)

    label2 = tk.Label(window_delete, text="Choose Pages To Delete", font=('Arial', 40, 'bold'), fg='red', bg='black')
    label2.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

    choose_button = tk.Button(window_delete, text="Choose File", font=('Arial', 15, 'bold'), fg='red', bg='black',
                              width=22,
                              activeforeground='white', activebackground='black', command=lambda: choose_files(1))
    choose_button.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Label do wyświetlenia nazw wybranych plików
    selected_files_label = tk.Label(window_delete, text="", font=('Arial', 12), fg='white', bg='black', justify=tk.LEFT)
    selected_files_label.place(relx=0.5, rely=0.55, anchor=tk.CENTER)

    radio_button1 = tk.Radiobutton(window_delete, text="Specific sites they can be \n seperate by comma",
                                   variable=radio_var, value=1,
                                   font=('Arial', 15, 'bold'), fg='red', bg='black',
                                   command=lambda: radio_button_selected(1))
    radio_button2 = tk.Radiobutton(window_delete, text="many sites in order \n e.g 1-10", variable=radio_var, value=2,
                                   font=('Arial', 15, 'bold'), fg='red', bg='black',
                                   command=lambda: radio_button_selected(2))

    radio_button1.place(relx=0.2, rely=0.3, anchor=tk.CENTER)
    radio_button2.place(relx=0.8, rely=0.3, anchor=tk.CENTER)
    print('działam')

    window_delete.mainloop()

def change_to_word():
    global selected_files_label

    def convert_to_word():
        output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Word Files", "*.docx")],
                                                        initialfile='Name of a file to save')
        cv = Converter(file_list)
        cv.convert(output_file_path)  # all pages by default
        cv.close()




    window.withdraw()
    radio_var = tk.IntVar()
    radio_var.set(2)

    window_word = tk.Tk()
    window_word.title("Delete Pages")
    window_word.configure(bg="black")
    window_width = 900
    window_height = 500
    center_window(window_word, window_width, window_height)
    label2 = tk.Label(window_word, text="Choose Pages To Delete", font=('Arial', 40, 'bold'), fg='red', bg='black')
    label2.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

    choose_button = tk.Button(window_word, text="Choose File", font=('Arial', 15, 'bold'), fg='red', bg='black',
                              width=22,
                              activeforeground='white', activebackground='black', command=lambda: choose_files(1))
    choose_button.place(relx=0.5, rely=0.3, anchor=tk.CENTER)


    convert_button = tk.Button(window_word, text="Convert to word", font=('Arial', 15, 'bold'), fg='red', bg='black',
                              width=22,
                              activeforeground='white', activebackground='black',command=convert_to_word)
    convert_button.place(relx=0.5, rely=0.45, anchor=tk.CENTER )

    # Label do wyświetlenia nazw wybranych plików
    selected_files_label = tk.Label(window_word, text="", font=('Arial', 12), fg='white', bg='black', justify=tk.LEFT)
    selected_files_label.place(relx=0.5, rely=0.55, anchor=tk.CENTER)


    print('działam')

    window_word.mainloop()

def App():
    global window
    window = tk.Tk()
    window.title("PDF Simple Manipulation")

    window_width = 900
    window_height = 500
    center_window(window, window_width, window_height)
    window.configure(bg="black")

    label2 = tk.Label(window, text="PDF Changer", font=('Arial', 40, 'bold'), fg='red', bg='black')
    label2.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

    merge_button = tk.Button(window, text="Merge", font=('Arial', 15, 'bold'), fg='red', bg='black', width=22,
                             activeforeground='white', activebackground='black', command=merge)
    merge_button.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    delete_button = tk.Button(window, text="Delete Pages", font=('Arial', 15, 'bold'), fg='red', bg='black', width=22,
                              activeforeground='white', activebackground='black', command=delete_pages)
    delete_button.place(relx=0.5, rely=0.45, anchor=tk.CENTER)

    transfer_text_button = tk.Button(window, text="Transfer Text with picture", font=('Arial', 15, 'bold'), fg='red',
                                     bg='black'
                                     , width=22, activeforeground='white', activebackground='black', )
    transfer_text_button.place(relx=0.5, rely=0.75, anchor=tk.CENTER)

    change_to_word_button = tk.Button(window, text="Change to word file", font=('Arial', 15, 'bold'), fg='red',
                                     bg='black'
                                     , width=22, activeforeground='white', activebackground='black', command=change_to_word)
    change_to_word_button.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

    window.mainloop()


App()
