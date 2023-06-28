from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path


def main():
    root = Tk()
    app = Window(root)
    root.mainloop()


class Window():
    def __init__(self, master):
        self.master = master
        self.master.title("PDF - Extract & Merge")
        self.master.iconbitmap(fr"Z:\Python_Projects\PDF_Editor\python.ico")
        self.master.geometry('800x200')

        self.RadioFrame = LabelFrame(self.master, text="Functions", width=250, height=150)
        self.RadioFrame.place(x=10, y=10)

        Options = {
            "Extract All": "extractall",
            "Extract Range": "extractrange",
            "Merge": "merge"
        }

        selection = StringVar()
        selection.set("extractall")

        for text, value in Options.items():
            Radiobutton(self.RadioFrame, text=text, variable=selection, value=value, command=lambda: Clicked(selection.get())).pack(anchor=W)

        def Clicked(variable):
            if selection.get() == "extractall":
                self.ExtractAllFrame = LabelFrame(self.master, text="Extract All Pages", width=250, height=150)
                self.ExtractAllFrame.place(x=125, y=10)
                frame = self.ExtractAllFrame
                def extractall():
                    try:
                        # Fetch the filepath
                        absfilepath = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=[("PDF files", "*.pdf")])
                        filepath = absfilepath.replace('/', '\\')
                        p = Path(filepath)
                        outputdir = str(p.parent)
                        filename_only = str(p.name).split('.')[0]
                        fileext = str(p.suffix)
                        Label(frame, text=f"Selected File:- {filepath}").grid(row=5, column=0, sticky='w')
                        Label(frame, text=f"Output Directory:- {outputdir}").grid(row=6, column=0, sticky='w')
                        # Extract all pages of the PDF file
                        input_pdf = PdfReader(filepath)
                        for i, page in enumerate(input_pdf.pages):
                            output = PdfWriter()
                            output.add_page(page)
                            extracted_filepath = fr"{outputdir}\{filename_only}_{i+1}{fileext}"
                            with open(extracted_filepath, "wb") as output_stream:
                                output.write(output_stream)
                        messagebox.showinfo(title="Success", message="Task completed successfully.")
                    except Exception as e:
                        messagebox.showerror(title="ERROR", message=fr"Error:- {e}")
                Button(frame, text="Select File", command=lambda: extractall()).grid(row=2, column=1, padx=5, pady=5)
                try:
                    self.ExtractRangeFrame.destroy()
                    self.MergeFrame.destroy()
                except AttributeError:
                    pass

            elif selection.get() == "extractrange":
                self.ExtractRangeFrame = LabelFrame(self.master, text="Extract Custom Range", width=250, height=150)
                self.ExtractRangeFrame.place(x=125, y=10)
                frame = self.ExtractRangeFrame
                Label(frame, text="Step# 1 --> Start Page#").grid(row=0, column=0)
                e1 = Entry(frame)
                e1.grid(row=0, column=1)
                e1.insert(0, "2")
                Label(frame, text="Step# 2 --> End Page#").grid(row=1, column=0)
                e2 = Entry(frame)
                e2.grid(row=1, column=1)
                e2.insert(0, "3")
                def extractrange():
                    try:
                        # Fetch the filepath
                        absfilepath = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=[("PDF files", "*.pdf")])
                        filepath = absfilepath.replace('/', '\\')
                        p = Path(filepath)
                        outputdir = str(p.parent)
                        filename_only = str(p.name).split('.')[0]
                        fileext = str(p.suffix)
                        Label(frame, text=f"Selected File:- {filepath}").grid(row=5, column=0, sticky='w')
                        Label(frame, text=f"Output Directory:- {outputdir}").grid(row=6, column=0, sticky='w')
                        # Extract the pages from the give range
                        input_pdf = PdfReader(filepath)
                        output = PdfWriter()
                        rangestart = int(e1.get())
                        rangeend = int(e2.get())
                        for i in range(rangestart, (rangeend + 1)):
                            output.add_page(input_pdf.pages[i-1])
                        extracted_filepath = fr"{outputdir}\{filename_only}_{rangestart}-{rangeend}{fileext}"
                        with open(extracted_filepath, "wb") as output_stream:
                            output.write(output_stream)
                        Label(frame, text=f"Extracted File:- {extracted_filepath}", fg="green").grid(row=7, column=0, sticky='w')
                        messagebox.showinfo(title="Success", message="Task completed successfully.")
                    except Exception as e:
                        messagebox.showerror(title="ERROR", message=fr"Error:- {e}")
                Label(frame, text="Step# 3 --> ").grid(row=2, column=0)
                Button(frame, text="Select File", command=lambda: extractrange()).grid(row=2, column=1, padx=5, pady=5)
                try:
                    self.ExtractAllFrame.destroy()
                    self.MergeFrame.destroy()
                except AttributeError:
                    pass

            elif selection.get() == "merge":
                self.MergeFrame = LabelFrame(self.master, text="Merge PDF", width=250, height=150)
                self.MergeFrame.place(x=125, y=10)
                frame = self.MergeFrame
                Label(frame, text="NOTE:- The selected files will be sorted in an ascending oder before merging,\nso please name them in the order you want them merged.", fg='red').grid(row=1, column=0)
                def mergeall():
                    try:
                        # Fetch the folderpath
                        filepath_tuple = filedialog.askopenfilenames(initialdir="/", title="Select files", filetypes=[("PDF files", "*.pdf")])
                        filepath = filepath_tuple[0].replace('/', '\\')
                        p = Path(filepath)
                        outputdir = str(p.parent)
                        Label(frame, text=f"Output Directory:- {outputdir}").grid(row=6, column=0, sticky='w')
                        # Merge all the PDF files in the directory
                        output = PdfMerger()
                        for i, filepath in enumerate(filepath_tuple):
                            newpath = filepath_tuple[i].replace('/', '\\')
                            file = PdfReader(newpath)
                            output.append(file)
                        merged_filepath = fr"{outputdir}\Merged.pdf"
                        with open(merged_filepath, "wb") as output_stream:
                            output.write(output_stream)
                        Label(frame, text=f"Merged File:- {merged_filepath}", fg="green").grid(row=7, column=0, sticky='w')
                        messagebox.showinfo(title="Success", message="Task completed successfully.")
                    except Exception as e:
                        messagebox.showerror(title="ERROR", message=fr"Error:- {e}")
                Button(frame, text="Select Files", command=lambda: mergeall()).grid(row=2, column=0, padx=5, pady=5)
                try:
                    self.ExtractAllFrame.destroy()
                    self.ExtractRangeFrame.destroy()
                except AttributeError:
                    pass


# root.mainloop()
if __name__ == '__main__':
    main()