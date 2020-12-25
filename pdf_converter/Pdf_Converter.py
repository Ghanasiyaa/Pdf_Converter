from tkinter import *
import sys
import os
import win32com.client
from win32com import client
from appJar import gui
from pathlib import Path
from PIL import *
root = Tk()
root.title("WPE >> PDF")
root.iconbitmap('C:/Users/Ghanasiyaa/OneDrive/Pictures/pdflogo.ico')

def openWord():

    wdFormatPDF = 17

    def Word_to_pdf():
        def word_to_pdf(input_file, output_file):
            in_file = input_file
            out_file = str(output_file) + '.pdf'

            word = win32com.client.Dispatch('Word.Application')
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
            print(".DOCX to PDF conversion sucessful and Saved")
            if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
                app.stop()

        def validate_inputs(src_file, dest_dir, out_file):

            errors = False
            error_msgs = []
            if Path(src_file).suffix.upper() != ".DOCX":
                errors = True
                error_msgs.append("Please select a .DOC or .DOCX   input file")

            if not (Path(dest_dir)).exists():
                errors = True
                error_msgs.append("Please Select a valid output directory")

            # Check for a file name
            if len(out_file) < 1:
                errors = True
                error_msgs.append("Please enter a file name")

            return (errors, error_msgs)

        def press(button):
            if button == "Process":
                src_file = app.getEntry("Input_File")
                dest_dir = app.getEntry("Output_Directory")
                out_file = app.getEntry("Output_name")
                errors, error_msg = validate_inputs(src_file, dest_dir, out_file)
                if errors:
                    app.errorBox("Error", "\n".join(error_msg), parent=None)
                else:
                    word_to_pdf(src_file, Path(dest_dir, out_file))
            else:
                app.stop()

        app = gui("WORD :>) PDF Converter", useTtk=True)
        app.setTtkTheme('clam')
        app.setSize(500, 200)

        # Add the interactive components
        app.addLabel("Choose Source Word File to convert ")
        app.addFileEntry("Input_File")

        app.addLabel("Select Output Directory")
        app.addDirectoryEntry("Output_Directory")

        app.addLabel("Output file name")
        app.addEntry("Output_name")

        app.addButtons(["Process", "Quit"], press)
        app.go()

    Word_to_pdf()


def openPpt():

    def Powerpoint_to_pdf():

        def ppt_to_pdf(input_file, output_file, formatType=32):
            in_file = input_file
            out_file = str(output_file)  # desktop\file.pptx
            out_file += ".pdf"
            powerpoint = win32com.client.Dispatch("Powerpoint.Application")
            pdf = powerpoint.Presentations.Open(in_file, WithWindow=False)
            pdf.SaveAs(out_file, 32)
            pdf.Close()
            powerpoint.Quit()
            print("PPTX to PDF conversion sucessful and Saved")
            if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
                app.stop()

        def validate_inputs(src_file, dest_dir, out_file):

            errors = False
            error_msgs = []
            if Path(src_file).suffix.upper() != ".PPTX":
                errors = True
                error_msgs.append("Please select a .PPTX input file")

            if not (Path(dest_dir)).exists():
                errors = True
                error_msgs.append("Please Select a valid output directory")

            # Check for a file name
            if len(out_file) < 1:
                errors = True
                error_msgs.append("Please enter a file name")

            return (errors, error_msgs)

        def press(button):
            if button == "Process":
                src_file = app.getEntry("Input_File")
                dest_dir = app.getEntry("Output_Directory")
                out_file = app.getEntry("Output_name")
                errors, error_msg = validate_inputs(src_file, dest_dir, out_file)
                if errors:
                    app.errorBox("Error", "\n".join(error_msg), parent=None)
                else:
                    ppt_to_pdf(src_file, Path(dest_dir, out_file))
            else:
                app.stop()

        app = gui("POWERPOINT :>) PDF Converter", useTtk=True)
        app.setTtkTheme('clam')
        app.setSize(500, 200)

        # Add the interactive components
        app.addLabel("Choose Source Powerpoint project File to convert ")
        app.addFileEntry("Input_File")

        app.addLabel("Select Output Directory")
        app.addDirectoryEntry("Output_Directory")

        app.addLabel("Output file name")
        app.addEntry("Output_name")

        app.addButtons(["Process", "Quit"], press)
        app.go()

    Powerpoint_to_pdf()

def openExcel():


    def Excel_to_pdf():
        def excel_to_pdf(input_file, output_file):
            xlApp = client.Dispatch("Excel.Application")
            books = xlApp.Workbooks.Open(input_file)
            ws = books.Worksheets[0]
            ws.Visible = 1
            out_file = str(output_file)  # desktop\file.pptx
            out_file += ".pdf"
            ws.ExportAsFixedFormat(0, out_file)
            print("PPTX to PDF conversion sucessful and Saved")
            if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
                app.stop()

        def validate_inputs(src_file, dest_dir, out_file):

            errors = False
            error_msgs = []
            if Path(src_file).suffix.upper() != ".XLSX":
                errors = True
                error_msgs.append("Please select a .PPTX input file")

            if not (Path(dest_dir)).exists():
                errors = True
                error_msgs.append("Please Select a valid output directory")

            # Check for a file name
            if len(out_file) < 1:
                errors = True
                error_msgs.append("Please enter a file name")

            return (errors, error_msgs)

        def press(button):
            if button == "Process":
                src_file = app.getEntry("Input_File")
                dest_dir = app.getEntry("Output_Directory")
                out_file = app.getEntry("Output_name")
                errors, error_msg = validate_inputs(src_file, dest_dir, out_file)
                if errors:
                    app.errorBox("Error", "\n".join(error_msg), parent=None)
                else:
                    excel_to_pdf(src_file, Path(dest_dir, out_file))
            else:
                app.stop()

        app = gui("EXCEL :>) PDF Converter", useTtk=True)
        app.setTtkTheme('clam')
        app.setSize(500, 200)

        # Add the interactive components
        app.addLabel("Choose Source Excel File to convert ")
        app.addFileEntry("Input_File")

        app.addLabel("Select Output Directory")
        app.addDirectoryEntry("Output_Directory")

        app.addLabel("Output file name")
        app.addEntry("Output_name")

        app.addButtons(["Process", "Quit"], press)
        app.go()

    Excel_to_pdf()

label1=Label(root,text="choose a word  file :").grid(row=0,column=0,padx=5,pady=5)
label2=Label(root,text="choose a  ppt  file :").grid(row=1,column=0,padx=5,pady=5)
label3=Label(root,text="choose a excel file :").grid(row=2,column=0,padx=5,pady=5)

button1 = Button(root,text="word  >> pdf",fg="blue",command=openWord).grid(row=0,column=1,padx=5,pady=5)
button2 = Button(root,text=" ppt  >> pdf",fg="red",command=openPpt).grid(row=1,column=1,padx=5,pady=5)
button3 = Button(root,text="excel >> pdf",fg="green",command=openExcel).grid(row=2,column=1,padx=5,pady=5)

root.geometry('300x200')
root.resizable(width=False, height=False)
root.mainloop()
