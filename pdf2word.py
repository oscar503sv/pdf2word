import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.visible = 0

pdfs_path = "input_path" # path where .pdf files are stored
for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")):
    print("***************************************************************")
    print("documento encontrado: "+doc)
    filename = doc.split('\\')[-1]
    in_file = os.path.abspath(doc)
    print("convirtiendo documento a .docx...")
    wb = word.Documents.Open(in_file)
    out_file = os.path.abspath("output_path"+filename[0:-4]+ ".docx".format(i))
    print("Documento de salida:\n",out_file)
    wb.SaveAs2(out_file, FileFormat=16) 
    print("Exitoso.")
    wb.Close()
word.Quit()