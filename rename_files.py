from docx import Document
import os

mydir = 'C:\\Users\\User\\Desktop\\folder'

for file in os.listdir(mydir):
    document = Document(mydir + "\\" + file)
    name = document.paragraphs[2].text

    old_file = os.path.join(mydir, file)
    new_file = os.path.join(mydir, f"{name}.docx")
    os.rename(old_file, new_file)
