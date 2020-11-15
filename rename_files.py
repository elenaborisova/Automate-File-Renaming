from docx import Document
import os
import tkinter
import tkinter.filedialog
import shutil


def choose_directory():
    root = tkinter.Tk()
    root.withdraw()
    dir_name = tkinter.filedialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
    if len(dir_name) > 0:
        print(f"You chose {dir_name}")

    return dir_name


def create_new_dir(chosen_dir):
    new_dir = os.path.join(chosen_dir, "Renamed Documents")
    if not os.path.exists(new_dir):
        os.mkdir(new_dir)

    return new_dir


def copy_all_files(chosen_dir, new_dir):
    all_files = os.listdir(chosen_dir)

    for file in all_files:
        full_file_name = os.path.join(chosen_dir, file)
        if os.path.isfile(full_file_name):
            shutil.copy(full_file_name, new_dir)


def rename_files_in_dir(new_dir):
    for file in os.listdir(new_dir):
        document = Document(new_dir + "\\" + file)
        name = document.paragraphs[2].text

        old_file_name = os.path.join(new_dir, file)
        new_file_name = os.path.join(new_dir, f"{name}.docx")
        os.rename(old_file_name, new_file_name)


chosen_dir = choose_directory()
new_dir = create_new_dir(chosen_dir)
copy_all_files(chosen_dir, new_dir)
rename_files_in_dir(new_dir)
