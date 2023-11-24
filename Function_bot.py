import os


def delete_xlsx_files():
    root_folder = os.getcwd()  # Getting the current working folder

    for file in os.listdir(root_folder):  # Sorting through all the files in the folder
        if file.endswith(".xlsx"):  # If the file has an extension .xlsx
            file_path = os.path.join(root_folder, file)  # Getting the full path to the file
            os.remove(file_path)  # Deleting the file
