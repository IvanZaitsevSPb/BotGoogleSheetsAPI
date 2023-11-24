import os

folder_path = os.path.dirname(__file__) + '\\'


def get_file_names(folder_path):
    file_names = []
    for file_name in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path , file_name)) and file_name.endswith(".xlsx"):
            file_names.append(file_name)
    return file_names

