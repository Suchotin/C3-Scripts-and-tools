import os
from datetime import datetime, timedelta


def find_excels(root_folder, days_threshold = 1):
    file_paths = []

    for root, dirs, files in os.walk(root_folder):
        # Check if the current directory is a second-level subdirectory
        if os.path.relpath(root, root_folder).count(os.sep) == 1:
            for file in files:
                file_path = os.path.join(root, file)
                modification_time = datetime.fromtimestamp(os.path.getmtime(file_path))

                # Check if the file has a .xlsx extension and was modified within the given period
                if file.lower().endswith('.xlsx') and file.lower().startswith('offer') and datetime.now() - modification_time <= timedelta(
                        days=days_threshold):
                    file_paths.append(file_path)

    return file_paths


# Example usage:
root_folder_path = r'C:\Users\osukhotin\YandexDisk\C3 Presale\=КП'
result_paths = find_excels(root_folder_path)
print("File Paths:")
print(result_paths)
print(len(result_paths))
