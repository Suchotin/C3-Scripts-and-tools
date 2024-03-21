import os
import re
import shutil


def extract_digits(text):
    return re.findall(r'\b(\d{6})(?:\.\b|\b)', text)  # Match sequences of 4 to 6 digits that may or may not be followed by a separate point


def find_files_with_numbers(root_folder, numbers):
    combined_paths = {}

    for root, dirs, files in os.walk(root_folder):
        if os.path.relpath(root, root_folder).count(os.sep) == 1:
            for file in files:
                if file.lower().endswith('.xlsx') and file.lower().startswith('offer'):
                    file_path = os.path.join(root, file)
                    digits = extract_digits(file)
                    for number in numbers:
                        if str(number) in digits:
                            if number not in combined_paths:
                                combined_paths[number] = [(file_path, file)]
                            else:
                                combined_paths[number].append((file_path, file))

    return combined_paths


def copy_files_to_target_folder(file_paths, target_folder):
    exclude_paths = {number: paths for number, paths in file_paths.items() if len(paths) > 1}
    for number, paths in exclude_paths.items():
        print(f"Number {number}:")
        for file_path, file_name in paths:
            print(f"  {file_name}")
    for number, paths in file_paths.items():
        for file_path, file_name in paths:
            target_path = os.path.join(target_folder, file_name)
            shutil.copy(file_path, target_path)


def main():

    # Define the txt file
    txt_filepath = r'C:\Users\osukhotin\Desktop\ВСЕ КП ЗА 2023.txt'

    # Read numbers from a text file
    with open(txt_filepath) as f:
        numbers = [int(line.strip()) for line in f.readlines()]

    # Define the root folder
    root_folder = r'C:\Users\osukhotin\YandexDisk\C3 Presale\=КП'

    # Find files with numbers in their names
    combined_files = find_files_with_numbers(root_folder, numbers)

    # Define the target folder
    target_folder = r'C:\Users\osukhotin\Desktop\АРХИВ 2023\15.02.24'

    # Copy files to the target folder
    copy_files_to_target_folder(combined_files, target_folder)

    lost_numbers = [number for number in numbers if number not in combined_files]
    print(f"Numbers without matching files: {lost_numbers}")


if __name__ == "__main__":
    main()
