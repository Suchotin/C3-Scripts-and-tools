import os
import re


def extract_digits(text):
    return re.findall(r'\b(\d{6})(?:\.\b|\b)',
                      text)  # Match sequences of 4 to 6 digits that may or may not be followed by a separate point


def find_numbers_with_no_files(numbers, target_folder):
    target_files = set()
    for root, dirs, files in os.walk(target_folder):
        for file in files:
            digits = extract_digits(file)
            for number in numbers:
                if str(number) in digits:
                    target_files.add(number)

    return [number for number in numbers if number not in target_files]


def main():

    # Define the txt file
    txt_filepath = r'C:\Users\osukhotin\Desktop\ВСЕ КП ЗА 2023.txt'

    # Read numbers from a text file
    with open(txt_filepath) as f:
        numbers = [int(line.strip()) for line in f.readlines()]

    # Define the target folder
    target_folder = r'C:\Users\osukhotin\Desktop\АРХИВ 2023\15.02.24'

    # Ensure target folder exists
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)


    # Find numbers with no corresponding files in the target folder
    numbers_with_no_files = find_numbers_with_no_files(numbers, target_folder)
    if numbers_with_no_files:
        print("Numbers with no corresponding files in the target folder:")
        for number in numbers_with_no_files:
            print(number)


if __name__ == "__main__":
    main()
