import json
import os
import re
import tkinter as tk
from difflib import SequenceMatcher
from tkinter import filedialog, ttk

from folder_create import create_invalid_folder
from lambda_function import run_lambda
from marksheet_excel import create_excel

subjects = [
    "006 ENGLISH",
    "006 ENGLISH (F. L. )",
    "050 MATHEMATICS",
    "052 CHEMISTRY",
    "053 CHEMISTRY PRACT",
    "054 PHYSICS",
    "055 PHYSICS PRACT",
    "331 COMPUTER",
    "332 COMPUTER PRACT",
    "SANSKRIT",
]

word_marks = {
    "ZERO": 0,
    "ONE": 1,
    "TWO": 2,
    "THREE": 3,
    "FOUR": 4,
    "FIVE": 5,
    "SIX": 6,
    "SEVEN": 7,
    "EIGHT": 8,
    "NINE": 9,
    "ELEVEN": 11,
    "TWELVE": 12,
    "THIRTEEN": 13,
    "FOURTEEN": 14,
    "FIFTEEN": 15,
    "SIXTEEN": 16,
    "SEVENTEEN": 17,
    "EIGHTEEN": 18,
    "NINETEEN": 19,
    "TWENTY": 20,
    "THIRTY": 30,
    "FORTY": 40,
    "FIFTY": 50,
    "SIXTY": 60,
    "SEVENTY": 70,
    "EIGHTY": 80,
    "NINETY": 90,
    "HUNDRED": 100,
}

total = "THEORY TOTAL ON WHICH GRADE IS CALCULATED"
total_marks = []
seat_no = []


def similar(a, b):
    return SequenceMatcher(None, a.upper(), b.upper()).ratio()


def expected_word(test, only_sub):
    sim_idx = []
    threshold = 0.6
    for i in only_sub:
        sim_idx.append(similar(test, i))
    temp = max(sim_idx)
    if temp >= threshold:
        max_idx = sim_idx.index(temp)
        return only_sub[max_idx]
    else:
        return False


def expected_marks(test, word_marks):
    sim_idx = []
    threshold = 0.5
    for i in word_marks.keys():
        sim_idx.append(similar(test, i))
    # print('sim_idx',sim_idx)
    temp = max(sim_idx)
    # print('temp',temp)
    if temp >= threshold:
        max_idx = sim_idx.index(temp)
        # print('max_idx',max_idx)
        key = list(word_marks)[max_idx]
        return word_marks[key]
    else:
        return False


def word_to_marks(s):
    token = list(s.split())
    # print('token',token)
    num = []
    for i in token:
        e = expected_marks(i, word_marks)
        # print('e',e)
        if e or str(e).isnumeric():
            num.append(e)
    # print('num',num)
    if len(num) == 3:
        return int("".join(map(str, num)))
    # print('num',num)


def detect_total(s):
    token = list(s.split())
    flag = False
    for i in token:
        if similar(i, "ONLY") > 0.7:
            flag = True
            break
    if flag and len(token) > 4:
        return True
    else:
        return False


def total_word_to_marks(s):
    token = list(s.split())
    # print('token',token)
    num = []
    for i in token:
        e = expected_marks(i, word_marks)
        # print('e',e)
        if e:
            num.append(e)
    total_marks = num[0] * num[1] + num[2] + num[3]
    return total_marks


def is_word(str):
    lst = list(str)
    threshold = 0.7
    res = []
    for i in lst:
        if i.isalpha():
            res.append(True)
        else:
            res.append(False)
    # print('res',res)
    truth_value = 0
    if len(res) != 0:
        truth_value = res.count(True) / len(res)
    if truth_value >= threshold:
        return True
    else:
        return False


def is_num(str, threshold):
    lst = list(str)
    res = []
    for i in lst:
        if i.isnumeric():
            res.append(True)
        else:
            res.append(False)
    # print('res',res)
    truth_value = 0
    if len(res) != 0:
        truth_value = res.count(True) / len(res)
    if truth_value >= threshold:
        return True
    else:
        return False


def get_raw_data(response):
    # print("\nText\n========")
    text = ""
    for item in response["Blocks"]:
        if item["BlockType"] == "LINE":
            # print ('\033[94m' + item["Text"] + '\033[0m')
            text = text + "\n" + item["Text"]

    # print(text)
    lines = re.split(r"\n", text)
    return lines


def clean_data(lines):
    sub_re = re.compile(r"^\d\d\d\s\D\D")
    seatno_re = re.compile(r"^\D\s\d\d\d\d\d\d$")
    # print('----------result-----------')
    cleaned_data = []
    i = 0
    while i < len(lines):
        if sub_re.search(lines[i]):
            # print(lines[i])
            cleaned_data.append(lines[i])
            # marks
            i += 1
            j = 0
            turn = 1
            while j < turn:
                if is_num(lines[i], 0.5):
                    # print('j is_num')
                    # print(lines[i])
                    cleaned_data.append(lines[i])
                    i += 1
                    if is_num(lines[i], 0.5):
                        # print('k is_num')
                        # print(lines[i])
                        cleaned_data.append(lines[i])
                        i += 1
                    # k = 0
                    # while k < 3:
                    #     if is_num(lines[i], 0.5):
                    #         print('k is_num')
                    #         print(lines[i])
                    #         cleaned_data.append(lines[i])
                    #         i += 1
                    #         break
                    #     k += 1
                    # break
                j += 1

            # print('before l word',lines[i])
            l = 0
            while l < 3:
                if is_word(lines[i]):
                    lst = list(lines[i].split())
                    if len(lst) == 3:
                        # print('l isword')
                        # print(lines[i])
                        cleaned_data.append(lines[i])
                        break
                l += 1
                i += 1
        elif similar(lines[i], total) > 0.7:
            total_marks.append(lines[i])
        elif lines[i] == "650":
            for u in range(i - 3, i + 4):
                total_marks.append(lines[u])

        elif seatno_re.search(lines[i]):
            seat_no.append(lines[i])

        i += 1

    # print('cleaned data', cleaned_data)
    return cleaned_data


def get_final_clean_data(cleaned_data):
    # print('-------------final----------------------')
    final = []
    i = 0
    while i < len(cleaned_data):
        exp = expected_word(cleaned_data[i], subjects)
        if exp:
            # print('exp', exp)
            final.append(exp)
            i += 1
            # print('before isnum', cleaned_data[i])
            if is_num(cleaned_data[i + 1], 0.7):
                # print('inside if isnum', cleaned_data[i])
                num_marks = cleaned_data[i + 1]
                final.append(int(num_marks))
                i += 2
            else:
                # print('inside else isnum', cleaned_data[i])
                k = 0
                p = i
                while k < 2:
                    if is_num(cleaned_data[p], 0.7):
                        l = list(cleaned_data[p].split())
                        if len(l) >= 2:
                            final.append(int(l[-1]))
                            p += 1
                            i += 1
                            break
                    p += 1
                    k += 1

            b = 0
            while b < 4:
                # print('before isword', cleaned_data[i])
                if is_word(cleaned_data[i]):
                    wordmarks = word_to_marks(cleaned_data[i])
                    final.append(wordmarks)
                    i += 1
                    break
                i += 1
                b += 1

            # print('before j=0',cleaned_data[i])
            # j = 0
            # while j < 3:
            #     # print('inside while j',j,cleaned_data[i])
            #     if is_word(cleaned_data[i]):
            #         # print('inside isword ',cleaned_data[i])
            #         wordmarks = word_to_marks(cleaned_data[i])
            #         final.append(wordmarks)
            #         i += 1
            #         break
            #     j += 1
            #     i += 1
        else:
            i += 1

    # print('final', final)
    return final


def get_result(final):
    result = []
    validate = True
    m = 0
    while m < len(final):
        if isinstance(final[m], str):
            result.append(final[m])
            m += 1
        elif isinstance(final[m], int) and isinstance(final[m + 1], int):
            if final[m] == final[m + 1]:
                result.append(final[m])
            else:
                validate = False
            m += 2
        else:
            result.append(final[m])
            m += 1

    # print('result', result)
    return result, validate


def get_total_marks(total_marks, result):
    # print('------------------------------------------------------')
    # print(total_marks)
    validate = True
    detected_marks = None
    calculated_marks = 0
    for i in total_marks:
        if i.isdigit() and i != "650":
            detected_marks = i
    # print('detected_marls', detected_marks)
    for i in result:
        if isinstance(i, int):
            calculated_marks += i

    # print('calculated_marks', calculated_marks)

    if int(detected_marks) == calculated_marks:
        result.append("Total")
        result.append(calculated_marks)
    else:
        validate = False
    return result, validate


def get_seat_no():
    return seat_no


def get_json(marks, seat, val_result, val_total, filename):
    sample = [
        "english",
        0,
        "maths",
        0,
        "chem",
        0,
        "chemP",
        0,
        "phy",
        0,
        "phyP",
        0,
        "comp",
        0,
        "compP",
        0,
        "total",
        0,
    ]
    # marks = ['006 ENGLISH (F. L. )', 61, '050 MATHEMATICS', 80, '052 CHEMISTRY', 69, '053 CHEMISTRY PRACT', 44,
    #          '054 PHYSICS', 62, '055 PHYSICS PRACT', 45, '331 COMPUTER', 53, '332 COMPUTER PRACT', 44, 'Total', 458]
    #
    # seat = ['B 106346']
    temp = {}
    for i in range(0, len(marks), 2):
        temp[sample[i]] = marks[i + 1]
    if val_total and val_result:
        temp["validate"] = True
    else:
        temp["validate"] = False
        temp["filename"] = filename
    # print(temp)
    d = {seat[-1]: temp}

    # print(d)
    return d


def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        print(f"Selected folder: {folder_path}")
        folder_path_label.config(text=f"Selected Folder: {folder_path}")
        upload_folder(folder_path)


def upload_folder(folder_path):
    loading_screen = tk.Toplevel(root)
    loading_screen.title("Loading")
    loading_screen.geometry("200x100")
    loading_label = tk.Label(
        loading_screen, text="Processing folder...", font=("Arial", 12)
    )
    loading_label.pack(pady=10)
    root.update()
    file_names = [
        f
        for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f))
    ]

    excel_dict = {}
    invalid_file_list = []

    for file in file_names:
        response = run_lambda(folder_path + "/" + file)
        lines = get_raw_data(response)
        try:
            cleaned_data = clean_data(lines)
            final_clean_data = get_final_clean_data(cleaned_data)
            result, val_result = get_result(final_clean_data)
            result_updated, val_total = get_total_marks(total_marks, result)
            marks_json = get_json(
                result_updated, get_seat_no(), val_result, val_total, file
            )

            if not marks_json[get_seat_no()[-1]]["validate"]:
                invalid_file_list.append(file)

            excel_dict.update(marks_json)
        except:
            invalid_file_list.append(file)
            marks_json = {f"error_{file}": {"validate": False, "filename": file}}
            excel_dict.update(marks_json)

    if checked_invalid.get():
        create_invalid_folder(file_list=invalid_file_list)

    create_excel(excel_dict)
    # Destroy loading screen
    loading_screen.destroy()
    success_label = tk.Label(
        root, text="Excel sheet created succesfully", font=("Arial", 12)
    )
    success_label.configure(
        bg="#F8F8FF", fg="#4B0082"
    )  # Set background and foreground color
    success_label.pack(pady=10)


# Create main window
root = tk.Tk()
root.title("Upload Folder")
root.geometry("600x400")  # Set window size
root.configure(bg="#F8F8FF")  # Set background color

checkbox_frame = tk.Frame(root)
checkbox_frame.pack()

# Create a boolean variable to hold the state of the checkbox
checked_invalid = tk.BooleanVar()

# Create the checkbox
checkbox = tk.Checkbutton(
    checkbox_frame, text="Create folder for invalid marksheet", variable=checked_invalid
)
checkbox.pack(side=tk.LEFT)

checked_preprocess = tk.BooleanVar()

# Create the checkbox
checkbox = tk.Checkbutton(
    checkbox_frame, text="Preprocess", variable=checked_preprocess
)
checkbox.pack(side=tk.LEFT)

# Create label for displaying selected folder path
folder_path_label = tk.Label(root, text="No folder selected", font=("Arial", 12))
folder_path_label.configure(
    bg="#F8F8FF", fg="#4B0082"
)  # Set background and foreground color
folder_path_label.pack(pady=10)

# Create browse button
browse_button = tk.Button(
    root,
    text="Start",
    command=browse_folder,
    bg="#6495ED",
    fg="white",
    font=("Arial", 12),
)
browse_button.pack(pady=10)

# Run the application
root.mainloop()
