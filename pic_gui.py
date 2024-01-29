import cv2
import os, sys
import numpy as np
import tkinter as tk
from tkinter import filedialog
from pic2file import detect_and_transform
from PIL import Image, ImageTk
from tkinter import font
from user2pic import select_four_points, transform_perspective
from docx import Document
import re
from docx.shared import Cm
image_path2 = None
save_pic = None


def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]


def file2doc():
    input_files = entry_input_folder_file2doc.get()
    print(input_files)
    output_doc_name = entry_output_docname.get()
    # Get the folder name from the user
    folder_name = input_files

    # Get the list of image files in the folder
    image_files = [file for file in os.listdir(folder_name) if file.lower().endswith(('.jpg', '.jpeg', '.png'))]

    # Sort the image files using natural sorting
    sorted_image_files = sorted(image_files, key=natural_sort_key)

    # Create a new Word document
    doc = Document()

    # Set the margins of the document (1cm top and bottom, 2cm left and right)
    section = doc.sections[0]  # Assuming there is only one section in the document
    section.page_width = Cm(21)  # A4 width
    section.page_height = Cm(29.7)  # A4 height

    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # Iterate through the sorted image files and add them to the document
    for image_file in sorted_image_files:

        # Get the full path of the image file
        image_path = os.path.join(folder_name, image_file)

        # Add a new paragraph and insert the image with specified width and height in centimeters
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        picture = run.add_picture(image_path, width=Cm(15.29), height=Cm(4.75))

        # Set line spacing to 0
        paragraph_format = paragraph.paragraph_format
        paragraph_format.line_spacing = 1
        # Set space after to 0
        paragraph_format.space_after = Cm(0)

        # Save the Word document
    doc.save(f'{output_doc_name}.doc')
    print('save!')


def browse_input_folder():
    folder_path = filedialog.askdirectory()
    entry_input_folder.delete(0, tk.END)
    entry_input_folder.insert(0, folder_path)

def browse_input_4point():
    # Use a StringVar to store the selected image filename
    selected_image_filename = tk.StringVar()
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
    if file_path:
        selected_image_filename.set(os.path.basename(file_path))
    entry_input_folder_file2doc.delete(0, tk.END)
    entry_input_folder_file2doc.insert(0, selected_image_filename)

def browse_input_folder_doc():
    folder_path = filedialog.askdirectory()
    entry_input_folder_file2doc.delete(0, tk.END)
    entry_input_folder_file2doc.insert(0, folder_path)

def start_processing():
    image_files = entry_input_folder.get()
    output_folder = entry_output_folder.get()
    print(image_files)
    # clean box
    result_text.config(state=tk.NORMAL)
    result_text.delete(1.0, tk.END)

    print_message = detect_and_transform(image_files, output_folder)
    print(print_message)
    if not print_message:
        result_text.insert(tk.END, "圖片處理出錯")  # error message
    for message in print_message:
        result_text.insert(tk.END, message + "\n")

    result_text.config(state=tk.DISABLED, font=font.Font(size=14))

    # update pic
    update_image_list(output_folder)


def update_image_list(output_folder):
    # clean pic column
    listbox_images.delete(0, tk.END)

    image_files = [f for f in os.listdir(output_folder) if f.endswith(('.png', '.jpg', '.jpeg'))]

    for image_file in image_files:
        listbox_images.insert(tk.END, image_file)


def show_selected_image(event):
    selected_index = listbox_images.curselection()
    if selected_index:
        selected_file = listbox_images.get(selected_index)
        image_path = os.path.join(entry_output_folder.get(), selected_file)
        show_image(image_path)

def show_image(image_path):
    image = Image.open(image_path)
    image.thumbnail((600, 200))
    photo = ImageTk.PhotoImage(image)

    # update Label
    image_label.config(image=photo)
    image_label.image = photo


def browse_image():
    global image_path2
    image_path2 = filedialog.askopenfilename(filetypes=[("JPEG files", "*.jpg"), ("All files", "*.*")])
    if image_path2:
        selected_image_filename.set(os.path.basename(image_path2))

    if image_path2:
        image = cv2.imread(image_path2)
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(image)
        img.thumbnail((600, 200))
        img = ImageTk.PhotoImage(img)
        panel.configure(image=img)
        panel.image = img
    else:
        print("請選擇要處理的圖片")  # select pic


def process_image():
    global image_path2, save_pic
    if image_path2:
        image = cv2.imread(image_path2)
        selected_points = select_four_points(image)
        target_size = (600, 200)

        transformed_image = transform_perspective(image, selected_points, target_size)
        transformed_image = cv2.cvtColor(transformed_image, cv2.COLOR_BGR2RGB)
        save_pic = transformed_image

        img = Image.fromarray(transformed_image)
        img.thumbnail((600, 200))
        img = ImageTk.PhotoImage(img)

        panel.configure(image=img)
        panel.image = img
    else:
        print("請選擇要處理的圖片")  # select pic


def save_image():
    global save_pic, image_path2
    print(type(save_pic))
    if save_pic is not None: 
        outp = entry_output_folder.get()
        st_path=[]
        for i in range(len(image_path2)-1,0, -1):
            if image_path2[i]=='/':
                st_path = image_path2[:i+1]+'transformed_'+image_path2[i+1:]
                break
        print(st_path)
        output_path = os.path.join(outp, st_path)
        cv2.imwrite(output_path, cv2.cvtColor(save_pic, cv2.COLOR_BGR2RGB))
        print("確認儲存...")
        print(type(save_pic))
    else:
        print('no pic')


root = tk.Tk()

label_yipin= tk.Label(root, text="Designed by Yipin Peng")
label_input_folder = tk.Label(root, text="輸入圖片的資料夾位置:")  # file address
entry_input_folder = tk.Entry(root, width=40)
button_browse_input = tk.Button(root, text="瀏覽", command=browse_input_folder)  # browser

label_output_folder = tk.Label(root, text="輸出資料夾的名稱:")  # file name for adjusted pic
entry_output_folder = tk.Entry(root, width=40)
# button_browse_output = tk.Button(root, text="瀏覽", command=browse_output_folder)

button_start = tk.Button(root, text="開始處理", command=start_processing)  # process
# file2doc
label_input_folder＿file2doc = tk.Label(root, text="輸入轉正圖片的資料夾位置:")
entry_input_folder_file2doc = tk.Entry(root, width=40)
button_browse_input_file2doc = tk.Button(root, text="瀏覽", command=browse_input_folder_doc)

# final_doc_name
label_doc_name = tk.Label(root, text="輸出word檔名稱:")  # name for word file
entry_output_docname = tk.Entry(root, width=40)
button_file_org = tk.Button(root, text="彙整至word", command=file2doc)  # organize in word

result_text = tk.Text(root, height=15, width=50)
result_text.insert(tk.END, "\n\n\n處理結果將顯示在這裡\n")  # the results will be showed here
result_text.config(state=tk.DISABLED, font = font.Font(size=14))  # 讓 Text 元件無法編輯

listbox_images = tk.Listbox(root, selectmode=tk.SINGLE, height=15, width=25)
listbox_images.bind("<Double-Button-1>", show_selected_image)

image_label = tk.Label(root)
image_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# place
label_yipin.grid(row=7, column=0, padx=5, pady=10, sticky=tk.N)
label_input_folder.grid(row=0, column=0, padx=5, pady=10, sticky=tk.E)
entry_input_folder.grid(row=0, column=1, padx=5, pady=10, sticky=tk.W)
button_browse_input.grid(row=0, column=2, padx=5, pady=10)

label_output_folder.grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
entry_output_folder.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)

label_input_folder＿file2doc.grid(row=5, column=3, padx=5, pady=10, sticky=tk.E)
entry_input_folder_file2doc.grid(row=5, column=4, padx=5, pady=10, sticky=tk.W)
button_browse_input_file2doc.grid(row=5, column=5, padx=5, pady=10)

label_doc_name.grid(row=6, column=3, padx=5, pady=10, sticky=tk.E)
entry_output_docname.grid(row=6, column=4, padx=5, pady=10, sticky=tk.W)
button_file_org.grid(row=6, column=5, padx=5, pady=10)

button_start.grid(row=1, column=2, pady=20)
# place Text
result_text.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
# place Listbox
listbox_images.grid(row=3, column=1, columnspan=3, padx=10, pady=10)
# users select 4point
label2 = tk.Label(root, text="選擇要處理的圖片:")  # select pic

label2.grid(row=0, column=2, columnspan=2, padx=5, pady=5, sticky=tk.E)


selected_image_filename = tk.StringVar()
entry_input_folder_4point = tk.Entry(root, textvariable=selected_image_filename, width=30)
entry_input_folder_4point.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)

browse_button = tk.Button(root, text="瀏覽", command=browse_image)  # browser
browse_button.grid(row=0, column=5, columnspan=3, padx=10, pady=10)
panel = tk.Label(root)
panel.grid(row=3, column=3, columnspan=3, padx=50, pady=10)

process_button = tk.Button(root, text="處理圖片", command=process_image)  # process pic
process_button.grid(row=4, column=3, columnspan=3, padx=0, pady=10)

button_save = tk.Button(root, text="儲存圖片", state=tk.NORMAL, command=save_image)  # save pic
button_save.grid(row=4, column=4, columnspan=3, padx=10, pady=10)

root.mainloop()
