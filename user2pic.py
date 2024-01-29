import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk


def select_four_points(image):
    # Registering Mouse Click Events on an Image
    points = []
    image2= image.copy()

    def click_event(event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            points.append([x, y])
            cv2.circle(image2, (x, y), 10, (0, 0, 255), -1)
            cv2.imshow('Select Points', image2)

    cv2.imshow('Select Points', image)
    cv2.setMouseCallback('Select Points', click_event)

    while len(points) < 4:
        cv2.waitKey(1)

    cv2.destroyAllWindows()
    return np.array(points, dtype=np.float32)


def transform_perspective(image, src_points, target_size):
    # define 4 points
    target_points = np.array([[0, 0], [target_size[0], 0], [target_size[0], target_size[1]], [0, target_size[1]]],
                             dtype=np.float32)

    # Calculate the perspective transformation matrix
    M = cv2.getPerspectiveTransform(src_points, target_points)

    # Perspective transformation
    transformed_image = cv2.warpPerspective(image, M, target_size)

    return transformed_image


class ImageProcessingApp:
    def __init__(self, master):
        self.master = master
        master.title("Image Processing GUI")

        self.label = tk.Label(master, text="選擇要處理的圖片:")
        self.label.pack(pady=10)

        self.browse_button = tk.Button(master, text="瀏覽", command=self.browse_image)
        self.browse_button.pack(pady=10)

        self.panel = tk.Label(master)
        self.panel.pack(padx=10, pady=10)

        self.process_button = tk.Button(master, text="處理圖片", command=self.process_image)
        self.process_button.pack(pady=10)

        self.image_path = None

    def browse_image(self):
        self.image_path = filedialog.askopenfilename(filetypes=[("JPEG files", "*.jpg"), ("All files", "*.*")])
        if self.image_path:
            image = cv2.imread(self.image_path)
            image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(image)
            img.thumbnail((600, 150))
            img = ImageTk.PhotoImage(img)
            self.panel.configure(image=img)
            self.panel.image = img  # 保留對圖片的引用
        else:
            print("請選擇要處理的圖片")

    def process_image(self):
        if self.image_path:
            image = cv2.imread(self.image_path)
            selected_points = select_four_points(image)
            target_size = (600, 200)
            transformed_image = transform_perspective(image, selected_points, target_size)

            transformed_image = cv2.cvtColor(transformed_image, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(transformed_image)
            img.thumbnail((600, 200))
            img = ImageTk.PhotoImage(img)
            self.panel.configure(image=img)
            self.panel.image = img  # 保留對圖片的引用
        else:
            print("請選擇要處理的圖片")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageProcessingApp(root)
    root.mainloop()
