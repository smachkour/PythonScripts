import pytesseract
import tkinter as tk
from PIL import ImageGrab, Image
import clipboard
import time
import sys
import cv2
import numpy as np
from pytesseract import Output

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image(img):
    # Convert the image to gray scale
    img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Performing OTSU threshold
    ret, thresh1 = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)

    # Specify structure shape and kernel size. 
    rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 18))

    # Appplying dilation on the threshold image
    dilation = cv2.dilate(thresh1, rect_kernel, iterations = 1)

    # Finding contours
    contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    # Creating a copy of image
    img_copy = img.copy()

    return Image.fromarray(img_copy)

class App():
    def __init__(self, root):
        self.is_running = True
        self.root = root
        self.root.title("Image to Text Converter")
        self.root.geometry('300x200')

        self.start_button = tk.Button(root, text='Start', command=self.start)
        self.start_button.pack()

        self.stop_button = tk.Button(root, text='Stop', command=self.stop, state=tk.DISABLED)
        self.stop_button.pack()

        self.console = tk.Text(root)
        self.console.pack()

        sys.stdout = TextRedirector(self.console)

        self.check_clipboard()

    def start(self):
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        print('Running...')

    def stop(self):
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        print('Stopped')

    def check_clipboard(self):
        if self.is_running:
            try:
                img = ImageGrab.grabclipboard()
                if img is not None:
                    try:
                        img = preprocess_image(img)
                        result = pytesseract.image_to_string(img, lang='eng', config='--oem 3')
                        clipboard.copy(result)
                        print('Image to text complete, you can now paste.')
                    except pytesseract.TesseractError as e:
                        print("Tesseract Error:", e)
                    except Exception as e:
                        print("General Error:", e)
            except Exception as e:
                print("Error grabbing clipboard:", e)

        self.root.after(1000, self.check_clipboard)


class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, str):
        self.widget.insert(tk.END, str)
        self.widget.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
