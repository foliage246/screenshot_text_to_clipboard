# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import messagebox
from PIL import ImageGrab
import pytesseract
import time
import sys
import os
import webbrowser

# --- Universal Asset Path Finder ---
# This function is crucial for finding assets in both development and packaged modes.
def get_asset_path(file_name):
    """
    Gets the absolute path to an asset file.
    This works for both running as a script and as a packaged --onefile PyInstaller exe.
    """
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # sets the _MEIPASS attribute to the path of the temporary folder.
        base_path = sys._MEIPASS
    else:
        # In a normal environment, the base path is the script's directory.
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, file_name)

# --- Tesseract Path Configuration ---
def get_tesseract_path():
    """Gets the execution path for Tesseract, handling packaged and dev environments."""
    if getattr(sys, 'frozen', False):
        # For the packaged app, find the Tesseract-OCR folder inside the bundle.
        tesseract_folder = get_asset_path("Tesseract-OCR")
        return os.path.join(tesseract_folder, "tesseract.exe")
    else:
        # For development, use the hardcoded path.
        if os.name == 'nt':
            return r'C:\Users\s1yeh\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
        else:
            return 'tesseract' # For macOS/Linux

# Set the command path for pytesseract
tesseract_cmd_path = get_tesseract_path()
if tesseract_cmd_path and os.path.exists(tesseract_cmd_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path
else:
    # If Tesseract is not found, don't set it yet; the error will be caught during execution.
    pass


class SelectionWindow:
    """Creates a semi-transparent top-level window for selecting a screen area."""
    def __init__(self, master_app):
        self.master_app = master_app
        self.master_app.root.withdraw() # Hide the main window
        
        time.sleep(0.2)

        self.selection_window = tk.Toplevel(self.master_app.root)
        self.selection_window.attributes("-fullscreen", True)
        self.selection_window.attributes("-alpha", 0.3)
        self.selection_window.attributes("-topmost", True)
        self.selection_window.overrideredirect(True)

        self.canvas = tk.Canvas(self.selection_window, cursor="cross", bg="grey")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.selection_window.bind("<Escape>", self.cancel_selection)

    def on_button_press(self, event):
        self.start_x = self.canvas.winfo_pointerx()
        self.start_y = self.canvas.winfo_pointery()
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x = self.canvas.winfo_pointerx()
        cur_y = self.canvas.winfo_pointery()
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x = self.canvas.winfo_pointerx()
        end_y = self.canvas.winfo_pointery()
        
        self.selection_window.destroy()
        
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        
        if x2 - x1 > 10 and y2 - y1 > 10: 
            self.master_app.capture_and_ocr((x1, y1, x2, y2))
        else:
            self.master_app.show_main_window("Selection cancelled")

    def cancel_selection(self, event):
        self.selection_window.destroy()
        self.master_app.show_main_window("Selection cancelled")


class OcrApp:
    """The main application class."""
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot to Clipboard")
        
        try:
            icon_path = get_asset_path('sam_tool_screenshot_text_to_clipboard.png')
            if os.path.exists(icon_path):
                photo = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(True, photo)
        except Exception as e:
            print(f"Could not set window icon: {e}")

        self.root.geometry("380x230") # Adjust window size for new components
        self.root.eval('tk::PlaceWindow . center')

        self.label = tk.Label(root, text="Click the button to start capturing text", padx=10, pady=10, font=("Arial", 10))
        self.label.pack(expand=True)

        self.capture_button = tk.Button(root, text="Select Area & Recognize Text", command=self.start_selection, font=("Arial", 12, "bold"), width=30, height=2)
        self.capture_button.pack(pady=5)
        
        # --- Frame for checkboxes ---
        checkbox_frame = tk.Frame(root)
        checkbox_frame.pack(pady=5)

        self.preserve_layout_var = tk.BooleanVar(value=False)
        self.layout_checkbox = tk.Checkbutton(
            checkbox_frame,
            text="Preserve Layout",
            variable=self.preserve_layout_var,
            font=("Arial", 10)
        )
        self.layout_checkbox.pack(side=tk.LEFT, padx=10)

        # *** NEW: Add a checkbox to control the "Always on Top" state ***
        self.always_on_top_var = tk.BooleanVar(value=True) # Checked by default
        self.on_top_checkbox = tk.Checkbutton(
            checkbox_frame,
            text="Always on Top",
            variable=self.always_on_top_var,
            font=("Arial", 10),
            command=self.toggle_always_on_top
        )
        self.on_top_checkbox.pack(side=tk.LEFT, padx=10)

        # --- A frame to manage the bottom button and status label ---
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        self.status_label = tk.Label(bottom_frame, text="Ready", fg="blue", font=("Arial", 9))
        self.status_label.pack(side=tk.LEFT)

        self.coffee_button = tk.Button(
            bottom_frame,
            text="Buy me a coffee ☕️",
            command=self.open_paypal_link,
            font=("Arial", 9)
        )
        self.coffee_button.pack(side=tk.RIGHT)

        # Set initial "Always on Top" state
        self.toggle_always_on_top()

    def toggle_always_on_top(self):
        """Updates the window's always-on-top attribute based on the checkbox."""
        is_on_top = self.always_on_top_var.get()
        self.root.attributes("-topmost", is_on_top)

    def open_paypal_link(self):
        """Opens the PayPal donation link in a new browser tab."""
        paypal_url = "https://buymeacoffee.com/foliage246"
        try:
            webbrowser.open_new_tab(paypal_url)
        except Exception as e:
            messagebox.showerror("Failed to Open", f"Could not open link:\n{e}")

    def start_selection(self):
        SelectionWindow(self)

    def capture_and_ocr(self, bbox):
        try:
            if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
                 raise FileNotFoundError

            self.show_main_window("Recognizing, please wait...", "orange")
            self.root.update_idletasks()

            screenshot = ImageGrab.grab(bbox=bbox)
            
            final_text = ""
            status_message = ""

            if self.preserve_layout_var.get():
                # --- (Mode 1) Preserve Layout with Advanced Spacing Logic ---
                data = pytesseract.image_to_data(screenshot, lang='eng+chi_tra', output_type=pytesseract.Output.DICT)
                
                n_boxes = len(data['level'])
                if n_boxes == 0:
                    final_text = ""
                else:
                    lines = {}
                    for i in range(n_boxes):
                        if int(data['conf'][i]) > 30 and data['text'][i].strip():
                            line_key = (data['block_num'][i], data['par_num'][i], data['line_num'][i])
                            if line_key not in lines:
                                lines[line_key] = []
                            
                            word_info = {'left': data['left'][i], 'width': data['width'][i], 'text': data['text'][i]}
                            lines[line_key].append(word_info)

                    reconstructed_lines = []
                    for key in sorted(lines.keys()):
                        words = sorted(lines[key], key=lambda w: w['left'])
                        if not words: continue

                        total_width = sum(w['width'] for w in words)
                        total_chars = sum(len(w['text']) for w in words)
                        avg_char_width = total_width / total_chars if total_chars > 0 else 1

                        line_str = ""
                        last_word_right = 0
                        
                        first_word_left = words[0]['left']
                        if first_word_left > avg_char_width:
                            num_indent_spaces = round(first_word_left / avg_char_width)
                            line_str += ' ' * int(num_indent_spaces)
                        
                        for word in words:
                            gap = word['left'] - last_word_right
                            if last_word_right > 0 and gap > avg_char_width:
                                num_spaces = round(gap / avg_char_width)
                                line_str += ' ' * int(num_spaces)
                            
                            line_str += word['text']
                            last_word_right = word['left'] + word['width']
                        
                        reconstructed_lines.append(line_str)
                    final_text = "\n".join(reconstructed_lines)
                status_message = "Success! Layout preserved and copied."
            else:
                # --- (Mode 2) Plain Text Extraction ---
                final_text = pytesseract.image_to_string(screenshot, lang='eng+chi_tra')
                status_message = "Success! Plain text copied."

            self.root.clipboard_clear()
            self.root.clipboard_append(final_text)
            self.show_main_window(status_message, "green")
        
        except FileNotFoundError:
             messagebox.showerror("Error", "Tesseract-OCR not found or path is not set.\nIf this is a packaged app, ensure the Tesseract-OCR folder is in the same directory as the .exe file.")
             self.show_main_window("Error: Tesseract not found", "red")
        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred:\n{e}")
            self.show_main_window("An error occurred", "red")
        
    def show_main_window(self, status_text, color="blue"):
        self.status_label.config(text=status_text, fg=color)
        self.root.deiconify()

if __name__ == "__main__":
    root = tk.Tk()
    app = OcrApp(root)
    root.mainloop()
