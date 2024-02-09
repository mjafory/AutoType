import pyautogui
import tkinter as tk
from tkinter import filedialog
from docx import Document
import fitz  # PyMuPDF
import threading
import time
import os
import sys
import pyautogui


pyautogui.FAILSAFE = False  # Disable PyAutoGUI fail-safe
autotype_flag = False  # Flag to control autotyping loop
pause_flag = False  # Flag to control autotyping pause
stop_event = threading.Event()  # Event to signal stopping autotyping



def resource_path(relative):
    return os.path.join(
        os.environ.get(
            "_MEIPASS2",
            os.path.abspath(".")
        ),
        relative
    )

def read_text_from_file(file_path):
    _, file_extension = file_path.lower().rsplit('.', 1)

    if file_extension == 'txt':
        with open(file_path, 'r') as file:
            return file.read()
    elif file_extension == 'docx':
        doc = Document(file_path)
        return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    elif file_extension == 'pdf':
        pdf_document = fitz.open(file_path)
        text = ''
        for page_number in range(pdf_document.page_count):
            page = pdf_document[page_number]
            text += page.get_text()
        return text
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

def autotype_from_text(text, delay=5, interval=0.1):
    global pause_flag, stop_event
    for char in text:
        while pause_flag:
            time.sleep(1)  # Check for pause every second
        if stop_event.is_set():
            break  # Stop autotyping if the event is set
        pyautogui.write(char)
        time.sleep(interval)

def autotype_from_file(file_path, delay=5, interval=0.1):
    text = read_text_from_file(file_path)
    autotype_from_text(text, delay, interval)
    finished_label.config(text="Finished!", fg="green")  # Set label text and color when autotyping is finished
    finished_text.set(text)  # Set the finished text variable

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("Word Documents", "*.docx"), ("PDF Files", "*.pdf")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def start_autotype():
    global autotype_flag, stop_event
    autotype_flag = True
    stop_event.clear()  # Clear the stop event flag
    text = text_area.get("1.0", tk.END)  # Get text from the text area
    if not text.strip():  # If the text area is empty, use the file path
        file_path = file_entry.get()
        threading.Thread(target=autotype_from_file, args=(file_path,), kwargs={'delay': float(delay_var.get()), 'interval': float(speed_var.get())}).start()
    else:
        threading.Thread(target=autotype_from_text, args=(text,), kwargs={'delay': float(delay_var.get()), 'interval': float(speed_var.get())}).start()
    start_button.config(bg='green')  # Change start button color to green
    stop_button.config(bg='SystemButtonFace')  # Reset stop button color
    pause_button.config(bg='SystemButtonFace')  # Reset pause button color

def stop_autotype():
    global autotype_flag, stop_event
    autotype_flag = False
    stop_event.set()  # Set the stop event flag
    stop_button.config(bg='red')  # Change stop button color to red
    start_button.config(bg='SystemButtonFace')  # Reset start button color
    pause_button.config(bg='SystemButtonFace')  # Reset pause button color

def pause_autotype():
    global pause_flag
    pause_flag = not pause_flag
    if pause_flag:
        pause_button.config(bg='yellow')  # Change pause button color to yellow
    else:
        pause_button.config(bg='SystemButtonFace')  # Reset pause button color
        if not autotype_flag:  # Check if autotyping is not running
            time.sleep(5)  # Wait for 5 seconds before restarting when resuming
            pyautogui.click()  # Click on the current mouse position (assumes the cursor is already over the input field)
            start_autotype()

def toggle_always_on_top():
    is_always_on_top = root.attributes('-topmost')
    root.attributes('-topmost', not is_always_on_top)

def delayed_start_autotype():
    print("Switch to the input field. Autotyping will start in a few seconds...")
    time.sleep(5)  # Adjust the default delay as needed
    pyautogui.click()  # Click on the current mouse position (assumes the cursor is already over the input field)
    start_autotype()

def on_entry_click(event):
    if file_entry.get() == 'Enter file path...':
        file_entry.delete(0, tk.END)
        file_entry.config(fg='black')  # Change text color to black

def on_entry_leave(event):
    if not file_entry.get():
        file_entry.insert(0, 'Enter file path...')
        file_entry.config(fg='grey')  # Change text color to grey

def import_text():
    input_text = text_area.get("1.0", tk.END)
    with open("imported_text.txt", "w") as file:
        file.write(input_text)

def clear_text():
    text_area.delete("1.0", tk.END)
    finished_label.config(text="")  # Clear finished label
    finished_text.set("")  # Clear finished text variable


# GUI Setup
root = tk.Tk()
root.title("Autotyping App")


# Set the fixed size
root.geometry("430x500")
# Disallow resizing
root.resizable(width=False, height=False)


# Add Logo to Title Bar
logo_path = resource_path("logo.PNG")
logo = tk.PhotoImage(file=logo_path)
root.iconphoto(True, logo)

# File Selection
file_frame = tk.Frame(root)
file_frame.grid(row=0, column=0, columnspan=2, pady=10)

file_entry = tk.Entry(file_frame, width=35, fg='blue')
file_entry.insert(0, 'Enter file path...')
file_entry.bind('<FocusIn>', on_entry_click)
file_entry.bind('<FocusOut>', on_entry_leave)
file_entry.grid(row=0, column=0, padx=5, ipadx=5, ipady=5)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.grid(row=0, column=1, padx=5)

file_entry.configure(font=("Georgia", 11), bg="#f0f0f0", fg="black", relief="solid", borderwidth=2)
browse_button.configure(font=("Georgia", 11), bg="#4CAF50", fg="white", relief="raised", borderwidth=2)

# Text Area for Copy/Paste
text_area_frame = tk.Frame(root)
text_area_frame.grid(row=1, column=0, columnspan=2, pady=10)

text_area_label = tk.Label(text_area_frame, text="Type or Paste Text:")
text_area_label.grid(row=0, column=0, pady=(0, 5))

text_area = tk.Text(text_area_frame, height=10, width=50, bd=2, relief="solid")
text_area.grid(padx=15, pady=5)

# Import Text Button
import_text_button = tk.Button(text_area_frame, text="Save Text", command=import_text)
import_text_button.grid(row=2, column=0, padx=(15, 0), pady=2, sticky="w")  # Adjust the padx parameter and use sticky="w"

# Clear Text Button
clear_button = tk.Button(text_area_frame, text="Clear Text", command=clear_text)
clear_button.grid(row=2, column=0, padx=(100, 0), pady=2, sticky="w")  # Adjust the padx parameter and use sticky="w"



# Speed Control
speed_label = tk.Label(root, text="Speed (interval between keystrokes):")
speed_label.grid(row=2, column=0, padx=(15, 0), pady=10, sticky=tk.W)

speed_var = tk.StringVar()
speed_entry = tk.Entry(root, textvariable=speed_var, width=10)
speed_entry.insert(0, "0.1")  # Default speed
speed_entry.grid(row=2, column=1, pady=10, padx=(0, 20),)

# Delay Control
delay_label = tk.Label(root, text="Start Delay (seconds):")
delay_label.grid(row=3, column=0, padx=(15, 0),pady=10, sticky=tk.W)

delay_var = tk.StringVar()
delay_entry = tk.Entry(root, textvariable=delay_var, width=10)
delay_entry.insert(0, "5")  # Default delay
delay_entry.grid(row=3, column=1, pady=10, padx=(0, 20))

# Start Autotype Button
start_button = tk.Button(root, text="Start", command=delayed_start_autotype)
start_button.grid(row=4, column=0, pady=2,sticky="w", padx=(130, 0))

# Stop Autotype Button
stop_button = tk.Button(root, text="Stop", command=stop_autotype)
stop_button.grid(row=4, column=0, pady=2, sticky="w",padx=(185, 0))

# Pause Autotype Button
pause_button = tk.Button(root, text="Pause/Resume", command=pause_autotype)
pause_button.grid(row=4, column=0, pady=2, sticky="w",padx=(245, 0))

# Finished Label
finished_label = tk.Label(root, text="", fg="black")
finished_label.grid(row=6, column=0, columnspan=2, pady=10)

# Always on Top Checkbox
always_on_top_var = tk.IntVar()
always_on_top_checkbox = tk.Checkbutton(root, text="Always on Top", variable=always_on_top_var, command=toggle_always_on_top)
always_on_top_checkbox.grid(row=7, column=0, pady=(0, 5), sticky="w",padx=(80,0)) # Use sticky="w" to align to the left

# Curtesy Label
curtesy_label = tk.Label(root, text="Made By M Jafory", fg="brown")
curtesy_label.grid(row=7, column=0, pady=(0, 5), sticky="w",padx=(230,0))  # Use sticky="w" to align to the left


# Variable to hold finished text
finished_text = tk.StringVar()

root.mainloop()
