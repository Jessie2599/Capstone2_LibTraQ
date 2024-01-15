import logging
import tkinter as tk
from datetime import datetime
from tkinter import ttk, Entry, messagebox, END, scrolledtext
from typing import io
import matplotlib.pyplot as plt
import pymysql
from PIL import Image
from PIL import ImageTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import io
from PIL import ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES=True

recorded_data = []
date_time_label = None
library_id_no_entry = None
purpose_label = None
purposes_frame = None
entry_enabled = True
selected_purpose = None
db = None
cursor = None

current_date = datetime.now()

def connection():
    conn = pymysql.connect(host='localhost', user='root', password="", database='libtraq_db')
    return conn

def on_closing():
    pass

if 8 <= current_date.month <= 12:
    academic_year = str(current_date.year) + "-" + str(current_date.year + 1)
    semester = "1st Semester"
elif 1 <= current_date.month <=7:
    academic_year = str(current_date.year - 1) + "-" + str(current_date.year)
    semester = "2nd Semester"
elif current_date.month == 1:
    academic_year = str(current_date.year - 1) + "-" + str(current_date.year)
    semester = "2nd Semester"

def clear_entry():
    library_id_no_entry.delete(0, END)

def update_datetime():
    global date_time_label
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    day_of_week = datetime.now().strftime("%A")
    date_time_label.config(text=f"{day_of_week}, {current_datetime}")
    date_time_label.after(1000, update_datetime)

def clear_and_restart():
    global library_id_no_entry, purpose_label, purposes_frame, entry_enabled, selected_purpose
    library_id_no_entry.config(state=tk.NORMAL)
    library_id_no_entry.delete(0, tk.END)
    purpose_label.config(text="Purpose:")
    for widget in purposes_frame.winfo_children():
        widget.destroy()
    entry_enabled = True
    selected_purpose = None

def restart():
    clear_and_restart()
    display_purposes(None)

def show_confirmation_window(library_id_no, first_name, middle_name, last_name, course, current_datetime, purpose):
    confirmation_window = tk.Toplevel()
    confirmation_window.title("Confirmation Window")
    confirmation_window.protocol("WM_DELETE_WINDOW", on_closing)

    confirmation_message = f"Library ID Number: {library_id_no}\nFirst Name: {first_name}\nMiddle Name: {middle_name}\nLast Name: {last_name}\nCourse: {course}\nDate and Time: {current_datetime}\nPurpose: {purpose}"

    confirmation_label = tk.Label(confirmation_window, text=confirmation_message, padx=20, pady=20)
    confirmation_label.pack()

    photo_box = tk.Canvas(confirmation_window, width=220, height=200, bg="white", highlightbackground="black", highlightthickness=2)
    photo_box.pack()

    connect = pymysql.connect(host='localhost', user='root', password="", database='libtraq_db')
    cursor = connect.cursor()

    cursor.execute("SELECT photo FROM student WHERE library_id_no = %s", (library_id_no,))
    result = cursor.fetchone()

    if result:
        image_data = result[0]
        image = Image.open(io.BytesIO(image_data))
        image.thumbnail((220, 200))

        photo = ImageTk.PhotoImage(image)
        photo_box.image = photo
        photo_box.create_image(0, 0, anchor=tk.NW, image=photo)

    connect.close()

    def close_confirmation_window():
        confirmation_window.destroy()

    confirmation_window.after(3000, close_confirmation_window)

def save_and_confirm_purpose(purpose):
    global library_id_no_entry, entry_enabled, selected_purpose, academic_year, semester
    library_id_no = library_id_no_entry.get()
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if library_id_no:
        recorded_data.append([library_id_no, current_datetime, purpose, academic_year, semester])
        conn = connection()
        cursor = conn.cursor()

        try:
            cursor.execute("SELECT first_name, middle_name, last_name, course FROM student WHERE library_id_no = %s", (library_id_no,))
            result = cursor.fetchone()

            if result:
                first_name, middle_name, last_name, course = result
                cursor.execute("INSERT INTO `library_attendance1` (`library_id_no`, `first_name`, `middle_name`, `last_name`, `course`, `purpose`, `date_and_time`, `academic_year`, `semester`) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)", (library_id_no, first_name, middle_name, last_name, course, purpose, current_datetime, academic_year, semester))
                conn.commit()
                conn.close()
                show_confirmation_window(library_id_no, first_name, middle_name, last_name, course, current_datetime, purpose)
                clear_entry()
                clear_and_restart()
            else:
                messagebox.showerror("Error", "Library ID Number not found in the 'student' table.")
        except pymysql.Error as e:
            messagebox.showerror("Database Error", f"Error: {e}")

def check_library_id_no(library_id_no):
    connect = pymysql.connect(host='localhost', user='root', password="", database='libtraq_db')
    cursor = connect.cursor()
    cursor.execute("SELECT library_id_no, status FROM student WHERE library_id_no = %s", (library_id_no,))
    result = cursor.fetchone()
    connect.close()

    logging.debug(f"Result from database: {result}")

    if result is not None:
        if result[1] == 'Active':
            return True
        else:
            return False
    else:
        return False

def display_purposes(event):
    global library_id_no_entry, entry_enabled, selected_purpose
    library_id_no = library_id_no_entry.get()

    if library_id_no and entry_enabled:
        if check_library_id_no(library_id_no):
            purposes = [
                "A. Read a Book",
                "B. Read a Journal",
                "C. Read a Periodical ",
                "D. Borrow a Book ",
                "E. Return a Book",
                "F. Connect to the Internet",
                "G. Research",
                "H. Appointment with the Librarian",
                "I. Utilization of Undergrad Thesis",
                "J. Others"
            ]
            for widget in purposes_frame.winfo_children():
                widget.destroy()

            for i, purpose in enumerate(purposes, start=1):
                purpose_button = tk.Button(purposes_frame, text=purpose, font=("Rockwell ", 11), width=34, anchor="w", command=lambda purpose=purpose: save_and_confirm_purpose(purpose))
                purpose_button.pack(fill="x")

            library_id_no_entry.config(state=tk.DISABLED)
            entry_enabled = False
        else:
            messagebox.showerror("Error", "Your account is inactive. Please contact the library administrator for assistance.")
            return
def handle_purpose_key(key, purpose):
    global selected_purpose
    selected_purpose = purpose
    save_and_confirm_purpose(purpose)

def fetch_data_and_create_graph():
    try:
        conn = connection()
        cursor = conn.cursor()
        cursor.execute("SELECT course, COUNT(*) FROM library_attendance1 GROUP BY course")
        results = cursor.fetchall()
        conn.close()

        if not results:
            messagebox.showerror("Error", "No attendance data found.")
            return

        courses, counts = zip(*results)
        total_count = sum(counts)

        percentages = [count / total_count * 100 for count in counts]

        fig, ax = plt.subplots(figsize=(6, 4))
        bars = ax.bar(courses, counts, color='gray')

        ax.set_xlabel('Course')
        ax.set_ylabel('Attendance Count')
        ax.set_title('Attendance by Course')
        ax.tick_params(axis='x', rotation=45)

        for bar, percentage in zip(bars, percentages):
            height = bar.get_height()
            ax.annotate(f'{percentage:.2f}%', (bar.get_x() + bar.get_width() / 2, height), ha='center', va='bottom')

        canvas = FigureCanvasTkAgg(fig, master=home)
        canvas.get_tk_widget().place(x=20, y=320)
        canvas.draw()

        plt.close(fig)

    except pymysql.Error as e:
        messagebox.showerror("Database Error", f"Error: {e}")

def generate_leaderboard():
    try:
        conn = connection()
        cursor = conn.cursor()
        cursor.execute("SELECT first_name, middle_name, last_name, course, purpose, COUNT(*) as attendance_count FROM library_attendance1 GROUP BY first_name, middle_name, last_name, course, purpose ORDER BY attendance_count DESC LIMIT 10")
        results = cursor.fetchall()
        conn.close()

        if not results:
            messagebox.showerror("Error", "No attendance data found.")
            return

        leaderboard_label = tk.Label(home, text="Top Ten Visitors", font=("Arial", 15, "bold"))
        leaderboard_label.place(x=20, y=60)

        leaderboard_tree = ttk.Treeview(home, columns=("Name", "Course", "Purpose", "Attendance"), show="headings", height=10)
        leaderboard_tree.heading("Name", text="Name", anchor="center")
        leaderboard_tree.heading("Course", text="Course", anchor="center")
        leaderboard_tree.heading("Purpose", text="Purpose", anchor="center")
        leaderboard_tree.heading("Attendance", text="Attendance", anchor="center")

        style = ttk.Style()
        style.configure("Treeview.Heading")

        leaderboard_tree.column("Name", width=200, anchor='center')
        leaderboard_tree.column("Course", width=120, anchor='center')
        leaderboard_tree.column("Purpose", width=170, anchor='center')
        leaderboard_tree.column("Attendance", width=130, anchor='center')

        leaderboard_tree.place(x=20, y=90)

        for index, (first_name, middle_name, last_name, course, purpose, attendance_count) in enumerate(results, start=1):
            name = f"{first_name} {middle_name} {last_name}"
            leaderboard_tree.insert("", "end", values=(name, course, purpose, attendance_count))

    except pymysql.Error as e:
        messagebox.showerror("Database Error", f"Error: {e}")

home = tk.Tk()
home.title("LibTraQ: Library Tracker and Monitoring System using QR Code")
home.geometry("1366x768")
home.resizable(height=False, width=False)

def open_marquee_window():
    pin_window = tk.Toplevel()
    pin_window.title("Enter PIN")
    pin_window.geometry("200x100")
    pin_window.resizable(height=False, width=False)

    pin_label = tk.Label(pin_window, text="Enter PIN:", font=("Arial", 12))
    pin_label.pack()

    pin_entry = tk.Entry(pin_window, show="*", font=("Arial", 12))
    pin_entry.pack()


    def check_pin_and_open_marquee():
        # Check the PIN entered here
        if pin_entry.get() == "admin":  # Replace "1234" with your actual PIN
            pin_window.destroy()
            marquee_window = tk.Toplevel()
            marquee_window.title("Library Announcement")
            marquee_window.geometry("300x200")
            marquee_window.resizable(height=False, width=False)

            marquee_label = tk.Label(marquee_window, text="Enter Announcement:", font=("Rockwell", 16))
            marquee_label.pack()

            marquee_entry = scrolledtext.ScrolledText(marquee_window, font=("Arial", 16), height=5)
            marquee_entry.pack()

            def display_marquee_message():
                message = marquee_entry.get("1.0", tk.END)
                marquee_message.config(text=message)
                marquee_window.destroy()

            submit_button = tk.Button(marquee_window, text="Submit", font=("Rockwell", 16), command=display_marquee_message)
            submit_button.pack()


    submit_pin_button = tk.Button(pin_window, text="Submit", font=("Arial", 12), command=check_pin_and_open_marquee)
    submit_pin_button.pack()

date_time_label = tk.Label(home, text="", font=("Arial", 18))
date_time_label.place(relx=0.03, rely=1.0, anchor=tk.SW)
update_datetime()

canvas = tk.Canvas(home, width=1600, height=1500)
canvas.pack()

my_tree = ttk.Treeview(home, height=20)
my_tree['columns'] = ("Library ID Number", "First Name", "Middle Name", "Last Name", "Course", "Purpose", "Date & Time")

my_tree.column("#0", width=0, stretch=tk.NO)
my_tree.column("Library ID Number", anchor="center", width=120)
my_tree.column("First Name", anchor="center", width=120)
my_tree.column("Middle Name", anchor="center", width=120)
my_tree.column("Last Name", anchor="center", width=120)
my_tree.column("Course", anchor="center", width=110)
my_tree.column("Purpose", anchor="center", width=150)
my_tree.column("Date & Time", anchor="center", width=150)

my_tree.heading("Library ID Number", text="Library ID Number", anchor="center")
my_tree.heading("First Name", text="First Name", anchor="center")
my_tree.heading("Middle Name", text="Middle Name", anchor="center")
my_tree.heading("Last Name", text="Last Name", anchor="center")
my_tree.heading("Course", text="Course", anchor="center")
my_tree.heading("Purpose", text="Purpose", anchor="center")
my_tree.heading("Date & Time", text="Date & Time", anchor="center")

my_tree.pack()

background_image = tk.PhotoImage(file="images/student_background.png")
background_label = tk.Label(canvas, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

library_id_no_label = tk.Label(home, text="Library ID Number:", font=("Rockwell", 20))
library_id_no_label.place(x=730, y=228)

library_id_no_entry = Entry(home, font=("Rockwell", 20), width=12, highlightthickness=3)
library_id_no_entry.place(x=980, y=228)
library_id_no_entry.focus()

library_id_no_entry.bind("<Return>", display_purposes)

purpose_label = tk.Label(home, text="Purpose:", font=("Rockwell", 20))
purpose_label.place(x=730, y=270)

purposes_frame = tk.Frame(home, highlightthickness=3)
purposes_frame.place(x=850, y=270)

entry_enabled = True
selected_purpose = None

home.bind("<space>", lambda event: restart())

home.bind("<Control-g>", lambda event: fetch_data_and_create_graph())
home.bind("<Control-a>", lambda event: open_marquee_window())
home.bind("<Control-l>", lambda event: generate_leaderboard())

icon_image = Image.open("images/announcement_icon.png")
original_width, original_height = icon_image.size
aspect_ratio = original_width / original_height
new_width = 120
new_height = 70
icon_image = icon_image.resize((new_width, new_height), Image.LANCZOS)
icon_photo = ImageTk.PhotoImage(icon_image)

marquee_button = tk.Button(home, image=icon_photo, command=open_marquee_window, height=new_height, width=new_width)
marquee_button.image = icon_photo
marquee_button.place(x=390, y=10)

marquee_message = tk.Label(home, text="", font=("Rockwell", 16), fg="black", wraplength=1000)
marquee_message.place(x=730, y=600)

icon_image = Image.open("images/leaderboard_icon.png")
original_width, original_height = icon_image.size
aspect_ratio = original_width / original_height
new_width = 120
new_height = 70
icon_image = icon_image.resize((new_width, new_height), Image.LANCZOS)
icon_photo = ImageTk.PhotoImage(icon_image)

fetch_and_generate_button = tk.Button(home, image=icon_photo, command=lambda: [fetch_data_and_create_graph(), generate_leaderboard()], height=new_height, width=new_width)
fetch_and_generate_button.image = icon_photo
fetch_and_generate_button.place(x=520, y=10)

update_datetime()
home.mainloop()
