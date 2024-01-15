import tkinter as tk
import cv2
from PIL import Image, ImageTk

def open_camera_and_capture():
    # Open camera
    cap = cv2.VideoCapture(0)  # 0 corresponds to the default camera (you may need to change this if you have multiple cameras)

    if not cap.isOpened():
        print("Error: Could not open camera.")
        return

    # Capture a photo
    ret, frame = cap.read()

    if ret:
        # Convert the image from BGR to RGB
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # Display the captured image in the photo box
        photo = ImageTk.PhotoImage(Image.fromarray(frame_rgb))
        photo_box.config(image=photo)
        photo_box.image = photo  # to prevent garbage collection

    # Release the camera
    cap.release()

# Create the main window
add_record_window = tk.Tk()
add_record_window.title("Camera Example")

# Create a PhotoBox to display the captured image
photo_box = tk.Label(add_record_window)
photo_box.place(x=150, y=10)

# Create a button to open the camera and capture a photo
camera_button = tk.Button(add_record_window, text="Camera", font=("Arial", 18), command=open_camera_and_capture)
camera_button.place(x=38, y=520)

# Start the Tkinter event loop
add_record_window.mainloop()
