import tkinter as tk
from tkinter import ttk
import cv2
from PIL import Image, ImageTk

class CameraApp:
    def __init__(self, window, window_title):
        self.window = window
        self.window.title(window_title)

        self.video_source = 0  # Use the default camera (you can change it based on your camera index)

        self.vid = cv2.VideoCapture(self.video_source)

        self.canvas = tk.Canvas(window, width=self.vid.get(cv2.CAP_PROP_FRAME_WIDTH), height=self.vid.get(cv2.CAP_PROP_FRAME_HEIGHT))
        self.canvas.pack()

        self.btn_camera = ttk.Button(window, text="Camera", command=self.toggle_camera)
        self.btn_camera.pack(padx=10, pady=5)

        self.is_camera_open = False

        self.update()
        self.window.mainloop()

    def toggle_camera(self):
        if self.is_camera_open:
            self.stop_camera()
        else:
            self.start_camera()

    def start_camera(self):
        self.vid = cv2.VideoCapture(self.video_source)
        self.is_camera_open = True
        self.update()

    def stop_camera(self):
        if self.vid.isOpened():
            self.vid.release()
            self.is_camera_open = False

    def update(self):
        if self.is_camera_open:
            ret, frame = self.vid.read()
            if ret:
                photo = convert_frame_to_photo(frame)
                canvas.create_image(0, 0, anchor=tk.NW, image=photo)
                self.window.after(10, self.update)

    def convert_frame_to_photo(self, frame):
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        photo = ImageTk.PhotoImage(image=Image.fromarray(frame))
        return photo

if __name__ == "__main__":
    root = tk.Tk()
    app = CameraApp(root, "Tkinter Camera App")
