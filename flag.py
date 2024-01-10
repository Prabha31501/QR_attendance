# Libraries for this project
#-----------------Sending email--------------#
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
#-----------Creating window application-----------#
import customtkinter as ctk
import tkinter as tk
#----------------Adding image in the application--------------#
from PIL import ImageTk, Image
from tkinter import Tk, Frame
#-------------Generating QR code-----------------------#
import qrcode
from pyzbar import pyzbar
from pyzbar.pyzbar import decode
#------------------Open camera-------------------#
import cv2
#--------------------Date and time-----------------#
from datetime import datetime
import time
import pandas as pd
#_------------------Google spread sheet----------------#
import gspread
#-----------Message pop-up----------------#
import tkinter.messagebox as tkmb
import numpy as np
#----------Library files for Generating pdf------------------#
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage, Spacer, Paragraph
from tkinter import PhotoImage
from reportlab.lib.styles import getSampleStyleSheet
import re
import PySimpleGUI as sg
import os

#-----------------------Creating the window---------------------------#
ctk.set_appearance_mode("dark")
# Selecting color theme - blue, green, dark-blue
ctk.set_default_color_theme("green")
root = ctk.CTk()
root.title("Inphase Power Technologies")
root.geometry("450x560")
root.resizable(False,False)
frame_1 = ctk.CTkFrame(root, width=450, height=560)
frame_1.place(x=0, y=0)
icon = Image.open('pk.ico')
photo = ImageTk.PhotoImage(icon)

#--------------This line is used to change the icon------------------------#
root.wm_iconphoto(False, photo)
root.iconbitmap('pk.ico')

#Impoting the  image of inphase
inphase_logo=os.path.join(os.path.dirname(__file__),'test.png')
image=ctk.CTkImage(light_image=Image.open(inphase_logo),size=(280,140))
image_lable=ctk.CTkLabel(frame_1,image=image, text="").place(x=100,y=10)
# ------------------------------------ Open Google spread sheet ------------------------#
gc = gspread.service_account(filename='nsdc.json')
spreadsheet = gc.open("NSDC- Attendance")
wks1 = spreadsheet.get_worksheet(0)
wks2 = spreadsheet.get_worksheet(1)
wks3 = spreadsheet.get_worksheet(2)
wks4 = spreadsheet.get_worksheet(3)

#------------------------------------- Current date and time -------------------------#
# Get the current date and time
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
current_date = now.strftime("%Y-%m-%d")

#------------------Admin login button functions starts------------------#
def admin_log():

    #pass_frame = ctk.CTkFrame(frame_1, width=450, height=560).place(x=0, y=0)
    pass_frame_2 = ctk.CTkFrame(root, width=450, height=560).place(x=0, y=0)
    admin_label = ctk.CTkLabel(pass_frame_2, image=image, text="").place(x=100,y=10)
    wecome = ctk.CTkLabel(pass_frame_2, text="Admin login").place(x=150, y=180)
    user_label= ctk.CTkLabel(pass_frame_2, text="Enter the user name").place(x=37,y=220)
    pass_label = ctk.CTkLabel(pass_frame_2, text="Enter the password  ").place(x=37, y=262)
    user_entry= ctk.CTkEntry(pass_frame_2, textvariable=l, placeholder_text="Username").place(x=200, y=221)
    pass_entry = ctk.CTkEntry(pass_frame_2,textvariable=l2, show="*",placeholder_text="Password").place(x=200, y=263)
    ent_btn = ctk.CTkButton(pass_frame_2, text="Login", command=enter).place(x=120, y=320)
    back_btn = ctk.CTkButton(pass_frame_2, text="Back",command=back).place(x=280, y=320)

#------------------------------------Admin login button function ends-------------------------#

#-----------------------------password enter button function starts---------------------#
def enter():
    #global enter
    user_1=l.get()
    pass_1=l2.get()
    l.set("")
    l2.set("")
    rows_1 = wks3.get_all_values()
    cell = wks3.find(user_1)
    row_index = cell.row
    row_values = wks3.row_values(row_index)
    if user_1==row_values[1] and pass_1==row_values[2]:
        msg = f'Login Successfull \n Welcome {row_values[0]}'
        tkmb.showinfo(title='Information',message=msg)

        frame_1 = ctk.CTkFrame(root, width=450, height=560).place(x=0, y=0)
        inphase_lable = ctk.CTkLabel(frame_1, image=image, text="").place(x=100,y=10)
        #Admin_btn = ctk.CTkButton(frame_1, text="Admin", command=admin_log).place(x=180, y=200)
        register_btn = ctk.CTkButton(frame_1, text="Register", command=new_register,state="normal").place(x=180, y=250)
        entry_btn = ctk.CTkButton(frame_1, text="Entry", command=entry).place(x=180, y=300)
        logout_btn= ctk.CTkButton(frame_1, text="Logout", command=logout).place(x=180, y=350)
        de_user= ctk.CTkButton(frame_1, text="User Access", command=remo).place(x=180, y=400)
    else:
        msg = f'Login failed! \n Enter the valid user name and password'
        tkmb.showinfo(title='Information', message=msg)
#------------------------Password enter button function ends-----------------------#

#-----------------------------------New register button function starts----------------------#
def new_register():

    #reg_frame= ctk.CTkFrame(frame_1, width=450, height=560).place(x=0, y=0)
    reg_frame_2 = ctk.CTkFrame(root, width=450, height=560).place(x=0, y=0)
    in_lab= ctk.CTkLabel(reg_frame_2, image=image, text="").place(x=100, y=20)
    nr_lab= ctk.CTkLabel(reg_frame_2, text="New Register").place(x=150, y=180)

    #Labels for input
    name_lab= ctk.CTkLabel(reg_frame_2, text="Enter the name").place(x=37,y=220)
    email_lab = ctk.CTkLabel(reg_frame_2, text="Enter the Email").place(x=37,y=260)
    number_lab = ctk.CTkLabel(reg_frame_2, text="Enter the number").place(x=37,y=300)
    empid_lab = ctk.CTkLabel(reg_frame_2, text="Enter the Emp.Id").place(x=37,y=340)

    #user input entry to collect the datas
    name_entry = ctk.CTkEntry(reg_frame_2, textvariable=l3, placeholder_text="Enter name").place(x=200, y=221)
    email_entry = ctk.CTkEntry(reg_frame_2, textvariable=l4, placeholder_text="Enter email").place(x=200, y=261)
    number_entry = ctk.CTkEntry(reg_frame_2, textvariable=l5,  placeholder_text="Enter number").place(x=200, y=301)
    empid_entry = ctk.CTkEntry(reg_frame_2, textvariable=l6,  placeholder_text="Enter Emp.Id").place(x=200, y=341)

    create_btn= ctk.CTkButton(reg_frame_2, text="Create", command=create).place(x=100, y=381)
    back_btn = ctk.CTkButton(reg_frame_2, text="Back", command=back_1).place(x=290, y=381)
#---------------------------------New register button function ends----------------------#

#----------------------------------Create button function starts------------------------#
def create():
    global wks2
    #collecting the input
    text = l3.get()
    text2 = l4.get()
    text3 = l5.get()
    text4 = l6.get()
     #Adding the data in QR code
    QRcode = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
    data = text + ',' + text2 + ',' + text3 + ',' + text4
    QRcode.add_data(data)
    QRcode.make()

    # Generate QR code with black color with white background
    QRcolor = 'black'
    img = QRcode.make_image(fill_color=QRcolor, back_color="white").convert('RGB')
    pd=img.save(text4 + ".png")
    pd_2=img.save(text4 +".pdf")

    # Creating PDF
    data = [
        ["Name", f'{text}'],
        ["Emp.Id", f'{text4}'],
        ["Email Id", f'{text2}'],
        ["Mobile NO.", f'{text3}']
    ]

    # Set up PDF document
    file_path_pdf = text + ".pdf"
    #file_path_pdf_2 = text + "_2.pdf"
    #file_path_pdf_3 = text + "_3.pdf"
    doc = SimpleDocTemplate(file_path_pdf, pagesize=A4)
    #doc_2 = SimpleDocTemplate(file_path_pdf_2, pagesize=landscape(A4))
    #doc_3 = SimpleDocTemplate(file_path_pdf_3, pagesize=A4)

    # Build the PDF document
    elements = []

    # Use the alias for the reportlab Image class
    extra_image_path = 'test.png'  # Replace with the actual path to your image
    extra_img = ReportLabImage(extra_image_path, width=300, height=50)  # Set the width as the page width
    elements.append(extra_img)

    # Add Spacer to create empty space below the extra image (adjust coordinates as needed)
    #elements.append(Spacer(1, 1))

    # Add Spacer to create empty space above the table (adjust coordinates as needed)
    elements.append(Spacer(1, 10))

    # Set up table with data
    table_data = [data[0]] + data[1:]
    table = Table(table_data)

    # Apply the style to the table
    style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black)
    ])
    table.setStyle(style)

    # Add the table to the elements
    elements.append(table)
    text_content = (f'On behalf of Inphase, I am extremely excited to share with '
                    f'you the offer letter for the role of R&D Engineer. Your passion and skills '
                    f'are the perfect fit for the company. You will be a part of the team starting '
                    f'from 03/01/2024. As for your  QR access, it is attached to this email. Please dont share this QR code to anyone'
                    f'If you need any guidence about this QR code. Kindly contact nsdccourse@gmail.com'
                    f'thank you')
    text_paragraph = Paragraph(text_content, getSampleStyleSheet()['BodyText'])
    elements.append(text_paragraph)

    # Add Spacer to create empty space below the table (adjust coordinates as needed)
    elements.append(Spacer(1, 15))

    # Use the alias for the reportlab Image class
    img_path = text4 + '.png'
    img = ReportLabImage(img_path, width=100, height=100)  # Set the width and height as needed
    elements.append(img)

    doc.build(elements)


    #------------------------------- Store new registered data in Gsheet --------------------------#
    wks2 = gc.open("NSDC- Attendance").worksheet("Database")
    wks2.append_row([text, text2, text3,text4, current_date, current_time])
    wks2.format('A1:F1', {'textFormat': {'bold': True}})

    # ------------------------------------------- Send Qr to entered E-mail -----------------------------------------
    # E mail the QR code
    msg = MIMEMultipart()
    msg['Subject'] = 'QR code access | Inphase'
    body_text = f'hi {text},\n{a2[0]}\n{a2[1]}\n{a2[2]}\n{a2[3]}'
    msg.attach(MIMEText(body_text, 'plain'))

    # open the file in binary
    binary_image = open(text + '.pdf', 'rb')
    binary_image_1 = open(text4 + '.png', 'rb')
    file_name = text + '.pdf'
    file_name_1 = text4 + '.png'
    payload = MIMEBase('application', 'octate-stream', Name=file_name)
    payload_1 = MIMEBase('application', 'octate-stream', Name=file_name_1)
    payload.set_payload((binary_image).read())
    payload_1.set_payload((binary_image_1).read())

    # encoding the binary into base64
    encoders.encode_base64(payload)
    encoders.encode_base64(payload_1)

    # add header with pdf name
    payload.add_header('Content-Decomposition', f'attachment', filename=file_name)
    msg.attach(payload)
    payload_1.add_header('Content-Decomposition', f'attachment', filename=file_name_1)
    msg.attach(payload_1)
    msg['To'] = text2
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login('nsdccourse@gmail.com', 'zcsn voxd zxod aabf')
    s.sendmail('nsdccourse@gmail.com', msg['To'], msg.as_string())
    s.quit()
    #msg = f'QR code Sent Successfull to Registered Email \n Check the inbox {text2}'
    #showinfo(title='Information', message=msg)
    ctk.CTkLabel(root, text="QR code generated and send through email").place(x=10,y=450)
    l3.set("")
    l4.set("")
    l5.set("")
    l6.set("")
    #new_register()

#------------------------Create button function ends------------------------#

#----------------------------User entry button function starts------------------#
def entry():

    entry_frame = ctk.CTkFrame(frame_1, width=450, height=560)
    entry_frame.place(x=0, y=0)
    label = ctk.CTkLabel(entry_frame, text="Press 'Snap' to capture a snapshot.")
    label.pack()
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    current_date = now.strftime("%Y-%m-%d")
    gc = gspread.service_account('nsdc.json')
    spreadsheet = gc.open("NSDC- Attendance")
    wks1 = spreadsheet.get_worksheet(0)

    snapshot = None
    snap_button_pressed = False

    # import cv2
    # import time
    # from pyzbar.pyzbar import decode

    snapshot = None
    snap_button_pressed = False

    def on_mouse(event, x, y, flags, param):
        global snapshot, snap_button_pressed
        if event == cv2.EVENT_LBUTTONDOWN:
            # Check if the click is within the Snap button region
            if 500 <= x <= 600 and 20 <= y <= 70:
                snap_button_pressed = True
                #print("Snap button pressed!")
    def trace_qr_code():
        global snapshot, snap_button_pressed
        # Open the default camera (index 0)
        cap = cv2.VideoCapture(0)
        # Wait for 5 seconds before starting QR code scanning
        time.sleep(5)
        # Create the main window
        cv2.namedWindow("Traced QR Code")
        snap_button_pressed = False  # Initialize snap_button_pressed

        while True:
            # Read a frame from the webcam
            ret, frame = cap.read()
            # Convert the frame to grayscale
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            # Display the Snap button
            cv2.rectangle(frame, (500, 20), (600, 70), (255, 255, 255), -1)
            cv2.putText(frame, "Snap", (530, 55), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 0), 2, cv2.LINE_AA)
            # Check for mouse events
            cv2.setMouseCallback("Traced QR Code", on_mouse)
            # Capture a snapshot if the Snap button is pressed
            if snap_button_pressed:
                snapshot = frame.copy()
                # Decode QR codes in the snapshot
                qr_codes = decode(cv2.cvtColor(snapshot, cv2.COLOR_BGR2GRAY))
                # Iterate through detected QR codes
                for qr_code in qr_codes:
                    rect = qr_code.rect
                    # Extract the coordinates of the bounding box
                    x, y, w, h = rect.left, rect.top, rect.width, rect.height
                    # Draw the bounding box
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    # Extract and print QR code data
                    data_1 = qr_code.data.decode('utf-8')
                    # Close the window after decoding the QR code
                    cv2.destroyAllWindows()
                    cap.release()
                    return data_1 # Exit the function immediately

                snap_button_pressed = False  # Reset the button state

            # Display the frame with traced QR codes
            cv2.imshow("Traced QR Code", frame)

            # Break the loop if 'q' key is pressed
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        # Release the webcam and close the window (in case 'q' is pressed)
        cap.release()
        cv2.destroyAllWindows()
    # Call the function to start webcam QR code scanning
    decoded_objects=trace_qr_code()

    # Check if any QR codes were found
    def process_attendance(data_1, current_date, current_time):
        if data_1:
            data = data_1.split(',')
            compare_number = data[2]
            # Check deactivation
            if not deactivated(compare_number):
                # Check attendance
                if update_out_time(data, compare_number, current_date, current_time):
                    entry_type = "Out-Time Entry"
                else:
                    add_attendance_entry(data, current_date, current_time, "In-Time Entry")
                    entry_type = "In-Time Entry"

                #display_attendance_info(data, current_date, current_time, entry_type)
                root.after(5000, back)
            else:
                msg = f'{data[0]} Your ID is De-Activated\nKindly contact admin'
                tkmb.showinfo(title='Information', message=msg)
                back()
        else:
            print("No qr code detected!")
            back()

    def deactivated(compare_number):
        rows_2 = wks4.get_all_values()
        return any(compare_number == daa[2] for daa in rows_2)

    def update_out_time(data, compare_number, current_date, current_time):
        rows = wks1.get_all_values()
        for row in rows[1:]:
            if row[2] == compare_number and row[4] == current_date:
                wks1.update_cell(rows.index(row) + 1, 7, current_time)
                msg = (f'\t\tOut-time Entry \n Name : \t\t{data[0]} \n Mobile Number : \t{data[1]} \n E-Mail Id : \t{data[2]} \n '
                       f'Emp. Id : \t{data[3]} \n Date : \t\t{current_date} \n In-Time : \t{row[5]} \n Out-Time : \t{current_time}')
                tkmb.showinfo(title='Information', message=msg)
                back()
                return True
        return False

    def add_attendance_entry(data, current_date, current_time, entry_type):
        wks1.append_row([data[0], data[1], data[2], data[3], current_date, current_time])
        entry_out_label = ctk.CTkLabel(entry_frame, ).place(x=0, y=0)
        msg = (f'\t\tIntime Entry \n Name : \t\t{data[0]} \n Mobile Number : \t{data[1]} \n E-Mail Id : \t{data[2]} \n '
               f'Emp. Id : \t{data[3]} \n Date : \t\t{current_date} \n In-Time : \t{current_time}')
        tkmb.showinfo(title='Information', message=msg)
        back()


    process_attendance(decoded_objects, current_date, current_time)
#----------------------Entry button function ends------------------#


#------------------------------User access button function starts--------------------------#
def remo():
    global remo_frame_2
    remo_frame = ctk.CTkFrame(frame_1, width=450, height=560).place(x=0, y=0)
    remo_frame_2 = ctk.CTkFrame(remo_frame, width=450, height=560).place(x=0, y=0)
    inphase_lable = ctk.CTkLabel(remo_frame_2, image=image, text="").place(x=100, y=20)
    lab_1= ctk.CTkLabel(remo_frame_2, text="ID Activation / Deactivation").place(x=150, y=200)
    lab_2= ctk.CTkLabel(remo_frame_2, text="Enter the Number").place(x=10, y=250)
    ent_1= ctk.CTkEntry(remo_frame_2, textvariable=l7, placeholder_text="Enter the mobile number....").place(x=170, y=251)
    ser_btn= ctk.CTkButton(remo_frame_2, text="Search", command=search).place(x=75,y=300)
    ba_btn= ctk.CTkButton(remo_frame_2, text="Back", command=back_1).place(x=250,y=300)
#------------------------------User access button function ends--------------------------#

#------------------------------------Search button function starts-----------------#
def search():
    global search
    ch_num = l7.get()
    cell = wks2.find(ch_num)
    row_index = cell.row
    row_values = wks2.row_values(row_index)
    l7.set("")
    # Display the row values in the result_label
    res_lab = ctk.CTkLabel(remo_frame_2, text=str(row_values[0]) + " " + str(row_values[1]) + " " + str(row_values[2]) + " " + str(row_values[3])).place(x=30, y=340)
    # Use lambda to pass arguments to deactivate
    deav_btn = ctk.CTkButton(remo_frame_2, text="Deactivate", command=lambda: deactivate(row_values)).place(x=75, y=380)
    av_btn = ctk.CTkButton(remo_frame_2, text="Activate", command=lambda: activate(row_values,ch_num)).place(x=200, y=380)
    return ch_num

# -------------------------------Activate button function starts------------------------------#
def activate(row_values, ch_num):
    # Open the workbook and the "Deactivated" worksheet
    workbook = gc.open("NSDC- Attendance")
    deactivated_sheet = workbook.worksheet("Deactivated")
    # Get all records from the "Deactivated" worksheet
    data = deactivated_sheet.get_all_records()
    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(data)
    # Ensure 'Mobile Number' column is treated as string
    df['Mobile Number'] = df['Mobile Number'].astype(str)
    # Convert ch_num to string
    ch_num = str(ch_num)
    # Find the rows that match the given mobile number
    search_results = df[df['Mobile Number'].str.contains(ch_num)]
    # Get the indices of the rows to delete
    indices_to_delete = search_results.index
    # Sort the indices in descending order
    indices_to_delete = sorted(indices_to_delete, reverse=True)
    # Delete the rows from the Google Sheets workbook using delete_rows
    for index in indices_to_delete:
        deactivated_sheet.delete_rows(index + 2)
    msg_1 = f'{row_values[0]}, is activated!'
    tkmb.showinfo(title='Information', message=msg_1)
    back_1()
# -------------------------------Activate button function ends------------------------------#

#-------------------------------De-Activate button function starts -------------------------------#
def deactivate(row_values):
    wks3 = gc.open("NSDC- Attendance").worksheet("Deactivated")
    wks3.append_row([row_values[0], row_values[1], row_values[2], row_values[3], current_date, current_time])
    wks3.format('A1:F1', {'textFormat': {'bold': True}})
    msg_1 = f'{row_values[0]},is deactivated !'
    tkmb.showinfo(title='Information', message=msg_1)
    back_1()
# -------------------------------De-Activate button function ends------------------------------#

#---------------------------------Log-Out button function starts-------------------------------#
def logout():
    msg_1 = f'Logut Successfull \n Thank you '
    tkmb.showinfo(title='Information', message=msg_1)
    back()

p = ""
l = tk.StringVar()
l2 = tk.StringVar()
l3 = tk.StringVar()
l4 = tk.StringVar()
l5 = tk.StringVar()
l6 = tk.StringVar()
l7 = tk.StringVar()

# Back button function with logout
def back():
    frame_1 = ctk.CTkFrame(root, width=450, height=560).place(x=0, y=0)
    inphase_lable = ctk.CTkLabel(frame_1, image=image, text="").place(x=100, y=20)
    Admin_btn = ctk.CTkButton(frame_1, text="Admin", command=admin_log).place(x=300, y=170)
    #register_btn = ctk.CTkButton(frame_1, text="Register", command=new_register, state="disabled").place(x=180, y=250)
    entry_btn =ctk.CTkButton(frame_1, text="Entry", command=entry).place(x=180, y=220)

# Back button function without log out
def back_1():
    frame_1 = ctk.CTkFrame(root, width=450, height=560).place(x=0, y=0)
    inphase_lable = ctk.CTkLabel(frame_1, image=image, text="").place(x=100, y=20)
    #Admin_btn = Button(frame_1, text="Admin", command=admin_log).place(x=180, y=200)
    register_btn = ctk.CTkButton(frame_1, text="Register", command=new_register, state="normal").place(x=180, y=250)
    entry_btn = ctk.CTkButton(frame_1, text="Entry", command=entry).place(x=180, y=300)
    logout_btn = ctk.CTkButton(frame_1, text="Logout", command=logout).place(x=180, y=350)
    de_user = ctk.CTkButton(frame_1, text="User Access", command=remo).place(x=180, y=400)
base2=open(r"message.txt")
read_2=base2.read()
a2=read_2.split(',')

# creating buttons
Admin_btn = ctk.CTkButton(frame_1, text="Admin", command=admin_log).place(x=300, y=170)
#register_btn = Button(frame_1, text="Register", command=new_register,state="disabled").place(x=180, y=250)
entry_btn = ctk.CTkButton(frame_1, text="Entry", command=entry).place(x=180, y=220)
# CLosed the window loop
root.mainloop()

#demo for praba - 9-Jan-2024