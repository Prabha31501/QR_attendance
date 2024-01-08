# Libraries for this project
import PySimpleGUI as sg
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from tkinter import *
from PIL import ImageTk, Image
import qrcode
import cv2
from pyzbar import pyzbar
from pyzbar.pyzbar import decode
from datetime import datetime
import time
import pandas as pd
import gspread
from tkinter.messagebox import showinfo
import numpy as np
#----------Library files for Generating pdf------------------#
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage, Spacer, Paragraph
from tkinter import PhotoImage
from reportlab.lib.styles import getSampleStyleSheet
import re

#-----------------------------------Creating the window-------------------------#
root = Tk()
root.title("Inphase Power Technologies")
root.geometry("450x560")
root.resizable(False,False)
frame_1 = Frame(root, width=450, height=560, bg="white")
frame_1.place(x=0, y=0)
icon = Image.open('pk.ico')
photo = ImageTk.PhotoImage(icon)
root.wm_iconphoto(False, photo) #This line is used to change the icon
root.iconbitmap('pk.ico')

#Impoting the  images
inphase_logo = ImageTk.PhotoImage(Image.open("test.png"))

#Inphase logo attached
inphase_lable=Label(frame_1,image=inphase_logo,bg="White").place(x=100,y=20)


# ------------------------------------------------- Open Google spread sheet ---------------------------------------
gc = gspread.service_account(filename='nsdc.json')
spreadsheet = gc.open("NSDC- Attendance")
wks1 = spreadsheet.get_worksheet(0)
wks2 = spreadsheet.get_worksheet(1)
wks3 = spreadsheet.get_worksheet(2)
wks4 = spreadsheet.get_worksheet(3)


# ------------------------------------------------- Current date and time -------------------------
# Get the current date and time
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
current_date = now.strftime("%Y-%m-%d")


#button functions
def admin_log():
    global pass_frame
    pass_frame = Frame(frame_1, width=450, height=560, bg="white").place(x=0, y=0)
    pass_frame_2 = Frame(pass_frame, width=450, height=560, bg="white").place(x=0, y=0)
    admin_label = Label(pass_frame_2, image=inphase_logo, text="bg", bg="White").place(x=100, y=20)
    wecome = Label(pass_frame_2, text="Admin login", bg="white",fg="red", font=('Timesnewroman', 16, "bold")).place(x=150, y=180)
    user_label= Label(pass_frame_2, text="Enter the user name", bg="white", font=('Timesnewroman', 10, "bold")).place(x=37,y=220)
    pass_label = Label(pass_frame_2, text="Enter the password  ",bg="white",font=('Timesnewroman', 10, "bold")).place(x=37, y=262)
    user_entry= Entry(pass_frame_2, textvariable=l,font=('Timesnewroman', 11, "")).place(x=200, y=221)
    pass_entry = Entry(pass_frame_2, show="*", textvariable=l2, font=('Timesnewroman', 11, "")).place(x=200, y=263)
    ent_btn = Button(pass_frame_2, text="Login",bg="White", command=enter, relief=RIDGE).place(x=188, y=320)
    back_btn = Button(pass_frame_2, text="Back",command=back, bg="White", relief=RIDGE).place(x=278, y=320)

# New register button function
def new_register():
    reg_frame= Frame(frame_1, width=450, height=560, bg="white").place(x=0, y=0)
    reg_frame_2 = Frame(reg_frame, width=450, height=560, bg="white").place(x=0, y=0)
    in_lab= Label(reg_frame_2, image=inphase_logo, text="bg", bg="White").place(x=100, y=20)
    nr_lab= Label(reg_frame_2, text="New Register", bg="white", fg="red", font=('Timesnewroman', 16, "bold")).place(x=150, y=180)

    #Labels for input
    name_lab= Label(reg_frame_2, text="Enter the name", bg="white", font=('Timesnewroman', 10, "bold")).place(x=37,y=220)
    email_lab = Label(reg_frame_2, text="Enter the Email", bg="white", font=('Timesnewroman', 10, "bold")).place(x=37,y=260)
    number_lab = Label(reg_frame_2, text="Enter the number", bg="white", font=('Timesnewroman', 10, "bold")).place(x=37,y=300)
    empid_lab = Label(reg_frame_2, text="Enter the Emp.Id", bg="white", font=('Timesnewroman', 10, "bold")).place(x=37,y=340)

    #user input entry to collect the datas
    name_entry = Entry(reg_frame_2, textvariable=l3, font=('Timesnewroman', 11, "")).place(x=200, y=221)
    email_entry = Entry(reg_frame_2, textvariable=l4, font=('Timesnewroman', 11, "")).place(x=200, y=261)
    number_entry = Entry(reg_frame_2, textvariable=l5, font=('Timesnewroman', 11, "")).place(x=200, y=301)
    empid_entry = Entry(reg_frame_2, textvariable=l6, font=('Timesnewroman', 11, "")).place(x=200, y=341)

    create_btn= Button(reg_frame_2, text="Create", command=create, bg="white", relief=RIDGE).place(x=240, y=381)
    back_btn = Button(reg_frame_2, text="Back", command=back_1, bg="White", relief=RIDGE).place(x=290, y=381)

# Create button function
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
    doc = SimpleDocTemplate(file_path_pdf, pagesize=A4)
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

    # ------------------------------------------------- Store new registered data in Gsheet ------------------
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
    msg = f'QR code Sent Successfull to Registered Email \n Check the inbox {text2}'
    showinfo(title='Information', message=msg)

    l3.set("")
    l4.set("")
    l5.set("")
    l6.set("")
    new_register()
    #create_btn = Button(reg_frame_2, text="Create", command=create, bg="white", relief=RIDGE).place(x=240, y=381)
    #back_btn = Button(reg_frame_2, text="Back", command=back, bg="White", relief=RIDGE).place(x=290, y=381)

# User entry button function
def entry():
    entry_frame = Frame(frame_1, width=450, height=560, bg="White")
    entry_frame.place(x=0, y=0)
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    current_date = now.strftime("%Y-%m-%d")
    gc = gspread.service_account('nsdc.json')
    spreadsheet = gc.open("NSDC- Attendance")
    wks1 = spreadsheet.get_worksheet(0)
    # Initialize webcam
    cap = cv2.VideoCapture(0)
    #cv2.namedWindow("Camera Feed", cv2.WINDOW_NORMAL)

    # time delay 5 seconds
    # time.sleep(5)
    # Capture frame
    _, frame = cap.read()
    #cv2.imshow("Camera Feed", frame)  # Display the frame in the "Camera Feed" window
    #cv2.waitKey(5)

    cv2.destroyAllWindows()

    # Find QR codes in the frame
    decoded_objects = pyzbar.decode(frame)
    # Check if any QR codes were found
    if decoded_objects:
        data1 = decoded_objects[0].data.decode("utf-8")
        data = data1.split(',')
        compare_number = str(data[2])
        compare_date = str(current_date)
        rows = wks1.get_all_values() # Attendance sheet
        rows_2=wks4.get_all_values() # Deactivated sheet
        rows_3=wks2.get_all_values() # Database Sheet
        for daa in rows_2:
            fin=re.findall(daa[2],compare_number)
            #print(fin)
        if not fin:
            #print("Not in fin")
                #print(daa[2])
            for row in rows[1:]:
                #print(row)
                if row[2] == compare_number and row[4] == compare_date:
                    print(row[2], row[4])
                    print("Number and Date matched!")
                    OUT_TIME = current_time
                    wks = gc.open("NSDC- Attendance").worksheet("Attendance")
                    wks1.update_cell(rows.index(row) + 1, 7, OUT_TIME)
                    wks1.format('A1:G1', {'textFormat': {'bold': True}})
                    entry_out_label = Label(entry_frame).place(x=0, y=0)
                    Entry_detail = Label(entry_frame, text="Today Out-Time Entry", bg="white",font=('Timesnewroman', 20, 'bold')).place(x=118, y=40)
                    name_label = Label(entry_frame, text="Name  : " + data[0], bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=100)
                    number_label = Label(entry_frame, text="E-Mail  : " + data[1], bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=140)
                    mail_label = Label(entry_frame, text="Mobile  : " + data[2], bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=180)
                    emp_labble = Label(entry_frame, text="Emp.Id  :" + data[3], bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=220)
                    date_label = Label(entry_frame, text="Date  : " + current_date, bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=260)
                    In_label = Label(entry_frame, text="In - Time  : " + row[4], bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=300)
                    Out_label = Label(entry_frame, text="Out - Time  : " + current_time, bg="white",font=('Timesnewroman', 14, 'bold')).place(x=40, y=340)
                    root.after(5000, back)
                    break
                else:
                    print(row[2], row[4])
                    print("Number and Date not matched!")
                    # store the qr values and in-time
                    wks = gc.open("NSDC- Attendance").worksheet("Attendance")
                    #  Append Data from qr code
                    wks1.append_row([data[0], data[1], data[2], data[3], current_date, current_time])
                    wks1.format('A1:G1', {'textFormat': {'bold': True}})
                    entry_out_label = Label(entry_frame, ).place(x=0, y=0)
                    Entry_detail = Label(entry_frame, text="Today In-Time Entry", bg="white",font=('Timesnewroman', 20, 'bold')).place(x=118, y=40)
                    name_label = Label(entry_frame, text="Name  : " + data[0], bg="white",font=('Timesnewroman', 12, 'bold')).place(x=40, y=100)
                    number_label = Label(entry_frame, text="Mobile number  : " + data[1], bg="white",font=('Timesnewroman', 12, 'bold')).place(x=40, y=140)
                    mail_label = Label(entry_frame, text="E-Mail  : " + data[2], bg="white",font=('Timesnewroman', 12, 'bold')).place(x=40, y=180)
                    emp_labble = Label(entry_frame, text="Emp.Id  :" + data[3], bg="white", font=('Timesnewroman', 12, 'bold')).place(x=40, y=220)
                    date_label = Label(entry_frame, text="Date  : " + current_date, bg="white",font=('Timesnewroman', 12, 'bold')).place(x=40, y=260)
                    In_label = Label(entry_frame, text="In - Time  : " + current_time, bg="white",font=('Timesnewroman', 12, 'bold')).place(x=40, y=300)
                    root.after(5000, back)
                break
        else:
            msg = f'{daa[0]} Your Id is De-Activated \n Kindly contact admin'
            showinfo(title='Information', message=msg)
            back()

    else:
        a = back()

    # Release webcam
    cap.release()

# password enter button function
def enter():
    global enter
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
        showinfo(title='Information',message=msg)

        frame_1 = Frame(root, width=450, height=560, bg="white").place(x=0, y=0)
        inphase_lable = Label(frame_1, image=inphase_logo, bg="White").place(x=100, y=20)
        Admin_btn = Button(frame_1, text="Admin", command=admin_log, bg="White", relief=RIDGE).place(x=180, y=200)
        register_btn = Button(frame_1, text="Register", command=new_register, bg="White", relief=RIDGE,state="normal").place(x=180, y=250)
        entry_btn = Button(frame_1, text="Entry", command=entry, bg="White", relief=RIDGE).place(x=180, y=300)
        logout_btn= Button(frame_1, text="Logout", command=logout,bg="White", relief=RIDGE).place(x=180, y=350)
        de_user= Button(frame_1, text="User Access", command=remo,bg="white", relief=RIDGE).place(x=180, y=400)
    else:
        msg = f'Login failed! \n Enter the valid user name and password'
        showinfo(title='Information', message=msg)

# User access button function
def remo():
    global remo_frame_2
    remo_frame = Frame(frame_1, width=450, height=560, bg="white").place(x=0, y=0)
    remo_frame_2 = Frame(remo_frame, width=450, height=560, bg="white").place(x=0, y=0)
    inphase_lable = Label(remo_frame_2, image=inphase_logo, bg="White").place(x=100, y=20)
    lab_1= Label(remo_frame_2, text="ID Activation / Deactivation", bg="white", font=('Timesnewroman', 14, 'bold')).place(x=50, y=150)
    lab_2= Label(remo_frame_2, text="Enter the Number", bg="white", font=('Timesnewroman',10,'bold')).place(x=10, y=220)
    ent_1= Entry(remo_frame_2, text="Number", textvariable=l7, bg="white", font=('Timesnewroman',10,'bold')).place(x=170, y=221)
    ser_btn= Button(remo_frame_2, text="Search", command=search,bg="white", font=('Timesnewroman', 10 ,'bold')).place(x=170,y=270)
    ba_btn= Button(remo_frame_2, text="Back", command=back_1,bg="white", font=('Timesnewroman', 10 ,'bold')).place(x=240,y=270)

# Search button
def search():
    global search
    ch_num = l7.get()
    cell = wks2.find(ch_num)
    row_index = cell.row
    row_values = wks2.row_values(row_index)
    l7.set("")
    # Display the row values in the result_label
    res_lab = Label(remo_frame_2, text=str(row_values[0]) + " " + str(row_values[1]) + " " + str(row_values[2]) + " " + str(row_values[3]), bg="white").place(x=0, y=340)
    # Use lambda to pass arguments to deactivate
    deav_btn = Button(remo_frame_2, text="Deactivate", bg="white", command=lambda: deactivate(row_values)).place(x=160, y=380)
    av_btn = Button(remo_frame_2, text="Activate", bg="white", command=lambda: activate(row_values,ch_num)).place(x=300, y=380)
    return ch_num

# Activate button function
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
    showinfo(title='Information', message=msg_1)
    back_1()


# De-Activate button function
def deactivate(row_values):
    wks3 = gc.open("NSDC- Attendance").worksheet("Deactivated")
    wks3.append_row([row_values[0], row_values[1], row_values[2], row_values[3], current_date, current_time])
    wks3.format('A1:F1', {'textFormat': {'bold': True}})
    msg_1 = f'{row_values[0]},is deactivated !'
    showinfo(title='Information', message=msg_1)
    back_1()
# Log-Out button function
def logout():
    msg_1 = f'Logut Successfull \n Thank you '
    showinfo(title='Information', message=msg_1)
    back()

p = ""
l = StringVar()
l2 = StringVar()
l3 = StringVar()
l4 = StringVar()
l5 = StringVar()
l6 = StringVar()
l7 = StringVar()

# Back button function with logout
def back():
    frame_1 = Frame(root, width=450, height=560, bg="white").place(x=0, y=0)
    inphase_lable = Label(frame_1, image=inphase_logo, bg="White").place(x=100, y=20)
    Admin_btn = Button(frame_1, text="Admin", command=admin_log, bg="White", relief=RIDGE).place(x=180, y=200)
    register_btn = Button(frame_1, text="Register", command=new_register, bg="White", relief=RIDGE, state="disabled").place(x=180, y=250)
    entry_btn = Button(frame_1, text="Entry", command=entry, bg="White", relief=RIDGE).place(x=180, y=300)

# Back button function without log out
def back_1():
    frame_1 = Frame(root, width=450, height=560, bg="white").place(x=0, y=0)
    inphase_lable = Label(frame_1, image=inphase_logo, bg="White").place(x=100, y=20)
    Admin_btn = Button(frame_1, text="Admin", command=admin_log, bg="White", relief=RIDGE).place(x=180, y=200)
    register_btn = Button(frame_1, text="Register", command=new_register, bg="White", relief=RIDGE, state="normal").place(x=180, y=250)
    entry_btn = Button(frame_1, text="Entry", command=entry, bg="White", relief=RIDGE).place(x=180, y=300)
    logout_btn = Button(frame_1, text="Logout", command=logout, bg="White", relief=RIDGE).place(x=180, y=350)
    de_user = Button(frame_1, text="User Access", command=remo, bg="white", relief=RIDGE).place(x=180, y=400)
base2=open(r"message.txt")
read_2=base2.read()
a2=read_2.split(',')

# creating buttons
Admin_btn = Button(frame_1, text="Admin", command=admin_log, bg="White", relief=RIDGE).place(x=180, y=200)
register_btn = Button(frame_1, text="Register", command=new_register, bg="White", relief=RIDGE,state="disabled").place(x=180, y=250)
entry_btn = Button(frame_1, text="Entry", command=entry, bg="White", relief=RIDGE).place(x=180, y=300)
# CLosed the window loop
root.mainloop()