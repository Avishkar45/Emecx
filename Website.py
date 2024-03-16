import os
import re
import csv
import ssl
import smtplib
import tkinter as tk
from tkinter import ttk
from datetime import date
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from PIL import ImageTk, Image
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


#############################################################################################

## __________________________@@@@@@@@@@@@@@@@@@@Code By Omecx(परमं अस्तित्वम्)@@@@@@@@@@@@@@@@@@@__________________________##

html_value = ""
tp_file = None


# COMMON FUNCTIONS
def generate_preview(progress_bar, percent_label, email_sender, subject, body_entry, func):
    # Place the progress bar
    progress_bar.place(x=150, y=450, height=30, width=600)
    percent_label.place(x=760, y=450, height=30, width=100)

    # Get the input values from the user
    sender = email_sender.get()
    subject_line = subject.get()

    if isinstance(body_entry, str):
        body = body_entry
    elif hasattr(body_entry, 'get'):
        body = body_entry.get('1.0', 'end')
    else:
        body = "THE EMAIL BODY"

    # Create the message preview
    preview = f"From: {sender}\n"
    preview += f"To: <recipient email>\n"
    preview += f"Subject: {subject_line}\n\n"
    preview += f"{body}"

    # create a dummy root window
    root = tk.Tk()
    root.withdraw()

    # Display the preview in a new window
    preview_window = tk.Toplevel()
    preview_window.configure(bg="black")

    # get the screen width and height
    screen_width = preview_window.winfo_screenwidth()
    screen_height = preview_window.winfo_screenheight()

    # calculate the x and y position of the Toplevel window to center it
    x = (screen_width - preview_window.winfo_reqwidth()) / 2
    y = (screen_height - preview_window.winfo_reqheight()) / 2

    preview_window.title("Message Preview")
    preview_window.geometry("+%d+%d" % (x, y))

    preview_label = tk.Label(preview_window, text=preview, font=('orbitron', 13), bg="#2D548F")
    preview_label.pack(padx=20, pady=20)

    approve_button = tk.Button(preview_window, text="Approve", font=('orbitron', 13), command=lambda: (preview_window.destroy(), func()), bg="#677CC6")
    approve_button.pack(side="left", padx=20, pady=10)

    cancel_button = tk.Button(preview_window, text="Cancel", font=('orbitron', 13), command=lambda: preview_window.destroy(), bg="#677CC6")
    cancel_button.pack(side="right", padx=20, pady=10)

    # start the event loop
    preview_window.mainloop()

    root.destroy()


def load_dates(rep):
    """Loads The Dates for The Reminder Page"""
    try:
        with open(rep, 'r') as file:
            reader = csv.DictReader(file)
            dates = []
        for row in reader:
            due_date = row["due_date"]
            reminder_date = row["reminder_date"]
            dates.append((due_date, reminder_date))
    except Exception as e:
        print(e)
        return dates


def open_file():
    """Open a file dialog and return the selected file path."""
    root = tk.Tk()
    root.withdraw()
    try:
        rep = filedialog.askopenfilename(filetypes=[("Csv Files", "*.csv"), ("Template", "*.txt"), ("All files", "*")])
        if rep:
            return rep
        else:
            messagebox.showerror("Error", "No file selected.")
            return None
    except FileNotFoundError:
        messagebox.showerror("Error", "File not found.")
        return None
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return None


def browse_folder():
    """Open a folder dialog and return the selected folder path."""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory()
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
    return folder_path


def open_files():
    """Open a file dialog and return the selected file path."""
    root = tk.Tk()
    root.withdraw()
    try:
        rep = filedialog.askopenfilenames(filetypes=[("All files", "*")])
        if rep:
            return rep
        else:
            messagebox.showerror("Error", "No file selected.")
            return None
    except FileNotFoundError:
        messagebox.showerror("Error", "File not found.")
        return None
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return None


def read_template():
    """Read the email template from a file and return its contents and variables."""
    global tp_file

    try:
        if tp_file is None:
            messagebox.showinfo("TEMPLATE SELECTION", "Please Choose the Template File")
            rep = open_file()
        else:
            rep = tp_file

        with open(rep) as file:
            html = file.read()
            variables = re.findall(r"{(.*?)}", html)

            return html, variables
    except Exception as e:
        messagebox.showerror("ERROR FOR TEMPLATE FILE", str(e))


def find_attachment_file(attachment_filename, script_dir):
    # Check if attachment_filename is already an absolute path
    if os.path.isabs(attachment_filename):
        return attachment_filename

    # Check if attachment_filename is located in the  directory as the script (The selected Folder)
    attachment_path = os.path.join(script_dir, attachment_filename)
    if os.path.exists(attachment_path):
        return attachment_path

    # Search for attachment_filename in all subdirectories of the script directory
    for root, dirs, files in os.walk(script_dir):
        if attachment_filename in files:
            return os.path.join(root, attachment_filename)

    # Attachment file not found
    return None


def create_info_level(title, info):
    toplevel = tk.Toplevel()
    toplevel.title(title)
    
    # get the screen width and height
    screen_width = toplevel.winfo_screenwidth()
    screen_height = toplevel.winfo_screenheight()

    # calculate the x and y position of the root window to center it
    x = (screen_width - 400) / 2
    y = (screen_height - 400) / 2

    # Attributes OF the Window
    toplevel.geometry('400x400+%d+%d' % (x, y))
    toplevel.configure(bd=5, bg='#677CC6', borderwidth=1, highlightbackground="white")
    lbl_info = tk.Label(toplevel, text=info, height=5, width=20, font=('orbitron', 13), fg='white', bg="black",
                        bd=5)
    lbl_info.place(x=70, y=50)
    btn_cls = tk.Button(toplevel, text='OK', bd=5, height=1, width=5, font=('orbitron', 13), fg='#677CC6',
                        command=lambda: toplevel.destroy(), bg='black', relief='raised')
    btn_cls.place(x=170, y=250)
    toplevel.mainloop()


# The main App Class ($$Shizun$$)
class Emecx(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        # get the screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # calculate the x and y position of the root window to center it
        x = (screen_width - 1280) / 2
        y = (screen_height - 720) / 2

        # Attributes OF the Window
        self.geometry("1280x720+%d+%d" % (x, y))
        self.title("Emecx (Automated Email Sender)")
        self.resizable(width=True, height=True)

        # Image for the Toolbar Button
        img1 = Image.open("images/button (1).png")
        img1 = img1.resize((158, 53), Image.LANCZOS)
        self.photoimg01 = ImageTk.PhotoImage(img1)

        img2 = Image.open("images/button (2).png")
        img2 = img2.resize((158, 53), Image.LANCZOS)
        self.photoimg02 = ImageTk.PhotoImage(img2)

        img3 = Image.open("images/button (3).png")
        img3 = img3.resize((158, 53), Image.LANCZOS)
        self.photoimg03 = ImageTk.PhotoImage(img3)

        img4 = Image.open("images/button (4).png")
        img4 = img4.resize((158, 53), Image.LANCZOS)
        self.photoimg04 = ImageTk.PhotoImage(img4)

        img5 = Image.open("images/button (5).png")
        img5 = img5.resize((158, 53), Image.LANCZOS)
        self.photoimg05 = ImageTk.PhotoImage(img5)

        img6 = Image.open("images/button (6).png")
        img6 = img6.resize((158, 53), Image.LANCZOS)
        self.photoimg06 = ImageTk.PhotoImage(img6)

        img7 = Image.open("images/button (7).png")
        img7 = img7.resize((90, 28), Image.LANCZOS)
        self.photoimg07 = ImageTk.PhotoImage(img7)

        # Image for the SEND MAIL Button
        img8 = Image.open("images/button (8).png")
        img8 = img8.resize((142, 48), Image.LANCZOS)
        self.photoimg08 = ImageTk.PhotoImage(img8)

        # Image for the PAges Frame
        imgF1 = Image.open("images/Frame.png")
        imgF1 = imgF1.resize((1010, 530), Image.LANCZOS)
        self.photoimg1 = ImageTk.PhotoImage(imgF1)

        # Image for Toolbar Frame
        imgF2 = Image.open("images/Frame (1).png")
        imgF2 = imgF2.resize((210, 530), Image.LANCZOS)
        self.photoimg2 = ImageTk.PhotoImage(imgF2)

        # The IMage to be Used for the Container
        imgF3 = Image.open("images/Frame 1.png")
        imgF3 = imgF3.resize((1280, 720), Image.LANCZOS)
        self.photoimg3 = ImageTk.PhotoImage(imgF3)

        # The container for the PAges Frames
        container = tk.Frame(self)
        container.place(x=0, y=0, height=720, width=1280)

        # Style for progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Horizontal.TProgressbar", troughcolor="black", bordercolor="green", background="#677CC6",
                        lightcolor="white", darkcolor="blue")

        lbl_container = tk.Label(container, image=self.photoimg3)
        lbl_container.place(x=0, y=0, height=720, width=1280)

        # The Instance of ToolBAr Class (the frame Changer)
        t1 = Toolbar(container, self)
        t1.place(x=0, y=100, height=550, width=220)

        # Dictionary To contain The Frames
        self.frames = {}

        # Loop to Assign the KEy and Items to The Dictionary (KEy = name of the class of PAge)(Item = frame created)
        for F in (BulkPage, AttachmentPage, OptionPage, TemplatePage, ReminderPage, HelpPage):
            frame = F(container, self)

            self.frames[F] = frame

            frame.place(x=210, y=100, height=550, width=1030)

        self.change(BulkPage)

    # The Method for the Frame change (PAges)
    def change(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def runf(self, func, parent_frame):
        # Disable all buttons and entries in the parent frame
        for widget in parent_frame.winfo_children(self):
            if isinstance(widget, tk.Entry) or isinstance(widget, tk.Button):
                widget.config(state=tk.DISABLED)

        # Call the specified function
        func()

        # Enable all buttons and entries in the parent frame
        for widget in parent_frame.winfo_children(self):
            if isinstance(widget, tk.Entry) or isinstance(widget, tk.Button):
                widget.config(state=tk.NORMAL)

########################################################################################################################


# Our Toolbar Class
class Toolbar(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10

        self.place(x=0, y=100, height=550, width=220)

        lbl_toolbar = tk.Label(self, image=handler.photoimg2)
        lbl_toolbar.place(x=0, y=0, height=530, width=200)

        # Button 1(Home)
        home_btn = tk.Button(self, text='', image=handler.photoimg01, command=lambda: handler.change(BulkPage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=40)

        # Button 2(Bulk Emails)
        home_btn = tk.Button(self, text='', image=handler.photoimg02, command=lambda: handler.change(AttachmentPage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=120)

        # Button 3(Email With Attachments)
        home_btn = tk.Button(self, text='', image=handler.photoimg03, command=lambda: handler.change(OptionPage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=200)

        # Button 4(Email With Template)
        home_btn = tk.Button(self, text='', image=handler.photoimg04, command=lambda: handler.change(TemplatePage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=280)

        # Button 5(Schedule Email)
        home_btn = tk.Button(self, text='', image=handler.photoimg05, command=lambda: handler.change(ReminderPage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=360)

        # Button 6(Reminder Emails)
        home_btn = tk.Button(self, text='', image=handler.photoimg06, command=lambda: handler.change(HelpPage), borderwidth=0)
        home_btn.place(height=50, width=150, x=25, y=440)

########################################################################################################################


# The BulkPAge Class
class BulkPage(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10
        self.place(x=210, y=100, height=550, width=1030)

        # Creating the String Variables
        email_sender = tk.StringVar()
        email_pass = tk.StringVar()
        subject = tk.StringVar()
        # body_var = tk.StringVar()

        # The Frame Label and Title Label
        lbl_bulkpage = tk.Label(self, image=handler.photoimg1)
        lbl_bulkpage.place(x=0, y=0, height=530, width=1010)

        # The Frame Title
        lbl_page = tk.Label(self, text="BULK EMAILS", font=('orbitron', 15))
        lbl_page.place(x=305, y=10, height=50, width=400)

        # Email Label and Entry
        email_label = tk.Label(self, text="EMAIL", font=('orbitron', 13))
        email_label.place(x=30, y=80, height=30, width=130)

        email_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_sender)
        email_entry.place(x=180, y=80, height=30, width=300)

        # Password Label and Entry
        pass_label = tk.Label(self, text="PASSWORD", font=('orbitron', 13))
        pass_label.place(x=30, y=120, height=30, width=130)

        pass_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_pass)
        pass_entry.place(x=180, y=120, height=30, width=300)

        # Subject LAbel and Entry
        subject_label = tk.Label(self, text="SUBJECT", font=('orbitron', 13))
        subject_label.place(x=30, y=160, height=30, width=130)

        subject_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=subject)
        subject_entry.place(x=180, y=160, height=30, width=300)

        # Body LAbel and Entry
        body_label = tk.Label(self, text="BODY", font=('orbitron', 13))
        body_label.place(x=30, y=200, height=30, width=130)

        body_entry = tk.Text(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 10), wrap='word')
        body_entry.place(x=180, y=200, height=120, width=300)

        # Submit Button For Preview
        send_email = tk.Button(self, text='', image=handler.photoimg08, command=lambda: handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, body_entry, lambda: from_csv_simple()), BulkPage), borderwidth=0)
        send_email.place(x=200, y=350, height=50, width=150)

        # ProgressBar For the Email sending Process
        progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate",
                                       style="Horizontal.TProgressbar")

        percent_label = tk.Label(self, text="0%", font=("Arial", 12), fg="white", bg="#677CC6")

        # Function
        def from_csv_simple():
            counter = 0

            # Send the emails
            sender = email_sender.get()
            password = email_pass.get()
            subject_line = subject.get()
            text = body_entry.get('1.0', 'end')

            # The Csv File for Email
            messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
            rep = open_file()
            with open(rep, 'r', encoding='utf-8-sig') as file:
                line_count = sum(1 for _ in csv.reader(file)) - 1
                file.seek(0)
                reader = csv.reader(file)
                headers = next(reader)

                try:
                    for row in reader:
                        index = headers.index("Recipients")
                        recipient_email = row[index]

                        bulk_send(recipient_email, text, sender, password, subject_line)
                        counter += 1

                        progress_bar["value"] = round((counter / line_count) * 100)
                        progress_bar.update()
                        percent_label.config(text=f"{(progress_bar['value'])}%")
                        percent_label.update_idletasks()
                except Exception as e:
                    messagebox.showerror("Error\tAn error occurred.\n", str(e))
                else:
                    create_info_level("Email Sent", "Thank For USing Emecx")

        def bulk_send(receiver_email, body, sender_email, password, subject_line):
            try:
                # Start Creation of Message
                message = MIMEMultipart("alternative")
                message["Subject"] = subject_line
                message["From"] = sender_email
                message["To"] = receiver_email

                part1 = MIMEText(body, 'plain')

                # Add HTML/plain-text parts to MIMEMultipart message
                # The email client will try to render the last part first
                # message.attach(part1)
                message.attach(part1)

                # Create secure connection with server and send email
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                    server.login(sender_email, password)
                    body = message.as_string()
                    server.sendmail(sender_email, receiver_email, body)
            except Exception as e:
                print(e)
                messagebox.showerror("Error\tAn error occurred.\n", str(e))
            else:
                print("Job Done")


########################################################################################################################


# The AttachmentPage Class
class AttachmentPage(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10
        self.place(x=210, y=100, height=550, width=1030)

        # Creating the String Variables
        email_sender = tk.StringVar()
        email_pass = tk.StringVar()
        subject = tk.StringVar()

        # The Frame Label and Title Label
        lbl_bulkpage = tk.Label(self, image=handler.photoimg1)
        lbl_bulkpage.place(x=0, y=0, height=530, width=1010)

        # The Frame Title
        lbl_page = tk.Label(self, text="BULK ATTACHMENT", font=('orbitron', 15))
        lbl_page.place(x=305, y=10, height=50, width=400)

        # Email Label and Entry
        email_label = tk.Label(self, text="EMAIL", font=('orbitron', 13))
        email_label.place(x=30, y=80, height=30, width=130)

        email_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_sender)
        email_entry.place(x=180, y=80, height=30, width=300)

        # Password Label and Entry
        pass_label = tk.Label(self, text="PASSWORD", font=('orbitron', 13))
        pass_label.place(x=30, y=120, height=30, width=130)

        pass_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_pass)
        pass_entry.place(x=180, y=120, height=30, width=300)

        # Subject LAbel and Entry
        subject_label = tk.Label(self, text="SUBJECT", font=('orbitron', 13))
        subject_label.place(x=30, y=160, height=30, width=130)

        subject_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=subject)
        subject_entry.place(x=180, y=160, height=30, width=300)

        # Body LAbel and Entry
        body_label = tk.Label(self, text="BODY", font=('orbitron', 13))
        body_label.place(x=30, y=200, height=30, width=130)

        body_entry = tk.Text(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 10), wrap='word')
        body_entry.place(x=180, y=200, height=120, width=300)

        # Submit Button For Preview
        send_email = tk.Button(self, text='', image=handler.photoimg08, command=lambda: handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, body_entry, lambda: from_csv_simple()), OptionPage), borderwidth=0)
        send_email.place(x=200, y=350, height=50, width=150)

        # ProgressBar For the Email sending Process
        progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate",
                                       style="Horizontal.TProgressbar")

        percent_label = tk.Label(self, text="0%", font=("Arial", 12), fg="white", bg="#677CC6")

        # Function
        def from_csv_simple():
            counter = 0

            # Send the emails
            sender = email_sender.get()
            password = email_pass.get()
            subject_line = subject.get()
            text = body_entry.get('1.0', 'end')

            # The Csv File for Email
            messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
            rep = open_file()
            with open(rep, 'r', encoding='utf-8-sig') as file:
                line_count = sum(1 for _ in csv.reader(file)) - 1
                file.seek(0)
                reader = csv.reader(file)
                headers = next(reader)
                # Attachments to be selected
                messagebox.showinfo("Email Attachment/s", "Choose the Common Attachment/s")
                file_paths = open_files()

                try:
                    for row in reader:
                        index = headers.index("Recipients")
                        recipient_email = row[index]

                        bulk_send(recipient_email, text, sender, password, subject_line, file_paths)
                        counter += 1

                        progress_bar["value"] = round((counter / line_count) * 100)

                        progress_bar.update()
                        percent_label.config(text=f"{(progress_bar['value'])}%")
                        percent_label.update_idletasks()

                except Exception as e:
                    messagebox.showerror("Error\tAn error occurred.\n", str(e))
                else:
                    create_info_level("Email Sent", "Thank For USing Emecx")

        def bulk_send(receiver_email, body, sender_email, password, subject_line, file_paths):
            try:
                # Start Creation of Message
                message = MIMEMultipart("alternative")
                message["Subject"] = subject_line
                message["From"] = sender_email
                message["To"] = receiver_email

                part1 = MIMEText(body, 'plain')

                # Add HTML/plain-text parts to MIMEMultipart message
                # The email client will try to render the last part first
                # message.attach(part1)
                message.attach(part1)

                # For Multiple files
                for file_path in file_paths:
                    with open(file_path, 'rb') as file:
                        attachment = MIMEApplication(file.read(), _subtype=os.path.splitext(file_path)[1][1:])
                        attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                        message.attach(attachment)

                # Create secure connection with server and send email
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                    server.login(sender_email, password)
                    body = message.as_string()
                    server.sendmail(sender_email, receiver_email, body)
            except Exception as e:
                print(e)
            else:
                print("JOB DONE")

########################################################################################################################


class OptionPage(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10
        self['bg'] = 'black'

        self.place(x=210, y=100, height=550, width=1030)

        # Creating the String Variables
        email_sender = tk.StringVar()
        email_pass = tk.StringVar()
        subject = tk.StringVar()
        vart = tk.IntVar()

        # The Frame Label and Title Label
        lbl_optionPage = tk.Label(self, image=handler.photoimg1)
        lbl_optionPage.place(x=0, y=0, height=530, width=1010)

        # The Frame Title
        lbl_page = tk.Label(self, text="OPTION BODY", font=('orbitron', 15))
        lbl_page.place(x=305, y=10, height=50, width=400)

        # Email Label and Entry
        email_label = tk.Label(self, text="EMAIL", font=('orbitron', 13))
        email_label.place(x=30, y=80, height=30, width=130)

        email_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_sender)
        email_entry.place(x=180, y=80, height=30, width=300)

        # Password Label and Entry
        pass_label = tk.Label(self, text="PASSWORD", font=('orbitron', 13))
        pass_label.place(x=30, y=120, height=30, width=130)

        pass_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_pass)
        pass_entry.place(x=180, y=120, height=30, width=300)

        # Subject LAbel and Entry
        subject_label = tk.Label(self, text="SUBJECT", font=('orbitron', 13))
        subject_label.place(x=30, y=160, height=30, width=130)

        subject_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=subject)
        subject_entry.place(x=180, y=160, height=30, width=300)

        body_lbl = tk.Label(self, text='SOME EMAIL BODY OPTIONS', font=('orbitron', 13))
        body_lbl.place(x=200, y=200, height=30, width=400)

        # ProgressBar For the Email sending Process
        progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate", style="Horizontal.TProgressbar")

        percent_label = tk.Label(self, text="0%", font=("Arial", 12), fg="white", bg="#677CC6")

        # Function
        def from_csv_simple():
            counter = 0

            try:
                # Send the emails
                sender = email_sender.get()
                password = email_pass.get()
                subject_line = subject.get()
                text = body_sel()

                messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
                rep = open_file()

                # The File names
                messagebox.showinfo("Email Attachment", "Choose the Common Attachment")
                file_path = open_file()

                with open(rep, 'r', encoding='utf-8-sig') as file:
                    line_count = sum(1 for _ in csv.reader(file)) - 1
                    file.seek(0)
                    reader = csv.reader(file)
                    headers = next(reader)

                    for row in reader:
                        index = headers.index("Recipients")
                        recipient_email = row[index]

                        bulk_send(recipient_email, text, sender, password, subject_line, file_path)
                        counter += 1

                        progress_bar["value"] = round((counter / line_count) * 100)
                        progress_bar.update()
                        percent_label.config(text=f"{(progress_bar['value'])}%")
                        percent_label.update_idletasks()
            except Exception as e:
                messagebox.showerror("Error\tAn error occurred.\n", str(e))
            else:
                create_info_level("Email Sent", "Thank For USing Emecx")

        def bulk_send(receiver_email, body, sender_email, password, subject_line, file_path):
            try:
                # Start Creation of Message
                message = MIMEMultipart("alternative")
                message["Subject"] = subject_line
                message["From"] = sender_email
                message["To"] = receiver_email

                part1 = MIMEText(body, 'plain')

                # Add HTML/plain-text parts to MIMEMultipart message
                # The email client will try to render the last part first
                # message.attach(part1)
                message.attach(part1)

                with open(file_path, 'rb') as file:
                    attachment = MIMEApplication(file.read(), _subtype=os.path.splitext(file_path)[1][1:])
                    attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                    message.attach(attachment)

                # Create secure connection with server and send email
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                    server.login(sender_email, password)
                    body = message.as_string()
                    server.sendmail(sender_email, receiver_email, body)
            except Exception as e:
                print(e)

        def body_sel():
            var = vart.get()
            if var == 1:
                body = "Dear Customer, \n Due Unforeseen Circumstances we have to Shutdown our Service.\n We Will Deal with this as soon as possible.\n Sorry for the Inconvenience."
            elif var == 2:
                body = "Dear Parent, \nTomorrow we have Scheduled a Book Fair At your Wards School.\n Please Make Sure You visit.\n we have a lot of good books"
            elif var == 3:
                body = ''
            else:
                body = ""

            # Store the value of body in a global variable
            global body_value
            body_value = body

            return body

        R1 = tk.Radiobutton(self, text="Option 1", variable=vart, value=1, command=lambda: body_sel())
        R1.place(x=30, y=250, height=30, width=200)

        R2 = tk.Radiobutton(self, text="Option 2", variable=vart, value=2, command=lambda: body_sel())
        R2.place(x=250, y=250, height=30, width=200)

        R3 = tk.Radiobutton(self, text="Option 3", variable=vart, value=3, command=lambda: body_sel())
        R3.place(x=470, y=250, height=30, width=200)

        # Submit Button For Preview
        send_email = tk.Button(self, text='', image=handler.photoimg08, command=lambda: handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, body_value, lambda: from_csv_simple()), BulkPage), borderwidth=0)
        send_email.place(x=200, y=350, height=50, width=150)


########################################################################################################################

# The TemplatePage Class
class TemplatePage(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10
        self.place(x=210, y=100, height=550, width=1030)

        email_sender = tk.StringVar()
        email_pass = tk.StringVar()
        subject = tk.StringVar()

        # The Frame Label and Title Label
        lbl_templatePage = tk.Label(self, image=handler.photoimg1)
        lbl_templatePage.place(x=0, y=0, height=530, width=1010)

        # The Frame Title
        lbl_page = tk.Label(self, text="CUSTOM BODY", font=('orbitron', 15))
        lbl_page.place(x=305, y=10, height=50, width=400)

        # Email Label and Entry
        email_label = tk.Label(self, text="EMAIL", font=('orbitron', 13))
        email_label.place(x=30, y=80, height=30, width=130)

        email_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_sender)
        email_entry.place(x=180, y=80, height=30, width=300)

        # Password Label and Entry
        pass_label = tk.Label(self, text="PASSWORD", font=('orbitron', 13))
        pass_label.place(x=30, y=120, height=30, width=130)

        pass_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_pass)
        pass_entry.place(x=180, y=120, height=30, width=300)

        # Subject Label and Entry
        email_subject = tk.Label(self, text="SUBJECT", font=('orbitron', 13))
        email_subject.place(x=30, y=160, height=30, width=130)

        subject_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=subject)
        subject_entry.place(x=180, y=160, height=30, width=300)

        # Email Sending Button
        send_email = tk.Button(self, text='', image=handler.photoimg08, command=lambda: choice_send(), borderwidth=0)
        send_email.place(x=400, y=375, height=50, width=150)

        selected_button = tk.IntVar()
        selected_buttonx = tk.IntVar()

        body_var = tk.StringVar()

        lbl_body = tk.Label(self, text="EMAIL BODY FORMAT", font=('orbitron', 13), bd=3)
        lbl_body.place(x=200, y=200, height=30, width=400)

        collapse_radio1 = tk.Radiobutton(self, text="TEXT BODY", variable=selected_button, value=0, font=('orbitron', 10), bg='white')
        collapse_radio1.place(x=100, y=240, height=30, width=200)

        collapse_radio2 = tk.Radiobutton(self, text="BODY FROM TEMPLATE", variable=selected_button, value=1, font=('orbitron', 10), bg='white')
        collapse_radio2.place(x=350, y=240, height=30, width=200)

        btn_collapse_all = tk.Button(self, text='', image=handler.photoimg07, command=lambda: close_all(), bd=0)
        btn_collapse_all.place(x=600, y=240, height=30, width=90)

        lbl_attachment = tk.Label(self, text="ATTACHMENT TYPE", font=('orbitron', 13), bd=3)
        lbl_attachment.place(x=200, y=280, height=30, width=400)

        collapse_radio1 = tk.Radiobutton(self, text="COMMON ATTACHMENT", variable=selected_buttonx, value=0, font=('orbitron', 10), bg='white')
        collapse_radio1.place(x=100, y=320, height=30, width=200)

        collapse_radio2 = tk.Radiobutton(self, text="FILES FROM DIRECTORY", variable=selected_buttonx, value=1, font=('orbitron', 10), bg='white')
        collapse_radio2.place(x=350, y=320, height=30, width=200)

        # ProgressBar For the Email sending Process
        progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate",
                                       style="Horizontal.TProgressbar")

        percent_label = tk.Label(self, text="0%", font=("Arial", 12), fg="white", bg="#677CC6")

        # The Text Body Frame
        Bodyframe = tk.LabelFrame(self, bd=2)

        # The Template Body Frame
        Templateframe = tk.LabelFrame(self, bd=2)

        lbl_title = tk.Label(Bodyframe, text='TEXT BODY', font=('orbitron', 13))
        lbl_body = tk.Label(Bodyframe, text='BODY', font=('orbitron', 12), background='#ffffff')
        body_entry = tk.Text(Bodyframe)

        lbl_title_tp = tk.Label(Templateframe, text='TEMPLATE BODY', font=('orbitron', 12))
        lbl_file = tk.Label(Templateframe, text='CHOOSE THE TEMPLATE FILE', font=('orbitron', 12))
        file_name_tp = tk.Text(Templateframe, font=('orbitron', 12), wrap=tk.NONE)
        btn_file = tk.Button(Templateframe, text='', image=handler.photoimg07, bd=0, command=lambda: file_name())
        scrollbar = tk.Scrollbar(Templateframe, orient='horizontal', command=file_name_tp.xview)
        scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        file_name_tp.config(xscrollcommand=scrollbar.set)

        def file_name():
            file_name_tp.configure(state="normal")
            file_name_tp.delete('1.0', 'end')
            global tp_file
            tp_file = open_file()
            file_name_tp.insert('1.0', tp_file)
            file_name_tp.configure(state="disabled")
            file_name_tp.update_idletasks()

        def collapse_rise_frame(*args):
            if selected_button.get() == 0:
                if not Bodyframe.winfo_ismapped():
                    Bodyframe.place(x=30, y=280, height=200, width=600)
                Templateframe.place_forget()
            elif selected_button.get() == 1:
                if not Bodyframe.winfo_ismapped():
                    Templateframe.place(x=30, y=280, height=200, width=600)
                Bodyframe.place_forget()

        def close_all():
            for frame in [Bodyframe, Templateframe]:
                frame.place_forget()

        selected_button.trace_add("write", collapse_rise_frame)

        lbl_title.place(x=225, y=10, height=20, width=150)
        lbl_body.place(x=10, y=40, height=20, width=80)
        body_entry.place(x=100, y=40, height=140, width=475)

        lbl_title_tp.place(x=225, y=10, height=20, width=150)
        lbl_file.place(x=150, y=40, height=20, width=300)
        btn_file.place(x=255, y=80, height=30, width=90)
        file_name_tp.place(x=15, y=120, height=30, width=550)

        def choice_send():
            if selected_button.get() == 0:
                handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, body_entry, lambda: from_csv_send_body()), TemplatePage)
            elif selected_button.get() == 1:
                html, variables = read_template()
                handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, html, lambda: from_csv_send_template()), TemplatePage)

        def send_email_body(sender, password, recipient, subject_msg, body, attachment_path):
            """Send an email with the specified parameters."""
            message = MIMEMultipart("alternative")
            message["From"] = sender
            message["To"] = recipient
            message["Subject"] = subject_msg

            # if body option is selected then get body from the designated Stringvar()
            part1 = MIMEText(body, 'plain')
            message.attach(part1)

            # Add the attachment
            try:
                with open(attachment_path, 'rb') as file:
                    attachment = MIMEApplication(file.read(), _subtype="ext")
                    attachment.add_header('Content-Disposition', 'attachment',
                                          filename=os.path.basename(attachment_path))
                    message.attach(attachment)

                # Send the email
                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as server:
                    server.login(sender, password)
                    server.sendmail(sender, recipient, message.as_string())

            except Exception as e:
                print(e, 'ERROR')

        def from_csv_send_template():
            counter = 0

            executed = False

            messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
            rep = open_file()

            with open(rep, 'r', encoding='utf-8-sig') as file:
                line_count = sum(1 for _ in csv.reader(file)) - 1
                file.seek(0)
                reader = csv.reader(file)
                headers = next(reader)
                variable_dicts = []

                # Read the email template and variables
                html, variables = read_template()

                if selected_buttonx.get() == 1:
                    # Common folder for the files to Easier accessibility
                    messagebox.showinfo("Email Attachment", "Choose the Common Attachment Folder")
                    script_dir = browse_folder()

                # Send the emails
                sender = email_sender.get()
                password = email_pass.get()
                subject_line = subject.get()

                try:
                    for row in reader:
                        recipient_idx = headers.index("Recipients")
                        recipient = row[recipient_idx]

                        filename_idx = headers.index("filename")
                        filename = row[filename_idx]
                        row_dict = {}
                        for variable in variables:
                            if variable in headers:
                                index = headers.index(variable)
                                row_dict[variable] = row[index]
                        variable_dicts.append(row_dict)

                        # Single attachment file if option 3 is selected
                        if selected_buttonx.get() == 0 and not executed:
                            messagebox.showinfo("Email Attachment", "Choose the Common Attachment")
                            attachment_path = open_file()
                            executed = True
                        elif selected_buttonx.get() == 0 and executed:
                            pass
                        # if the option 4 is selected then we need the common directory and call the find file function
                        elif selected_buttonx.get() == 1:
                            attachment_path = find_attachment_file(filename, script_dir)
                        # Send email to each recipient
                        send_email_template(sender, password, recipient, subject_line, html, attachment_path,
                                            variable_dicts)
                        counter += 1
                        progress_bar["value"] = round((counter / line_count) * 100)
                        progress_bar.update()
                        percent_label.config(text=f"{(progress_bar['value'])}%")
                        percent_label.update_idletasks()
                except Exception as e:
                    messagebox.showerror("Error\tAn error occurred.\n", str(e))
                else:
                    create_info_level("Email Sent", "Thank For USing Emecx")

        def send_email_template(sender, password, recipient, subject_msg, body, attachment_path, variable_dicts):
            """Send an email with the specified parameters."""
            message = MIMEMultipart("alternative")
            message["From"] = sender
            message["To"] = recipient
            message["Subject"] = subject_msg

            # Replace the email body variables with their values
            for variable_dict in variable_dicts:
                for variable in variable_dict:
                    body = body.replace(f"{{{variable}}}", variable_dict[variable])
                part2 = MIMEText(body, 'html')

                message.attach(part2)
            variable_dicts.clear()

            # Add the attachment
            try:
                with open(attachment_path, 'rb') as file:
                    attachment = MIMEApplication(file.read(), _subtype="ext")
                    attachment.add_header('Content-Disposition', 'attachment',
                                          filename=os.path.basename(attachment_path))
                    message.attach(attachment)

                # Send the email
                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as server:
                    server.login(sender, password)
                    server.sendmail(sender, recipient, message.as_string())

            except Exception as e:
                messagebox.showerror("Error\tAn error occurred.\n", str(e))

        def from_csv_send_body():

            counter = 0

            executed = False

            messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
            rep = open_file()

            with open(rep, 'r', encoding='utf-8-sig') as file:
                line_count = sum(1 for _ in csv.reader(file)) - 1

                file.seek(0)
                reader = csv.reader(file)
                headers = next(reader)

                if selected_buttonx.get() == 1:
                    # Common folder for the files to Easier accessibility
                    messagebox.showinfo("Email Attachment", "Choose the Common Attachment Folder")
                    script_dir = browse_folder()

                # Send the emails
                sender = email_sender.get()
                password = email_pass.get()
                subject_line = subject.get()

                try:
                    for row in reader:
                        recipient_idx = headers.index("Recipients")
                        recipient = row[recipient_idx]

                        filename_idx = headers.index("filename")
                        filename = row[filename_idx]

                        # Single attachment file if option 3 is selected
                        if selected_buttonx.get() == 0 and not executed:
                            messagebox.showinfo("Email Attachment", "Choose the Common Attachment")
                            attachment_path = open_file()
                            executed = True
                        elif selected_buttonx.get() == 0 and executed:
                            pass
                        # if the option 4 is selected then we need the common directory and call the find file function
                        elif selected_buttonx.get() == 1:

                            attachment_path = find_attachment_file(filename, script_dir)

                        # Send email to each recipient
                        body = body_var.get()
                        send_email_body(sender, password, recipient, subject_line, body, attachment_path)

                        counter += 1

                        progress_bar["value"] = round((counter / line_count) * 100)
                        progress_bar.update()
                        percent_label.config(text=f"{(progress_bar['value'])}%")
                        percent_label.update_idletasks()
                except Exception as e:
                    messagebox.showerror("Error\tAn error occurred.\n", str(e))
                else:
                    create_info_level("Email Sent", "Thank For USing Emecx")


########################################################################################################################


# The ReminderPage Class
class ReminderPage(tk.LabelFrame):
    def __init__(self, parent, handler):
        tk.LabelFrame.__init__(self, parent)

        self['bd'] = 10
        self.place(x=210, y=100, height=550, width=1030)

        email_sender = tk.StringVar()
        email_pass = tk.StringVar()
        subject = tk.StringVar()

        # The Frame Label and Title Label
        lbl_reminderPage = tk.Label(self, image=handler.photoimg1)
        lbl_reminderPage.place(x=0, y=0, height=530, width=1010)

        # The Frame Title
        lbl_page = tk.Label(self, text="REMAINDER EMAILS", font=('orbitron', 15))
        lbl_page.place(x=305, y=10, height=50, width=400)

        # Email Label and Entry
        email_label = tk.Label(self, text="EMAIL", font=('orbitron', 13))
        email_label.place(x=30, y=80, height=30, width=130)

        email_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_sender)
        email_entry.place(x=180, y=80, height=30, width=300)

        # Password Label and Entry
        pass_label = tk.Label(self, text="PASSWORD", font=('orbitron', 13))
        pass_label.place(x=30, y=120, height=30, width=130)

        pass_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=email_pass)
        pass_entry.place(x=180, y=120, height=30, width=300)

        # Subject Label and Entry
        email_subject = tk.Label(self, text="SUBJECT", font=('orbitron', 13))
        email_subject.place(x=30, y=160, height=30, width=130)

        subject_entry = tk.Entry(self, bg="#D9D9D9", fg="#000716", font=('orbitron', 13), textvariable=subject)
        subject_entry.place(x=180, y=160, height=30, width=300)

        # Email Sending Button
        send_email = tk.Button(self, text='', image=handler.photoimg08, command=lambda: handler.runf(lambda: generate_preview(progress_bar, percent_label, email_sender, subject, html_value, lambda: from_csv_send_reminder()), ReminderPage), borderwidth=0)
        send_email.place(x=400, y=375, height=50, width=150)

        # ProgressBar For the Email sending Process
        progress_bar = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate",
                                       style="Horizontal.TProgressbar")

        percent_label = tk.Label(self, text="0%", font=("Arial", 12), fg="white", bg="#677CC6")

        def send_email_template(sender, password, recipient, subject_msg, body, attachment_path, variable_dicts):
            """Send an email with the specified parameters."""
            message = MIMEMultipart("alternative")
            message["From"] = sender
            message["To"] = recipient
            message["Subject"] = subject_msg

            # Replace the email body variables with their values
            for variable_dict in variable_dicts:
                for variable in variable_dict:
                    body = body.replace(f"{{{variable}}}", variable_dict[variable])
                part2 = MIMEText(body, 'html')

                message.attach(part2)
            variable_dicts.clear()

            # Add the attachment
            try:
                with open(attachment_path, 'rb') as file:
                    attachment = MIMEApplication(file.read(), _subtype="ext")
                    attachment.add_header('Content-Disposition', 'attachment',
                                          filename=os.path.basename(attachment_path))
                    message.attach(attachment)

                # Send the email
                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as server:
                    server.login(sender, password)
                    server.sendmail(sender, recipient, message.as_string())

            except Exception as e:
                messagebox.showerror("Error\tAn error occurred.\n", str(e))

        def from_csv_send_reminder():
            counter = 0

            messagebox.showinfo("DATASHEET SELECTION", "Please Choose the Csv File")
            rep = open_file()

            with open(rep, 'r', encoding='utf-8-sig') as file:
                line_count = sum(1 for _ in csv.reader(file)) - 1

                file.seek(0)
                reader = csv.reader(file)
                headers = next(reader)

                variable_dicts = []

                present = date.today()

                global tp_file
                tp_file = None

                # Read the email template and variables
                html, variables = read_template()

                # Common folder for the files to Easier accessibility
                messagebox.showinfo("Email Attachment", "Choose the Common Attachment Folder")
                script_dir = browse_folder()

                # Send the emails
                sender = email_sender.get()
                password = email_pass.get()
                subject_line = subject.get()

                try:
                    for row in reader:
                        recipient_idx = headers.index("Recipients")
                        recipient = row[recipient_idx]

                        filename_idx = headers.index("filename")
                        filename = row[filename_idx]

                        reminder_date_idx = headers.index("reminder_date")
                        reminder_date_string = row[reminder_date_idx]

                        reminder_date = datetime.strptime(reminder_date_string, '%d-%m-%Y').date()

                        status_idx = headers.index("status")
                        status = row[status_idx]

                        row_dict = {}
                        for variable in variables:
                            if variable in headers:
                                index = headers.index(variable)
                                row_dict[variable] = row[index]
                        variable_dicts.append(row_dict)

                        attachment_path = find_attachment_file(filename, script_dir)
                        if (present >= reminder_date) and (status == 'N'):
                            # Send email to each recipient
                            send_email_template(sender, password, recipient, subject_line, html, attachment_path,
                                                variable_dicts)

                            counter += 1

                            progress_bar["value"] = round((counter / line_count) * 100)
                            progress_bar.update()
                            percent_label.config(text=f"{(progress_bar['value'])}%")
                            percent_label.update_idletasks()
                        else:
                            line_count -= 1
                            progress_bar["value"] = round((counter / line_count) * 100)
                            progress_bar.update()
                            percent_label.config(text=f"{(progress_bar['value'])}%")
                            percent_label.update_idletasks()
                            pass
                except Exception as e:
                    messagebox.showerror("Error\tAn error occurred.\n", str(e))
                else:
                    create_info_level("Email Sent", "Thank For USing Emecx")

########################################################################################################################


class HelpPage(tk.Frame):
    def __init__(self, parent, handler):
        tk.Frame.__init__(self, parent)

        self['bd'] = 10
        self.place(x=210, y=100, height=550, width=1030)

        # Create a Canvas widget
        canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        canvas.pack(side="left", fill=tk.BOTH, expand=True)

        # Add a Scrollbar widget for the canvas
        scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side="right", fill=tk.Y)

        # Configure the canvas to use the scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Create a frame inside the canvas for all the widgets
        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor='nw')

        imgc = Image.open("images/Frame (c).png")
        self.photoimgc = ImageTk.PhotoImage(imgc)

        lbl = tk.Label(frame, image=self.photoimgc)
        lbl.place(x=0, y=0, relwidth=1, relheight=1)

        # Generate App Password for Google Authentication
        label1 = tk.Label(frame, text="1. Generate App Password for Google Authentication", font=('orbitron', 15), bg="#2D548F")
        label1.grid(padx=10, pady=10, row=0, column=0)
        text1 = tk.Text(frame, height=10, width=60, font=('orbitron', 14), wrap='word')
        text1.insert(tk.END,
                     "To generate an App Password for Google Authentication, follow these steps:\n\n1. Go to your Google Account settings.\n2. In the Security section, click on 'App Passwords'.\n3. Select the app and device you want to generate the password for.\n4. Follow the instructions to generate and use the App Password.")
        text1.configure(state='disabled')
        text1.grid(padx=10, pady=10, row=1, column=0)

        # Designing of CSV sheet
        label2 = tk.Label(frame, text="2. Designing of CSV sheet", font=('orbitron', 15), bg="#2D548F")
        label2.grid(padx=10, pady=10, row=2, column=0)
        text2 = tk.Text(frame, height=10, width=60, font=('orbitron', 14), wrap='word')
        text2.insert(tk.END,
                     "To design a CSV sheet with comma delimited format and recipient emails as titles, follow these steps:\n\n1. Open a new spreadsheet in Microsoft Excel or Google Sheets.\n2. In the first row, enter the following titles for each column: 'Recipient Emails', 'Subject', 'Body'.\n3. Add your email recipients' addresses in the 'Recipient Emails' column, and your desired email subject and body in the 'Subject' and 'Body' columns, respectively.\n4. Save the file as a CSV file, and make sure to select the option to save it in comma delimited format.")
        text2.configure(state='disabled')
        text2.grid(padx=10, pady=10, row=3, column=0)

        # Template file of HTML
        label3 = tk.Label(frame, text="3. Template file of HTML", font=('orbitron', 15), bg="#2D548F")
        label3.grid(padx=10, pady=10, row=4, column=0)
        text3 = tk.Text(frame, height=10, width=60, font=('orbitron', 14), wrap='word')
        text3.insert(tk.END,
                     "To create a template file of HTML for your email, follow these steps:\n\n1. Open a text editor or HTML editor.\n2. Start with the basic HTML structure, including <html>, <head>, and <body> tags.\n3. Add any necessary CSS styling to the <head> section.\n4. Add the email content you want to include in the <body> section.\n5. Save the file with a .html extension.")
        text3.configure(state='disabled')
        text3.grid(padx=10, pady=10, row=5, column=0)

        label3 = tk.Label(frame, text="CONTACT: raidenrage456@gmail.com\n MOBILE: 900000000", font=('orbitron', 15), bg="#2D548F")
        label3.grid(padx=10, pady=10, row=6, column=0)

        label3 = tk.Label(frame, text="CONTACT: avishkardhembare45@gmail.com\n MOBILE: 900000000", font=('orbitron', 15), bg="#2D548F")
        label3.grid(padx=10, pady=10, row=7, column=0)


if __name__ == '__main__':
    Omecx_app = Emecx()
    Omecx_app.mainloop()
