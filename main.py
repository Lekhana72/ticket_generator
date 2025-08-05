import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import pandas as pd
import yagmail
import os
import sys







# === CONFIG ===
TEMPLATE_IMAGE = "background.png"  # Optional background (leave blank if none)
FONT_PATH = "Pixelscapes.ttf"      # Make sure this font file is in the same folder
TICKETS_FOLDER = "tickets"
SENDER_EMAIL = "flashbeee.2k25@gmail.com"
APP_PASSWORD = "pzof pruq hdop iiim"

# === Validations BEFORE GUI starts ===
if not os.path.exists(FONT_PATH):
    messagebox.showerror("Error", f"Font file '{FONT_PATH}' not found.")
    sys.exit()

if TEMPLATE_IMAGE and not os.path.exists(TEMPLATE_IMAGE):
    messagebox.showerror("Error", f"Background image '{TEMPLATE_IMAGE}' not found.")
    sys.exit()

if not SENDER_EMAIL or not APP_PASSWORD:
    messagebox.showerror("Error", "Please set your email and app password.")
    sys.exit()

os.makedirs(TICKETS_FOLDER, exist_ok=True)

# === Start GUI ===
root = tk.Tk()
root.title("Farewell Invitation Sender")
root.geometry("800x750")

# === Background ===
if os.path.exists(TEMPLATE_IMAGE):
    bg = Image.open(TEMPLATE_IMAGE).resize((800, 750))
    bg_img = ImageTk.PhotoImage(bg)
    bg_label = tk.Label(root, image=bg_img)
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# === Input Fields ===
labels = {}
entries = {}
fields = ["Venue", "Date", "Time", "Address", "Welcome Message"]
y_positions = [350, 420, 490, 560, 630]

for i, field in enumerate(fields):
    labels[field] = tk.Label(root, text=f"{field}:", bg="white", font=("Comic Sans MS", 14, "bold"))
    labels[field].place(x=100, y=y_positions[i])
    entries[field] = tk.Entry(root, width=40, font=("Comic Sans MS", 14))
    entries[field].place(x=250, y=y_positions[i])

# === Excel Upload ===
EXCEL_FILE = ""

def upload_file():
    global EXCEL_FILE
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file:
        EXCEL_FILE = file
        file_label.config(text=os.path.basename(file))

upload_btn = tk.Button(root, text="Upload Excel", command=upload_file, bg="#FFB6C1", font=("Comic Sans MS", 10, "bold"))
upload_btn.place(x=100, y=300)

file_label = tk.Label(root, text="No file selected", bg="white", font=("Comic Sans MS", 10))
file_label.place(x=250, y=300)

# === Status Label ===
status_label = tk.Label(root, text="Status: Waiting", bg="white", font=("Comic Sans MS", 12), fg="green")
status_label.place(x=100, y=250)

# === Generate Ticket ===
def generate_ticket(name):
    img = Image.new("RGB", (600, 900), "#ffb6c1")  # pink background
    draw = ImageDraw.Draw(img)

    try:
        font_title = ImageFont.truetype(FONT_PATH, 40)
        font_normal = ImageFont.truetype(FONT_PATH, 24)
        font_small = ImageFont.truetype(FONT_PATH, 16)
    except:
        font_title = font_normal = font_small = ImageFont.load_default()

    # Header
    draw.text((165, 60), "INVITATION", fill="black", font=font_title)

    # Categories
    categories = ["games", "lunch", "musical", "DJ", "etc"]
    for i, cat in enumerate(categories):
        x = 50 + i * 100  # reduced spacing between categories
        draw.text((x, 110), cat, fill="black", font=font_small)
        draw.rectangle([x, 135, x + 12, 148], outline="black", width=1)  # smaller rectangle


    draw.line((50, 170, 550, 170), fill="black", width=1)


    draw.text((50, 200), f"Event : FLASHBEEE", fill="black", font=font_small)
    draw.text((50, 225), f"Name : {name}", fill="black", font=font_small)
    draw.text((50, 250), f"Venue : {entries['Venue'].get()}", fill="black", font=font_small)
    draw.text((50, 275), f"Date : {entries['Date'].get()} at {entries['Time'].get()}", fill="black", font=font_small)
    draw.text((50, 300), f"Address : {entries['Address'].get()}", fill="black", font=font_small)
   
    draw.text((50, 340), f"{entries['Welcome Message'].get()}", fill="black", font=font_small)
    # draw a dashed line - - - - - like this
    draw.line((50, 365, 550, 365), fill="black", width=1)

    # Aligned event schedule block
    y_start = 390
    line_spacing = 25

    activities = [
        "Welcome Drinks",
        "Game 1",
        "Performance",
        "Game 2",
        "Lunch Break",
        "Paper Dance",
        "Performance",
        "Game 3",
        "Closing Ceremony"
    ]

    times = [
        "10:30 AM",
        "11:00 AM",
        "12:00 PM",
        "12:10 PM",
        "01:00 PM",
        "02:00 PM",
        "02:30 PM",
        "02:40 PM",
        "05:30 PM"
    ]

    # Draw column headers (optional)
    # draw.text((50, y_start - 30), "Activity", fill="black", font=font_small)
    # draw.text((400, y_start - 30), "Time", fill="black", font=font_small)

    for i in range(len(activities)):
        draw.text((50, y_start + i * line_spacing), activities[i], fill="black", font=font_small)
        draw.text((450, y_start + i * line_spacing), times[i], fill="black", font=font_small)

    # Mood meter and footer
    draw.text((50, 630), "Mood-o-meter : ", fill="black", font=font_small)
    
    font_star = ImageFont.truetype("arial.ttf", size=20)  # Works on most systems
    # Draw stars for mood
    draw.text((230, 630), "☆ ☆ ☆ ☆ ☆", fill="black", font=font_star)
    

    draw.line((50, 660, 550, 660), fill="black", width=1)
    
    #draw.rectangle([50, 690, 550, 730], fill="black")
    # Draw label before QR
    draw.text((230, 670), "Scan for venue 📍", fill="black", font=font_small)

    # Paste QR from file
    paste_existing_qr(img, "venue_qr.png", position=(255, 690), size=(100, 100))




    draw.text((170, 795), "2927 2 017 02 209605 1 5080", fill="black", font=font_small)

    draw.text((230, 850), "SEE YOU THERE", fill="black", font=font_small)


    filename = os.path.join(TICKETS_FOLDER, f"{name}_ticket.png")
    img.save(filename)
    return filename

# === Add QR Code Function ===
from PIL import Image, ImageDraw, ImageFont

def paste_existing_qr(ticket_img, qr_path, position=(400, 620), size=(100, 100)):
    # Open the QR code image
    qr_img = Image.open(qr_path).convert("RGB")

    # Resize if needed
    qr_img = qr_img.resize(size)

    # Paste QR onto the ticket
    ticket_img.paste(qr_img, position)

    return ticket_img



# === Send Invitations ===
def send_invitations():
    if EXCEL_FILE == "":
        messagebox.showerror("Error", "Please upload an Excel file.")
        return

    try:
        status_label.config(text="Status: Processing...", fg="orange")
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip()

        if 'Name' not in df.columns or 'Email' not in df.columns:
            messagebox.showerror("Error", "Excel must have 'Name' and 'Email' columns.")
            return

        df['Name'] = df['Name'].astype(str)
        df['Email'] = df['Email'].astype(str)

        yag = yagmail.SMTP(SENDER_EMAIL, APP_PASSWORD)

        for _, row in df.iterrows():
            name = row['Name']
            email = row['Email']

            ticket_file = generate_ticket(name)

            contents = [
                f"Dear {name},\n\nYou are invited to our farewell party!\n\nVenue: {entries['Venue'].get()}\nDate: {entries['Date'].get()} at {entries['Time'].get()}\nAddress: {entries['Address'].get()}\n\n{entries['Welcome Message'].get()}\n\nSee you there! 🎉",
                ticket_file
            ]

            yag.send(to=email, subject="Farewell Party Invitation", contents=contents)
            print(f"✅ Sent to: {name} ({email})")

        status_label.config(text="Status: Invitations Sent!", fg="green")
        messagebox.showinfo("Done", "All invitations sent successfully!")

    except Exception as e:
        status_label.config(text="Status: Error", fg="red")
        messagebox.showerror("Error", str(e))

# === Send Button ===
send_btn = tk.Button(root, text="Send Invitations", command=send_invitations, bg="#FF69B4", font=("Comic Sans MS", 12, "bold"))
send_btn.place(x=100, y=700)

# === Run App ===
root.mainloop()
