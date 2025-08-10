# ticket_generator
A Python tool to generate personalized event tickets from Excel data using PIL. Adds a venue QR code and supports auto email delivery via SMTP. Ideal for student events, farewells, or small gatherings.

🎟️ Event Ticket Generator with QR Code (Python + PIL)
This project is a Python-based tool that generates customized black-and-white event tickets using participant details from an Excel sheet. It incorporates image processing, font styling, and dynamic QR code embedding for venue access.

[![Lekhana-ticket-3.png](https://i.postimg.cc/zBkxjsBj/Lekhana-ticket-3.png)](https://postimg.cc/mzt3Z67z)

✨ Features:
📥 Reads names and email addresses from Excel (.xlsx) using pandas.

🖼️ Generates personalized tickets using PIL (Python Imaging Library).

🖨️ Supports pixel fonts and elegant ticket layout.

📍 Embeds QR code of the event venue for quick access via Google Maps or other apps.

📧 Automates ticket delivery via email (Gmail SMTP with app password).

📁 Organizes output as individual .png ticket files.

🔧 Tech Stack:
Python 3

PIL / Pillow

Pandas

smtplib (Gmail SMTP)

qrcode (for venue QR code)

openpyxl (for Excel support)

📌 How It Works:
Add your participant list in Excel format (Name, Email).

Customize event details like date, venue, and address.

Generate a QR code of the venue (optional but recommended).

Run the script to:

Create ticket images.

Paste the QR code.

Optionally send tickets via email.

📂 Output:
All tickets are saved as .png files in the tickets/ folder and can be shared or printed.
