import os
import tarfile
import json
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

entry_folder = None  # Define entry_folder as a global variable

def extract_json(tarball_path):
    hostname_data = None
    with tarfile.open(tarball_path, "r") as tar:
        for member in tar.getmembers():
            if member.name.endswith("device_information.json"):
                    f = tar.extractfile(member)
                    json_data = json.load(f)
                    hostname_data = json_data["hostname"]
                    print(f"hn data: {hostname_data}")
            if member.isfile() and os.path.basename(member.name) == "raw_datastore.json":
                json_data = tar.extractfile(member).read()
                data = json.loads(json_data)
                raw_data_avg = data.get("raw_data_avg")
                print(raw_data_avg)
                if raw_data_avg:
                    duplicates = find_duplicates(raw_data_avg)
                    print(duplicates)
                    if not duplicates:
                        generate_pdf(hostname_data, tarball_path, "No Duplicates Found")
                    else:
                        generate_pdf(hostname_data, tarball_path, f"Duplicates found for these Standards - {duplicates}")
                else:
                    generate_pdf(hostname_data, tarball_path, "No raw_data_avg found")

def find_duplicates(raw_data_avg):
    duplicates = []
    seen = set()
    for key, value in raw_data_avg.items():
        if isinstance(value, list):
            sublist_str = ','.join(map(str, value))
            if sublist_str in seen:
                duplicates.append(key)
            else:
                seen.add(sublist_str)
    return duplicates

def generate_pdf(hostname_data, tarball_path, message):
    folder_path = os.path.dirname(tarball_path)
    pdf_filename = os.path.join(folder_path, os.path.splitext(os.path.basename(tarball_path))[0] + ".pdf")
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    
    # Insert Company Logo
    logo_path = "C:\\Users\\UKumar\\OneDrive - Genomadix\\Desktop\\Genomadix\\MES-Extractors\\QC Check Tool for Duplicate Values\\genomadix-logo-colour-no-tagline.jpg"  # Provide the path to company logo image file
    if os.path.exists(logo_path):
        logo = ImageReader(logo_path)
        c.drawImage(logo, 200, 700, width=3*inch, height=1*inch)

    # Set font and font size for the heading
    c.setFont("Helvetica-Bold", 20)
    
    # Draw the heading "Device QC Check"
    c.drawString(150, 660, f"Device QC Check for {hostname_data}")
    
    # Set font and font size for the date and signature
    c.setFont("Helvetica", 15)
    
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Move to the next line
    c.drawString(80, 550, f"Date and time of QC Check: {current_date}")

    c.drawString(80, 500, f"Hostname: {hostname_data}")
    
    # Draw the message (e.g., "Verification Successful")
    c.drawString(80, 450, f"Outcome: {message}")

    c.line(80, 120, 230, 120)

    c.line(345, 120, 500, 120)
    
    # Add text for the reviewer's signature
    c.drawString(80, 100, "Operator's Signature:")

    c.drawString(350, 100, "Reviewer's Signature:")
    
    # Save the PDF file
    c.save()

def select_folder():
    global entry_folder  # Reference the global entry_folder variable
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder_path)

def process_tarballs():
    global entry_folder  # Reference the global entry_folder variable
    folder_path = entry_folder.get()
    if folder_path:
        for file_name in os.listdir(folder_path):
            if file_name.endswith(".gz"):
                tarball_path = os.path.join(folder_path, file_name)
                extract_json(tarball_path)
        messagebox.showinfo("Done", "Processing completed.")

def main():
    global entry_folder  # Define entry_folder as global within main()
    root = tk.Tk()
    root.title("QC Check Tool")
    root.geometry("400x200")  # Adjust the dimensions as needed
    
    # Entry to display selected folder path
    entry_folder = tk.Entry(root, width=50)
    entry_folder.pack(side="top", pady=5)
    
    # Frame to contain buttons
    button_frame = tk.Frame(root)
    button_frame.pack(side="top", pady=5)
    
    # Folder Selection Button
    btn_select_folder = tk.Button(button_frame, text="Select Tarball Folder", command=select_folder)
    btn_select_folder.pack(side="left", padx=5)
    
    # Process Tarballs Button
    btn_process = tk.Button(button_frame, text="Process Tarballs", command=process_tarballs)
    btn_process.pack(side="left", padx=5)
    
    root.mainloop()

if __name__ == "__main__":
    main()
