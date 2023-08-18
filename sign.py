import fitz
from PIL import Image
import os
import shutil
import subprocess
import time
import datetime

# Set the log file path and name
log_path = "input_log.txt"

# Check if the log file already exists
if not os.path.exists(log_path):
    # If the log file doesn't exist, create it and write the header row
    with open(log_path, "w") as f:
        f.write("Date,Time,UserName\n")

# Get the current date and time
now = datetime.datetime.now()
date = now.date().isoformat()
mytime = now.time().strftime("%H:%M:%S")

# Define the input and output directories
input_dir = ".\\"
unsigned_dir = ".\\unsigned"
signed_dir = ".\\signed"

# Define the path for the signature file
#signature_path = ".\\signature.png"

# Define the signature positions and page ranges for each user
user_signatures = {
    "USER1": {
        "page_range": (0, 0),
        "signature_pos": (15, 720, 15 + 100, 720 + 50),
    },
    "USER2": {
        "page_range": (0, 1),
        "signature_pos": (15, 690, 15 + 100, 690 + 50),
    },
    "USER3": {
        "page_range": (0, 1),
        "signature_pos": (50, 430, 50 + 100, 430 + 50),
    },
}

# Prompt the user to enter their username
valid_usernames = ["USER1", "USER2", "USER3"]
username = input("Enter your username: ")

while username not in valid_usernames:
    print("Invalid username. Please try again.")
    username = input("Enter your username: ")
username_or = username
# Prompt the user to sign on behalf of someone else
while True:
    sign_on_behalf = input("Do you want to sign on behalf of someone else? (1 for Yes, 0 for No): ")
    if sign_on_behalf in ["0", "1"]:
        sign_on_behalf = int(sign_on_behalf)
        break
    else:
        print("Invalid input. Please enter 1 or 0.")

# If the user wants to sign on behalf of someone else
if sign_on_behalf == 1:
    # Prompt the user to enter the username of the person they're signing on behalf of
    behalf_username = input("Enter the username of the person you're signing on behalf of (USER1, USER2, or USER3): ")
    while behalf_username not in valid_usernames:
        behalf_username = input("Invalid username. Please enter a valid username (USER1, USER2, or USER3): ")
    # If the entered username is valid
    if behalf_username in user_signatures:
        # Retrieve the signature position and page range for the user being signed on behalf of
        signature_pos = user_signatures[behalf_username]["signature_pos"]
        page_range = user_signatures[behalf_username]["page_range"]
        if username == "USER1":
            signature_path = r""
        elif username == "USER2":
            signature_path = r""
        else:
            signature_path = r"" 
        username = behalf_username
    else:
        print("Invalid username.")
        sign_on_behalf = 0

# If the user does not want to sign on behalf of someone else
if sign_on_behalf == 0:
    # Set the signature positions and page ranges based on the user's own information
    signature_pos = user_signatures[username]["signature_pos"]
    page_range = user_signatures[username]["page_range"]
    if username == "USER1":
        signature_path = r""
    elif username == "USER2":
        signature_path = r""
    else:
        signature_path = r"" 
    behalf_username = "N/A"
# Initialize a variable to count the number of PDF files processed
pdf_files_processed = 0

# Initialize counts for signed and unsigned files
signed_count = 0
unsigned_count = 0

# Iterate over each PDF file in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith(".pdf"):
        # Preview the PDF file in AcroRd32 Reader and prompt the user to decide whether to sign it or not
        try:
            subprocess.check_output(["TASKLIST", "/FI", "IMAGENAME eq AcroRd32.exe"], shell=True)
            os.startfile(os.path.join(input_dir, filename))
            # check for valid value for sign
            while True:
                sign = input("Do you want to sign the file? (1 for Yes, 0 for No): ")
                if sign in ['0', '1']:
                    sign = int(sign)
                    break
                else:
                    print("Invalid input. Please enter 1 for Yes or 0 for No.")
            # Append the user input and timestamp to the log file
            with open(log_path, "a") as f:
                f.write(f"{date},{mytime},{username_or}, {sign_on_behalf}, {behalf_username}, {sign}\n")
        except subprocess.CalledProcessError:
            print("AcroRd32 Reader is not running.")
            sign = 0
        # Open the PDF file using PyMuPDF if the user wants to sign it
        if sign == 1:
            with fitz.open(os.path.join(input_dir, filename)) as input_pdf:
                output_pdf = fitz.open()

                # Add the signature to the specified pages of the PDF file
                for i, page in enumerate(input_pdf):
                    if user_signatures[username]["page_range"][0] <= i <= user_signatures[username]["page_range"][1]:
                        # Set the signature position based on the user's page range
                        if username == "USER1":
                            signature_pos = (15, 720, 15 + 100, 720 + 50)
                        elif username in ["USER2", "USER3"]:
                            if i == 0:
                                signature_pos = (15, 690, 15 + 100, 690 + 50)
                            elif i == 1:
                                signature_pos = (50, 530, 50 + 100, 530 + 50)
                        # Add the signature image to the page as a small rectangular shape at the specified position
                        page_mediabox = page.mediabox
                        signature_image = fitz.Pixmap(signature_path)
                        page.insert_image(
                            rect=fitz.Rect(*signature_pos),
                            pixmap=signature_image,
                            keep_proportion=True,
                        )
                    output_pdf.insert_pdf(input_pdf, from_page=i, to_page=i)

                # Save the signed PDF file to the signed directory
                output_file_path = os.path.join(signed_dir, filename)
                output_pdf.save(output_file_path)
                signed_count += 1
        else:
            # Copy the unsigned PDF file to the unsigned directory
            shutil.copyfile(os.path.join(input_dir, filename), os.path.join(unsigned_dir, filename))
            unsigned_count += 1
        # Close the preview of the PDF file
        subprocess.check_output(["TASKKILL", "/F", "/IM", "AcroRd32.exe"], shell=True)
        
        # Wait for 1 seconds
        time.sleep(1)

        #Delete the preview file
        os.remove(os.path.join(input_dir, filename))
        
        # Increment the PDF files processed counter
        pdf_files_processed += 1  
              
# Display the counts of signed and unsigned files
print(" ")
print(" ")
print(f"{signed_count} PDF files have been signed.")
print(f"{unsigned_count} PDF files have been left unsigned.")
print(f"{pdf_files_processed} PDF files have been processed.")
input("Press any key to exit.")  # Wait for user input before closing the console
