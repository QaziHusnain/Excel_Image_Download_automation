import os
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog

def download_images_and_rename(input_excel):
    # Read the Excel file
    df = pd.read_excel(input_excel)

    # Create a folder to save the images
    output_folder = "downloaded_images"
    os.makedirs(output_folder, exist_ok=True)

    for index, row in df.iterrows():
        product_code = row.iloc[0]
        for i in range(1, len(row)):
            if pd.notna(row.iloc[i]):
                variant_code = f"PT{i:02d}"
                image_url = row.iloc[i]
                image_extension = image_url.split('.')[-1]
                new_filename = f"{product_code}.{variant_code}.{image_extension}"

                # Download and save the image
                response = requests.get(image_url, stream=True)
                with open(os.path.join(output_folder, new_filename), 'wb') as file:
                    for chunk in response.iter_content(chunk_size=128):
                        file.write(chunk)

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry_var.set(file_path)

def start_download():
    input_excel_file = entry_var.get()
    if os.path.exists(input_excel_file):
        download_images_and_rename(input_excel_file)
        result_label.config(text="Download completed.")
    else:
        result_label.config(text="Invalid file path.")

if __name__ == "__main__":
    # GUI setup
    root = tk.Tk()
    root.title("Image Downloader")

    # Input File Entry
    entry_var = tk.StringVar()
    input_label = tk.Label(root, text="Select Excel File:")
    input_entry = tk.Entry(root, textvariable=entry_var, width=40)
    input_entry.grid(row=0, column=1, padx=10, pady=10)
    input_label.grid(row=0, column=0, padx=10, pady=10)

    # Browse Button
    browse_button = tk.Button(root, text="Browse", command=browse_excel_file)
    browse_button.grid(row=0, column=2, padx=10, pady=10)

    # Download Button
    download_button = tk.Button(root, text="Start Download", command=start_download)
    download_button.grid(row=1, column=1, pady=20)

    # Result Label
    result_label = tk.Label(root, text="")
    result_label.grid(row=2, column=1, pady=10)

    root.mainloop()
