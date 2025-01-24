import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def excel_to_vcf(excel_file, output_dir='vcf_output'):
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Read Excel file
    df = pd.read_excel(excel_file)
    
    # Assume first column is name, second is phone number
    Name = df.columns[0]
    ContactNo = df.columns[1]
    
    # Create VCF file
    vcf_filename = os.path.join(output_dir, 'contacts.vcf')
    
    with open(vcf_filename, 'w', encoding='utf-8') as vcf_file:
        for index, row in df.iterrows():
            name = str(row[Name]).strip()
            phone = str(row[ContactNo]).strip()
            
            # Skip rows with invalid phone numbers
            if not phone or phone == 'nan':
                continue
            
            vcf_file.write('BEGIN:VCARD\n')
            vcf_file.write('VERSION:3.0\n')
            vcf_file.write(f'FN:{name}\n')
            vcf_file.write(f'TEL:{phone}\n')
            vcf_file.write('END:VCARD\n')
    
    return vcf_filename

class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to VCF Converter")
        self.root.geometry("400x200")
        
        # Create and pack widgets
        self.label = tk.Label(root, text="Select Excel File to Convert", pady=20)
        self.label.pack()
        
        self.browse_button = tk.Button(root, text="Browse Excel File", command=self.browse_file)
        self.browse_button.pack(pady=20)
        
        self.status_label = tk.Label(root, text="", wraplength=350)
        self.status_label.pack(pady=20)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
                output_file = excel_to_vcf(file_path)
                messagebox.showinfo("Success", f"VCF file created successfully!\nLocation: {output_file}")
                self.status_label.config(text=f"Converted: {os.path.basename(file_path)}")
            
def main():
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()