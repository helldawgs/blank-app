import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import win32com.client
import datetime
import os

class DamagedTrailerReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UPS Damaged Trailer Reporter")
        
        # Set window size
        self.window_width = 800
        self.window_height = 600
        
        # Configure background
        self.configure(bg='#f5f5f5')
        
        # Create main frame
        self.main_frame = ttk.Frame(self, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Initialize variables
        self.file_path = tk.StringVar()
        
        # Yard code mapping
        self.yard_mapping = {
            'AFW5': 'AFW2', 'AGS5': 'AGS3', 'AVP9': 'AVP9', 'AVPA': 'AVP3',
            'AZA9': 'SDL2', 'AZAB': 'AZA4', 'BDL6': 'BDL6', 'BDU5': 'BDU2',
            'BFI7': 'BFI7', 'BUFA': 'BUF9', 'CAEA': 'CAE1', 'CHAA': 'CHA1',
            'CLT6': 'CLT6', 'CMHA': 'CMH6', 'COS5': 'JHW1', 'CVG2': 'CVG2',
            'DENB': 'DEN7', 'DFWA': 'DFW9', 'DPA7': 'DPA7', 'DTW3': 'DTW3',
            'DTW9': 'DET7', 'EWRB': 'EWR7', 'EWRC': 'MMU9', 'FAR1': 'FAR1',
            'FTWB': 'FTW2', 'HGRA': 'HGR2', 'HLAA': 'LGB9', 'HOU3': 'HOU3',
            'HOU7': 'HOU7', 'HOUB': 'HOU9', 'HSV2': 'HSV2', 'IAH1': 'IAH1',
            'INDB': 'IND8', 'JAX9': 'CRG1', 'LAS2': 'LAS2', 'LEX1': 'LEX1',
            'LEX2': 'LEX2', 'LUK7': 'LUK7', 'MCI3': 'MCI3', 'MCI7': 'MCI7',
            'MDT5': 'MDT8', 'MDW8': 'MDW8', 'MEM8': 'MEM8', 'MGE8': 'PDK2',
            'MIAA': 'MIA7', 'MQJA': 'MQJ2', 'MSP6': 'MSP6', 'MSP7': 'STP2',
            'MTN6': 'MTN3', 'MTNB': 'MTN7', 'OAKB': 'OAK7', 'OKCA': 'OKC9',
            'ONTA': 'ONT2', 'ONTD': 'BUR7', 'ORD9': 'RFD7', 'PBIA': 'PBI2',
            'PDXA': 'HIO9', 'PDXB': 'PDX6', 'PGAA': 'PGA1', 'PHLA': 'PHL7',
            'PHXA': 'PCA1', 'PHX6': 'PHX6', 'PILA': 'MEM3', 'PIT9': 'PIT4',
            'RFD4': 'RFD4', 'RICB': 'RIC3', 'RICA': 'RIC9', 'RNT9': 'AUN2',
            'RSW5': 'LAL4', 'SAN5': 'SDM4', 'SAT6': 'SAT6', 'SAT9': 'SAT7',
            'SCK3': 'SCK3', 'SDF8': 'SDF8', 'SDFA': 'SDF9', 'SLCA': 'SLC3',
            'SMF7': 'SMF7', 'SMFA': 'SMF9', 'STL9': 'BLV2', 'STLA': 'STL6',
            'TCYA': 'OAK9', 'TEN1': 'TEN1', 'TPAA': 'TPA6', 'TTNA': 'TTN2',
            'TUS1': 'TUS1', 'YYZ1': 'YYZ1', 'YYZ7': 'YYZ7', 'YYZ9': 'YYZ9',
            'YYC1': 'YYC1', 'YVR3': 'YVR3', 'YYC6': 'YYC6', 'YGK1': 'YGK1'
        }
        
        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#f5f5f5') 
        style.configure('TLabel', background='#f5f5f5', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10))
        style.configure('TLabelframe', background='#f5f5f5')
        style.configure('TLabelframe.Label', background='#f5f5f5', font=('Segoe UI', 10, 'bold'))

    def create_widgets(self):
        # Title
        title_label = ttk.Label(
            self.main_frame,
            text="UPS Damaged Trailer Report Generator",
            font=('Segoe UI', 14, 'bold')
        )
        title_label.pack(pady=20)

        # File Selection Frame
        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT)

        # Generate Button
        generate_btn = ttk.Button(
            self.main_frame,
            text="Generate Reports",
            command=self.create_damaged_trailer_emails
        )
        generate_btn.pack(pady=20)

        # Status Area
        status_frame = ttk.LabelFrame(self.main_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(status_frame, height=8, width=50)
        self.status_text.pack(fill=tk.BOTH, expand=True)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.status_text.delete(1.0, tk.END)
            self.status_text.insert(tk.END, f"File selected: {file_path}\n")

    def create_damaged_trailer_emails(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Find specific mailbox
            for account in namespace.Accounts:
                if account.SmtpAddress == "3P-damaged-trailer-repairs@amazon.com":
                    specific_account = account
                    break
            else:
                messagebox.showerror("Error", "Could not find the specified mailbox.")
                return
            
            drafts_folder = specific_account.DeliveryStore.GetDefaultFolder(16)
            
        except Exception as e:
            messagebox.showerror("Error", f"Outlook error: {str(e)}")
            return

        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        try:
            df = pd.read_excel(self.file_path.get(), sheet_name="sheet1")
        except Exception as e:
            messagebox.showerror("Error", f"Excel error: {str(e)}")
            return

        dict_data = {}
        unique_sites = {}

        # Process rows
        for i in range(3, len(df)):
            if pd.notna(df.iloc[i, 2]):
                original_site = df.iloc[i, 1]
                site = self.yard_mapping.get(original_site, original_site)
                
                entry_str = f"""
                <tr style='border-bottom: 1px solid #eee;'>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(df.iloc[i, 2])}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(original_site)}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(df.iloc[i, 7])}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(df.iloc[i, 10])}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(df.iloc[i, 6])}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'>{self.handle_null_or_empty(df.iloc[i, 12])}</td>
                    <td style='padding: 12px; border: 1px solid #ddd; color: #333;'></td>
                </tr>
                """

                if site not in dict_data:
                    dict_data[site] = entry_str
                else:
                    dict_data[site] += entry_str

                if site not in unique_sites:
                    unique_sites[site] = 1
                else:
                    unique_sites[site] += 1

        # Create emails
        email_count = 0
        for site_key in dict_data:
            mail = drafts_folder.Items.Add()
            mail.To = f"{site_key}-repairs@amazon.com"
            mail.CC = f"ups-dispatch-amazon-{site_key}@ups.com; UPS-Automotive-Amazon-{site_key}@ups.com"
            mail.Subject = f"UPS Damaged Trailer Report - {site_key} - {datetime.date.today().strftime('%m/%d/%Y')}"

            html_header = f"""
            <html><body style='font-family: Segoe UI, Arial, sans-serif; background-color: #f5f5f5; color: #333; line-height: 1.6;'>
            <div style='max-width: 900px; margin: auto; padding: 25px; background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
            <h2 style='text-align: center; background-color: #232F3E; color: white; padding: 15px; margin: 0; border-radius: 6px; font-size: 24px;'>
            UPS Damaged Trailer Report - {site_key}</h2>
            """

            html_table = f"""
            <div style='margin: 25px 0;'><table style='width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 20px; font-size: 14px;'>
            <thead><tr style='background-color: #FF9900;'>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>Trailer ID</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>Yardcode</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>Trailer Status</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>Location</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>Tag Dwell</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>YMs Notes</th>
            <th style='padding: 12px; border: 1px solid #ddd; border-bottom: 2px solid #232F3E; text-align: left; font-weight: 600;'>UPS Status Updates</th>
            </tr></thead>
            <tbody style='background-color: #ffffff;'>{dict_data[site_key]}</tbody></table></div>
            """

            html_footer = f"""
            <div style='margin-top: 30px; padding: 20px; background-color: #f8f9fa; border-radius: 6px; border-left: 4px solid #FF9900;'>
            <div style='font-size: 14px; color: #444;'>
            <p style='margin: 0 0 10px 0;'><strong style='color: #232F3E;'>Date:</strong> {datetime.date.today().strftime('%m/%d/%Y')}</p>
            <p style='margin: 0 0 10px 0;'><strong style='color: #232F3E;'>Total Damaged Trailers:</strong> {unique_sites[site_key]}</p></div>
            <div style='margin-top: 20px; padding-top: 20px; border-top: 1px solid #dee2e6;'>
            <h3 style='color: #232F3E; margin: 0 0 10px 0; font-size: 16px;'>Action Needed:</h3>
            <p style='margin: 0; color: #555; font-size: 14px;'>Please review the damaged trailers listed above and take appropriate action. 
            For any questions or concerns, please reach out to your management team.</p></div></div></div></body></html>
            """

            mail.HTMLBody = html_header + html_table + html_footer
            mail.Save()
            
            email_count += 1
            self.status_text.insert(tk.END, f"Created email draft for site: {site_key}\n")
            self.update_idletasks()

        if email_count > 0:
            messagebox.showinfo("Success", f"{email_count} email draft(s) have been created and saved in the Drafts folder.")
        else:
            messagebox.showwarning("Warning", "No email drafts were created. Please check your data.")

    def handle_null_or_empty(self, value, default_value=""):
        if pd.isna(value) or value == "null":
            return default_value
        return str(value)

if __name__ == "__main__":
    app = DamagedTrailerReportApp()
    app.mainloop()
