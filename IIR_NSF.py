import pandas as pd
from fpdf import FPDF
import os
from zipfile import ZipFile

#For SF Message remove the comments below

# def create_report(dispute_id, dispute_date, issuer, trx_amount, trx_date):
#     report = f"""
#     Initial Investigation Report ({dispute_id})
    
#     Case {dispute_id} was reported on {dispute_date} by {issuer}. Funds equal to PKR {trx_amount}/- were received on {trx_date}. Due to timely action by us against the beneficiary account, the funds are on hold.
    
#     Kindly email at " " for further information.
#     """
#     return report

#NSF Template

def create_report(dispute_id, dispute_date, issuer, trx_amount, trx_date):
    report = f"""
    Initial Investigation Report ({dispute_id})

    Case {dispute_id} was reported on {dispute_date}, by {issuer}. Funds equal to PKR {trx_amount}/- were received on {trx_date}. 
    Timely action was taken against the beneficiary, and their account was blocked. Funds were not held with us.

    Kindly email at " " for further information.
    """
    return report

file_path = "Data_File.xlsx"  # Path
df = pd.read_excel(file_path)
df = df.applymap(str)

df['CustomerDisputeDate'] = pd.to_datetime(df['CustomerDisputeDate']).dt.strftime('%Y-%m-%d %I:%M %p')
df['TrxDate'] = pd.to_datetime(df['TrxDate']).dt.strftime('%Y-%m-%d %I:%M %p')

for index, row in df.iterrows():
    print("\n--- Report Preview ---")
    report_text = create_report(
        row['DisputeID'], row['CustomerDisputeDate'], row['Issuer'], row['TrxAmount'], row['TrxDate']
    )
    print(report_text)

output_folder = "Reports"
os.makedirs(output_folder, exist_ok=True)

pdfs = []  
for index, row in df.iterrows():
    dispute_id = row['DisputeID']
    report_text = create_report(
        dispute_id, row['CustomerDisputeDate'], row['Issuer'], row['TrxAmount'], row['TrxDate']
    )
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, txt=report_text, align="C")
    
    pdf_file = os.path.join(output_folder, f"Report_{dispute_id}.pdf")
    pdf.output(pdf_file)
    pdfs.append(pdf_file)

print(f"\nPDFs saved in folder: {output_folder}")

zip_file = "Reports.zip"
with ZipFile(zip_file, 'w') as zipf:
    for pdf in pdfs:
        zipf.write(pdf, os.path.basename(pdf))

print(f"\nAll PDFs compressed into: {zip_file}")
