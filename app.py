import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64
import glob
import time

# Page Configuration
st.set_page_config(
    page_title="Petrol Claim Portal",
    page_icon="⛽",
    layout="wide"
)

# Sidebar: API Key and Portal Selection
st.sidebar.title("⛽ Petrol Claim Portal")
st.sidebar.markdown("---")

# Get Groq API Key from Streamlit Secrets or Sidebar Input
groq_api_key = ""
if "GROQ_API_KEY" in st.secrets:
    groq_api_key = st.secrets["GROQ_API_KEY"]
else:
    groq_api_key = st.sidebar.text_input("Groq API Key", type="password", help="Get your free API Key from console.groq.com")

st.sidebar.markdown("---")
portal_mode = st.sidebar.radio("Select Portal", ["👥 Colleague Portal", "🔑 Owner Portal"])

# Find Excel template automatically in the directory
def get_template_path():
    xlsx_files = glob.glob("*.xlsx")
    if xlsx_files:
        return xlsx_files[0]
    return "Fuel Reimbursement Claim Form new- Original - Copy.xlsx.xlsx"

# Parse single receipt using Groq Vision API (Strictly capturing credit card / actual payment amount)
def parse_receipt_with_groq(image_bytes, api_key):
    try:
        client = Groq(api_key=api_key)
        encoded_image = base64.b64encode(image_bytes).decode('utf-8')
        
        prompt = """
        You are an expert financial OCR assistant for petrol receipts in Malaysia.
        Analyze this petrol receipt image and extract the following 3 core fields accurately in JSON format:
        {
          "receipt_no": "Receipt number, invoice number, or transaction reference number (string or null)",
          "litres": 0.00,
          "amount_rm": "The exact credit card charged amount, cash payment amount, or final amount paid by card. Do NOT pick total price before discount, and do NOT pick subsidy or MADANI amounts. Pick the exact final amount charged to the payment method."
        }
        Return ONLY valid JSON. If any field is not found, put null for receipt_no and 0.00 for numbers.
        """
        
        chat_completion = client.chat.completions.create(
            model="qwen/qwen3.8-27b",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{encoded_image}"
                            }
                        }
                    ]
                }
            ],
            response_format={"type": "json_object"}
        )
        return json.loads(chat_completion.choices[0].message.content)
    except Exception as e:
        return {"receipt_no": "-", "litres": 0.0, "amount_rm": 0.0}

# Generate Excel Claim File supporting multiple receipts (filling multiple rows starting from row 15)
def generate_excel_claim(profile_data, claim_data_list):
    template_path = get_template_path()
    wb = openpyxl.load_workbook(template_path)
    
    if "Revised 2" in wb.sheetnames:
        ws = wb["Revised 2"]
    else:
        ws = wb.active

    # Fill Employee Information
    ws['D10'] = profile_data['name']
    ws['K10'] = profile_data['designation']
    ws['D11'] = profile_data['employee_no']
    ws['K11'] = profile_data['department']
    ws['D12'] = profile_data['vehicle_no']
    ws['K12'] = profile_data['monthly_limit']
    ws['D13'] = profile_data['month']

    # Fill Multiple Receipts starting from row 15 downwards
    start_row = 15
    for idx, item in enumerate(claim_data_list):
        current_row = start_row + idx
        ws.cell(row=current_row, column=7).value = item['receipt_no']
        ws.cell(row=current_row, column=9).value = item['litres']
        ws.cell(row=current_row, column=11).value = item['amount_rm']

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==================== 1. Colleague Portal ====================
if portal_mode == "👥 Colleague Portal":
    st.title("👥 Fuel Reimbursement Claim Form - Colleague Portal")
    st.markdown("Please enter your personal details and upload multiple fuel receipts. The system will automatically process all receipts and generate your claim Excel sheet.")

    with st.form("colleague_form"):
        col1, col2 = st.columns(2)
        with col1:
            c_name = st.text_input("Employee Name")
            c_emp_no = st.text_input("Employee No.")
            c_dept = st.text_input("Department")
        with col2:
            c_designation = st.text_input("Designation")
            c_vehicle = st.text_input("Vehicle No.")
            c_limit = st.number_input("Monthly Claim Limit (RM)", value=500.0)
        
        c_month = st.text_input("Claim for the Month of (e.g., October 2026)")
        uploaded_files = st.file_uploader("Upload Receipt Images (Multiple allowed)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        submitted = st.form_submit_button("🤖 Auto-Parse All Receipts & Generate Claim")

    if submitted:
        if not groq_api_key:
            st.error("Please configure your Groq API Key in Streamlit Secrets or enter it in the sidebar!")
        elif not uploaded_files:
            st.error("Please upload at least one receipt image!")
        else:
            total_files = len(uploaded_files)
            progress_bar = st.progress(0, text=f"Initializing batch processing for {total_files} receipts...")
            
            extracted_claims = []
            for i, file in enumerate(uploaded_files):
                progress_percent = int(((i) / total_files) * 90) + 5
                progress_bar.progress(progress_percent, text=f"Processing receipt {i+1} of {total_files} ({file.name})...")
                
                image_bytes = file.getvalue()
                receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                
                if receipt_info:
                    extracted_claims.append({
                        "receipt_no": receipt_info.get("receipt_no", "-"),
                        "litres": float(receipt_info.get("litres", 0) or 0),
                        "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                    })
            
            progress_bar.progress(95, text="Populating multi-row Excel sheet...")
            time.sleep(0.3)
            
            profile = {
                "name": c_name,
                "employee_no": c_emp_no,
                "department": c_dept,
                "designation": c_designation,
                "vehicle_no": c_vehicle,
                "monthly_limit": c_limit,
                "month": c_month
            }
            
            excel_file = generate_excel_claim(profile, extracted_claims)
            progress_bar.progress(100, text="Completed!")
            time.sleep(0.2)
            
            st.success(f"Successfully processed {len(extracted_claims)} receipts!")
            st.download_button(
                label="📥 Download Claim Excel Sheet",
                data=excel_file,
                file_name=f"Petrol_Claim_{c_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==================== 2. Owner Portal ====================
elif portal_mode == "🔑 Owner Portal":
    st.title("🔑 Owner Portal")
    
    password = st.text_input("Enter Password", type="password")
    
    if password == "P@ssw0rd":
        st.success("Authentication successful! Welcome back.")
        
        st.markdown("### 📋 Preset Profile Information")
        col1, col2 = st.columns(2)
        with col1:
            my_name = st.text_input("Employee Name", value="SOO WAI WING")
            my_emp_no = st.text_input("Employee No.", value="AJMS0849")
            my_dept = st.text_input("Department", value="IS&T")
        with col2:
            my_designation = st.text_input("Designation", value="IS&T SM")
            my_vehicle = st.text_input("Vehicle No.", value="VKG4113")
            my_limit = st.number_input("Monthly Claim Limit (RM)", value=800.0)
            
        my_month = st.text_input("Claim for the Month of", value="October 2026")
        
        st.markdown("---")
        my_uploaded_files = st.file_uploader("Upload Fuel Receipts (Multiple allowed)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="my_receipts")
        
        if st.button("🚀 Process My Claims"):
            if not groq_api_key:
                st.error("Please configure your Groq API Key in Streamlit Secrets or enter it in the sidebar!")
            elif not my_uploaded_files:
                st.error("Please upload at least one receipt!")
            else:
                total_files = len(my_uploaded_files)
                progress_bar = st.progress(0, text=f"Initializing batch OCR for {total_files} receipts...")
                
                extracted_claims = []
                for i, file in enumerate(my_uploaded_files):
                    progress_percent = int(((i) / total_files) * 90) + 5
                    progress_bar.progress(progress_percent, text=f"Processing receipt {i+1} of {total_files} ({file.name})...")
                    
                    image_bytes = file.getvalue()
                    receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                    
                    if receipt_info:
                        extracted_claims.append({
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": float(receipt_info.get("litres", 0) or 0),
                            "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                        })
                
                progress_bar.progress(95, text="Generating Excel with multi-row entries...")
                time.sleep(0.3)
                
                profile = {
                    "name": my_name,
                    "employee_no": my_emp_no,
                    "department": my_dept,
                    "designation": my_designation,
                    "vehicle_no": my_vehicle,
                    "monthly_limit": my_limit,
                    "month": my_month
                }
                
                excel_file = generate_excel_claim(profile, extracted_claims)
                progress_bar.progress(100, text="Done!")
                time.sleep(0.2)
                
                st.success(f"Successfully processed {len(extracted_claims)} receipts!")
                st.download_button(
                    label="📥 Download My Claim Excel",
                    data=excel_file,
                    file_name=f"Petrol_Claim_{my_name}_{my_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    elif password != "":
        st.error("Incorrect password! Please try again.")
    else:
        st.info("Please enter the password to access your secure portal.")
