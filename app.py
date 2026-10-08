import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64
import glob

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

# Parse receipt using Groq Vision API (Using 11b-vision-preview for broad compatibility)
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
          "amount_rm": 0.00
        }
        Return ONLY valid JSON. If any field is not found, put null for receipt_no and 0.00 for numbers.
        """
        
        chat_completion = client.chat.completions.create(
            model="llama-3.2-11b-vision-preview", # Fully supported multi-modal vision model on Groq
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
        st.error(f"Error communicating with Groq API: {e}")
        return None

# Generate Excel Claim File
def generate_excel_claim(profile_data, claim_data):
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
    ws['D13'] = claim_data['month']

    # Fill Receipt and Claim Details (Row 15)
    ws.cell(row=15, column=7).value = claim_data['receipt_no']
    ws.cell(row=15, column=9).value = claim_data['litres']
    ws.cell(row=15, column=11).value = claim_data['amount_rm']

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==================== 1. Colleague Portal ====================
if portal_mode == "👥 Colleague Portal":
    st.title("👥 Fuel Reimbursement Claim Form - Colleague Portal")
    st.markdown("Please enter your personal details and upload your fuel receipt. The system will automatically extract details and generate your claim Excel sheet.")

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
        uploaded_file = st.file_uploader("Upload Receipt Image", type=["jpg", "jpeg", "png"])
        
        submitted = st.form_submit_button("🤖 Auto-Parse & Generate Claim")

    if submitted:
        if not groq_api_key:
            st.error("Please configure your Groq API Key in Streamlit Secrets or enter it in the sidebar!")
        elif not uploaded_file:
            st.error("Please upload a receipt image first!")
        else:
            with st.spinner("AI is processing your receipt..."):
                image_bytes = uploaded_file.getvalue()
                receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                
                if receipt_info:
                    st.success("Receipt parsed successfully!")
                    st.json(receipt_info)
                    
                    profile = {
                        "name": c_name,
                        "employee_no": c_emp_no,
                        "department": c_dept,
                        "designation": c_designation,
                        "vehicle_no": c_vehicle,
                        "monthly_limit": c_limit
                    }
                    
                    claim = {
                        "month": c_month,
                        "receipt_no": receipt_info.get("receipt_no", "-"),
                        "litres": float(receipt_info.get("litres", 0) or 0),
                        "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                    }
                    
                    excel_file = generate_excel_claim(profile, claim)
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
        my_uploaded_file = st.file_uploader("Upload Fuel Receipt", type=["jpg", "jpeg", "png"], key="my_receipt")
        
        if st.button("🚀 Process My Claim"):
            if not groq_api_key:
                st.error("Please configure your Groq API Key in Streamlit Secrets or enter it in the sidebar!")
            elif not my_uploaded_file:
                st.error("Please upload a receipt!")
            else:
                with st.spinner("Groq AI is analyzing your receipt..."):
                    image_bytes = my_uploaded_file.getvalue()
                    receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                    
                    if receipt_info:
                        st.success("Extraction Result:")
                        st.json(receipt_info)
                        
                        profile = {
                            "name": my_name,
                            "employee_no": my_emp_no,
                            "department": my_dept,
                            "designation": my_designation,
                            "vehicle_no": my_vehicle,
                            "monthly_limit": my_limit
                        }
                        
                        claim = {
                            "month": my_month,
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": float(receipt_info.get("litres", 0) or 0),
                            "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                        }
                        
                        excel_file = generate_excel_claim(profile, claim)
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
