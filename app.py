import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64
import glob
import datetime

# Page Configuration
st.set_page_config(
    page_title="Petrol Claim Portal",
    page_icon="⛽",
    layout="wide"
)

# Sidebar: API Key and Portal Selection
st.sidebar.title("⛽ Petrol Claim Portal")
st.sidebar.markdown("---")

groq_api_key = ""
if "GROQ_API_KEY" in st.secrets:
    groq_api_key = st.secrets["GROQ_API_KEY"]
else:
    groq_api_key = st.sidebar.text_input("Groq API Key", type="password", help="Get your free API Key from console.groq.com")

st.sidebar.markdown("---")
portal_mode = st.sidebar.radio("Select Portal", ["👥 Colleague Portal", "🔑 Owner Portal"])

# Automatically find Excel template
def get_template_path():
    xlsx_files = glob.glob("*.xlsx")
    if xlsx_files:
        return xlsx_files[0]
    return "Fuel Reimbursement Claim Form new- Original - Copy.xlsx.xlsx"

# Automatically calculate last month (e.g., September 2026 if current is October 2026)
def get_last_month_str():
    today = datetime.date.today()
    first_day = today.replace(day=1)
    last_month = first_day - datetime.timedelta(days=1)
    return last_month.strftime("%B %Y")

# Parse receipt focusing purely on Receipt No and Litres (Amount is calculated precisely via 1.99 rate)
def parse_receipt_with_groq(image_bytes, api_key):
    try:
        client = Groq(api_key=api_key)
        encoded_image = base64.b64encode(image_bytes).decode('utf-8')
        
        prompt = """
        You are an expert financial OCR assistant for petrol receipts in Malaysia.
        Analyze this petrol receipt image and extract ONLY the following 2 core fields accurately in JSON format:
        {
          "receipt_no": "The transaction reference number or receipt number usually located near the top (e.g., 8425439_20260830_IPFI295)",
          "litres": 0.00
        }
        Return ONLY valid JSON. If any field is not found, put null for receipt_no and 0.00 for litres.
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
        
        content = chat_completion.choices[0].message.content.strip()
        if content.startswith("```json"):
            content = content[7:]
        if content.startswith("```"):
            content = content[3:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()
        
        return json.loads(content)
    except Exception as e:
        return {"receipt_no": "-", "litres": 0.0}

# Generate Excel Claim File with duplicate check and strict Litres x 1.99 calculation
def generate_excel_claim(profile_data, raw_claim_list):
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

    # Anti-Duplicate & Filtering logic (preventing duplicate receipt numbers)
    seen_receipts = set()
    unique_claims = []
    
    for item in raw_claim_list:
        r_no = item.get('receipt_no', '-')
        if r_no and r_no != '-' and r_no != 'Error':
            if r_no in seen_receipts:
                continue # Skip duplicate receipt
            seen_receipts.add(r_no)
        
        litres = float(item.get('litres', 0) or 0)
        # Strict calculation: Claim Amount = Litres * 1.99 (rounded to 2 decimal places)
        amount_rm = round(litres * 1.99, 2)
        
        unique_claims.append({
            "Filename": item.get("Filename", "Receipt"),
            "receipt_no": r_no,
            "litres": litres,
            "amount_rm": amount_rm
        })

    # Fill multiple receipts starting from row 15 downwards
    start_row = 15
    for idx, item in enumerate(unique_claims):
        current_row = start_row + idx
        
        # Receipt No (Column 7)
        ws.cell(row=current_row, column=7).value = item['receipt_no']
        
        # Litres (Column 9)
        ws.cell(row=current_row, column=9).value = item['litres']
        
        # Claim Amount calculated precisely at 1.99 per litre (Column 11)
        ws.cell(row=current_row, column=11).value = item['amount_rm']

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, unique_claims

last_month_value = get_last_month_str()

# ==================== 1. Colleague Portal ====================
if portal_mode == "👥 Colleague Portal":
    st.title("👥 Fuel Reimbursement Claim Form - Colleague Portal")
    st.markdown("Please enter your personal details and **upload up to 20 receipt images**. Automatic duplicate prevention and RM1.99/litre calculation applied.")

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
        
        c_month = st.text_input("Claim for the Month of", value=last_month_value, key="colleague_month_input")
        uploaded_files = st.file_uploader("Upload Receipt Images (Multiple allowed, up to 20)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        submitted = st.form_submit_button("🤖 Batch Process & Generate Claim")

    if submitted:
        if not groq_api_key:
            st.error("Please configure your Groq API Key!")
        elif not uploaded_files:
            st.error("Please upload at least one receipt image!")
        else:
            total_files = len(uploaded_files)
            raw_claims = []
            
            with st.status(f"Processing {total_files} receipts securely...", expanded=True) as status:
                for i, file in enumerate(uploaded_files):
                    status.update(label=f"Processing receipt {i+1} of {total_files} ({file.name})...")
                    
                    image_bytes = file.getvalue()
                    receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                    
                    if receipt_info:
                        raw_claims.append({
                            "Filename": file.name,
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": receipt_info.get("litres", 0)
                        })
                
                status.update(label="Applying duplicate checks and RM1.99 calculation...", state="complete", expanded=False)
            
            profile = {
                "name": c_name,
                "employee_no": c_emp_no,
                "department": c_dept,
                "designation": c_designation,
                "vehicle_no": c_vehicle,
                "monthly_limit": c_limit,
                "month": c_month
            }
            
            excel_file, processed_claims = generate_excel_claim(profile, raw_claims)
            
            st.markdown("### 📊 Batch Extraction Preview (Anti-Duplicate & Auto-Calculated)")
            st.dataframe(pd.DataFrame(processed_claims), use_container_width=True)
            
            st.success(f"Successfully processed {len(processed_claims)} unique receipts (Duplicates removed if any).")
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
            
        my_month = st.text_input("Claim for the Month of", value=last_month_value, key="owner_month_input")
        
        st.markdown("---")
        my_uploaded_files = st.file_uploader("Upload Fuel Receipts (Multiple allowed, up to 20)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="my_receipts")
        
        if st.button("🚀 Batch Process My Receipts"):
            if not groq_api_key:
                st.error("Please configure your Groq API Key!")
            elif not my_uploaded_files:
                st.error("Please upload at least one receipt!")
            else:
                total_files = len(my_uploaded_files)
                raw_claims = []
                
                with st.status(f"Processing {total_files} receipts securely...", expanded=True) as status:
                    for i, file in enumerate(my_uploaded_files):
                        status.update(label=f"Processing receipt {i+1} of {total_files} ({file.name})...")
                        
                        image_bytes = file.getvalue()
                        receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                        
                        if receipt_info:
                            raw_claims.append({
                                "Filename": file.name,
                                "receipt_no": receipt_info.get("receipt_no", "-"),
                                "litres": receipt_info.get("litres", 0)
                            })
                    
                    status.update(label="Applying duplicate checks and RM1.99 calculation...", state="complete", expanded=False)
                
                profile = {
                    "name": my_name,
                    "employee_no": my_emp_no,
                    "department": my_dept,
                    "designation": my_designation,
                    "vehicle_no": my_vehicle,
                    "monthly_limit": my_limit,
                    "month": my_month
                }
                
                excel_file, processed_claims = generate_excel_claim(profile, raw_claims)
                
                st.markdown("### 📊 Batch Extraction Preview (Anti-Duplicate & Auto-Calculated)")
                st.dataframe(pd.DataFrame(processed_claims), use_container_width=True)
                
                st.success(f"Successfully processed {len(processed_claims)} unique receipts (Duplicates removed if any).")
                st.download_button(
                    label="📥 Download My Claim Excel",
                    data=excel_file,
                    file_name=f"Petrol_Claim_{my_name}_{my_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    elif password != "":
        st.error("Incorrect password! Please try again.")
    else:
        st.info("请输入密码以进入你的专属后台。")
