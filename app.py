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

# Parse receipt focusing on Date, Receipt No, and Litres
def parse_receipt_with_groq(image_bytes, api_key):
    try:
        client = Groq(api_key=api_key)
        encoded_image = base64.b64encode(image_bytes).decode('utf-8')
        
        prompt = """
        You are an expert financial OCR assistant for petrol receipts in Malaysia.
        Analyze this petrol receipt image and extract the following 3 fields accurately in JSON format:
        {
          "date": "The transaction date on the receipt in DD/MM/YYYY or YYYY-MM-DD format (string or null)",
          "receipt_no": "The transaction reference number or receipt number usually located near the top (e.g., 8425439_20260830_IPFI295)",
          "litres": 0.00
        }
        Return ONLY valid JSON. If any field is not found, put null for date and receipt_no, and 0.00 for litres.
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
        return {"date": "-", "receipt_no": "-", "litres": 0.0}

# Generate Excel Claim File with smart deduplication and RM1.99 calculation
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

    # Smart Deduplication: Only consider exact duplicate if Date AND Litres AND Receipt No are all identical.
    # Different date or different litres means they are separate valid transactions even if receipt_no had an issue.
    seen_transactions = set()
    unique_claims = []
    
    for item in raw_claim_list:
        date_str = str(item.get('date', '-')).strip()
        r_no = str(item.get('receipt_no', '-')).strip()
        if not r_no or r_no == '' or r_no == 'nan':
            r_no = '-'
            
        try:
            litres = float(item.get('litres', 0) or 0)
        except:
            litres = 0.0
            
        # Unique signature: receipt_no + date + litres
        signature = f"{r_no}_{date_str}_{litres}"
        
        # If receipt_no is '-', we don't treat it as duplicate unless date and litres are also 100% identical to an existing one
        if r_no != '-' and r_no != 'Error':
            if signature in seen_transactions:
                continue
            seen_transactions.add(signature)
        
        amount_rm = round(litres * 1.99, 2)
        
        unique_claims.append({
            "Filename": item.get("Filename", "Receipt"),
            "date": date_str,
            "receipt_no": r_no,
            "litres": litres,
            "amount_rm": amount_rm
        })

    # Fill multiple receipts starting from row 15 downwards
    start_row = 15
    for idx, item in enumerate(unique_claims):
        current_row = start_row + idx
        
        ws.cell(row=current_row, column=7).value = item['receipt_no']
        ws.cell(row=current_row, column=9).value = item['litres']
        ws.cell(row=current_row, column=11).value = item['amount_rm']

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, unique_claims

last_month_value = get_last_month_str()

# ==================== 1. Colleague Portal ====================
if portal_mode == "👥 Colleague Portal":
    st.title("👥 Fuel Reimbursement Claim Form - Colleague Portal")
    st.markdown("Please enter your personal details and **upload up to 20 receipt images**. Different dates or litres are treated as unique transactions. You can edit any missing reference numbers in the preview table.")

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
                            "date": receipt_info.get("date", "-"),
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": receipt_info.get("litres", 0)
                        })
                
                status.update(label="Done!", state="complete", expanded=False)
            
            st.session_state.raw_claims = raw_claims
            st.session_state.profile = {
                "name": c_name,
                "employee_no": c_emp_no,
                "department": c_dept,
                "designation": c_designation,
                "vehicle_no": c_vehicle,
                "monthly_limit": c_limit,
                "month": c_month
            }

    if "raw_claims" in st.session_state and st.session_state.raw_claims:
        st.markdown("### 📊 Batch Extraction Preview & Editor")
        st.info("💡 **Tip**: Check the extracted **Date**, **Receipt No**, and **Litres**. If any receipt number was read as `-`, you can **directly click and type** the correct number in the table below!")
        
        _, initial_processed = generate_excel_claim(st.session_state.profile, st.session_state.raw_claims)
        
        edited_df = st.data_editor(
            pd.DataFrame(initial_processed),
            num_rows="dynamic",
            use_container_width=True,
            key="colleague_editor"
        )
        
        final_claims = edited_df.to_dict('records')
        excel_file, _ = generate_excel_claim(st.session_state.profile, final_claims)
        
        st.download_button(
            label="📥 Download Final Verified Claim Excel Sheet",
            data=excel_file,
            file_name=f"Petrol_Claim_{st.session_state.profile['name']}.xlsx",
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
                                "date": receipt_info.get("date", "-"),
                                "receipt_no": receipt_info.get("receipt_no", "-"),
                                "litres": receipt_info.get("litres", 0)
                            })
                    
                    status.update(label="Done!", state="complete", expanded=False)
                
                st.session_state.my_raw_claims = raw_claims
                st.session_state.my_profile = {
                    "name": my_name,
                    "employee_no": my_emp_no,
                    "department": my_dept,
                    "designation": my_designation,
                    "vehicle_no": my_vehicle,
                    "monthly_limit": my_limit,
                    "month": my_month
                }

        if "my_raw_claims" in st.session_state and st.session_state.my_raw_claims:
            st.markdown("### 📊 Batch Extraction Preview & Editor")
            st.info("💡 **Tip**: Check the extracted **Date**, **Receipt No**, and **Litres**. If any receipt number was read as `-`, you can **directly click and type** the correct number in the table below!")
            
            _, initial_processed_my = generate_excel_claim(st.session_state.my_profile, st.session_state.my_raw_claims)
            
            edited_df_my = st.data_editor(
                pd.DataFrame(initial_processed_my),
                num_rows="dynamic",
                use_container_width=True,
                key="owner_editor"
            )
            
            final_claims_my = edited_df_my.to_dict('records')
            excel_file_my, _ = generate_excel_claim(st.session_state.my_profile, final_claims_my)
            
            st.download_button(
                label="📥 Download My Verified Claim Excel",
                data=excel_file_my,
                file_name=f"Petrol_Claim_{st.session_state.my_profile['name']}_{st.session_state.my_profile['month']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    elif password != "":
        st.error("Incorrect password! Please try again.")
    else:
        st.info("请输入密码以进入你的专属后台。")
