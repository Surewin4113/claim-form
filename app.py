import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64
import glob
import datetime
import time
import re

# 页面配置
st.set_page_config(
    page_title="Petrol Claim Portal",
    page_icon="⛽",
    layout="wide"
)

# 侧边栏：API Key 与页面选择
st.sidebar.title("⛽ Petrol Claim Portal")
st.sidebar.markdown("---")

groq_api_key = ""
if "GROQ_API_KEY" in st.secrets:
    groq_api_key = st.secrets["GROQ_API_KEY"]
else:
    groq_api_key = st.sidebar.text_input("Groq API Key", type="password", help="请从 console.groq.com 获取 API Key")

st.sidebar.markdown("---")
portal_mode = st.sidebar.radio("选择访问页面", ["👥 同事报销页面", "🔑 个人专属后台"])

# 自动寻找 Excel 模板
def get_template_path():
    xlsx_files = glob.glob("*.xlsx")
    if xlsx_files:
        return xlsx_files[0]
    return "Fuel Reimbursement Claim Form new- Original - Copy.xlsx.xlsx"

# 自动计算上个月份
def get_last_month_str():
    today = datetime.date.today()
    first_day = today.replace(day=1)
    last_month = first_day - datetime.timedelta(days=1)
    return last_month.strftime("%B %Y")

# 带 429 自动重试与限速保护的解析函数
def parse_receipt_with_groq(image_bytes, api_key):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            client = Groq(api_key=api_key)
            encoded_image = base64.b64encode(image_bytes).decode('utf-8')
            
            prompt = """
            You are an expert financial OCR assistant for petrol receipts in Malaysia.
            Analyze this petrol receipt image and extract the following 3 fields accurately:
            1. date (transaction date in DD/MM/YYYY or YYYY-MM-DD format)
            2. receipt_no (transaction reference number or receipt number near the top, e.g. 8425439_20260830_IPFI295)
            3. litres (numeric value greater than 0)
            
            You MUST output ONLY a valid JSON object in this exact format, with no other text:
            {"date": "YYYY-MM-DD", "receipt_no": "...", "litres": 0.00}
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
                ]
            )
            
            content = chat_completion.choices[0].message.content.strip()
            
            match = re.search(r'\{.*\}', content, re.DOTALL)
            if match:
                content = match.group(0)
            
            data = json.loads(content)
            r_no = str(data.get("receipt_no", "")).strip()
            date_val = str(data.get("date", "")).strip()
            
            try:
                litres = float(data.get("litres", 0) or 0)
            except:
                litres = 0.0
            
            if litres > 0:
                return {
                    "date": date_val if date_val and date_val not in ["null", "None", "", "-"] else "-",
                    "receipt_no": r_no if r_no and r_no not in ["null", "None", "", "-"] else "-",
                    "litres": litres
                }
            else:
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                else:
                    return {
                        "date": date_val if date_val else "-",
                        "receipt_no": r_no if r_no else "-",
                        "litres": litres
                    }
        except Exception as e:
            error_msg = str(e)
            # 如果触发了 429 频次限制，自动多等待几秒再重试
            if "429" in error_msg or "rate_limit" in error_msg.lower():
                if attempt < max_retries - 1:
                    time.sleep(5) # 触发限速时多暂停 5 秒
                    continue
            
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
            else:
                return {"date": "-", "receipt_no": f"Error: {error_msg[:30]}", "litres": 0.0}
    return {"date": "-", "receipt_no": "-", "litres": 0.0}

# 生成 Excel 报销表（智能去重 + 严格按 1 litre = RM1.99 计算）
def generate_excel_claim(profile_data, raw_claim_list):
    template_path = get_template_path()
    wb = openpyxl.load_workbook(template_path)
    
    if "Revised 2" in wb.sheetnames:
        ws = wb["Revised 2"]
    else:
        ws = wb.active

    ws['D10'] = profile_data['name']
    ws['K10'] = profile_data['designation']
    ws['D11'] = profile_data['employee_no']
    ws['K11'] = profile_data['department']
    ws['D12'] = profile_data['vehicle_no']
    ws['K12'] = profile_data['monthly_limit']
    ws['D13'] = profile_data['month']

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
            
        signature = f"{r_no}_{date_str}_{litres}"
        
        if r_no != '-' and not r_no.startswith("Error"):
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

# ==================== 1. 同事报销页面 ====================
if portal_mode == "👥 同事报销页面":
    st.title("👥 同事油费报销申请表")
    st.markdown("请填写您的个人信息并上传多张收据。系统已内置限速保护，防止触发 429 频率限制。")

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
        uploaded_files = st.file_uploader("上传多张收据图片 (支持多选，最多20张)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        submitted = st.form_submit_button("🤖 批量智能识别并生成报销表")

    if submitted:
        if not groq_api_key:
            st.error("请先配置 Groq API Key！")
        elif not uploaded_files:
            st.error("请至少上传一张收据图片！")
        else:
            total_files = len(uploaded_files)
            raw_claims = []
            
            with st.status(f"正在安全处理共 {total_files} 张收据（内置限速保护）...", expanded=True) as status:
                for i, file in enumerate(uploaded_files):
                    status.update(label=f"正在处理第 {i+1} 张 / 共 {total_files} 张 ({file.name})...")
                    
                    image_bytes = file.getvalue()
                    receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                    
                    if receipt_info:
                        raw_claims.append({
                            "Filename": file.name,
                            "date": receipt_info.get("date", "-"),
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": receipt_info.get("litres", 0)
                        })
                    # 每张图片之间暂停 2.5 秒，完美避开 Groq 429 限速
                    time.sleep(2.5)
                
                status.update(label="处理完成！", state="complete", expanded=False)
            
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
        st.markdown("### 📊 批量识别明细预览与编辑")
        st.info("💡 **提示**：如果某张收据信息有误，可直接在下方表格中点击修改，修改后会自动重新计算金额。")
        
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
            label="📥 下载最终核对后的报销 Excel 表格",
            data=excel_file,
            file_name=f"Petrol_Claim_{st.session_state.profile['name']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==================== 2. 个人专属后台 ====================
elif portal_mode == "🔑 个人专属后台":
    st.title("🔑 Soo Wai Wing 的专属报销后台")
    
    password = st.text_input("请输入访问密码", type="password")
    
    if password == "P@ssw0rd":
        st.success("密码验证成功！欢迎回来。")
        
        st.markdown("### 📋 预设档案信息")
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
        my_uploaded_files = st.file_uploader("上传多张油站收据 (支持多选，最多20张)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="my_receipts")
        
        if st.button("🚀 批量安全处理我的所有收据"):
            if not groq_api_key:
                st.error("请先配置 Groq API Key！")
            elif not my_uploaded_files:
                st.error("请至少上传一张收据！")
            else:
                total_files = len(my_uploaded_files)
                raw_claims = []
                
                with st.status(f"正在安全处理共 {total_files} 张收据（内置限速保护）...", expanded=True) as status:
                    for i, file in enumerate(my_uploaded_files):
                        status.update(label=f"正在处理第 {i+1} 张 / 共 {total_files} 张 ({file.name})...")
                        
                        image_bytes = file.getvalue()
                        receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                        
                        if receipt_info:
                            raw_claims.append({
                                "Filename": file.name,
                                "date": receipt_info.get("date", "-"),
                                "receipt_no": receipt_info.get("receipt_no", "-"),
                                "litres": receipt_info.get("litres", 0)
                            })
                        time.sleep(2.5) # 限速保护
                    
                    status.update(label="处理完成！", state="complete", expanded=False)
                
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
            st.markdown("### 📊 批量识别明细预览与编辑")
            st.info("💡 **提示**：如果某张收据信息有误，可直接在下方表格中点击修改，修改后会自动重新计算金额。")
            
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
                label="📥 下载我的专属报销 Excel 表格",
                data=excel_file_my,
                file_name=f"Petrol_Claim_{st.session_state.my_profile['name']}_{st.session_state.my_profile['month']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    elif password != "":
        st.error("密码错误！请重新输入。")
    else:
        st.info("请输入密码以进入你的专属后台。")
