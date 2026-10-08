import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64
import glob
import time

# 页面配置
st.set_page_config(
    page_title="Petrol Claim Portal",
    page_icon="⛽",
    layout="wide"
)

# 侧边栏：API Key 与 页面选择
st.sidebar.title("⛽ Petrol Claim Portal")
st.sidebar.markdown("---")

groq_api_key = ""
if "GROQ_API_KEY" in st.secrets:
    groq_api_key = st.secrets["GROQ_API_KEY"]
else:
    groq_api_key = st.sidebar.text_input("Groq API Key", type="password", help="请从 console.groq.com 获取免费 API Key")

st.sidebar.markdown("---")
portal_mode = st.sidebar.radio("选择访问页面", ["👥 同事报销页面", "🔑 个人专属后台"])

# 自动寻找 Excel 模板
def get_template_path():
    xlsx_files = glob.glob("*.xlsx")
    if xlsx_files:
        return xlsx_files[0]
    return "Fuel Reimbursement Claim Form new- Original - Copy.xlsx.xlsx"

# 单张收据解析函数（精准提取信用卡支付金额，避开补贴/总价）
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
          "amount_rm": "The exact credit card charged amount, cash payment amount, or final amount paid by card. Do NOT pick total price before discount, and do NOT pick subsidy or MADANI amounts. Pick the exact final amount charged to portal."
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
        return {"receipt_no": "Error", "litres": 0.0, "amount_rm": 0.0}

# 生成支持多行填充的 Excel 报销表
def generate_excel_claim(profile_data, claim_data_list):
    template_path = get_template_path()
    wb = openpyxl.load_workbook(template_path)
    
    if "Revised 2" in wb.sheetnames:
        ws = wb["Revised 2"]
    else:
        ws = wb.active

    # 填入员工个人信息
    ws['D10'] = profile_data['name']
    ws['K10'] = profile_data['designation']
    ws['D11'] = profile_data['employee_no']
    ws['K11'] = profile_data['department']
    ws['D12'] = profile_data['vehicle_no']
    ws['K12'] = profile_data['monthly_limit']
    ws['D13'] = profile_data['month']

    # 从第 15 行开始，循环填入多张收据数据
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

# ==================== 1. 同事报销页面 ====================
if portal_mode == "👥 同事报销页面":
    st.title("👥 同事油费报销申请表")
    st.markdown("请填写您的个人信息并**一次性上传最多 20 张收据**。系统将逐张安全解析并在下方展示识别明细。")

    with st.form("colleague_form"):
        col1, col2 = st.columns(2)
        with col1:
            c_name = st.text_input("Employee Name (员工姓名)")
            c_emp_no = st.text_input("Employee No. (员工编号)")
            c_dept = st.text_input("Department (部门)")
        with col2:
            c_designation = st.text_input("Designation (职位)")
            c_vehicle = st.text_input("Vehicle No. (车牌号)")
            c_limit = st.number_input("Monthly Claim Limit (RM)", value=500.0)
        
        c_month = st.text_input("Claim for the Month of (报销月份，例: October 2026)")
        uploaded_files = st.file_uploader("上传多张油站收据图片 (支持多选，最多20张)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        
        submitted = st.form_submit_button("🤖 批量智能识别并生成报销表")

    if submitted:
        if not groq_api_key:
            st.error("请先配置 Groq API Key！")
        elif not uploaded_files:
            st.error("请至少上传一张收据图片！")
        else:
            total_files = len(uploaded_files)
            extracted_claims = []
            
            # 使用现代化折叠状态面板，展示 20 张收据的实时处理进度
            with st.status(f"正在批量处理 0 / {total_files} 张收据...", expanded=True) as status:
                for i, file in enumerate(uploaded_files):
                    status.update(label=f"正在处理第 {i+1} 张 / 共 {total_files} 张 ({file.name})...")
                    
                    image_bytes = file.getvalue()
                    receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                    
                    if receipt_info:
                        extracted_claims.append({
                            "Filename": file.name,
                            "receipt_no": receipt_info.get("receipt_no", "-"),
                            "litres": float(receipt_info.get("litres", 0) or 0),
                            "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                        })
                    
                    # 稳定限速：每张间隔 1 秒，确保 20 张顺利通过不报错
                    time.sleep(1)
                
                status.update(label=f"成功解析全部 {total_files} 张收据！", state="complete", expanded=False)
            
            # 展示识别结果明细表格，让你亲眼确认 20 张全部读出来了
            st.markdown("### 📊 批量识别明细预览")
            st.dataframe(pd.DataFrame(extracted_claims), use_container_width=True)
            
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
            
            st.success(f"成功将 {len(extracted_claims)} 张收据填入 Excel 表格第 15 行起的多行中！")
            st.download_button(
                label="📥 下载完整报销 Excel 表格",
                data=excel_file,
                file_name=f"Petrol_Claim_{c_name}.xlsx",
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
            
        my_month = st.text_input("Claim for the Month of", value="October 2026")
        
        st.markdown("---")
        my_uploaded_files = st.file_uploader("上传多张油站收据 (支持多选，最多20张)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="my_receipts")
        
        if st.button("🚀 批量处理我的所有收据"):
            if not groq_api_key:
                st.error("请先配置 Groq API Key！")
            elif not my_uploaded_files:
                st.error("请至少上传一张收据！")
            else:
                total_files = len(my_uploaded_files)
                extracted_claims = []
                
                with st.status(f"正在批量处理 0 / {total_files} 张收据...", expanded=True) as status:
                    for i, file in enumerate(my_uploaded_files):
                        status.update(label=f"正在处理第 {i+1} 张 / 共 {total_files} 张 ({file.name})...")
                        
                        image_bytes = file.getvalue()
                        receipt_info = parse_receipt_with_groq(image_bytes, groq_api_key)
                        
                        if receipt_info:
                            extracted_claims.append({
                                "Filename": file.name,
                                "receipt_no": receipt_info.get("receipt_no", "-"),
                                "litres": float(receipt_info.get("litres", 0) or 0),
                                "amount_rm": float(receipt_info.get("amount_rm", 0) or 0)
                            })
                        
                        time.sleep(1)
                    
                    status.update(label=f"成功解析全部 {total_files} 张收据！", state="complete", expanded=False)
                
                st.markdown("### 📊 批量识别明细预览")
                st.dataframe(pd.DataFrame(extracted_claims), use_container_width=True)
                
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
                
                st.success(f"成功将 {len(extracted_claims)} 张收据填入 Excel 表格！")
                st.download_button(
                    label="📥 下载我的专属多行报销 Excel",
                    data=excel_file,
                    file_name=f"Petrol_Claim_{my_name}_{my_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    elif password != "":
        st.error("密码错误！请重新输入。")
    else:
        st.info("请输入密码以进入你的专属后台。")
