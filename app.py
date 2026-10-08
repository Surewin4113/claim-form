import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
import io
import json
import base64

def parse_receipt_with_groq(image_bytes, api_key):
    client = Groq(api_key=api_key)
    
    # 将上传的图片转换为 Base64 编码
    encoded_image = base64.b64encode(image_bytes).decode('utf-8')
    
    # 专门针对马来西亚油站收据优化的 Prompt
    prompt = """
    You are an expert OCR and financial data extraction assistant for petrol receipts in Malaysia.
    Analyze this petrol receipt image and extract the following 3 core fields accurately:
    1. "receipt_no": The receipt number, invoice number, or transaction reference number (e.g., Receipt No, Txn ID, Ref No).
    2. "litres": The total volume of fuel pumped in litres (e.g., float number like 45.50). If not explicitly stated, calculate it or look for 'L' / 'Litres'.
    3. "amount_rm": The total amount paid in RM (Ringgit Malaysia), usually labeled as Total, Amount, Cash, or Credit Card amount (e.g., float number like 95.00).

    Return ONLY a valid JSON object in this exact format:
    {
      "receipt_no": "string or null",
      "litres": 0.00,
      "amount_rm": 0.00
    }
    Do not include any extra markdown formatting or conversational text outside the JSON.
    """
    
    try:
        chat_completion = client.chat.completions.create(
            model="llama-3.2-90b-vision-preview", # Groq 免费且强大的多模态视觉模型
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
        
        # 解析 AI 返回的 JSON 结果
        result_text = chat_completion.choices[0].message.content
        data = json.loads(result_text)
        return data
        
    except Exception as e:
        print(f"Error parsing receipt with Groq AI: {e}")
        return {
            "receipt_no": "",
            "litres": 0.0,
            "amount_rm": 0.0
        }
