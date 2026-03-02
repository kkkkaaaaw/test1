#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文自动生成程序
功能：读取输入数据，调用AI生成公文，发送邮件
"""

import os
import sys
import json
import time
import datetime
import smtplib
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import requests

# ==========================================
# 配置部分
# ==========================================

# 从环境变量读取配置
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_BASE_URL = " https://generativelanguage.googleapis.com/v1beta/models "

# 公文配置
DOCUMENT_TYPE = os.getenv("DOCUMENT_TYPE", "通知")
DOCUMENT_TITLE = os.getenv("DOCUMENT_TITLE", "关于XX工作的通知")
DOCUMENT_SUBJECT = os.getenv("DOCUMENT_SUBJECT", "工作安排")
DOCUMENT_SENDER = os.getenv("DOCUMENT_SENDER", "公司办公室")
DOCUMENT_RECEIVER = os.getenv("DOCUMENT_RECEIVER", "各部门、各子公司")
DOCUMENT_URGENCY = os.getenv("DOCUMENT_URGENCY", "普通")
DOCUMENT_SECURITY = os.getenv("DOCUMENT_SECURITY", "无")

# 输入文件
INPUT_DATA_FILE = os.getenv("INPUT_DATA_FILE", "input_data.txt")
EXTRA_INSTRUCTIONS = os.getenv("EXTRA_INSTRUCTIONS", "")

# 邮件配置
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
EMAIL_RECEIVERS = os.getenv("EMAIL_RECEIVERS", "")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.qq.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))

# 日期
TODAY_STR = datetime.datetime.now().strftime("%Y年%m月%d日")
TIMESTAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# ==========================================
# 工具函数
# ==========================================

def log_info(message):
    """记录信息日志"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[INFO] {timestamp} - {message}")

def log_error(message, error=None):
    """记录错误日志"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[ERROR] {timestamp} - {message}")
    if error:
        print(f"错误详情: {str(error)}")
        traceback.print_exc()

def check_required_env():
    """检查必需的环境变量"""
    required_vars = {
        "GEMINI_API_KEY": GEMINI_API_KEY,
        "EMAIL_SENDER": EMAIL_SENDER,
        "EMAIL_PASSWORD": EMAIL_PASSWORD,
        "EMAIL_RECEIVERS": EMAIL_RECEIVERS,
    }
    
    missing_vars = []
    for name, value in required_vars.items():
        if not value or value.strip() == "":
            missing_vars.append(name)
    
    if missing_vars:
        log_error(f"缺少必需的环境变量: {', '.join(missing_vars)}")
        log_error("请在GitHub Secrets中配置这些变量")
        return False
    
    return True

def read_input_data():
    """读取输入数据文件"""
    try:
        if not os.path.exists(INPUT_DATA_FILE):
            log_error(f"输入文件不存在: {INPUT_DATA_FILE}")
            return "无输入数据，请根据公文类型和标题生成相应内容。"
        
        with open(INPUT_DATA_FILE, 'r', encoding='utf-8') as f:
            content = f.read().strip()
        
        if not content:
            log_error("输入文件内容为空")
            return "无输入数据，请根据公文类型和标题生成相应内容。"
        
        log_info(f"成功读取输入文件: {INPUT_DATA_FILE}，长度: {len(content)}字符")
        return content
    except Exception as e:
        log_error(f"读取输入文件失败: {INPUT_DATA_FILE}", e)
        return "读取输入数据失败，请根据公文类型和标题生成相应内容。"

# ==========================================
# Gemini API调用函数
# ==========================================

def call_gemini_api(prompt, max_retries=3):
    """调用Gemini API生成内容"""
    if not GEMINI_API_KEY:
        log_error("Gemini API密钥未配置")
        return None
    
    url = f"{GEMINI_BASE_URL}/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    
    headers = {
        "Content-Type": "application/json",
    }
    
    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.3,
            "topP": 0.8,
            "topK": 40,
            "maxOutputTokens": 4096,
        },
        "safetySettings": [
            {
                "category": "HARM_CATEGORY_HARASSMENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_HATE_SPEECH",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            }
        ]
    }
    
    for attempt in range(max_retries):
        try:
            log_info(f"调用Gemini API (尝试 {attempt + 1}/{max_retries})...")
            
            response = requests.post(
                url,
                headers=headers,
                json=payload,
                timeout=60
            )
            
            if response.status_code == 200:
                result = response.json()
                if "candidates" in result and len(result["candidates"]) > 0:
                    text = result["candidates"]["content"]["parts"]["text"]
                    log_info("Gemini API调用成功")
                    return text
                else:
                    log_error(f"API响应格式异常: {result}")
            else:
                log_error(f"API调用失败，状态码: {response.status_code}")
                log_error(f"响应内容: {response.text}")
                
                if response.status_code == 429:
                    wait_time = (attempt + 1) * 10
                    log_info(f"API限流，等待{wait_time}秒后重试...")
                    time.sleep(wait_time)
                elif response.status_code >= 500:
                    wait_time = (attempt + 1) * 5
                    log_info(f"服务器错误，等待{wait_time}秒后重试...")
                    time.sleep(wait_time)
                else:
                    break  # 客户端错误，不重试
                    
        except requests.exceptions.Timeout:
            log_error(f"API请求超时 (尝试 {attempt + 1})")
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 5
                log_info(f"等待{wait_time}秒后重试...")
                time.sleep(wait_time)
        except requests.exceptions.ConnectionError:
            log_error(f"网络连接错误 (尝试 {attempt + 1})")
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 10
                log_info(f"等待{wait_time}秒后重试...")
                time.sleep(wait_time)
        except Exception as e:
            log_error(f"API调用异常 (尝试 {attempt + 1})", e)
            break
    
    return None

# ==========================================
# 公文生成函数
# ==========================================

def generate_document_prompt(input_data):
    """生成公文提示词"""
    
    # 根据公文类型选择不同的提示词模板
    prompt_templates = {
        "通知": """
你是一个专业的公文写作助手。请根据以下信息生成一份正式的通知：

【公文基本信息】
公文类型：通知
标题：{title}
发文单位：{sender}
主送机关：{receiver}
事由：{subject}
紧急程度：{urgency}
密级：{security}
日期：{date}

【公文格式要求】
1. 严格按照《党政机关公文格式》国家标准（GB/T 9704-2012）撰写
2. 结构必须包含：标题、主送机关、正文、发文机关署名、成文日期
3. 正文内容应分层次叙述，逻辑清晰
4. 语言要求：庄重、准确、简练、规范
5. 使用公文专用语：如"现将有关事项通知如下"、"特此通知"等

【输入材料】
{input_data}

【生成要求】
请基于以上输入材料，生成一份完整的通知。要求：
1. 标题使用二号小标宋体（在文本中用【标题】标记）
2. 主送机关使用三号仿宋体
3. 正文使用三号仿宋体，段落首行缩进2字符
4. 一级标题使用三号黑体，二级标题使用三号楷体
5. 落款包括发文机关和成文日期
6. 如有附件，在正文后注明

【输出格式】
请输出完整的公文内容，使用以下HTML格式确保格式规范：
<div style="font-family: '仿宋', 'SimSun', serif; line-height: 1.5;">
  <div style="text-align: center; font-size: 22px; font-weight: bold; margin-bottom: 30px;">【标题】</div>
  <div style="font-size: 16px; margin-bottom: 20px;">【主送机关】</div>
  <div style="font-size: 16px; text-indent: 2em; margin-bottom: 15px;">【正文内容】</div>
  <div style="margin-top: 50px; text-align: right;">
    <div style="font-size: 16px;">【发文机关】</div>
    <div style="font-size: 16px;">【成文日期】</div>
  </div>
</div>
        """,
        
        "报告": """
你是一个专业的公文写作助手。请根据以下信息生成一份正式的报告：

【公文基本信息】
公文类型：报告
标题：{title}
发文单位：{sender}
主送机关：{receiver}
事由：{subject}
日期：{date}

【公文格式要求】
1. 报告是向上级机关汇报工作、反映情况、提出建议的公文
2. 结构：标题、主送机关、正文（引言、主体、结尾）、落款
3. 正文主体部分应分条列项，层次分明
4. 结尾使用"特此报告"或"以上报告，请审阅"等专用语

【输入材料】
{input_data}

【生成要求】
请生成一份关于{subject}的工作报告，要求：
1. 开头简要说明报告缘由
2. 主体部分详细汇报工作情况、存在问题、下一步计划
3. 数据准确，事实清楚
4. 提出明确的工作建议或请求

【输出格式】
使用规范的公文HTML格式。
        """,
        
        "请示": """
你是一个专业的公文写作助手。请根据以下信息生成一份正式的请示：

【公文基本信息】
公文类型：请示
标题：{title}
发文单位：{sender}
主送机关：{receiver}
事由：{subject}
紧急程度：{urgency}
日期：{date}

【公文格式要求】
1. 请示是向上级请求指示、批准的公文
2. 必须一事一请示
3. 结构：请示缘由、请示事项、结语
4. 结语使用"妥否，请批示"或"以上请示，请予批准"等

【输入材料】
{input_data}

【生成要求】
请生成一份关于{subject}的请示，要求：
1. 充分说明请示的理由和依据
2. 明确具体的请示事项
3. 提出明确的请求
4. 语言恳切，理由充分

【输出格式】
使用规范的公文HTML格式。
        """,
        
        "函": """
你是一个专业的公文写作助手。请根据以下信息生成一份正式的函：

【公文基本信息】
公文类型：函
标题：{title}
发文单位：{sender}
主送机关：{receiver}
事由：{subject}
日期：{date}

【公文格式要求】
1. 函适用于不相隶属机关之间商洽工作、询问和答复问题
2. 语气要得体，根据内容确定语气
3. 结构：发函缘由、事项、结语
4. 结语使用"特此函商"、"请予函复"等

【输入材料】
{input_data}

【生成要求】
请生成一份关于{subject}的函，要求：
1. 开头说明发函缘由
2. 主体明确商洽或询问的事项
3. 结尾表达期望回复的意愿
4. 语气恰当，措辞准确

【输出格式】
使用规范的公文HTML格式。
        """,
        
        "纪要": """
你是一个专业的公文写作助手。请根据以下信息生成一份正式的会议纪要：

【公文基本信息】
公文类型：会议纪要
标题：{title}
发文单位：{sender}
事由：{subject}
日期：{date}

【公文格式要求】
1. 纪要应客观、准确、完整地记录会议情况和议定事项
2. 结构：会议概况、会议内容、议定事项、出席人员
3. 使用"会议认为"、"会议决定"、"会议要求"等纪要用语
4. 议定事项应分条列项，责任明确

【输入材料】
{input_data}

【生成要求】
请生成一份关于{subject}的会议纪要，要求：
1. 记录会议时间、地点、主持人、出席人员
2. 概括会议主要内容和讨论情况
3. 明确会议议定事项和分工
4. 列出需要落实的具体工作

【输出格式】
使用规范的公文HTML格式。
        """
    }
    
    # 获取对应类型的提示词模板
    template = prompt_templates.get(DOCUMENT_TYPE, prompt_templates["通知"])
    
    # 填充模板
    prompt = template.format(
        title=DOCUMENT_TITLE,
        sender=DOCUMENT_SENDER,
        receiver=DOCUMENT_RECEIVER,
        subject=DOCUMENT_SUBJECT,
        urgency=DOCUMENT_URGENCY,
        security=DOCUMENT_SECURITY,
        date=TODAY_STR,
        input_data=input_data
    )
    
    # 添加额外指令
    if EXTRA_INSTRUCTIONS and EXTRA_INSTRUCTIONS.strip():
        prompt += f"\n\n【额外指令】\n{EXTRA_INSTRUCTIONS}"
    
    # 添加严格的格式指令
    strict_instruction = f"""
【极端严格的格式指令】
1. 必须严格按照中国党政机关公文格式标准生成
2. 标题：{DOCUMENT_TITLE}
3. 主送机关：{DOCUMENT_RECEIVER}
4. 发文机关：{DOCUMENT_SENDER}
5. 成文日期：{TODAY_STR}
6. 紧急程度：{DOCUMENT_URGENCY}
7. 密级：{DOCUMENT_SECURITY}

【绝对禁止】
1. 禁止使用Markdown格式
2. 禁止添加任何解释性文字
3. 禁止改变公文的基本结构
4. 禁止使用口语化表达

【输出要求】
只输出公文正文的HTML格式内容，不要输出任何其他文字。
"""
    
    return prompt + "\n\n" + strict_instruction

def generate_document():
    """生成公文"""
    log_info("开始生成公文...")
    log_info(f"公文类型: {DOCUMENT_TYPE}")
    log_info(f"公文标题: {DOCUMENT_TITLE}")
    log_info(f"发文单位: {DOCUMENT_SENDER}")
    log_info(f"主送机关: {DOCUMENT_RECEIVER}")
    
    # 读取输入数据
    input_data = read_input_data()
    
    # 生成提示词
    prompt = generate_document_prompt(input_data)
    log_info(f"提示词长度: {len(prompt)}字符")
    
    # 调用Gemini API
    document_content = call_gemini_api(prompt)
    
    if not document_content:
        log_error("公文生成失败")
        # 生成一个简单的错误提示公文
        document_content = f"""
<div style="font-family: '仿宋', 'SimSun', serif; line-height: 1.5;">
  <div style="text-align: center; font-size: 22px; font-weight: bold; margin-bottom: 30px;">{DOCUMENT_TITLE}</div>
  <div style="font-size: 16px; margin-bottom: 20px;">{DOCUMENT_RECEIVER}：</div>
  <div style="font-size: 16px; text-indent: 2em; margin-bottom: 15px;">
    由于系统生成公文时遇到技术问题，未能生成完整的公文内容。请根据以下信息手动撰写：
  </div>
  <div style="font-size: 16px; text-indent: 2em; margin-bottom: 15px;">
    <strong>事由：</strong>{DOCUMENT_SUBJECT}<br/>
    <strong>输入材料：</strong>{input_data[:200]}...
  </div>
  <div style="margin-top: 50px; text-align: right;">
    <div style="font-size: 16px;">{DOCUMENT_SENDER}</div>
    <div style="font-size: 16px;">{TODAY_STR}</div>
  </div>
</div>
"""
    
    # 保存公文到文件
    output_filename = f"generated_document_{TIMESTAMP}.html"
    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(document_content)
        log_info(f"公文已保存到文件: {output_filename}")
        
        # 同时保存一个纯文本版本
        text_filename = f"generated_document_{TIMESTAMP}.txt"
        # 简单去除HTML标签
        import re
        text_content = re.sub(r'<[^>]+>', '', document_content)
        text_content = re.sub(r'\s+', ' ', text_content).strip()
        
        with open(text_filename, 'w', encoding='utf-8') as f:
            f.write(text_content)
        log_info(f"公文纯文本版本已保存到文件: {text_filename}")
        
        return document_content, output_filename
    except Exception as e:
        log_error("保存公文文件失败", e)
        return document_content, None

# ==========================================
# 邮件发送函数
# ==========================================

def send_email(document_content, attachment_path=None):
    """发送包含公文的邮件"""
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        log_error("邮件配置不完整，跳过发送")
        return False
    
    # 解析收件人列表
    receivers_list = []
    if EMAIL_RECEIVERS:
        receivers_list = [r.strip() for r in EMAIL_RECEIVERS.replace('，', ',').split(',') if r.strip()]
    
    if not receivers_list:
        receivers_list = [EMAIL_SENDER]  # 默认发给自己
    
    log_info(f"收件人列表: {', '.join(receivers_list)}")
    
    # 清理内容中的代码标记
    content = document_content.replace("```html", "").replace("```", "").strip()
    
    # 创建完整的HTML邮件内容
    full_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{DOCUMENT_TITLE}</title>
    <style>
        body {{
            font-family: '仿宋', 'SimSun', 'Microsoft YaHei', sans-serif;
            line-height: 1.8;
            color: #000;
            background-color: #f5f5f5;
            padding: 20px;
            max-width: 800px;
            margin: 0 auto;
        }}
        .document-container {{
            background-color: white;
            padding: 40px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            position: relative;
            margin-bottom: 30px;
        }}
        .document-header {{
            text-align: center;
            margin-bottom: 40px;
            border-bottom: 2px solid #000;
            padding-bottom: 20px;
        }}
        .document-title {{
            font-size: 22px;
            font-weight: bold;
            font-family: '小标宋', 'SimSun', serif;
            margin-bottom: 10px;
        }}
        .document-subtitle {{
            font-size: 16px;
            color: #666;
            margin-bottom: 5px;
        }}
        .document-receiver {{
            font-size: 16px;
            text-align: left;
            margin: 30px 0;
        }}
        .document-body {{
            font-size: 16px;
            margin-bottom: 20px;
        }}
        .document-signature {{
            text-align: right;
            margin-top: 100px;
        }}
        .sender {{
            font-size: 16px;
            margin-bottom: 10px;
        }}
        .date {{
            font-size: 16px;
        }}
        .urgency-badge {{
            position: absolute;
            top: 20px;
            right: 20px;
            padding: 5px 10px;
            background-color: #f0f0f0;
            border: 1px solid #ccc;
            font-size: 14px;
        }}
        .security-badge {{
            position: absolute;
            top: 20px;
            left: 20px;
            padding: 5px 10px;
            background-color: #f0f0f0;
            border: 1px solid #ccc;
            font-size: 14px;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            font-size: 12px;
            color: #999;
            border-top: 1px solid #eee;
            padding-top: 20px;
        }}
        .info-box {{
            background-color: #f8f9fa;
            border-left: 4px solid #007bff;
            padding: 15px;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="info-box">
        <strong>公文信息</strong><br/>
        类型：{DOCUMENT_TYPE} | 发文单位：{DOCUMENT_SENDER} | 生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    </div>
    
    <div class="document-container">
        {'' if DOCUMENT_SECURITY == '无' else f'<div class="security-badge">密级：{DOCUMENT_SECURITY}</div>'}
        {'' if DOCUMENT_URGENCY == '普通' else f'<div class="urgency-badge">紧急程度：{DOCUMENT_URGENCY}</div>'}
        
        <div class="document-header">
            <div class="document-title">{DOCUMENT_TITLE}</div>
            <div class="document-subtitle">公文类型：{DOCUMENT_TYPE}</div>
            <div class="document-subtitle">生成日期：{TODAY_STR}</div>
        </div>
        
        <div class="document-receiver">
            <strong>主送：</strong>{DOCUMENT_RECEIVER}
        </div>
        
        <div class="document-body">
            {content}
        </div>
        
        <div class="document-signature">
            <div class="sender">
                <strong>发文机关：</strong>{DOCUMENT_SENDER}
            </div>
            <div class="date">
                <strong>成文日期：</strong>{TODAY_STR}
            </div>
        </div>
    </div>
    
    <div class="footer">
        <p>本公文由AI自动生成，仅供参考。正式公文请按规范程序签发。</p>
        <p>生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p>如需修改或重新生成，请在GitHub Actions中重新运行工作流。</p>
    </div>
</body>
</html>
"""
    
    # 创建邮件
    msg = MIMEMultipart()
    msg['From'] = formataddr(("公文生成系统", EMAIL_SENDER))
    msg['To'] = ", ".join(receivers_list)
    msg['Subject'] = Header(f"【{DOCUMENT_TYPE}】{DOCUMENT_TITLE} - {TODAY_STR}", 'utf-8')
    
    # 添加HTML内容
    msg.attach(MIMEText(full_html, 'html', 'utf-8'))
    
    # 添加附件（如果有）
    if attachment_path and os.path.exists(attachment_path):
        try:
            with open(attachment_path, 'rb') as f:
                attachment = MIMEText(f.read().decode('utf-8'), 'base64', 'utf-8')
                attachment.add_header('Content-Disposition', 'attachment', 
                                     filename=os.path.basename(attachment_path))
                msg.attach(attachment)
            log_info(f"已添加附件: {attachment_path}")
        except Exception as e:
            log_error("添加附件失败", e)
    
    # 发送邮件
    max_retries = 2
    for attempt in range(max_retries):
        try:
            log_info(f"尝试发送邮件 (尝试 {attempt + 1}/{max_retries})...")
            
            if SMTP_PORT == 465:
                # 使用SSL
                server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
            else:
                # 使用STARTTLS
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
                server.starttls()
            
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, receivers_list, msg.as_string())
            server.quit()
            
            log_info("✅ 邮件发送成功")
            return True
            
        except Exception as e:
            log_error(f"邮件发送失败 (尝试 {attempt + 1})", e)
            if attempt < max_retries - 1:
                wait_time = 5
                log_info(f"等待{wait_time}秒后重试...")
                time.sleep(wait_time)
    
    log_error("邮件发送最终失败")
    return False

# ==========================================
# 主函数
# ==========================================

def main():
    """主函数"""
    log_info("=" * 60)
    log_info("公文生成系统启动")
    log_info(f"当前时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_info("=" * 60)
    
    # 检查必需的环境变量
    if not check_required_env():
        log_error("环境变量检查失败，程序退出")
        sys.exit(1)
    
    # 生成公文
    document_content, output_filename = generate_document()
    
    # 发送邮件
    send_success = send_email(document_content, output_filename)
    
    # 输出结果
    log_info("=" * 60)
    if send_success:
        log_info("🎉 公文生成与发送流程完成！")
    else:
        log_info("⚠️ 公文生成完成，但邮件发送失败")
    
    # 输出公文内容预览
    log_info("公文内容预览（前500字符）:")
    log_info("-" * 60)
    preview = document_content[:500] + "..." if len(document_content) > 500 else document_content
    print(preview)
    log_info("-" * 60)
    
    # 返回状态码
    if send_success:
        sys.exit(0)
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()
