import os
import sys
import datetime
import time
import json
import re
from openai import OpenAI
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import markdown

# ==========================================
# 1. 变量解析与环境加载
# ==========================================
# 公文类型配置
DOCUMENT_TYPE = os.getenv("DOCUMENT_TYPE", "通知")  # 通知、报告、请示、函、纪要等
DOCUMENT_TITLE = os.getenv("DOCUMENT_TITLE", "关于XX工作的通知")
DOCUMENT_SUBJECT = os.getenv("DOCUMENT_SUBJECT", "工作安排")  # 事由
DOCUMENT_SENDER = os.getenv("DOCUMENT_SENDER", "XX部门")  # 发文单位
DOCUMENT_RECEIVER = os.getenv("DOCUMENT_RECEIVER", "相关部门")  # 主送机关
DOCUMENT_URGENCY = os.getenv("DOCUMENT_URGENCY", "普通")  # 紧急程度：特急、急件、普通
DOCUMENT_SECURITY = os.getenv("DOCUMENT_SECURITY", "无")  # 密级：绝密、机密、秘密、无

# 输入数据来源：可以是文件路径或直接的环境变量
INPUT_DATA_FILE = os.getenv("INPUT_DATA_FILE", "input_data.txt")
INPUT_DATA_TEXT = os.getenv("INPUT_DATA_TEXT", "")

# AI配置
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash") 
GEMINI_REQUEST_DELAY = float(os.getenv("GEMINI_REQUEST_DELAY", "3.0"))

# 邮件配置
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVERS = os.getenv("EMAIL_RECEIVERS")
SMTP_SERVER = "smtp.qq.com"

# 日期信息
TODAY_STR = datetime.date.today().strftime("%Y年%m月%d日")
CURRENT_YEAR = datetime.date.today().year

# ==========================================
# 2. 数据读取函数
# ==========================================
def read_input_data():
    """
    读取输入数据，优先从文件读取，其次从环境变量读取
    """
    data_content = ""
    
    # 首先尝试从环境变量读取
    if INPUT_DATA_TEXT and INPUT_DATA_TEXT.strip():
        data_content = INPUT_DATA_TEXT.strip()
        print(f"从环境变量读取输入数据，长度: {len(data_content)} 字符")
    
    # 如果环境变量为空，尝试从文件读取
    elif os.path.exists(INPUT_DATA_FILE):
        try:
            with open(INPUT_DATA_FILE, 'r', encoding='utf-8') as f:
                data_content = f.read().strip()
            print(f"从文件 {INPUT_DATA_FILE} 读取输入数据，长度: {len(data_content)} 字符")
        except Exception as e:
            print(f"读取文件失败: {e}")
            data_content = "无输入数据"
    else:
        data_content = "无输入数据"
        print("未找到输入数据，使用默认值")
    
    return data_content

# ==========================================
# 3. 公文生成函数
# ==========================================
def generate_document(client, model_name, input_data):
    """
    根据输入数据和配置生成公文
    """
    
    # 根据公文类型选择不同的提示词模板
    prompt_templates = {
        "通知": """
        【公文类型】通知
        【发文单位】{sender}
        【主送机关】{receiver}
        【标题】{title}
        【事由】{subject}
        【紧急程度】{urgency}
        【密级】{security}
        【日期】{date}
        
        【公文格式要求】
        1. 严格按照《党政机关公文格式》国家标准（GB/T 9704-2012）撰写
        2. 结构必须包含：标题、主送机关、正文、发文机关署名、成文日期
        3. 正文内容应分层次叙述，逻辑清晰
        4. 语言要求：庄重、准确、简练、规范
        5. 使用公文专用语：如"现将有关事项通知如下"、"特此通知"等
        
        【输入材料】
        {input_data}
        
        【生成要求】
        请基于以上输入材料，生成一份完整的{type}。要求：
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
        【公文类型】报告
        【发文单位】{sender}
        【主送机关】{receiver}
        【标题】{title}
        【事由】{subject}
        【日期】{date}
        
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
        【公文类型】请示
        【发文单位】{sender}
        【主送机关】{receiver}
        【标题】{title}
        【事由】{subject}
        【紧急程度】{urgency}
        【日期】{date}
        
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
        【公文类型】函
        【发文单位】{sender}
        【主送机关】{receiver}
        【标题】{title}
        【事由】{subject}
        【日期】{date}
        
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
        【公文类型】会议纪要
        【发文单位】{sender}
        【标题】{title}
        【事由】{subject}
        【日期】{date}
        
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
        type=DOCUMENT_TYPE,
        sender=DOCUMENT_SENDER,
        receiver=DOCUMENT_RECEIVER,
        title=DOCUMENT_TITLE,
        subject=DOCUMENT_SUBJECT,
        urgency=DOCUMENT_URGENCY,
        security=DOCUMENT_SECURITY,
        date=TODAY_STR,
        input_data=input_data
    )
    
    # 添加严格的格式指令
    strict_format_instruction = f"""
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
    
    full_prompt = prompt + "\n\n" + strict_format_instruction
    
    time.sleep(GEMINI_REQUEST_DELAY)
    
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": full_prompt}],
            temperature=0.1
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"生成公文失败: {e}"

# ==========================================
# 4. 邮件发送函数
# ==========================================
def send_email(subject, content):
    """
    发送包含公文的邮件
    """
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        print("邮件配置不完整，跳过发送")
        return
    
    receivers_list = [EMAIL_SENDER]
    if EMAIL_RECEIVERS:
        receivers_list = [r.strip() for r in EMAIL_RECEIVERS.replace('，', ',').split(',') if r.strip()]
    
    # 清理内容中的代码标记
    content = content.replace("```html", "").replace("```", "")
    
    # 创建完整的HTML邮件内容
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
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
                text-indent: 2em;
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
        </style>
    </head>
    <body>
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
    
    # 可选：添加纯文本版本
    plain_text = f"""
    {DOCUMENT_TITLE}
    
    公文类型：{DOCUMENT_TYPE}
    主送机关：{DOCUMENT_RECEIVER}
    发文机关：{DOCUMENT_SENDER}
    成文日期：{TODAY_STR}
    紧急程度：{DOCUMENT_URGENCY}
    密级：{DOCUMENT_SECURITY}
    
    正文内容：
    {content}
    
    ---
    本公文由AI自动生成，仅供参考。
    """
    msg.attach(MIMEText(plain_text, 'plain', 'utf-8'))
    
    # 发送邮件
    try:
        print("尝试使用SSL（端口465）发送邮件...")
        server = smtplib.SMTP_SSL(SMTP_SERVER, 465, timeout=30)
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, receivers_list, msg.as_string())
        server.quit()
        print("✅ 公文发送成功（465端口）")
        return True
    except Exception as e1:
        print(f"⚠️ 465端口失败（{e1}），尝试备用STARTTLS（端口587）...")
        try:
            time.sleep(3)
            server = smtplib.SMTP(SMTP_SERVER, 587, timeout=30)
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, receivers_list, msg.as_string())
            server.quit()
            print("✅ 公文发送成功（587端口）")
            return True
        except Exception as e2:
            print(f"❌ 邮件发送最终失败: {e2}")
            return False

# ==========================================
# 5. 执行主流程
# ==========================================
if __name__ == "__main__":
    print(f"-> 启动公文生成系统，当前日期: {TODAY_STR}")
    print(f"-> 公文类型: {DOCUMENT_TYPE}")
    print(f"-> 公文标题: {DOCUMENT_TITLE}")
    print(f"-> 发文单位: {DOCUMENT_SENDER}")
    print(f"-> 主送机关: {DOCUMENT_RECEIVER}")
    
    # 初始化AI客户端
    print(f"-> 正在使用Gemini接口，模型: {GEMINI_MODEL}")
    client = OpenAI(
        api_key=GEMINI_API_KEY,
        base_url=" https://generativelanguage.googleapis.com/v1beta/openai/ ",
        timeout=600.0
    )
    model = GEMINI_MODEL
    
    # 读取输入数据
    print("-> 读取输入数据...")
    input_data = read_input_data()
    
    # 生成公文
    print("-> AI正在生成公文...")
    document_content = generate_document(client, model, input_data)
    
    # 保存公文到文件（可选）
    output_filename = f"generated_document_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(document_content)
        print(f"✅ 公文已保存到文件: {output_filename}")
    except Exception as e:
        print(f"⚠️ 保存文件失败: {e}")
    
    # 发送邮件
    print("-> 发送公文邮件...")
    email_subject = f"【{DOCUMENT_TYPE}】{DOCUMENT_TITLE}"
    send_success = send_email(email_subject, document_content)
    
    if send_success:
        print("🎉 公文生成与发送流程完成！")
    else:
        print("⚠️ 公文生成完成，但邮件发送失败")
    
    # 输出公文内容预览
    print("\n" + "="*60)
    print("公文内容预览（前500字符）:")
    print("="*60)
    preview = document_content[:500] + "..." if len(document_content) > 500 else document_content
    print(preview)
    print("="*60)
