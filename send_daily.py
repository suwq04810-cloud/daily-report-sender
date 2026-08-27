from datetime import datetime
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import openpyxl


def excel_to_exact_html(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb["1.日工作简报"] if "1.日工作简报" in wb.sheetnames else wb.active

    # 提取 A1 标题
    title = (
        str(sheet["A1"].value).strip()
        if sheet["A1"].value
        else f'工作&学习日报-{datetime.now().strftime("%m月%d日")}'
    )

    # 提取 A3 (今日概述) 和 A5 (明日计划) 内容
    today_work = sheet["A3"].value or ""
    today_work_html = str(today_work).strip().replace("\n", "<br>")

    tomorrow_plan = sheet["A5"].value or ""
    tomorrow_plan_html = str(tomorrow_plan).strip().replace("\n", "<br>")

    # 严格按照规范构建 HTML 表格
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        /* 一、全局表格规范 */
        .excel-table {{
            border-collapse: collapse;
            width: 100%;
            max-width: 800px;
            border: 1px solid #000000;
            table-layout: fixed;
            box-sizing: border-box;
        }}
        .excel-table td {{
            border: 1px solid #000000;
            box-sizing: border-box;
        }}

        /* 二、各行样式与字号颜色规范 */
        /* 第 1 行 - 大标题 */
        .title-row {{
            background-color: #8DB4E2;
            font-family: "SimSun", "宋体", serif;
            font-size: 18pt;
            font-weight: bold;
            color: #000000;
            text-align: center;
            vertical-align: middle;
            height: 46px;
            line-height: 46px;
            padding: 0;
        }}

        /* 第 2 行 & 第 4 行 - 分栏栏目头 */
        .section-header {{
            background-color: #DBE5F1;
            text-align: left;
            vertical-align: middle;
            height: 30px;
            line-height: 30px;
            padding: 0 8px;
        }}
        .section-title {{
            font-family: "SimSun", "宋体", serif;
            font-size: 12pt;
            font-weight: bold;
            color: #000000;
        }}
        .red-tip {{
            font-family: "SimSun", "宋体", serif;
            font-size: 10pt;
            font-weight: bold;
            color: #FF0000;
        }}

        /* 第 3 行 & 第 5 行 - 正文内容区 */
        .content-box {{
            background-color: #FFFFFF;
            font-family: "STXihei", "华文细黑", "Microsoft YaHei", "微软雅黑", sans-serif;
            font-size: 10pt;
            font-weight: bold;
            color: #000000;
            text-align: left;
            vertical-align: top;
            padding: 12px 10px;
            white-space: pre-wrap;
            line-height: 1.6;
        }}
        .today-content {{ height: 120px; min-height: 120px; }}
        .tomorrow-content {{ height: 90px; min-height: 90px; }}

        /* 第 6 行 - 底部封底条 */
        .bottom-bar {{
            background-color: #DBE5F1;
            height: 18px;
            padding: 0;
        }}

        /* 签名区 */
        .signature {{
            margin-top: 25px;
            font-size: 12px;
            color: #333333;
            border-top: 1px solid #CCCCCC;
            padding-top: 6px;
            width: 220px;
            font-family: Arial, sans-serif;
        }}
    </style>
    </head>
    <body style="margin: 0; padding: 10px; background-color: #FFFFFF;">
        <table class="excel-table">
            <!-- 第 1 行：大标题 -->
            <tr>
                <td class="title-row">{title}</td>
            </tr>

            <!-- 第 2 行：今日工作概述栏目头 -->
            <tr>
                <td class="section-header">
                    <span class="section-title">今日工作概述</span><span class="red-tip">（重点工作、重大风险说明 (加粗标红）、周边求助 (加粗标红），小于3条）</span>
                </td>
            </tr>

            <!-- 第 3 行：今日工作概述正文 -->
            <tr>
                <td class="content-box today-content">{today_work_html}</td>
            </tr>

            <!-- 第 4 行：明日工作计划栏目头 -->
            <tr>
                <td class="section-header">
                    <span class="section-title">明日工作计划</span><span class="red-tip">（重点工作、重大风险说明 (加粗标红）、周边求助 (加粗标红），小于3条）</span>
                </td>
            </tr>

            <!-- 第 5 行：明日工作计划正文 -->
            <tr>
                <td class="content-box tomorrow-content">{tomorrow_plan_html}</td>
            </tr>

            <!-- 第 6 行：底部封底条 -->
            <tr>
                <td class="bottom-bar"></td>
            </tr>
        </table>

        <div class="signature">
            {os.environ.get("MAIL_USER")}
        </div>
    </body>
    </html>
    """
    return title, html_content


def send_mail():
    excel_file = "daily_report.xlsx"
    if not os.path.exists(excel_file):
        print(f"未找到 {excel_file} 文件，跳过本次发送。")
        return

    subject, html_body = excel_to_exact_html(excel_file)

    # =========================================================================
    # 【收件人模式切换】：需要用哪个模式，就保持哪个生效，另一个注释掉
    # =========================================================================

    # 👉 模式一：【测试模式】（当前生效：只发给你自己，抄送为空）
    # to_list = ["wqsud@isoftstone.com"]
    # cc_list = []

    # 👉 模式二：【正式发送模式】（正式上线时：解开这里的注释，把上面的模式一注释掉）
    to_list = ["hdliuf@isoftstone.com"]
    cc_list = [
         "weiliuay@isoftstone.com",
         "zycaoc@isoftstone.com",
         "nawangam@isoftstone.com"
     ]

    # =========================================================================

    mail_user = os.environ.get("MAIL_USER")
    mail_pass = os.environ.get("MAIL_PASS")
    mail_host = os.environ.get("MAIL_HOST", "smtp.isoftstone.com")

    msg = MIMEMultipart()
    msg["From"] = f"wqsud <{mail_user}>"
    msg["To"] = ",".join(to_list)
    if cc_list:
        msg["Cc"] = ",".join(cc_list)
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # 465 端口 SSL 发送
    server = smtplib.SMTP_SSL(mail_host, 465)
    server.login(mail_user, mail_pass)
    server.sendmail(mail_user, to_list + cc_list, msg.as_string())
    server.quit()
    print(f"邮件「{subject}」已成功发出！发送列表：{to_list + cc_list}")


if __name__ == "__main__":
    send_mail()
