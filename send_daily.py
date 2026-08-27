from datetime import datetime
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import time
import openpyxl


# =========================================================================
# 1. 日报 HTML 生成（提取 '1.日工作简报'）
# =========================================================================
def generate_daily_html(wb):
    sheet = wb["1.日工作简报"] if "1.日工作简报" in wb.sheetnames else wb.active

    title = (
        str(sheet["A1"].value).strip()
        if sheet["A1"].value
        else f'工作&学习日报-{datetime.now().strftime("%m月%d日")}'
    )
    today_work = str(sheet["A3"].value or "").strip().replace("\n", "<br>")
    tomorrow_plan = str(sheet["A5"].value or "").strip().replace("\n", "<br>")

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        .daily-table {{
            border-collapse: collapse;
            width: 100%;
            max-width: 800px;
            border: 1px solid #000000;
            table-layout: fixed;
            box-sizing: border-box;
            font-family: SimSun, '宋体', serif;
        }}
        .daily-table td {{ border: 1px solid #000000; box-sizing: border-box; }}
        .title-row {{
            background-color: #8DB4E2;
            font-family: SimSun, "宋体", serif;
            font-size: 18pt;
            font-weight: bold;
            color: #000000;
            text-align: center;
            vertical-align: middle;
            height: 46px;
        }}
        .section-header {{
            background-color: #DBE5F1;
            text-align: left;
            vertical-align: middle;
            height: 30px;
            padding: 0 8px;
        }}
        .section-title {{ font-family: SimSun, "宋体", serif; font-size: 12pt; font-weight: bold; color: #000000; }}
        .red-tip {{ font-family: SimSun, "宋体", serif; font-size: 10pt; font-weight: bold; color: #FF0000; }}
        .content-box {{
            background-color: #FFFFFF;
            font-family: "STXihei", "华文细黑", "Microsoft YaHei", sans-serif;
            font-size: 10pt;
            font-weight: bold;
            color: #000000;
            text-align: left;
            vertical-align: top;
            padding: 12px 10px;
            white-space: pre-wrap;
            line-height: 1.6;
        }}
        .bottom-bar {{ background-color: #DBE5F1; height: 18px; }}
        .signature {{ margin-top: 25px; font-size: 12px; color: #333333; border-top: 1px solid #CCCCCC; padding-top: 6px; width: 220px; font-family: Arial, sans-serif; }}
    </style>
    </head>
    <body style="margin: 0; padding: 10px; background-color: #FFFFFF;">
        <table class="daily-table">
            <tr><td class="title-row">{title}</td></tr>
            <tr><td class="section-header"><span class="section-title">今日工作概述</span><span class="red-tip">（重点工作、重大风险说明 (加粗标红）、周边求助 (加粗标红），小于3条）</span></td></tr>
            <tr><td class="content-box" style="height: 120px;">{today_work}</td></tr>
            <tr><td class="section-header"><span class="section-title">明日工作计划</span><span class="red-tip">（重点工作、重大风险说明 (加粗标红）、周边求助 (加粗标红），小于3条）</span></td></tr>
            <tr><td class="content-box" style="height: 90px;">{tomorrow_plan}</td></tr>
            <tr><td class="bottom-bar"></td></tr>
        </table>
        <div class="signature">{os.environ.get("MAIL_USER")}</div>
    </body>
    </html>
    """
    return title, html_content


# =========================================================================
# 2. 周报 HTML 生成（按图一标准 1:1 像素级深蓝粗体与浓黑宋体重构）
# =========================================================================
def generate_weekly_html(wb):
    sheet = wb["周工作简报"] if "周工作简报" in wb.sheetnames else wb.active

    # 1. 提取基本信息
    r1_title = str(sheet["A1"].value or "能力提升-学习周报").strip()
    curr_prod = str(sheet["C2"].value or "云计算&中级").strip()
    target_prod = str(sheet["G2"].value or "云计算&高级").strip()
    report_period = str(
        sheet["J2"].value or "2026年8月24日-2026年8月28日"
    ).strip()
    plan_period = str(sheet["C3"].value or "20周").strip()
    user_name = str(sheet["F4"].value or "苏文强").strip()
    user_phone = str(sheet["H4"].value or "16611601206").strip()

    # 2. 提取正文内容
    this_week_work = (
        str(sheet["A6"].value or "").strip().replace("\n", "<br>")
    )
    next_week_plan = (
        str(sheet["A8"].value or "学习华为私有云部署")
        .strip()
        .replace("\n", "<br>")
    )

    subject = "苏文强学习周报"

    # 红斜杠占位符
    red_slash = '<span style="color: #FF0000; font-weight: bold; font-family: SimSun, \'宋体\';">/</span>'

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        .weekly-container {{
            border-collapse: collapse;
            width: 100%;
            max-width: 900px;
            border: 1px solid #000000;
            font-family: SimSun, '宋体', serif;
            table-layout: fixed;
            box-sizing: border-box;
        }}
        .weekly-container td {{
            border: 1px solid #000000;
            padding: 4px 6px;
            box-sizing: border-box;
            text-align: center;
            vertical-align: middle;
        }}
        /* 1. 顶头大标题：淡紫蓝底色 #CCCCFF，浓郁纯深蓝字 #0000FF，18pt 粗体 */
        .title-banner {{
            background-color: #CCCCFF;
            color: #0000FF !important;
            font-family: SimSun, '宋体', serif;
            font-size: 18pt;
            font-weight: 900;
            text-align: center;
            height: 44px;
            letter-spacing: 1px;
        }}
        /* 2. 基本信息表 (纯白底，宋体加粗) */
        .info-cell {{
            background-color: #FFFFFF;
            font-family: SimSun, '宋体', serif;
            font-size: 10.5pt;
            color: #000000;
        }}
        .bold-txt {{
            font-weight: bold;
        }}
        /* 3. 分栏标题条：纯深蓝 #0000FF，14pt 加粗，带下划线质感 */
        .section-bar {{
            background-color: #CCCCFF;
            color: #0000FF !important;
            font-family: SimSun, '宋体', serif;
            font-size: 14pt;
            font-weight: 900;
            text-align: center;
            height: 36px;
            text-decoration: underline;
        }}
        .section-bar-plain {{
            background-color: #CCCCFF;
            color: #0000FF !important;
            font-family: SimSun, '宋体', serif;
            font-size: 14pt;
            font-weight: 900;
            text-align: center;
            height: 36px;
        }}
        /* 4. 黄底正文内容：采用宋体 SimSun 10.5pt，解决发细发虚问题，笔画浓黑扎实 */
        .content-yellow {{
            background-color: #FFFF99;
            color: #000000;
            font-family: SimSun, '宋体', serif;
            font-size: 10.5pt;
            line-height: 1.6;
            text-align: left !important;
            vertical-align: top;
            padding: 10px 12px;
            white-space: pre-wrap;
        }}
        /* 表头黄底加粗格 */
        .header-yellow {{
            background-color: #FFFF99;
            color: #000000;
            font-family: SimSun, '宋体', serif;
            font-size: 10pt;
            font-weight: bold;
            text-align: center;
        }}
        .val-yellow {{
            background-color: #FFFF99;
            color: #000000;
            font-family: SimSun, '宋体', serif;
            font-size: 10pt;
            text-align: center;
        }}
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
        <table class="weekly-container">
            <!-- 1. 【大标题栏】 -->
            <tr>
                <td colspan="8" class="title-banner">{r1_title}</td>
            </tr>

            <!-- 2. 【基本信息表】第 1 行 -->
            <tr>
                <td class="info-cell bold-txt">当前产品&级别</td>
                <td class="info-cell">{curr_prod}</td>
                <td class="info-cell bold-txt">目标产品&级别</td>
                <td class="info-cell">{target_prod}</td>
                <td class="info-cell bold-txt">报告周期</td>
                <td colspan="3" class="info-cell" style="font-size: 9.5pt;">{report_period}</td>
            </tr>

            <!-- 2. 【基本信息表】第 2 行 -->
            <tr>
                <td class="info-cell bold-txt">计划周期</td>
                <td colspan="3" class="info-cell">{plan_period}</td>
                <td class="info-cell bold-txt">姓名</td>
                <td class="info-cell">{user_name}</td>
                <td class="info-cell bold-txt">电话</td>
                <td class="info-cell">{user_phone}</td>
            </tr>

            <!-- 2. 【基本信息表】第 3 行 -->
            <tr>
                <td colspan="4" class="info-cell bold-txt">交付工程师信息</td>
                <td colspan="2" class="info-cell">{user_name}</td>
                <td colspan="2" class="info-cell">{user_phone}</td>
            </tr>

            <!-- 3. 【分栏标题条 - 本周工作内容】 (深蓝加粗下划线) -->
            <tr>
                <td colspan="8" class="section-bar"><u>本周工作内容</u></td>
            </tr>

            <!-- 4. 【本周工作内容正文】 (宋体 10.5pt 浓黑靠左) -->
            <tr>
                <td colspan="8" class="content-yellow">
                    {this_week_work}
                </td>
            </tr>

            <!-- 5. 【分栏标题条 - 下周工作计划】 -->
            <tr>
                <td colspan="8" class="section-bar"><u>下周工作计划</u></td>
            </tr>

            <!-- 5. 【下周工作计划正文】 -->
            <tr>
                <td colspan="8" class="content-yellow bold-txt" style="min-height: 40px;">
                    {next_week_plan}
                </td>
            </tr>

            <!-- 6. 【项目风险问题】 -->
            <tr>
                <td colspan="8" class="section-bar-plain">项目风险问题</td>
            </tr>
            <!-- 风险表头 -->
            <tr>
                <td class="header-yellow" style="width: 7%;">序号</td>
                <td class="header-yellow" style="width: 12%;"><span style="color: #FF0000; font-weight: bold;">产生日期</span></td>
                <td class="header-yellow" style="width: 25%;">问题描述&影响</td>
                <td class="header-yellow" style="width: 10%;">紧急程度</td>
                <td class="header-yellow" style="width: 20%;">规避措施、解决进展</td>
                <td class="header-yellow" style="width: 8%;">责任人</td>
                <td class="header-yellow" style="width: 8%;">状态</td>
                <td class="header-yellow" style="width: 10%;">实际关闭日期</td>
            </tr>
            <!-- 风险数据行 -->
            <tr>
                <td class="val-yellow">1</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_yellow if 'red_yellow' in locals() else red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
            </tr>

            <!-- 7. 【求助】 -->
            <tr>
                <td colspan="8" class="section-bar-plain">求助</td>
            </tr>
            <!-- 求助表头 -->
            <tr>
                <td class="header-yellow" style="width: 7%;">序号</td>
                <td colspan="3" class="header-yellow">求助事宜</td>
                <td class="header-yellow" style="width: 12%;">求助人</td>
                <td class="header-yellow" style="width: 16%;">要求解决时间</td>
                <td colspan="2" class="header-yellow">备注</td>
            </tr>
            <!-- 求助数据行 -->
            <tr>
                <td class="val-yellow">1</td>
                <td colspan="3" class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td class="val-yellow">{red_slash}</td>
                <td colspan="2" class="val-yellow">{red_slash}</td>
            </tr>
        </table>

        <div class="signature">
            {os.environ.get("MAIL_USER")}
        </div>
    </body>
    </html>
    """
    return subject, html_content


# =========================================================================
# 3. 发送邮件底层函数
# =========================================================================
def send_email_message(subject, html_body, to_list, cc_list):
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

    server = smtplib.SMTP_SSL(mail_host, 465, timeout=30)
    server.login(mail_user, mail_pass)
    server.sendmail(mail_user, to_list + cc_list, msg.as_string())
    server.quit()
    print(f"✅ 邮件投递成功 -> 主题:「{subject}」 | 收件人: {to_list}")


# =========================================================================
# 4. 主流程调度
# =========================================================================
def main():
    excel_file = "daily_report.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ 未找到 {excel_file} 文件，跳过本次发送。")
        return

    wb = openpyxl.load_workbook(excel_file, data_only=True)

    # =========================================================================
    # 【收发模式配置】
    # =========================================================================

    # 👉 模式一：【测试模式】（测试时所有邮件只发给自己）
    daily_to = ["wqsud@isoftstone.com"]
    daily_cc = []
    weekly_to = ["wqsud@isoftstone.com"]
    weekly_cc = []

    # 👉 模式二：【正式发送模式】（正式上线时启用以下配置）
    # daily_to = ["hdliuf@isoftstone.com"]
    # daily_cc = ["weiliuay@isoftstone.com", "zycaoc@isoftstone.com", "nawangam@isoftstone.com"]
    # weekly_to = ["zycaoc@isoftstone.com"]
    # weekly_cc = []

    # =========================================================================

    # 1. 发送【日报】
    print("\n==============================================")
    print(" [1/2] 正在投递：今日工作日报 ...")
    daily_subj, daily_html = generate_daily_html(wb)
    send_email_message(daily_subj, daily_html, daily_to, daily_cc)
    print("==============================================")

    # 2. 发送【周报】
    # 💡 提示：当前设置为 True 供测试；正式使用时请改为 datetime.now().weekday() == 4
    is_friday = True  # <-- 测试双发开启

    if is_friday:
        print("\n⏳ 缓冲 4 秒，确保两封邮件独立送达...")
        time.sleep(4)

        print("==============================================")
        print(" [2/2] 正在投递：学习周报 ...")
        weekly_subj, weekly_html = generate_weekly_html(wb)
        send_email_message(weekly_subj, weekly_html, weekly_to, weekly_cc)
        print("==============================================")
    else:
        print("今日非周五，跳过周报。")


if __name__ == "__main__":
    main()
