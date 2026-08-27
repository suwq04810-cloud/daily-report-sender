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
        .excel-table {{
            border-collapse: collapse;
            width: 100%;
            max-width: 800px;
            border: 1px solid #000000;
            table-layout: fixed;
            box-sizing: border-box;
        }}
        .excel-table td {{ border: 1px solid #000000; box-sizing: border-box; }}
        .title-row {{
            background-color: #8DB4E2;
            font-family: "SimSun", "宋体", serif;
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
        .section-title {{ font-family: "SimSun", "宋体", serif; font-size: 12pt; font-weight: bold; color: #000000; }}
        .red-tip {{ font-family: "SimSun", "宋体", serif; font-size: 10pt; font-weight: bold; color: #FF0000; }}
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
        .bottom-bar {{ background-color: #DBE5F1; height: 18px; }}
        .signature {{ margin-top: 25px; font-size: 12px; color: #333333; border-top: 1px solid #CCCCCC; padding-top: 6px; width: 220px; font-family: Arial, sans-serif; }}
    </style>
    </head>
    <body style="margin: 0; padding: 10px; background-color: #FFFFFF;">
        <table class="excel-table">
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
# 2. 周报 HTML 生成（完整 1:1 提取 '周工作简报' 全表 14 行）
# =========================================================================
def generate_weekly_html(wb):
    sheet = wb["周工作简报"] if "周工作简报" in wb.sheetnames else wb.active

    # 1. 顶部基本信息提取
    r1_title = str(sheet["A1"].value or "能力提升-学习周报").strip()
    curr_prod = str(sheet["C2"].value or "云计算&中级").strip()
    target_prod = str(sheet["G2"].value or "云计算&高级").strip()
    report_period = str(
        sheet["J2"].value or datetime.now().strftime("%Y年%m月%d日")
    ).strip()
    plan_period = str(sheet["C3"].value or "20周").strip()
    user_name = str(sheet["F4"].value or "苏文强").strip()
    user_phone = str(sheet["H4"].value or "16611601206").strip()

    # 2. 核心内容与计划提取
    this_week_work = (
        str(sheet["A6"].value or "").strip().replace("\n", "<br>")
    )
    next_week_plan = (
        str(sheet["A8"].value or "学习华为私有云部署")
        .strip()
        .replace("\n", "<br>")
    )

    # 3. 风险与求助数据提取
    risk_desc = str(sheet["C11"].value or "/").strip()
    help_desc = str(sheet["B14"].value or "/").strip()

    subject = "苏文强学习周报"

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        .weekly-table {{
            border-collapse: collapse;
            width: 100%;
            max-width: 900px;
            border: 1px solid #000000;
            font-family: 'Microsoft YaHei', '微软雅黑', 'SimSun', '宋体', sans-serif;
            font-size: 10pt;
            table-layout: fixed;
            box-sizing: border-box;
        }}
        .weekly-table td {{
            border: 1px solid #000000;
            padding: 5px 6px;
            box-sizing: border-box;
            text-align: center;
            vertical-align: middle;
        }}
        /* 模块顶头大栏 (蓝紫底色 #CCCCFF / 深蓝字 #0000FF) */
        .header-banner {{
            background-color: #CCCCFF;
            font-size: 14pt;
            font-weight: bold;
            color: #0000FF;
            height: 36px;
        }}
        /* 浅冰蓝标签格 #DBE5F1 */
        .lbl-cell {{
            background-color: #DBE5F1;
            font-weight: bold;
            color: #000000;
        }}
        /* 白底数据格 #FFFFFF */
        .val-cell {{
            background-color: #FFFFFF;
            color: #000000;
        }}
        /* 黄底数据格 #FFFF99 */
        .yellow-val {{
            background-color: #FFFF99;
            color: #000000;
        }}
        /* 黄底大正文区 */
        .yellow-content {{
            background-color: #FFFF99;
            text-align: left;
            vertical-align: top;
            padding: 12px 10px;
            white-space: pre-wrap;
            line-height: 1.6;
            color: #000000;
            font-size: 10pt;
        }}
        .red-text {{
            color: #FF0000;
            font-weight: bold;
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
        <table class="weekly-table">
            <!-- Row 1: 顶部大标题 -->
            <tr>
                <td colspan="10" class="header-banner" style="font-size: 16pt;">{r1_title}</td>
            </tr>

            <!-- Row 2: 产品与周期 -->
            <tr>
                <td colspan="2" class="lbl-cell">当前产品<br>&级别</td>
                <td colspan="2" class="val-cell">{curr_prod}</td>
                <td colspan="2" class="lbl-cell">目标产品<br>&级别</td>
                <td colspan="2" class="val-cell">{target_prod}</td>
                <td class="lbl-cell">报告周期</td>
                <td class="val-cell" style="font-size: 9pt;">{report_period}</td>
            </tr>

            <!-- Row 3: 计划周期 -->
            <tr>
                <td colspan="2" class="lbl-cell">计划周期</td>
                <td colspan="2" class="val-cell">{plan_period}</td>
                <td colspan="2" class="lbl-cell">姓名</td>
                <td colspan="2" class="val-cell">{user_name}</td>
                <td class="lbl-cell">电话</td>
                <td class="val-cell">{user_phone}</td>
            </tr>

            <!-- Row 4: 交付工程师信息 -->
            <tr>
                <td colspan="4" class="lbl-cell" style="font-size: 11pt;">交付工程师信息</td>
                <td colspan="3" class="yellow-val" style="font-weight: bold;">{user_name}</td>
                <td colspan="3" class="yellow-val" style="font-weight: bold;">{user_phone}</td>
            </tr>

            <!-- Row 5: 本周工作内容 表头 -->
            <tr>
                <td colspan="10" class="header-banner">本周工作内容</td>
            </tr>

            <!-- Row 6: 本周工作内容 主体 -->
            <tr>
                <td colspan="10" class="yellow-content">
                    {this_week_work}
                </td>
            </tr>

            <!-- Row 7: 下周工作计划 表头 -->
            <tr>
                <td colspan="10" class="header-banner">下周工作计划</td>
            </tr>

            <!-- Row 8: 下周工作计划 主体 -->
            <tr>
                <td colspan="10" class="yellow-content" style="min-height: 80px;">
                    {next_week_plan}
                </td>
            </tr>

            <!-- Row 9: 项目风险问题 表头 -->
            <tr>
                <td colspan="10" class="header-banner">项目风险问题</td>
            </tr>

            <!-- Row 10: 风险表头 -->
            <tr>
                <td class="lbl-cell" style="width: 6%;">序号</td>
                <td class="lbl-cell red-text" style="width: 10%;">产生日期</td>
                <td colspan="2" class="lbl-cell">问题描述&影响</td>
                <td class="lbl-cell" style="width: 10%;">紧急程度</td>
                <td colspan="2" class="lbl-cell">规避措施、解决进展</td>
                <td class="lbl-cell" style="width: 10%;">责任人</td>
                <td class="lbl-cell" style="width: 8%;">状态</td>
                <td class="lbl-cell" style="width: 12%;">实际关闭日期</td>
            </tr>

            <!-- Row 11: 风险数据行 -->
            <tr>
                <td class="val-cell">1</td>
                <td class="val-cell red-text">/</td>
                <td colspan="2" class="val-cell">{risk_desc}</td>
                <td class="val-cell">/</td>
                <td colspan="2" class="val-cell">/</td>
                <td class="val-cell">/</td>
                <td class="val-cell">/</td>
                <td class="val-cell">/</td>
            </tr>

            <!-- Row 12: 求助 表头 -->
            <tr>
                <td colspan="10" class="header-banner">求助</td>
            </tr>

            <!-- Row 13: 求助表头 -->
            <tr>
                <td class="lbl-cell" style="width: 6%;">序号</td>
                <td colspan="3" class="lbl-cell">求助事宜</td>
                <td colspan="2" class="lbl-cell">求助人</td>
                <td colspan="2" class="lbl-cell">要求解决时间</td>
                <td colspan="2" class="lbl-cell">备注</td>
            </tr>

            <!-- Row 14: 求助数据行 -->
            <tr>
                <td class="val-cell">1</td>
                <td colspan="3" class="val-cell">{help_desc}</td>
                <td colspan="2" class="val-cell">/</td>
                <td colspan="2" class="val-cell">/</td>
                <td colspan="2" class="val-cell">/</td>
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

    server = smtplib.SMTP_SSL(mail_host, 465)
    server.login(mail_user, mail_pass)
    server.sendmail(mail_user, to_list + cc_list, msg.as_string())
    server.quit()
    print(f"✅ 邮件「{subject}」发送成功！收件人: {to_list}")


# =========================================================================
# 4. 主调度流程
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

    # 👉 模式二：【正式发送模式】（上线时启用以下配置）
    # daily_to = ["hdliuf@isoftstone.com"]
    # daily_cc = ["weiliuay@isoftstone.com", "zycaoc@isoftstone.com", "nawangam@isoftstone.com"]
    # weekly_to = ["zycaoc@isoftstone.com"]
    # weekly_cc = []

    # =========================================================================

    # 1. 发送日报（周一至周五每天必发）
    print("\n--- [1/2] 正在处理日报 ---")
    daily_subj, daily_html = generate_daily_html(wb)
    send_email_message(daily_subj, daily_html, daily_to, daily_cc)

    # 2. 判断是否为周五发送周报
    # 💡 提示：如果当前需要强制测试周报，保持下面 is_friday = True 即可；测试完改回 weekday() == 4
    is_friday = datetime.now().weekday() == 4
    is_friday = True  # <-- 当前开启强制测试周报，你点 Run workflow 会连发【日报+周报】！

    if is_friday:
        print("\n--- [2/2] 正在处理周报 ---")
        time.sleep(2)  # 间隔 2 秒确保邮件服务器稳定接收两封
        weekly_subj, weekly_html = generate_weekly_html(wb)
        send_email_message(weekly_subj, weekly_html, weekly_to, weekly_cc)
    else:
        print("今天不是周五，跳过周报发送。")


if __name__ == "__main__":
    main()
