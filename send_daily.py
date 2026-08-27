from datetime import datetime
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
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
# 2. 周报 HTML 生成（提取 '周工作简报'）
# =========================================================================
def generate_weekly_html(wb):
    sheet = wb["周工作简报"] if "周工作简报" in wb.sheetnames else wb.active

    # 提取工程师信息：姓名(F4)、电话(H4)
    name = str(sheet["F4"].value or "苏文强").strip()
    phone = str(sheet["H4"].value or "16611601206").strip()

    # 提取工作内容与计划 (A6 为本周工作，A8 为下周计划)
    this_week_work = (
        str(sheet["A6"].value or "").strip().replace("\n", "<br>")
    )
    next_week_plan = (
        str(sheet["A8"].value or "按项目计划推进").strip().replace("\n", "<br>")
    )

    # 主题命名
    today_str = datetime.now().strftime("%m月%d日")
    subject = f"【周工作简报】{name}-{today_str}"

    # 严格按照周报排版规范构建 HTML
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
            table-layout: fixed;
            box-sizing: border-box;
        }}
        .weekly-table td {{
            border: 1px solid #000000;
            padding: 6px 8px;
            box-sizing: border-box;
        }}
        /* 交付人员信息栏 */
        .info-header {{
            background-color: #DBE5F1;
            font-family: 'SimSun', '宋体', serif;
            font-size: 12pt;
            font-weight: bold;
            color: #000000;
            text-align: center;
            width: 25%;
        }}
        .info-val {{
            background-color: #FFFF99;
            font-family: 'Microsoft YaHei', '微软雅黑', sans-serif;
            font-size: 10pt;
            font-weight: bold;
            color: #000000;
            text-align: center;
        }}
        /* 模块大表头 (蓝紫色高亮) */
        .block-header {{
            background-color: #CCCCFF;
            font-family: 'Microsoft YaHei', '微软雅黑', sans-serif;
            font-size: 14pt;
            font-weight: bold;
            color: #0000FF;
            text-align: center;
            height: 36px;
        }}
        /* 周报内容汇总区 (淡黄底色、多行展示、顶端对齐) */
        .work-content {{
            background-color: #FFFF99;
            font-family: 'Microsoft YaHei', '微软雅黑', sans-serif;
            font-size: 10pt;
            color: #000000;
            text-align: left;
            vertical-align: top;
            padding: 12px 10px;
            white-space: pre-wrap;
            line-height: 1.6;
            min-height: 200px;
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
            <!-- 交付工程师信息栏 -->
            <tr>
                <td class="info-header" colspan="2">交付工程师信息</td>
                <td class="info-val" colspan="2">{name}</td>
                <td class="info-val" colspan="2">{phone}</td>
            </tr>

            <!-- 本周工作内容表头 -->
            <tr>
                <td class="block-header" colspan="6">本周工作内容</td>
            </tr>

            <!-- 本周工作内容主体 -->
            <tr>
                <td class="work-content" colspan="6">
                    {this_week_work}
                </td>
            </tr>

            <!-- 下周工作计划表头 -->
            <tr>
                <td class="block-header" colspan="6">下周工作计划</td>
            </tr>

            <!-- 下周工作计划主体 -->
            <tr>
                <td class="work-content" colspan="6" style="min-height: 100px;">
                    {next_week_plan}
                </td>
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
# 3. 发送邮件通用函数
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
    print(f"✅ 邮件「{subject}」发送成功！")


# =========================================================================
# 4. 主流程调度：周一至周四发日报，周五自动连发【日报+周报】
# =========================================================================
def main():
    excel_file = "daily_report.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ 未找到 {excel_file} 文件，跳过本次发送。")
        return

    wb = openpyxl.load_workbook(excel_file, data_only=True)

    # ------------------ 收件人模式配置 ------------------
    # 👉 模式一：【测试模式】（发给自己测试）
    to_list = ["wqsud@isoftstone.com"]
    cc_list = []

    # 👉 模式二：【正式发送模式】（上线时启用）
    # to_list = ["hdliuf@isoftstone.com"]
    # cc_list = ["weiliuay@isoftstone.com", "zycaoc@isoftstone.com", "nawangam@isoftstone.com"]
    # ----------------------------------------------------

    # 1. 发送日报（周一到周五每天必发）
    daily_subj, daily_html = generate_daily_html(wb)
    print(f"正在发送日报: {daily_subj} ...")
    send_email_message(daily_subj, daily_html, to_list, cc_list)

    # 2. 判断今天是否为周五 (weekday == 4 为周五)
    # 💡 提示：如果现在想测试周报，可以把 False 改成 True，测试完改回 weekday() == 4
    is_friday = datetime.now().weekday() == 4
    is_friday = True  # <-- 测试周报时，解开这行注释即可立即测试周报！

    if is_friday:
        print("检测到今天是周五，准备自动发送周报...")
        weekly_subj, weekly_html = generate_weekly_html(wb)
        send_email_message(weekly_subj, weekly_html, to_list, cc_list)
    else:
        print("今天不是周五，无需发送周报。")


if __name__ == "__main__":
    main()
