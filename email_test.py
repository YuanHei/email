import os
import xlrd

from email import encoders
from email.mime.base import MIMEBase
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

import smtplib


def read_file(file_path):
    file_list = []
    work_book = xlrd.open_workbook(file_path)
    sheet_data = work_book.sheet_by_name('Sheet1')
    print('now is process :', sheet_data.name)
    Nrows = sheet_data.nrows

    for i in range(1, Nrows):
        file_list.append(sheet_data.row_values(i))

    return file_list


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


'''加密发送文本邮件'''


def sendEmail(from_addr, password, smtp_server, file_list):
    for i in range(len(file_list)):
        try:
            person_info = file_list[i]
            order_num, job_num, person_name, department, merit_pay, payable, total_pay, att_deduction, money_award,\
            total_due, social_security, accumulation_fund, personal_income, dormitory_fee, communications_subsidies ,\
            travel_allowance, real_pay, email_add  = int(person_info[0]), int(person_info[1]), str(person_info[2]), \
                                                     str(person_info[3]), str(person_info[4]), str(person_info[5]), \
                                                     str(person_info[6]), str(person_info[7]), str(person_info[8]), \
                                                     str(person_info[9]), str(person_info[10]), str(person_info[11]),\
                                                     str(person_info[12]), str(person_info[13]), str(person_info[14]), \
                                                     str(person_info[15]), str(person_info[16]), str(person_info[17])

            if "." in department:
                department = department.split(".")[0]
            if "." in payable:
                payable = payable.split(".")[0]
            if "." in total_pay:
                total_pay = total_pay.split(".")[0]
            if "." in att_deduction:
                att_deduction = att_deduction.split(".")[0]
            if "." in money_award:
                money_award = money_award.split(".")[0]
            if "." in total_due:
                total_due = total_due.split(".")[0]
            if "." in social_security:
                social_security = social_security.split(".")[0]
            if "." in accumulation_fund:
                accumulation_fund = accumulation_fund.split(".")[0]
            if "." in personal_income:
                personal_income = personal_income.split(".")[0]
            if "." in dormitory_fee:
                dormitory_fee = dormitory_fee.split(".")[0]
            if "." in communications_subsidies:
                communications_subsidies = communications_subsidies.split(".")[0]
            if "." in travel_allowance:
                travel_allowance = travel_allowance.split(".")[0]
            if "." in real_pay:
                real_pay = real_pay.split(".")[0]


            html_content = \
                '''
                <html>
                <body>
                    <h3 align="center">2020年4月工资条详情</h3>
                    <p> <div face="Verdana" align="center">XXXXX有限公司</div></p>
                    <p>您好：</p>
                    <blockquote><p>2020年4月工资已发至工资卡，工资明细如下，请查收！</p></blockquote>
                    <blockquote><p>序号：{order_num}</p></blockquote>
                    <blockquote><p>工号：{job_num} </p></blockquote>
                    <blockquote><p><strong>姓名：{person_name} </strong></p></blockquote>
                    <blockquote><p>部门：{department}</p></blockquote>
                    <blockquote><p><strong>绩效工资：{merit_pay} </strong></p></blockquote>
                    <blockquote><p><strong>应发工资：{payable} </strong></p></blockquote>
                    <blockquote><p><strong>全勤奖：{total_pay} </strong></p></blockquote>
                    <blockquote><p>考勤扣款：{att_deduction} </p></blockquote>
                    <blockquote><p>奖金/补上月：{money_award} </p></blockquote>
                    <blockquote><p><strong>应发合计：{total_due} </strong></p></blockquote>
                    <blockquote><p>代扣个人社保：{social_security} </p></blockquote>
                    <blockquote><p>代扣个人公积金：{accumulation_fund} </p></blockquote>
                    <blockquote><p>代扣上月个税：{personal_income} </p></blockquote>
                    <blockquote><p>扣宿舍费：{dormitory_fee} </p></blockquote>
                    <blockquote><p>通讯补(免税)：{communications_subsidies} </p></blockquote>
                    <blockquote><p>车补(免税)：{travel_allowance} </p></blockquote>
                    <blockquote><p><strong>实发工资：{real_pay} </strong></p></blockquote>
                    
                    <blockquote><p>感谢如此特别而又优秀的你，Were are 伐木累！</p><blockquote>
    
    
                    <p align="right">财务部</p>  
                     <p align="right">2020年05月11日</p> 
    
                </body>
                </html>
                '''.format(order_num=order_num,job_num=job_num, person_name=person_name, department=department, merit_pay=merit_pay, payable=payable, total_pay=total_pay, att_deduction=att_deduction, money_award=money_award, total_due=total_due, social_security=social_security, accumulation_fund=accumulation_fund, personal_income=personal_income, dormitory_fee=dormitory_fee, communications_subsidies=communications_subsidies, travel_allowance=travel_allowance, real_pay=real_pay)

            msg = MIMEMultipart()
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))
            i = 1

            msg['From'] = _format_addr('XXXXX财务部 <%s>' % from_addr)
            msg['To'] = _format_addr(person_name + '<%s>' % email_add)
            msg['Subject'] = Header('2020年4月工资条', 'utf-8').encode()

            server = smtplib.SMTP(smtp_server, 25)
            server.starttls()  # 调用starttls()方法，就创建了安全连接
            server.login(from_addr, password)  # 登录邮箱服务器
            server.sendmail(from_addr, [email_add], msg.as_string())  # 发送信息
            server.quit()
            print("序号：" + str(order_num), "  姓名：", person_name, "已发送成功！")
        except Exception as e:
            print("发送失败" + e)


if __name__ == '__main__':
    root_dir = 'E:\\PyChram项目集合\\email\\email_dir'
    file_path = root_dir + "\\test.xlsx"
    from_addr = 'XXX@qq.com'  # 邮箱登录用户名
    password = 'XXXXX'  # 登录密码
    smtp_server = 'smtp.qq.com'  # 服务器地址，默认端口号25

    file_list = read_file(file_path)
    sendEmail(from_addr, password, smtp_server, file_list)
    print('ok')