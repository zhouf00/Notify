from pandas import read_excel, DataFrame
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def load_file(path):
    df = DataFrame(read_excel(path))
    return dict(df.values)


class SendMail(object):

    def __init__(self, *args):
        email_host, email_port, mail_name, mail_passwd, from_phone, from_name, text, header = args[0]
        server = smtplib.SMTP_SSL(email_host, email_port)
        server.login(mail_name, mail_passwd)
        self._server = server
        self._from = ("%s<%s>")%(from_name, mail_name)
        self._from_name = from_name
        self._from_phone = from_phone
        self._text = text
        self._header = header

    def send(self, attachment, to_user, to, cc, file, html):
        message = MIMEMultipart()
        message['Subject'] = self._text + self._header   # 标题
        message['From'] = self._from  # 发件人
        message['To'] = to    # 收件人
        message['Cc'] = cc    # 抄送人

        # 加入正文
        # 加入签名
        html_dict = {
            'to_user': to_user,
            'text': self._text,
            'from_name': self._from_name,
            'from_phone': self._from_phone,
            'from_mail': self._from
        }
        html = html.format(**html_dict)
        msgText = MIMEText(html, 'html', 'utf-8')
        message.attach(msgText)

        # 添加附件
        pdfpart = MIMEApplication(open(file, 'rb').read())
        # 设置附件头
        pdfpart.add_header('Content-Disposition', 'attachment', filename=('gbk', '', '%s.pdf'%attachment))
        message.attach(pdfpart)

        try:
            self._server.sendmail(self._from, to, message.as_string())
        except Exception as e:
            # print(e)
            return "<%s>发送失败\n" % (to_user)
        else:
            return "<%s>发送成功\n" % (to_user)

    def get_html(self, file):
        with open(file, 'r', encoding='utf8') as f:
            html = f.read()
            return html


def send_mail(*args):

    email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month,\
    row_name, row_email, pdf_name, pdf_file = args
    sub = "%s工资条-%s"
    server = smtplib.SMTP_SSL(email_host, email_port)
    server.login(mail_name, mail_passwd)
    user = ("%s<%s>")%(name, mail_name)
    message = MIMEMultipart()
    message['Subject'] = sub%(month, pdf_name)
    message['From'] = user
    message['To'] = row_email
    # message['Cc'] = cc    # 抄送人

    # 加入正文
    # 加入签名
    html = sign_names(row_name, month, name, mail_name, mail_name_phone)
    msgText = MIMEText(html,'html','utf-8')
    message.attach(msgText)

    # 添加附件
    pdfpart = MIMEApplication(open(pdf_file, 'rb').read())
    # 设置附件头
    pdfpart.add_header('Content-Disposition', 'attachment', filename=('gbk', '', pdf_name))
    message.attach(pdfpart)

    try:
        server.sendmail(user, row_email, message.as_string())
    except Exception as e:
        #print(e)
        return "<%s>发送失败\n"%(row_name)
    else:
        return "<%s>发送成功\n"%row_email

def sign_names(name, month, rec_name, rec_mail, phone):
    return """
    <div style='font-size:14px;' class='black'>
%s:
<br>  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
感谢您对公司做出的贡献和努力！
<br>
现向您发送%s工资条，详见附件，如有疑问请随时和我联系，谢谢！
</div>
  <br>
  <br>
  <br>
<hr style="width: 210px; height: 1px;" color="#b5c4df" size="1" align="left">
<div style="font-family: lucida Grande, Verdana;font-size:10.5pt;">
    <div class="MsoNormal" style="">
        <b>%s</b>
    </div>
    <div class="MsoNormal" style="">
        <img src="https://exmail.qq.com/cgi-bin/viewfile?type=signature&amp;picid=ZX0227-vDWs9uGRMpO3DyV5N3wCO8o&amp;uin=715200393">
    </div>
    <div style='color:gray;'>
      <div style="">浙江中自庆安新能源技术有限公司 | <span style='font-size:9pt;'>Zhejiang WindIT Technology Co., Ltd.</span></div>
      <div style="font-size:9pt;line-height: 15.75pt;">Email：&nbsp;<a href="mailto:%s" target="_blank"
          style="outline: none; color: rgb(44, 74, 119);">%s</a>&nbsp;Cell：+86 %s</div>
      <div style="font-size:9pt;line-height: 15.75pt;">Phone：+86 571 2899 5830 Fax：+86 571 2899 5841</div>
      <div style="">地址：浙江杭州经济技术开发区6号路260号中自科技园1幢5楼 |
          <span style='font-size:9pt;'>Add: F5, B1, Chitic Science &amp;
          Technology Park, No.260,6th Road,Xiasha, EDA, Hangzhou, Zhejiang, China (310018)
          </span></div>
    </div>
    <div class="black" style="font-size:10.5pt;">
        设备健康成就工业智慧！
        <span style='font-size:9pt;'>
            | Enabling the industrialintelligence!
        </span>
    </div>
  </div>
  <div class="black" style=""><br></div>
  <div style="font-family:lucida Grande, Verdana;font-size: 9pt;color: gray;">保密声明:
    本电子邮件和任何附件仅供所指定涉及的个人或实体机构单独使用,
    并且可能包含有特权保密、机密和严禁披露或未经授权使用的信息。如果您不是本电子邮件的指定接收人,则特此声明，发件人严禁您使用、传播、分发或复制本电子邮件或本电子邮件所包含的信息。如果您误收到此传输，请通过回复电子邮件通知发件人，并从系统中删除所有副本。谢谢您的合作！&nbsp;<br><span
      style='font-size:8pt;'>Confidentiality
      Notice: This e-mail and any attachments are intended for the sole use of the individual or entity to which it is
      addressed and may contain information which is privileged, confidential and prohibited from disclosure or
      unauthorized use. If you are not the intended recipient of this email, you are hereby notified that any use,
      dissemination,distribution, or copying of this email or the information contained in this email is strictly
      prohibited by the sender. If you have received this transmission in error, please notify the sender by return
      email
      and delete all copies from your system. Thank you for your cooperation!</span>
  </div>
    """%(name, month, rec_name, rec_mail, rec_mail, phone)


if __name__ == '__main__':

    email_host = "smtp.exmail.qq.com"
    email_port = 465
    mail_name = "model@windit.com.cn"
    mail_passwd = "Zzqa2018"
    name = "文档管理"
    month = "11月"

    load_file('d:/study/Notify/conf/邮箱管理.xlsx')