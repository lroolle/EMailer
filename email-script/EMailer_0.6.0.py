 #!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Emailer
    Have FUuUN!
"""


import smtplib, re, datetime, codecs, time
from time import strftime, gmtime
from email.mime.text import MIMEText
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import parseaddr, formataddr


class Configer:
    """Configure class
        get_profile : get personal info
        get_addr    : get addresses from maillist
        convert_html: convert mail template to html
        timer       : count time
        timer_rem   : count remaining time
        spliter_msg : send message '== xxx =='
    """

     # init file location
    def __init__(self):
        self.config_file = '.\config.txt'
        self.addr_list = '.\mailList.txt'
        self.mail_set = '.\maillist_set.txt'
        self.addr_saved = '.\maillist_saved.txt'
        self.mail_text = '.\Content_Files\mail_text.html'
        self.smtp_server = 'smtp.viewwiden.com'

    def get_profile(self, config_file):
        """Get personal profiles from 'config.txt':
            Name, Email, Password, Mobile, Subject
            return a dict of profiles
        """
        with codecs.open(self.config_file, 'r', encoding='utf-8') as con:
            name = con.readline().replace('Name:', '').rstrip()
            email = con.readline().replace('Email address:', '').rstrip()
            password = con.readline().replace('Password:', '').rstrip()
            mobile = con.readline().replace('Mobile:', '').rstrip()
            subject = con.readline().replace('Subject:', '').rstrip()
        if email == '' or password == '' or password == '' or mobile == '' or mobile == '':
            return True
        else:
            profile = {'Name': name, 'Email': email, 'Pswd': password, 'Mobi': mobile, 'Sbj': subject}
            return profile

    def __addr_re(self, maillist, msg=True):
        """Find all the email address in 'mailist'
            and give warning of wrong email address
            return a list of addresses
        """
        with codecs.open(maillist, 'r', encoding='utf-8') as fp:
            mailre = re.compile(r'[\w\.\-]+@[\w\.\-]+')
            addr = mailre.findall(fp.read())
            print(fp.read())
            addr = list(addr)
            addrs = list()
            for a in addr:
                if msg:
                    if not re.match(r'^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]'
                                    r'{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))'
                                    r'([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$', a):
                        print('| WARNING: wrong addr [%s]'%a)
                if a not in addrs:
                    addrs.append(a)
        return addrs


    def get_addr(self, addr_list, addr_saved, ms):
        """Get email addresses from maillist and compare to the
            mail_saved to exclude the existed address
            return list of address never exist in the 'maillist_saved.txt'
        """
        addr = self.__addr_re(self.addr_list)
        addr_s = self.__addr_re(self.addr_saved, False)
        print('\n| Found [%d] Address' %(len(addr)))
        count = 0
        addrs = list()
        with codecs.open(ms,'w', encoding='utf-8') as m:
            m.write('\nTIME: ' + datetime.datetime.now().strftime(""
                                                                  "%Y-%m-%d %H:%M:%S")+'\r\n')
            for a in addr:
                if a in addr_s:
                    count += 1
                else:
                    addrs.append(a)
                    m.write(a+'\r\n')
        print('|> [%d] Found exist' %count)
        print('|> [%d] Fresh to send' %(len(addr)-count))
        return addrs

    def __time_format(self, sec):
        """Format time
            return xx hour xx minutes xx seconds
        """
        tm = gmtime(sec)
        if sec >= 3600:
            h = int(strftime('%H',tm))
            m = int(strftime('%M',tm))
            if m == 0:
                return '%d hour'%h
            else:
                return '%d hour %d minutes'%(h,m)
        elif sec >= 60:
            s = int(strftime('%S',tm))
            m = int(strftime('%M',tm))
            if s == 0:
                return '%d minutes'%m
            else:
                return '%d minutes %d seconds'%(m,s)
        else:
            s = int(strftime('%S',tm))
            return '%d seconds'%s

    def timer(self, sec):
        """Count time by seconds
            return formatted time use time_format
        """
        tm = self.__time_format(sec)
        print('| Spend %s'%tm)

    def timer_rem(self, tm, count, leth):
        """Count remain time
            return formatted time use time_format
        """
        if len(tm) <= 10:
            
            re_tm = tm[len(tm)-1] - tm[0]
            re_tm = (re_tm / count)*(leth - count)
        else:
            re_tm = tm[len(tm)-1] - tm[len(tm)-10]
            re_tm = (re_tm / 10)*(leth - count)
        
        tm = self.__time_format(re_tm)
        print('| TIME: remaining about %s\n'%tm)

    def spliter_msg(self, msg, leth=72,  symbol='='):
        """To send a message like ====  xxxx  =====
            print a formatted message
        """
        a = int((leth-len(msg))/2)+1
        for i in range(a):
            print('=',end = '')
        print(' ',msg,' ',end = '')
        for j in range(a):
            print('=',end = '')
        print('\n')
        
    def convert_html(self,profile,mail_text,to_per='Sir/Madam'):
        """Convert the mail template to a html
            return a html string
        """
        name = profile['Name']
        email = profile['Email']
        mobi = profile['Mobi']
        a = '''<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name=Generator content="Microsoft Word 15 (filtered)">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:宋体;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:"\@宋体";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	text-align:justify;
	text-justify:inter-ideograph;
	font-size:10.5pt;
	font-family:"Calibri",sans-serif;}
a:link, span.MsoHyperlink
	{color:#0563C1;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{color:#954F72;
	text-decoration:underline;}
.MsoChpDefault
	{font-family:"Calibri",sans-serif;}
 /* Page Definitions */
 @page WordSection1
	{size:595.3pt 841.9pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;
	layout-grid:15.6pt;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
</head>
<body lang=ZH-CN link="#0563C1" vlink="#954F72" style='text-justify-trim:punctuation'>
<div class=WordSection1 style='layout-grid:15.6pt'>
<p class=MsoNormal style='line-height:115%'><b><span lang=EN-US
style='font-size:11.0pt;line-height:115%;color:black'>Dear '''
        b = ''',</span></b></p>
<p class=MsoNormal style='line-height:115%'><b><span lang=EN-US
style='font-size:11.0pt;line-height:115%;color:black'>&nbsp;</span></b></p>
<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%;color:black'>How are you?</span></p>
<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%;color:black'>This is '''
        c = ''' from Shanghai View Widen Co., Ltd. I would like to take this opportunity to
invite you to join one of the best business meeting features on Chinese
imported food and beverage business “The 11<sup>th</sup> China Imported Food
&amp; Beverage Business Match Meeting” which will be held on May 10<sup>th</sup>
&amp; 11<sup>th</sup>, 2016 in Shanghai Hongqiao State Guest Hotel. You may
find more details in the attached.</span></p>

<p class=MsoNormal style='line-height:115%'><span lang=EN-US style='font-size:
11.0pt;line-height:115%;color:black'>The event is consisted of 2 major
programs:</span></p>

<p class=MsoNormal style='line-height:115%'><span lang=EN-US style='font-size:
11.0pt;line-height:115%;color:black'>- Keynote speech delivered by government
officials, revealing the latest Chinese import regulations, market trends and
new business models.</span></p>

<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%;color:black'>- Business
match one-on-one meetings with professional Chinese buyers in 5-star
environment with no interference.</span></p>

<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%;color:black'>Delegates will
have the opportunity to meet 300 pre-select top notch Chinese market players !</span></p>

<p class=MsoNormal style='line-height:115%'><span lang=EN-US style='font-size:
11.0pt;line-height:115%;color:black'>Let me know if you or any of your
colleague would be interested in participating. I will send you a business
proposal to reveal participating conditions if you are interested.</span></p>

<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%;color:black'>More
information on&nbsp;</span><span lang=EN-US style='font-size:10.0pt;line-height:
115%'><a href="http://www.viewwiden.com/"
title="blocked::http://www.viewwiden.com/&#10;http://www.viewwiden.com/&#10;blocked::http://www.viewwiden.com/&#10;http://www.viewwiden.com/"><span
style='font-size:11.0pt;line-height:115%'>www.viewwiden.com</span></a></span></p>

<p class=MsoNormal style='margin-bottom:12.0pt;line-height:115%'><span
lang=EN-US style='font-size:11.0pt;line-height:115%'>Cheers!</span></p>

<p class=MsoNormal align=left style='text-align:left;line-height:115%;
layout-grid-mode:char'><b><span lang=EN-US style='font-size:11.0pt;line-height:
115%;color:black'>'''
        d = '''</span></b></p>

<p class=MsoNormal align=left style='text-align:left;line-height:115%;
layout-grid-mode:char'><b><span lang=EN-US style='font-size:11.0pt;line-height:
115%;color:black'>Shanghai View Widen Consulting Co., Ltd.</span></b></p>

<p class=MsoNormal align=left style='text-align:left;line-height:115%;
layout-grid-mode:char'><span lang=EN-US style='font-size:9.0pt;line-height:
115%'>&nbsp;</span><span lang=EN-US style='font-size:9.0pt;line-height:115%'><img
border=0 width=132 height=27 src="cid:Logo"></span></p>

<p class=MsoNormal align=left style='text-align:left;line-height:115%;
layout-grid-mode:char'><span lang=EN-GB style='font-size:9.0pt;line-height:
115%;color:navy'>'</span><span lang=EN-GB style='font-size:9.0pt;line-height:
115%;color:gray'> +86(0)21-61472535</span><span lang=EN-GB style='font-size:
9.0pt;line-height:115%;color:#1F497D'>&nbsp;</span><span lang=EN-GB
style='font-size:9.0pt;line-height:115%;color:gray'>
/+86(0)21-61472535*8009(fax)/+86 '''
        e = '''(Mobile)</span><span lang=EN-US
style='font-size:9.0pt;line-height:115%'>&nbsp;<br>
</span><span lang=EN-GB style='font-size:9.0pt;line-height:115%;color:gray'>&nbsp;</span><span
lang=EN-US style='font-size:10.0pt;line-height:115%'><a
href="mailto:'''
        f = '''"><span lang=EN-GB style='font-size:9.0pt;
line-height:115%'>'''
        g = '''</span></a></span><span lang=EN-GB
style='font-size:9.0pt;line-height:115%;color:gray'> &nbsp;Website:</span><span
lang=EN-US style='font-size:10.0pt;line-height:115%'><a
href="http://www.viewwiden.com/" title="blocked::http://www.viewwiden.com/"><span
lang=EN-GB style='font-size:10.5pt;line-height:115%;color:gray'>www.viewwiden.com</span></a></span></p>

<p class=MsoNormal align=left style='text-align:left;line-height:115%;
layout-grid-mode:char'><span lang=EN-GB style='font-size:9.0pt;line-height:
115%;color:green'>P.S.</span><span lang=EN-GB style='font-size:9.0pt;
line-height:115%'> <i><span style='color:green'>Please consider the environment
before printing this email</span></i></span></p>

<p class=MsoNormal style='line-height:115%'><span lang=EN-US>&nbsp;</span></p>

<p class=MsoNormal style='line-height:115%'><span lang=EN-US style='font-size:
10.0pt;line-height:115%'>&nbsp;</span></p>
</div>
</body>
</html>'''

        fmail_text = a + to_per + b + name + c + name + \
                     d + mobi + e + email + f + email + g
        return fmail_text


class Mailer:
    
    """ Sending email main class """

    def __format_addr(self, s):
        """Format to Name<name.xxx@xxx.com>"""
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode('utf-8'), addr))

    def __send_mail(self, name, pswd, email, sbj, fmail_text,
                  to_addr, smtp_server,sendpdf=True):
        """Send email by smtp server
            return True if success, False if failed
        """
        msg = MIMEMultipart()
        msg['From'] = self.__format_addr((name+'<%s>' % email))
        msg['To'] = to_addr
        msg['Subject'] = Header(sbj, 'utf-8').encode()
        msg.attach(MIMEText(fmail_text, 'html', 'utf-8'))

        # Attach Logo
        with open('.\Content_Files\Logo.png', 'rb') as l:
            logo = MIMEImage(l.read())
            logo.add_header('Content-ID', '<Logo>')
            logo.add_header('X-Attachment-Id', 'Logo')
            msg.attach(logo)

        # Attach Pdf
        if sendpdf:
            with open('.\Content_Files\Intro.pdf','rb') as ip:
                intro = MIMEApplication(ip.read())
                fn = "The 11th China Imported Food & Beverage Business Meeting.pdf"
                intro.add_header('Content-Disposition', 'attachment', filename=fn)
                msg.attach(intro)
        # Connect to smtp server
        try:
            server = smtplib.SMTP(smtp_server, 25)
            #server.set_debuglevel(1)
            server.login(email, pswd)
            server.sendmail(email, [to_addr], msg.as_string())
            server.quit()
            return True
        except Exception as e:
            print('| ERR0R: ', e)
            return False

    def mailer(self, addrs):
        """Main function to send emails
            send emails to the maillist and save the
            addresses to the maillist_saved.
        """
        c = Configer()
        profile = c.get_profile(c.config_file)
        n = profile['Name']
        e = profile['Email']
        p = profile['Pswd']
        s = profile['Sbj']
        f = c.convert_html(profile=profile, mail_text=c.mail_text)
        count = 0
        count1 = 0
        with codecs.open(c.addr_saved, 'r+', encoding='utf-8') as a:
            a.read()
            a.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+' By.'+n+'\r\n')
        tm = [time.time()]
        c.spliter_msg(msg='START: sending emails to [%d] address(es)' % len(addrs))
        for addr in addrs:
            print('|', end='')
            for i in range(len(addr)+36):
                print('-', end='')
            print('\n', end='')
            print('| DEBUG: sending mail to [%s]' % addr)
            if self.__send_mail(name=n, pswd=p, email=e, sbj=s,
                              fmail_text=f, to_addr=addr, smtp_server=c.smtp_server):
                tm.append(time.time())
                count += 1
                print('| INFO: success to send mail to [%s]' % addr)
                print('| PERCENT: ******* %.1f%% complete!! (%d/%d)*********'
                      % ((count/len(addrs)*100), count, len(addrs)))
                if count > 0 and count != (len(addrs)):
                    c.timer_rem(tm, count, len(addrs))
                count1 += 1
                with codecs.open(c.addr_saved, 'r+', encoding='utf-8') as a:
                    if addr not in a:
                        a.read()
                        a.write(addr+'\r\n')
            else:
                tm.append(time.time())
                count += 1
                print('| ERROR: failed to send mail to [%s]' % addr)
                print('| PERCENT: ******* %.1f%% complete!! (%d/%d)*********'
                      % ((count/len(addrs)*100), count, len(addrs)))
                if count > 0 and count != (len(addrs)):
                    c.timer_rem(tm,count,len(addrs))
        c.spliter_msg(msg='FINISH: sending mails')
        print('| Success [%d]' % count1)
        print('| Failed [%d]'%(count-count1))


if __name__ == '__main__':
    st = time.time()
    m = Mailer()
    c = Configer()
    c.get_profile(c.config_file)
    config = c.get_profile(c.config_file)

    # Judge if config file is right
    if config is True:
        print('\n\n !!! Please Edit Your Config File !!!\n\n')
        time.sleep(3)
        exit()
    c.spliter_msg(msg='ADDRESS: get address and remove exist ones')
    addrs = c.get_addr(c.addr_list, c.addr_saved, c.mail_set)
    i = input('\n| Press Enter If U Want to Start Sending Emails： ')

    # Start to send emails
    if len(addrs) != 0 and i != 'q':
            m.mailer(addrs)
            et = time.time()
            c.timer(et-st)
            # input()
    else:
        print('\n| No Sending Emails...')
        time.sleep(3)
        exit()
