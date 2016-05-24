#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
*
* 有两个用处：
**  1. 把邮件地址提取出来
**  2. 把邮件地址提取出来并发送邮件
*
* aha !!!!!!
*
__version__   : 0.6.1
__author__    : Eric Wang
__lastupdate__: May 23 2016

"""

import smtplib
import re
import datetime
import codecs
import time
from email.mime.text import MIMEText
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import parseaddr, formataddr
from config import files


class Configer:
    """Configure class

    """

    def __init__(self):
        # init file location
        self.config = files['config']
        self.mailist = files['mailist']
        self.mailist_saved = files['mailist_saved']
        self.mailist_set = files['mailist_set']
        self.mail_text = files['mail_text']
        self.attachment = files['attachment']
        self.logo = files['logo']
        self.profile = None
        self.mail_content = None
        self.smtp_server = 'smtp.viewwiden.com'

    def get_profile(self):
        """Get personal profiles from 'config.txt':
            Name, Email, Password, Mobile, Subject
            return a dict of profiles
        """
        with codecs.open(self.config, 'r', encoding='utf-8') as con:
            name = con.readline().replace('Name:', '').rstrip()
            email = con.readline().replace('Email address:', '').rstrip()
            password = con.readline().replace('Password:', '').rstrip()
            mobile = con.readline().replace('Mobile:', '').rstrip()
            subject = con.readline().replace('Subject:', '').rstrip()

        if not subject:
            return False
        else:
            self.profile = {'Name': name,
                            'Email': email,
                            'Pswd': password,
                            'Mobi': mobile,
                            'Sbj': subject}
            return True

    def _addr_re(self, m, msg=True):
        """Find all the email address in 'mailist.txt'
            and give warning of wrong email address
            return a list of addresses
        """
        with codecs.open(m, 'r', encoding='utf-8') as fp:
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
                        print('| WARNING: wrong addr [%s]' % a)
                if a not in addrs:
                    addrs.append(a)
        return addrs

    def get_addr(self):

        """Get email addresses from maillist and compare to the
            mail_saved to exclude the existed address

            :return : addresses
        """
        addr = self._addr_re(self.mailist)
        addr_s = self._addr_re(self.mailist_saved, False)
        print('\n| Found [%d] Address' % (len(addr)))
        count = 0
        addrs = list()
        text = '\nTIME: ' + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + '\r\n'
        with codecs.open(self.mailist_set, 'w', encoding='utf-8') as m:
            m.write(text)
            for a in addr:
                if a in addr_s:
                    count += 1
                else:
                    addrs.append(a)
                    m.write(a + '\r\n')
        print('|> [%d] Found exist' % count)
        print('|> [%d] Fresh to send' % (len(addr) - count))
        return addrs

    @staticmethod
    def _time_format(sec):
        """Format time

        :return : xx hour xx minutes xx seconds
        """
        tm = time.gmtime(sec)
        if sec >= 3600:
            h = int(time.strftime('%H', tm))
            m = int(time.strftime('%M', tm))
            if m == 0:
                return '%d hour' % h
            else:
                return '%d hour %d minutes' % (h, m)
        elif sec >= 60:
            s = int(time.strftime('%S', tm))
            m = int(time.strftime('%M', tm))
            if s == 0:
                return '%d minutes' % m
            else:
                return '%d minutes %d seconds' % (m, s)
        else:
            s = int(time.strftime('%S', tm))
            return '%d seconds' % s

    @staticmethod
    def timer(sec):
        """Count time by seconds
            return formatted time use time_format
        """
        tm = Configer._time_format(sec)
        print('| Spend %s' % tm)

    @staticmethod
    def timer_rem(tm, count, leth):
        """Count remain time
            return formatted time use time_format
        """
        if len(tm) <= 10:

            re_tm = tm[len(tm) - 1] - tm[0]
            re_tm = (re_tm / count) * (leth - count)
        else:
            re_tm = tm[len(tm) - 1] - tm[len(tm) - 10]
            re_tm = (re_tm / 10) * (leth - count)

        tm = Configer._time_format(re_tm)
        print('| TIME: remaining about %s\n' % tm)

    @staticmethod
    def spliter_msg(msg, leth=72, symbol='='):
        """To send a message like ====  xxxx  =====
            print a formatted message
        """
        a = int((leth - len(msg)) / 2) + 1
        for i in range(a):
            print(symbol, end='')
        print(' ', msg, ' ', end='')
        for j in range(a):
            print(symbol, end='')
        print('\n')

    def convert_html(self):
        """Convert the mail template to a html
            return a html string
        """
        with codecs.open(self.mail_text, 'r') as fp:
            mail_content = fp.read()
        for k, v in self.profile.items():
            if '{%s}' % k in mail_content:
                mail_content = mail_content.replace('{%s}' % k, v)

        self.mail_content = mail_content


class Mailer:
    """ Sending email main class """

    def __init__(self, name, email, pswd, sbj, mail_text, smtp_server):
        self.name = name
        self.email = email
        self.pswd = pswd
        self.sbj = sbj
        self.mail_text = mail_text
        self.smtp_server = smtp_server
        self.server = None

    @staticmethod
    def _format_addr(s):
        """Format to Name<name.xxx@xxx.com>"""
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode('utf-8'), addr))

    def send_mail(self, to_addr):
        """Send email by smtp server

            :return : True if success, False if failed
        """
        try:
            self.check_server(self.server)
            msg = self.set_msg()
            self.server.sendmail(self.email, [to_addr], msg.as_string())
            return True
        except Exception as e:
            print('| ERR0R: ', e)
            return False

    def set_msg(self):
        msg = MIMEMultipart()
        msg['From'] = Mailer._format_addr((self.name + '<%s>' % self.email))
        msg['Subject'] = Header(self.sbj, 'utf-8').encode()
        if self.mail_text:
            msg.attach(MIMEText(self.mail_text, 'html', 'utf-8'))
        else:
            raise Exception('| ERROR: Check your mail template')

        # Attach Logo
        with open(files['logo'], 'rb') as l:
            logo = MIMEImage(l.read())
            logo.add_header('Content-ID', '<Logo>')
            logo.add_header('X-Attachment-Id', 'Logo')
            msg.attach(logo)

        # Attach Pdf
        try:
            with open(files['attachment'], 'rb') as ip:
                intro = MIMEApplication(ip.read())
                intro.add_header('Content-Disposition', 'attachment', filename=files['attachment_name'])
                msg.attach(intro)
        except Exception as e:
            print('| ERROR: Wrong Attachment')
            print(e)

        return msg

    def connect_server(self):
        # Connect to smtp server
        print('| INFO: connecting to server...')
        try:
            server = smtplib.SMTP(self.smtp_server, 25)
            # server.set_debuglevel(1)
            server.login(self.email, self.pswd)
            self.server = server
            return True
        except:
            print('| ERROR: Failed to connect to server!')
            print('\n !!! PLEASE CHECK YOUR ACCOUNT / PASSWORD / NETWORK !!!\n')
            return False

    def check_server(self, server):
        try:
            if not server or server.helo()[0] != 250:
                return self.connect_server()
        except:
            return self.connect_server()

    def quit(self):
        if self.server:
            print('| INFO: QUIT SERVER...')
            self.server.quit()


def main():
    """Main function to send emails
        send emails to the mailist and save the
        addresses to the mailist_saved.
    """
    c = Configer()
    if not c.get_profile():
        print('\n\n !!! PLEASE CHECK YOUR CONFIG FILE(config.txt) !!!\n\n')
        time.sleep(3)
        exit()

    p = c.profile
    c.convert_html()
    count_all, count_success = 0, 0

    with codecs.open(c.mailist_saved, 'r+', encoding='utf-8') as a:
        a.read()
        a.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' By.' + p['Name'] + '\r\n')

    tm = [time.time()]

    Configer.spliter_msg(msg='ADDRESS: get address and remove exist ones')
    addrs = c.get_addr()
    if not addrs:
        input('| >>> NO ADDERSS TO SEND, PRESS　ENTER TO EXIT!')
        exit()
    else:
        input('| >>> PRESS ENTER TO START!')

    Configer.spliter_msg(msg='START: sending emails to [%d] address(es)' % len(addrs))

    m = Mailer(p['Name'], p['Email'], p['Pswd'], p['Sbj'], c.mail_content, c.smtp_server)
    m.check_server(server=None)

    for addr in addrs:
        print('|', end='')
        for i in range(len(addr) + 36):
            print('-', end='')
        print('\n', end='')
        print('| DEBUG: sending mail to [%s]' % addr)
        if m.send_mail(addr):
            tm.append(time.time())
            count_all += 1

            print('| INFO: success to send mail to [%s]' % addr)
            print('| PERCENT: ******* %.1f%% complete!! (%d/%d)*********'
                  % ((count_all / len(addrs) * 100), count_all, len(addrs)))

            if count_all > 0 and count_all != (len(addrs)):
                c.timer_rem(tm, count_all, len(addrs))
            count_success += 1

            with codecs.open(c.mailist_saved, 'r+', encoding='utf-8') as a:
                if addr not in a:
                    a.read()
                    a.write(addr + '\r\n')
        else:
            tm.append(time.time())
            count_all += 1
            print('| ERROR: failed to send mail to [%s]' % addr)
            print('| PERCENT: ******* %.1f%% complete!! (%d/%d)*********'
                  % ((count_all / len(addrs) * 100), count_all, len(addrs)))
            if count_all > 0 and count_all != (len(addrs)):
                Configer.timer_rem(tm, count_all, len(addrs))

    Configer.spliter_msg(msg='FINISH: sending mails')
    print('| Success [%d]' % count_success)
    print('| Failed [%d]' % (count_all - count_success))
    m.quit()


if __name__ == '__main__':
    main()
