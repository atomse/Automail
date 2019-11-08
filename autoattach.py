# _*_ coding: utf-8 _*_


# ref: https://www.cnblogs.com/chouxianyu/p/11270101.html
# with modifications


import os
import sys
import shutil
import pickle
import argparse
import logging
import time

import poplib
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

import yaml
import chardet


BASEDIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILENAME = 'config.yaml'
CONFIG_TEMPLATE_FILENAME = 'config_template.yaml'

CONFIG_FILEPATH = os.path.join(BASEDIR, CONFIG_FILENAME)
CONFIG_TEMPLATE_FILEPATH = os.path.join(BASEDIR, CONFIG_TEMPLATE_FILENAME)

DEFAULT_METAPATH = os.path.join(BASEDIR, 'metadata')

logging.basicConfig()
if __name__ == '__main__':
    logger = logging.getLogger('automail')
else:
    logger = logging.getLogger(__name__)


def get_config(config_filename=CONFIG_FILEPATH) -> dict:
    """
    read config.yaml, which stores
        * email accounts and password
        * save_path
    """
    with open(config_filename) as fd:
        res = yaml.safe_load(fd)
    logger.debug(f'config: {res}')
    return res


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    return value


def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            if header == 'Subject':
                subject = decode_str(value)
                headers['Subject'] = subject
            if header == 'From':
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            if header == 'To':
                all_cc = value.split(',')
                to = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            if header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    cc_addr = u'%s <%s>' % (name, addr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc)
    return headers


def get_email_content(message, headers, save_path):
    attachments = []
    save_path = os.path.join(save_path, headers['From'])
    os.makedirs(save_path, exist_ok=True)
    for part in message.walk():
        filename = part.get_filename()
        if filename:
            filename = decode_str(filename)
            data = part.get_payload(decode=True)
            abs_filename = os.path.join(save_path, filename)
            attach = open(abs_filename, 'wb')
            attachments.append(filename)
            attach.write(data)
            attach.close()
    return attachments


def fetch_email_data(mails, last_mails=None, max_mail: int = None):
    if last_mails is None:
        for i in range(1, min(len(mails), max_mail or len(mails))+1):
            yield i
    else:
        assert isinstance(last_mails, list)
        for m in mails:
            num_m, id_m = m.decode().split()
            for oldm in last_mails:
                id_old = oldm.decode().split()[-1]
                if id_m == id_old:
                    already_fetched = True
                    break
            if already_fetched:
                continue
            yield int(num_m)


def fetch_email_account(address, password, pop3_server,
                        save_path, last_mails=None):
    # 账户信息
    # address = 'xxx@xxx.com.cn'
    # password = 'xxx'
    # pop3_server = 'xxx.xxx.com.cn'
    # 连接到POP3服务器，带SSL的:
    server = poplib.POP3_SSL(pop3_server)
    # 可以打开或关闭调试信息:
    server.set_debuglevel(0)
    # POP3服务器的欢迎文字:
    logger.debug(server.getwelcome())
    # 身份认证:
    server.user(address)
    server.pass_(password)
    # stat()返回邮件数量和占用空间:
    msg_count, msg_size = server.stat()
    logger.debug(f'message count: {msg_count}')
    logger.debug(f'message size: {msg_size} bytes')
    # b'+OK 237 174238271' list()响应的状态/邮件数量/邮件占用的空间大小
    resp, mails, octets = server.list()
    logger.debug(f"{resp}, {octets}")
    for i in fetch_email_data(mails, last_mails):
        resp, byte_lines, octets = server.retr(i)
        # 转码
        str_lines = []
        for x in byte_lines:
            # import pdb; pdb.set_trace()
            encoding = chardet.detect(x)['encoding'] or 'gbk'
            if encoding in ['windows-1255', 'iso-8859-']:
                encoding = 'gbk'
            logger.debug(f"{x}, encoding: {encoding}")
            str_lines.append(x.decode(encoding))
        # 拼接邮件内容
        msg_content = '\n'.join(str_lines)
        # 把邮件内容解析为Message对象
        msg = Parser().parsestr(msg_content)
        headers = get_email_headers(msg)
        attachments = get_email_content(msg, headers, save_path)
        logger.debug(f"-----------------------------'")
        if 'Subject' in headers:
            logger.debug(f"subject: {headers['Subject']}")
        if 'From' in headers:
            logger.debug(f"from: {headers['From']}")
        if 'To' in headers:
            logger.debug(f"to: {headers['To']}")
        if 'Cc' in headers:
            logger.debug(f"cc: {headers['Cc']}")
        if 'Date' in headers:
            logger.debug(f"date: {headers['Date']}")
        logger.debug(f"attachments:  {attachments}")
        logger.debug(f"-----------------------------'")
    server.quit()
    return mails


def load_last_mails(last_mails_path):
    """
    xx
    """
    if os.path.exists(last_mails_path):
        with open(last_mails_path, 'rb') as f:
            res = pickle.load(f)
        return res
    return None


def save_last_mails(last_mails, last_mails_path):
    """
    save mails data to metapath/{showname}
    """
    with open(last_mails_path, 'wb') as f:
        pickle.dump(last_mails, f)


def main(config_filename=CONFIG_FILEPATH):
    config = get_config(config_filename)
    metapath = config.get('metapath', DEFAULT_METAPATH)
    os.makedirs(metapath, exist_ok=True)
    attachpath = config.get('attachpath')
    nostop = config.get('nostop', False)
    interval = int(config.get('interval', 300))
    while True:
        for showname, email in config['emails'].items():
            address, password, servertype = email['address'], email['password'], \
                email.get('servertype', 'pop3')
            assert servertype == 'pop3', f"{servertype} not supported"
            address_suffix = address.split('@')[-1]
            server = config[servertype][address_suffix]
            last_mails_path = os.path.join(metapath, showname)
            last_mails = load_last_mails(last_mails_path)
            attachement_save_path = os.path.join(attachpath, showname)
            last_mails = fetch_email_account(address, password, server,
                                             attachement_save_path,
                                             last_mails=last_mails)
            save_last_mails(last_mails, last_mails_path)
        if not nostop:
            break
        time.sleep(interval)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--config', default=CONFIG_FILEPATH, type=str)
    parser.add_argument('--generate', action='store_true')
    parser.add_argument('--debug', action='store_true')
    args = parser.parse_args()
    if args.generate:
        shutil.copy(CONFIG_TEMPLATE_FILEPATH, CONFIG_FILENAME)
        sys.exit(0)
    if args.debug:
        logger.setLevel(logging.DEBUG)
    main(args.config)
