import base64

import quopri
from copy import copy
from datetime import datetime
from email import message_from_bytes
from email.header import decode_header
from typing import Union

from bs4 import BeautifulSoup


class DecodeException(Exception):
    pass


class CreateFolderFailed(Exception):
    pass


default_attach_filter = ('xls', 'xlsx', 'pdf', 'doc', 'docx', 'zip', 'rar')


def compose_multi_email_user_string(users: Union[str, list]):
    """
    ['aa@test.com.cn'] => 'aa@test.com.cn'
    ['aa@test.com.cn', 'bb@test.com.cn'] => 'aa@test.com.cn, bb@test.com.cn'
    :param users:
    :return:
    """
    if isinstance(users, list):
        ret = ', '.join(users) if len(users) > 1 else users[0]
    else:
        ret = users
    return ret


def get_decode_content(encode_str):
    """
    获取邮件解码后的内容
    :param encode_str:
    :return:
    """
    decode_str, encoding = decode_header(encode_str)[0]
    # decode 失败处理
    if decode_str == b'"' and '=' in encode_str:
        encode_str = encode_str.replace('"=', '=').replace('="', '=')
        decode_str, encoding = decode_header(encode_str)[0]

    if encoding == 'unknown-8bit' or encoding == 'gb2312':
        # 解码失败默认用 gbk
        encoding = 'gbk'

    if encoding is not None:
        decode_str = decode_str.decode(encoding)

    if '�' in decode_str:
        raise DecodeException('Decode_content_error.')
    return decode_str


def get_subject(msg):
    """
    获取邮件主题
    :param msg: Message 对象
    :return:
    """
    return get_decode_content(msg.get('Subject'))


def get_mail_date(msg):
    """
    获取邮件日期
    :param msg: Message 对象
    :return:
    """
    # https://www.simplifiedpython.net/python-convert-string-to-datetime/
    # 'Wed, 25 Sep 2019 20:59:23 +0800' ==> '25 Sep 2019 20:59:23'
    if msg.get('Date') is None:
        date_str = msg.get('received').split(',')[-1].split('+')[0].strip()
    else:
        date_str = msg.get('Date').split(',')[-1].split('+')[0].strip()
    date_ret = None
    for date_pattern in ('%d %b %Y %H:%M:%S', '%a %b %d %H:%M:%S %Y'):
        try:
            date_ret = datetime.strptime(date_str, date_pattern)
            if isinstance(date_ret, datetime):
                return date_ret
        except ValueError:
            continue
    if date_ret is None:
        raise ValueError('mail date pattern wrong. Check code.')


def get_sender(msg):
    """
    TODO 整理一下获取发件人邮箱，显示名称
    获取邮件发件人名称，不是发件人邮件。
    '"eservice@nbcb.cn" <eservice@nbcb.cn>' => eservice@nbcb.cn
    '<yywbfa@cmschina.com.cn>' => yywbfa
    '=?GB2312?B?uqPNqNakyK8=?= <ppos@htsec.com>' => 海通证券
    :param msg: Message 对象
    :return:
    """
    raw_sender = msg.get('From')
    if '?' in raw_sender:
        sender = get_decode_content(raw_sender)
    else:
        if '"' in raw_sender:
            start = raw_sender.find('"') + 1
            end = raw_sender.find('"', start)
            sender = raw_sender[start:end]
        else:
            sender = raw_sender.split('@')[0]
            sender = sender.strip('<>')
    return sender


def has_attachment(msg):
    """
    check whether has attachments
    :param msg:
    :return:
    """
    attachment_num = msg.get('X-ATTACHMENT-NUM')
    # plain text 也算 attachment
    if attachment_num < 2:
        return False
    else:
        return True


def get_attachment_name(part):
    """
    get chinese attachment name if part.get_filename() fails.
    :param part: Message object
    :return:
    """
    file_name = part.get_filename()
    try:
        file_name = get_decode_content(file_name)
        return file_name
    except (AttributeError, DecodeException):
        # Use shallow copy avoid string huge payload
        part_shallow = copy(part)
        part_shallow._payload = None
        part_str = part_shallow.as_string()
        part_str_elements = part_str.split('unknown-8bit')
        for item in part_str_elements:
            if '?b?' in item:
                file_name_b64 = item.split('?=')[0].strip().replace('?b?', '')
                for encoding in ('gbk', 'gb2312', 'utf-8'):
                    try:
                        file_name = base64.b64decode(file_name_b64).decode(encoding)
                    except UnicodeDecodeError:
                        continue
                file_name = file_name.split('=')[-1].replace('"', '')
                return file_name


def get_attachment(msg, file_filter=default_attach_filter):
    """
    获取邮件里的附件
    :param msg:
    :param file_filter:
    :return:
    """
    attachments = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart': continue
        if part.get_content_maintype() == 'text': continue
        if part.get('Content-Disposition') is None: continue
        file_name = get_attachment_name(part)
        file_extension = file_name.split('.')[-1]
        file = part.get_payload(decode=True)
        if file_extension not in file_filter:
            continue
        attachments.append({'name': file_name, 'file': file})
    return attachments


def get_charset(content_type):
    """
    获取 content-type 中的 charset
    :param content_type: “part.get('Content-Type')” 获取的 content-type
    :return:
    """
    charset = content_type.split('=')[1].lower()
    if 'gb2312' in charset or 'gbk' in charset:
        charset = 'gbk'
    elif 'utf-8' in charset:
        charset = 'utf-8'
    elif 'gb18030' in charset:
        charset = 'gb18030'
    else:
        charset = None
    return charset


def get_body_content(msg, html_parser='html.parser'):
    """
    获取邮件 body 信息
    :param msg:
    :param html_parser: BeautifulSoup features (html.parser, lxml, html5lib) see https://stackoverflow.com/a/60254943
    :return:
    """
    content = None
    for part in msg.walk():
        if part.get_content_type() in ("text/html", "text/plain"):
            content_encode = part.get_payload()
            content_transfer_encoding = part.get('Content-Transfer-Encoding')
            if part.get('Content-Type'):
                charset = get_charset(part.get('Content-Type'))
                if content_transfer_encoding == 'quoted-printable':
                    try:
                        content = quopri.decodestring(content_encode).decode(charset)
                    except LookupError:
                        content = quopri.decodestring(content_encode).decode('gbk')
                elif content_transfer_encoding == 'base64':
                    if charset is None:
                        # charset 为 None 的时候这部分数据可以不要了
                        continue
                    content = base64.b64decode(content_encode).decode(charset)
                elif content_transfer_encoding == '7bit' or content_transfer_encoding == '8bit':
                    content = content_encode
                else:
                    raise Exception('Content 无法处理。')

                if part.get_content_type() == 'text/html':
                    content = {'text': BeautifulSoup(content, features=html_parser).text, 'html': content}
                else:
                    content = {'text': content}

    return content


def eml_to_mail_info(eml, attach_filter=None, html_parser='html.parser'):
    """
    邮件 eml 格式转换成 mail 信息
    :param eml: 邮件 eml
    :param attach_filter: 附件过滤器
    :param html_parser: BeautifulSoup features (html.parser, lxml, html5lib) see https://stackoverflow.com/a/60254943
    :return: Mail object
    """
    email_message = message_from_bytes(eml)
    try:
        subject = get_subject(email_message)
    except Exception:
        raise DecodeException('Decode_failed/subject')

    try:
        mail_date = get_mail_date(email_message)
    except Exception:
        raise DecodeException('Decode_failed/mail_date')

    try:
        body = get_body_content(email_message, html_parser)
    except Exception:
        raise DecodeException('Decode_failed/body')

    try:
        attachments = get_attachment(email_message, file_filter=attach_filter) or None
    except Exception:
        raise DecodeException('Decode_failed/attachments')

    try:
        sender = get_sender(email_message)
    except Exception:
        raise DecodeException('Decode_failed/sender')

    return {'subject': subject, 'mail_date': mail_date, 'body': body, 'attachments': attachments, 'sender': sender}
