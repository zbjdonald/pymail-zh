import logging
import ssl

import requests
from imapclient import IMAPClient, SEEN
from imapclient.exceptions import IMAPClientError
from tqdm import tqdm

from pymail_zh.mail_utils import CreateFolderFailed, eml_to_mail_info, DecodeException

mail_ssl_context = ssl.create_default_context()
# don't check if certificate hostname doesn't match target hostname
mail_ssl_context.check_hostname = False
# don't check if the certificate is trusted by a certificate authority
mail_ssl_context.verify_mode = ssl.CERT_NONE


class Mail(object):
    """
    单封邮件对像
    """

    def __init__(self, uid, subject, mail_date, body, attachments, sender, eml):
        self.eml = eml
        self.subject = subject
        self.mail_date = mail_date
        self.body = body
        self.attachments = attachments
        self.uid = uid
        self.sender = sender


class MailClient(IMAPClient):
    """
    extend IMAPClient
    """

    def __init__(self, user_name, password, host, port=None, use_uid=True, use_ssl=True, stream=False,
                 ssl_context=mail_ssl_context, timeout=None):
        super(MailClient, self).__init__(host, port=port, use_uid=use_uid, ssl=use_ssl, stream=stream,
                                         ssl_context=ssl_context, timeout=timeout)
        self._user_name = user_name
        self._password = password
        self.req_session = requests.Session()

    def __enter__(self):
        self.login(self._user_name, self._password)
        return self

    def mark_unread(self, uid):
        """
        标记邮件为未读
        :param uid: mail uid
        :return:
        """
        self.remove_flags(uid, [SEEN])

    def mark_read(self, uid):
        """
        标记邮件为已读
        :param uid:
        :return:
        """
        self.add_flags(uid, [SEEN])

    def create_mail_folder(self, folder_name, max_try_iteration=10):
        """
        创建目录，支持 nested 目录。
        :param folder_name: 创建
        :param max_try_iteration: 创建 nested folder 时最大的尝试次数
        :return: create completed or create
        """
        paths_element = folder_name.split('/')
        possible_path = ['/'.join(paths_element[:i]) for i in reversed(range(1, len(paths_element) + 1))]
        final_pointer = 0
        total_iter_counts = 0
        while True:
            if total_iter_counts > max_try_iteration:
                raise Exception('Too many nested folder')
            try:
                msg = self.create_folder(possible_path[final_pointer])
                if msg.decode('utf-8') == 'CREATE completed':
                    if final_pointer == 0:
                        return 'create completed'
                    final_pointer = final_pointer - 1
            except (IMAPClientError, IndexError) as e:
                if isinstance(e, IndexError):
                    # possible_path[final_pointer] 有问题，创建失败。
                    return 'create failed'
                # 失败后，尝试创建父目录。
                final_pointer = final_pointer + 1
            total_iter_counts = total_iter_counts + 1

    def move_mail(self, uid, dest_path):
        """
        :param uid:
        :param dest_path: 准备移动到的 '/parent_folder/sub_folder/sub_folder'
        :return:
        """
        dest_path = dest_path.strip('/')
        try:
            self.copy(uid, dest_path)
        except (IMAPClientError,):
            # 可能是没有目标文件夹，或者此邮件就在这个目标文件夹里。
            if self.folder_exists(dest_path):
                return
            else:
                ret = self.create_mail_folder(dest_path)
                if ret == 'create completed':
                    self.copy(uid, dest_path)
                else:
                    raise CreateFolderFailed(f"创建邮箱文件夹 {dest_path} 失败")

    def handle_mails(self, func, folder_name='Inbox', criteria='UNSEEN', mails_count=50,
                     attach_filter=('xls', 'xlsx', 'pdf', 'doc', 'docx', 'rar', 'zip'), html_parser='html.parser',
                     **kwargs):
        """
        处理未读邮件，每次拿到邮件数据需要立即处理（下次拿到的 UID 可能会不一样，会影响需要 UID 的功能）
        :param func: 处理邮件的方法，input 为 MailAliyun 对象，
        :param folder_name: 邮件所在文件夹
        :param criteria: 搜索条件
        :param mails_count: 设置一次处理多少邮件
        :param attach_filter: 附件过滤器，只保留哪些附件
        :param html_parser: 提取 body 时，bs4 需要使用 html_parser，如果 html 提取有问题，请使用其他 parser
                            (html.parser, lxml, html5lib) see https://stackoverflow.com/a/60254943
        :param kwargs: 其他 func 需要用到的参数
        :return:
        """
        self.select_folder(folder_name)
        messages = self.search(criteria)
        if messages:
            if mails_count != 'ALL':
                messages = messages[:mails_count]
            for uid, message_data in tqdm(self.fetch(messages, 'RFC822').items()):
                eml = message_data[b'RFC822']
                try:
                    mail_info = eml_to_mail_info(eml, attach_filter=attach_filter)
                except DecodeException as e:
                    # 查看邮箱文件夹，例如 Decode_failed/subject 中包含邮件主题报错的邮件。
                    self.move_mail(uid, str(e))
                    break
                mail_info['uid'] = uid
                mail_info['eml'] = eml
                mail = Mail(**mail_info)
                func(self, mail, **kwargs)
        self.close_folder()


class MailQQ(MailClient):
    """
    QQ邮箱
    """

    def __init__(self, user_name, password, host='imap.qq.com', port=None, use_uid=True, use_ssl=True,
                 stream=False, ssl_context=mail_ssl_context, timeout=None):
        super(MailQQ, self).__init__(user_name, password, host=host, port=port, use_uid=use_uid,
                                     use_ssl=use_ssl, stream=stream, ssl_context=ssl_context, timeout=timeout)


class MailQiYeQQ(MailClient):
    """
    腾讯企业邮箱
    """

    def __init__(self, user_name, password, host='imap.exmail.qq.com', port=None, use_uid=True, use_ssl=True,
                 stream=False, ssl_context=mail_ssl_context, timeout=None):
        super(MailQiYeQQ, self).__init__(user_name, password, host=host, port=port, use_uid=use_uid,
                                         use_ssl=use_ssl, stream=stream, ssl_context=ssl_context, timeout=timeout)

    def move_mail(self, uid, dest_path):
        """
        :param uid:
        :param dest_path: 准备移动到的 '/parent_folder/sub_folder/sub_folder'
        :return:
        """
        dest_path = dest_path.strip('/')
        try:
            if not self.folder_exists(dest_path):
                if 'create failed' in self.create_mail_folder(dest_path):
                    raise CreateFolderFailed(f"创建邮箱文件夹 {dest_path} 失败")
            self.move(uid, dest_path)
        except (IMAPClientError,) as e:
            logging.error(e)


class MailAliyun(MailClient):
    """
    阿里云邮箱
    """

    def __init__(self, user_name, password, host='imap.aliyun.com', port=None, use_uid=True, use_ssl=True,
                 stream=False, ssl_context=mail_ssl_context, timeout=None):
        super(MailAliyun, self).__init__(user_name, password, host=host, port=port, use_uid=use_uid,
                                         use_ssl=use_ssl, stream=stream, ssl_context=ssl_context, timeout=timeout)


class MailQiYeAliyun(MailClient):
    """
    阿里云企业邮箱
    """

    def __init__(self, user_name, password, host='imap.qiye.aliyun.com', port=None, use_uid=True, use_ssl=True,
                 stream=False, ssl_context=mail_ssl_context, timeout=None):
        super(MailQiYeAliyun, self).__init__(user_name, password, host=host, port=port, use_uid=use_uid,
                                             use_ssl=use_ssl, stream=stream, ssl_context=ssl_context, timeout=timeout)
