import envelope


class Email:
    def __init__(self, sender, recipient, subject, message, reply_to="", copy_recipients=None,
                 blind_copy_recipients=None):
        if len(sender) == 0:
            raise Exception("Заполните отправителя!")
        if len(recipient) == 0:
            raise Exception("Заполните получателей!")
        if len(subject) == 0:
            raise Exception("Заполните тему письма!")
        if len(message) == 0:
            raise Exception("Заполните текст сообщения!")
        if len(reply_to) == 0:
            self.__reply_to = self.__sender
        else:
            self.__reply_to = reply_to
        self.__sender = sender
        self.__recipient = recipient
        if len(copy_recipients) > 0:
            self.__copy_recipients = copy_recipients
        else:
            self.__copy_recipients = None
        if len(blind_copy_recipients) > 0:
            self.__blind_copy_recipients = blind_copy_recipients
        else:
            self.__blind_copy_recipients = None
        self.__subject = subject
        self.__message = message
        self.__server = ''
        self.__port = 0
        self.__login = ''
        self.__password = ''
        try:
            args = {}
            self.combine_args(args, 'from_', self.__sender)
            self.combine_args(args, 'to', self.__recipient)
            self.combine_args(args, 'subject', self.__subject)
            self.combine_args(args, 'message', self.__message)
            self.combine_args(args, 'reply_to', self.__reply_to)
            self.combine_args(args, 'cc', self.__copy_recipients)
            self.combine_args(args, 'bcc', self.__blind_copy_recipients)
            self.__email_message = envelope.Envelope(**args)
        except Exception as E:
            print(f'Ошибка создания класса Mail: {E}')

    @staticmethod
    def combine_args(in_dict, key_name="default", key_value=None):
        if key_value is not None:
            if type(key_value) == dict:
                if len(key_value) > 0:
                    in_dict[key_name] = key_value
            else:
                in_dict[key_name] = key_value
        return in_dict

    @property
    def port(self):
        return self.__port

    @port.setter
    def port(self, value):
        self.__port = value

    @property
    def server(self):
        return self.__server

    @server.setter
    def server(self, value):
        self.__server = value

    @property
    def login(self):
        return self.__login

    @login.setter
    def login(self, value):
        self.__login = value

    @property
    def password(self):
        return self.__password

    @password.setter
    def password(self, value):
        self.__password = value

    def attach_file(self, filepath, inline=False):
        try:
            with open(filepath, 'rb') as f:
                file_name = filepath.split('\\')[-1].strip()
                self.__email_message.attach(name=file_name, attachment=f.read(), inline=inline)
                return True
        except Exception as E:
            print(f'Ошибка добавления вложения из файла {filepath}: {E}')
            return False

    @property
    def message(self):
        return self.__message

    @message.setter
    def message(self, value):
        self.__message = value

    def connect_smtp(self):
        try:
            self.__email_message.smtp(host=self.__server, port=self.__port, user=self.__login, password=self.__password)
            return True
        except Exception as E:
            print(f'Ошибка подключения к серверу: {E}')
            return False

    def send_email(self, send_mail=True):
        try:
            self.__email_message.send(send=send_mail)
            return True
        except Exception as E:
            print(f'Ошибка отправки письма: {E}')
            return False
