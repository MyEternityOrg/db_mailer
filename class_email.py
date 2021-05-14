import envelope


class Email:
    def __init__(self, sender, recipient, subject, message, reply_to=""):
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
        self.__subject = subject
        self.__message = message
        self.__server = ''
        self.__port = 0
        self.__login = ''
        self.__password = ''
        try:
            self.__email_message = envelope.Envelope(from_=self.__sender, to=self.__recipient, subject=self.__subject,
                                                     message=self.__message, reply_to=self.__reply_to)
        except Exception as E:
            print(f'Ошибка создания класса Mail: {E}')

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
