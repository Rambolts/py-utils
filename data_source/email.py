from exchangelib import DELEGATE, Account, Credentials, Configuration

class AccountEmail():
    """Classe de abstração da conexão com uma conta de e-mail"""
    
    def __init__(self, user, pwd, host):
        """
        Instancia um objeto de acesso a uma conta de e-mail.

        :param user: e-mail do usuário
        :param pwd: senha do usuário
        :param host: domínio da hospedagem
        """
        self.credentials = Credentials(
            username = user,
            password = pwd
        )
        self.config = Configuration(server=host, credentials=self.credentials)
        self.account = Account(
            primary_smtp_address = ''.join(user),
            config = self.config,
            autodiscover = True,
            access_type = DELEGATE
        )
        
    def access(self):
        """ 
        Retorna o acesso a conta.
        
        :returns: acesso.
        """
        return self.account