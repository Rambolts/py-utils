import warnings
import contextlib
import requests
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from urllib3.exceptions import InsecureRequestWarning

class SharepointDataSource():
    """ Representa a conexão ao Sharepoint """
    
    def __init__(self, url, client_id, client_secret):
        """
        Inicializa variáveis de conexão com o Sharepoint Online.

        :param url: url do sharepoint online ao qual deseja se autenticar
        :param client_id: client id da aplicação do sharepoint
        :param client_secret: client secret da aplicação do sharepoint
        """
        self._ctx = self.connect_to_sharepoint(url = url, 
                                               client_id = client_id, 
                                               client_secret = client_secret)

    @contextlib.contextmanager
    def no_ssl_verification(self):
        self.old_merge_environment_settings = requests.Session.merge_environment_settings
        self.opened_adapters = set()
        def merge_environment_settings(self, url, proxies, stream, verify, cert):
            self.opened_adapters.add(self.get_adapter(url))
            settings = self.old_merge_environment_settings(self, url, proxies, stream, verify, cert)
            settings['verify'] = False
            return settings       
        requests.Session.merge_environment_settings = merge_environment_settings
        try:
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', InsecureRequestWarning)
                yield
        finally:
            requests.Session.merge_environment_settings = self.old_merge_environment_settings
            for adapter in self.opened_adapters:
                try:
                    adapter.close()
                except:
                    pass

    def connect_to_sharepoint(self, url, **kwargs):
        with self.no_ssl_verification():
            if 'client_id' in kwargs:
                # com client_id e client_secret
                credentials = ClientCredential(kwargs['client_id'], kwargs['client_secret'])
                return ClientContext(url).with_credentials(credentials)

            if 'username' in kwargs: 
                # com username e password
                return ClientContext(url).with_user_credentials(kwargs['username'], kwargs['password'])
        return None

    def connect(self):
        return self._ctx