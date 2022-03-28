"""
    Classe utilizada para simplificar conexão com a API Client Docusign
"""

__author__ = 'Gustavo Magalhães'
__version__ = '1.0'

from docusign_esign.client.api_client import ApiClient

class DocusignDataSource():
    """ Representa a conexão ao DocuSign """

    def __init__(self, integration_key, user_id, authorization_server, base_path):
        """
        Inicializa variáveis de conexão com o Docusign.

        :param integration_key: chave de integração, valor GUID que identifica sua integração. 
        :param user_id: valor GUID que identifica exclusivamente um usuário DocuSign.
        :param authorization_server: server de autorização.
        :param base_path: link de requisição. www.docusign.net for production and demo.docusign.net for the developer sandbox.
        """
        with open('privatekey.txt', 'r') as pk:
            private_key = pk.read()
            
        self._api_client = ApiClient(host=base_path, oauth_host_name=authorization_server, base_path=base_path)

        # Getting Token
        self._token = self._api_client.request_jwt_user_token(
            client_id=integration_key,
            user_id=user_id, 
            oauth_host_name=authorization_server, 
            private_key_bytes=private_key, 
            expires_in=3600, 
            scopes=['signature', 'impersonation'])
        access_token = self._token.access_token

        # Getting Account Id
        self._user_info = self._api_client.get_user_info(access_token)
        self._account_id = self._user_info.accounts[0].account_id
        self._api_client.set_default_header(header_name='Authorization', header_value=f'Bearer {access_token}')

    def get_connection(self):
        return self._api_client, self._account_id