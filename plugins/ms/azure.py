import msal
from office365.graph_client import GraphClient, UserCollection, GroupCollection, DeltaCollection, SitesWithRoot
from retry import retry
from typing import List, Optional

from utils.requests import bypass_ssl

class AzurePlugin():
    """ Conector a serviços do Azure """

    def __init__(self, tenant: str, client_id: str, client_secret: str):
        """
        Construtor da classe

        :param tenant: tenant do azure que se deseja conectar
        :param client_id: client id da aplicação do azure
        :param client_secret: client secret da aplicação do azure
        """
        self._tenant = tenant
        self._client_id = client_id
        self._client_secret = client_secret
        self._internal_init()

    def _acquire_token_by_client_credentials(self):
        authority_url = 'https://login.microsoftonline.com/{0}'.format(self._tenant)
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id=self._client_id,
            client_credential=self._client_secret
        )
        return app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])

    @retry(ConnectionError, tries=3, delay=1)
    def _internal_init(self):
        """
        Inicializa variáveis de conexão com o Sharepoint Online e a lista-alvo.
        """
        with bypass_ssl():
            # se autentica
            self._client = GraphClient(self._acquire_token_by_client_credentials)

    @property
    def client(self) -> GraphClient:
        """
        Propriedade para o contexto de conexão com o Sharepoint Online

        :returns: context de conexão
        """
        return self._client
    
    @retry(ConnectionError, tries=3, delay=1)
    def get_users(self, fields: Optional[List[str]] = None, page_size: Optional[int] = 100) -> UserCollection:
        if fields:
            items = self._client.users.select(fields).paged(page_size).get().execute_query()
        else:
            items = self._client.users.paged(page_size).get().execute_query()
        
        return items
    
    @retry(ConnectionError, tries=3, delay=1)
    def get_groups(self, fields: Optional[List[str]] = None, page_size: Optional[int] = 100) -> GroupCollection:
        if fields:
            items = self._client.groups.select(fields).paged(page_size).get().execute_query()
        else:
            items = self._client.groups.paged(page_size).get().execute_query()
        
        return items
    
    @retry(ConnectionError, tries=3, delay=1)
    def get_applications(self, fields: Optional[List[str]] = None, page_size: Optional[int] = 100) -> DeltaCollection:
        if fields:
            items = self._client.applications.select(fields).paged(page_size).get().execute_query()
        else:
            items = self._client.applications.paged(page_size).get().execute_query()
        
        return items
    
    @retry(ConnectionError, tries=3, delay=1)
    def get_sites(self, fields: Optional[List[str]] = None, page_size: Optional[int] = 100) -> SitesWithRoot:
        if fields:
            items = self._client.sites.select(fields).paged(page_size).get().execute_query()
        else:
            items = self._client.sites.paged(page_size).get().execute_query()
        
        return items

