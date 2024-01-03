import pandas as pd
import ldap3
from ldap3.extend.microsoft.removeMembersFromGroups import ad_remove_members_from_groups
from ldap3.core.exceptions import LDAPException
from typing import List, Dict, Tuple, Any, Union, Optional

class ActiveDirectoryPlugin():
    """ Plugin para interagir com o LDAP da AG """

    @classmethod
    def authenticate(cls, user: str, password: str, server: Optional[str] = 'ldap://AGNET.local') -> bool:
        """
        Verifica se um usuário e senha estão corretos.

        :param user: usuário no formato AGDOMAIN\login
        :param password: senha
        :param server: servidor LDAP
        :returns: True quando a autenticação obteve êxito
        """
        try:
            with ldap3.Connection(server, user=user, password=password) as connection:
                return connection.result['description']=='success'
        except LDAPException:
            return False

    def __init__(self, 
                 user: str, 
                 password: str, 
                 server: Optional[str] = 'ldap://AGNET.local', 
                 base: Optional[str] = 'dc=AGNET,dc=local'):
        """
        Instancia objeto de conexão ao AD

        :param user: usuário
        :param password: senha
        :param server: servidor LDAP
        :param base: base de procura
        """
        self._server = server
        self._base = base
        self._user = user
        self._password = password
        self._connection = None

    def connect(self) -> ldap3.Connection:
        """
        Conecta ao AD

        :returns: objeto de conexão
        """
        # verifica se a conexão já foi estabelecida
        if self._connection is None or self._connection.closed:
            server = ldap3.Server(self._server)
            self._connection = ldap3.Connection(server, user=self._user, password=self._password, auto_bind=True)
        return self._connection

    def search(self, search_filter: str, attributes: Optional[List[str]] = ldap3.ALL_ATTRIBUTES) -> List[ldap3.Entry]:
        """
        Efetua busca de dados no AD.

        :param search_filter: filtro de pesquisa
        :param attributes: atributos que se deseja buscar
        :returns: lista de objetos do tipo ldap3.Entry
        """
        c = self.connect()
        c.search(search_base=self._base, search_filter=search_filter, search_scope=ldap3.SUBTREE, attributes=attributes)
        return c.entries

    def search_as_dataframe(self, search_filter: str, attributes: Optional[List[str]] = ldap3.ALL_ATTRIBUTES, export_dn: bool = False) -> pd.DataFrame:
        """
        Efetua busca de dados no AD e o retorno como DataFrame.

        :param search_filter: filtro de pesquisa
        :param attributes: atributos que se deseja buscar
        :param export_dn: quando True, o campo DN é retornado no dataframe na primeira coluna
        :returns: dataframe com os dados retornados
        """
        entries = self.search(search_filter, attributes)

        data = []
        for e in entries:
            record = {} if not export_dn else {'dn': e.entry_dn}
            record.update({att: e[att].value for att in e.entry_attributes})
            data.append(record)

        df = pd.DataFrame(data)
        if not isinstance(attributes, str):
            if export_dn:
                attributes = ['dn'] + attributes
            df = df[attributes]
        return df
    
    def modify(self, dn: Union[ldap3.Entry, str], changes: Dict[str, Tuple[str, Any]]) -> bool:
        """
        Modifica um registro no AD

        :param dn: pode ser um registro do AD ou uma string com o DN
        :param changes: dicionário contendo o atributo e a tupla (tipo de modificação, valor)
        :returns: True quando der certo
        """
        dn = dn.entry_dn if isinstance(dn, ldap3.Entry) else dn
        c = self.connect()
        return c.modify(dn, changes)
    
    def remove_members_from_groups(self, members_dn: List[Union[ldap3.Entry, str]], groups_dn: List[Union[ldap3.Entry, str]]) -> bool:
        """
        Remove usuários de uma lista de grupos

        :param members_dn: lista de registros do AD ou de strings com os DNs dos usuários
        :param groups_dn: lista de registros do AD ou de strings com os DNs dos grupos
        :returns: True quando der certo
        """
        if isinstance(members_dn, (str, ldap3.Entry)):
            members_dn = [members_dn]

        if isinstance(groups_dn, (str, ldap3.Entry)):
            groups_dn = [groups_dn]

        members_dn = [dn.entry_dn if isinstance(dn, ldap3.Entry) else dn for dn in members_dn]
        groups_dn = [dn.entry_dn if isinstance(dn, ldap3.Entry) else dn for dn in groups_dn]

        c = self.connect()
        return ad_remove_members_from_groups(c, members_dn, groups_dn, fix=False, raise_error=False)

        