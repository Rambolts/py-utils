import os
import re
import pandas as pd
import zipfile
from tqdm.auto import tqdm
from retry import retry
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.base_entity_collection import BaseEntityCollection
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import List as SharepointList
from office365.sharepoint.listitems.collection import ListItemCollection as SharepointListItemCollection
from office365.sharepoint.listitems.listitem import ListItem as SharepointListItem
from office365.sharepoint.folders.folder import Folder as SharepointFolder
from office365.sharepoint.folders.collection import FolderCollection as SharepointFolderCollection
from office365.sharepoint.files.file import File as SharepointFile
from office365.sharepoint.files.collection import FileCollection as SharepointFileCollection
from typing import List, Tuple, Dict, Any, Union

from utils.requests import bypass_ssl

def compare_sp_values(value1: Any, value2: Any) -> bool:
    """
    Compara dois valores de campos do Sharepoint para mostrar equivalência.

    :param value1: valor de entrada 1
    :param value2: valor de entrada 2
    :returns: True quando são equivalentes
    """
    # verifica se os campos value1 e value2 são strings e remove parte do horário da data
    if isinstance(value1, str):
        value1 = re.sub(r'T[\d\:]+Z', '', value1)
    if isinstance(value2, str):
        value2 = re.sub(r'T[\d\:]+Z', '', value2)

    # agrupa em um conjunto
    s = set([value1, value2])
    # se o conjunto possuir um único elemento ou for formado de None e string vazia, indica equivalência
    return (len(s) == 1) or (s == set([None, '']))

def list_items_to_dataframe(items: BaseEntityCollection) -> pd.DataFrame:
    """
    Exporta itens extraídos do Sharepoint em DataFrame

    :param items: coleção de items extraídos do Sharepoint
    :returns: DataFrame
    """
    data = [item.properties for item in items]
    return pd.DataFrame(data)

class SharepointListPlugin():
    """ Conector a listas do Sharepoint Online """

    def __init__(self, url: str, client_id: str, client_secret: str, list_title: str):
        """
        Construtor da classe

        :param url: url do sharepoint online onde se encontra a lista
        :param client_id: client id da aplicação do sharepoint
        :param client_secret: client secret da aplicação do sharepoint
        :param list_title: nome da lista
        """
        self._url = url
        self._client_id = client_id
        self._client_secret = client_secret
        self._internal_init(url, client_id, client_secret, list_title)

    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def _internal_init(self, url: str, client_id: str, client_secret: str, list_title: str):
        """
        Inicializa variáveis de conexão com o Sharepoint Online e a lista-alvo.

        :param url: url do sharepoint online onde se encontra a lista
        :param client_id: client id da aplicação do sharepoint
        :param client_secret: client secret da aplicação do sharepoint
        :param list_title: nome da lista
        """
        with bypass_ssl():
            # se autentica
            self._ctx_auth = ClientCredential(client_id, client_secret)
            # pega contexto
            self._ctx = ClientContext(url).with_credentials(self._ctx_auth)

            # carrega lista
            self._list_object = self._ctx.web.lists.get_by_title(list_title).get().execute_query()
            # carrega campos
            self._fields = self._list_object.fields.get().execute_query()

    @property
    def context(self) -> ClientContext:
        """
        Propriedade para o contexto de conexão com o Sharepoint Online

        :returns: context de conexão
        """
        return self._ctx

    @property
    def list_object(self) -> SharepointList:
        """
        Propriedade para recuperar o objeto da lista

        :returns: objeto da lista
        """
        return self._list_object
    
    @property
    def field_names(self) -> List[str]:
        return [f.internal_name for f in self._fields]
    
    @property
    def field_names_decoded(self) -> List[str]:
        return [f.title for f in self._fields]

    def validate_field_names(self, field_names: List[str]) -> bool:
        """
        Valida se todos os campos em uma lista estão presentes na lista

        :param field_names: lista com os nomes dos campos para validar (podem estar codificados ou não)
        :returns: True quando todos os campos estão presentes
        """
        field_names_encoded = [self.field_name_encoder(field) for field in field_names]
        all_fields = [field.internal_name for field in self._fields]
        return set(field_names_encoded) <= set(all_fields)

    def field_name_encoder(self, field_name: str) -> str:
        """
        Codifica o nome do campo

        :param field_name: apelido do campo
        :returns: nome interno do campo
        """
        for field in self._fields:
            if field.title.lower() == field_name.lower():
                return field.internal_name
        return field_name

    def field_name_decoder(self, field_name: str) -> str:
        """
        Decodifica o nome do campo

        :param field_name: nome interno do campo
        :returns: apelido do campo
        """
        for field in self._fields:
            if field.internal_name.lower() == field_name.lower():
                return field.title
        return field_name

    def get_field_type(self, field_name: str) -> str:
        """
        Retorna o tipo de dado do campo

        :param field_name: nome interno do campo
        :returns: tipo do campo como string
        """
        for field in self._fields:
            if field.internal_name.lower() == self.field_name_encoder(field_name).lower():
                return field.type_as_string
        return None

    def get_field_properties(self, field_name: str) -> Dict[str, Any]:
        """
        Retorna as propriedades do campo.

        :param field_name: nome interno do campo
        :returns: dicionário com as propriedades do campo
        """
        for field in self._fields:
            if field.internal_name.lower() == self.field_name_encoder(field_name).lower():
                return field.properties
        return {}

    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_items(self, fields: List[str] = None) -> SharepointListItemCollection:
        """
        Retorna itens da lista.

        :param fields: lista de campos para serem recuperados (podem estar codificados ou não)
        :returns: coleção de itens
        """
        fields_to_select = None

        if fields is not None:
            # compõe seleção de campos (inclui ID caso não seja informado)
            fields_to_select = [self.field_name_encoder(f) for f in set(['ID'] + fields)]

        # carrega itens
        if fields_to_select:
            items = self._list_object.items.select(fields_to_select).paged(100).get().execute_query()
        else:
            items = self._list_object.items.paged(100).get().execute_query()

        return items

    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_item(self, id_or_title: Union[int, str], id_field: str = None, fields: List[str] = None) -> SharepointListItem:
        """
        Retorna um item da lista de acordo com seu ID ou dado

        :param id_or_title: valor da chave primária
        :param id_field: campo que servirá como chave primária na busca
        :param fields: lista de campos para serem recuperados (podem estar codificados ou não)
        :returns: item relacionado à primeira ocorrência do filtro
        """
        fields_to_select = None

        if fields is not None:
            # compõe seleção de campos (inclui ID caso não seja informado)
            fields_to_select = [self.field_name_encoder(f) for f in set(['ID'] + fields)]

        # faz algumas validações quanto ao campo que servirá como chave primária
        if id_field is None:
            id_field = 'ID' if isinstance(id_or_title, int) else 'Title'
        else:
            id_field = self.field_name_encoder(id_field)

        # monta expressão do filtro
        if isinstance(id_or_title, int):
            filter_expression = f"{id_field} eq {id_or_title}"
        else:
            filter_expression = f"{id_field} eq '{id_or_title}'"

        # monta query para o item
        if fields_to_select:
            items = self._list_object.items.filter(filter_expression).select(fields_to_select).paged(100).get().execute_query()
        else:
            items = self._list_object.items.filter(filter_expression).paged(100).get().execute_query()

        # retorna o primeiro
        return items[0] if len(items) else None

    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def update_item(self, id_or_title: Union[int, str], data: Dict[str, Any]) -> SharepointListItem:
        """
        Atualiza item na lista

        :param id_or_title: item a ser atualizado (pode ser um id ou title)
        :param data: dicionário com os dados a serem atualizados (podem conter campos internos ou não)
        :returns: item modificado
        """
        # caso não seja um item, verifica se é um id ou string
        assert isinstance(id_or_title, int) or isinstance(id_or_title, str), 'Tipo incorreto de dado para "id_or_title"'
        item = self.get_item(id_or_title)

        # pequeno hack na lib
        item._resource_url = item.resource_url.replace('/items/', '/')

        # para cada item do dicionário
        for k, v in data.items():
            # codifica o campo
            field = self.field_name_encoder(k)
            field_type = self.get_field_type(field)

            type_mapper = {
                'Boolean': bool,
                'Number': float
            }

            # converte valor de acordo com seu tipo
            converter = type_mapper.get(field_type, None)
            if converter is not None:
                v = converter(v)

            # atualiza propriedade
            item.set_property(field, v)

        # atualiza item
        return item.update().execute_query()

    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def add_item(self, data: Dict[str, Any]) -> SharepointListItem:
        """
        Adiciona/ insere um novo item na lista

        :param data: dicionário contendo os dados para serem inseridos
        :returns: item adicionado
        """
        data_fixed = {}
        # para cada chave do dicionário
        for k, v in data.items():
            # codifica nome do campo
            field = self.field_name_encoder(k)
            field_type = self.get_field_type(field)

            type_mapper = {
                'Boolean': bool,
                'Number': float
            }

            # converte valor para o tipo do campo
            converter = type_mapper.get(field_type, None)
            if converter is not None:
                v = converter(v)

            data_fixed[field] = v

        # adiciona item
        return self._list_object.add_item(data_fixed).execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def delete_item(self, id_or_title: Union[int, str]) -> SharepointListItem:
        """
        Exclui item da lista do Sharepoint

        :param id_or_title: id ou título do item
        :returns: item excluído
        """
        # caso não seja um item, verifica se é um id ou string
        assert isinstance(id_or_title, int) or isinstance(id_or_title, str), 'Tipo incorreto de dado para "id_or_title"'

        item = self.get_item(id_or_title)
        return item.delete_object().execute_query()
    
class DocumentLibraryPlugin(SharepointListPlugin):
    """ Conector a bibliotecas de documentos do Sharepoint Online """

    def __init__(self, url: str, client_id: str, client_secret: str, list_title: str):
        """
        Construtor da classe

        :param url: url do sharepoint online onde se encontra a lista
        :param client_id: client id da aplicação do sharepoint
        :param client_secret: client secret da aplicação do sharepoint
        :param list_title: nome da lista
        """
        super().__init__(url, client_id, client_secret, list_title)
        # carrega diretório raiz
        self._root_folder = self._list_object.root_folder.get().execute_query()
    
    @property
    def root_folder(self) -> SharepointFolder:
        """
        Propriedade para recuperar o diretório raiz

        :returns: objeto da pasta
        """
        return self._root_folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_folder_by_path(self, 
                           path: str, 
                           parent_folder: SharepointFolder = None) -> SharepointFolder:
        """
        Recupera pasta através do caminho

        :param path: caminho da pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :returns: objeto da pasta do sharepoint
        """
        folder_names = path.split('/')

        folder = parent_folder or self._root_folder
        for folder_name in folder_names:
            if folder_name != '':
                folder = folder.folders.get_by_path(folder_name).get().execute_query()
    
        return folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_folder(self, 
                   folder_or_path: Union[str, SharepointFolder], 
                   parent_folder: SharepointFolder = None) -> SharepointFolder:
        """
        Recupera pasta

        :param folder_or_path: caminho da pasta ou objeto pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :returns: objeto da pasta do sharepoint
        """
        if isinstance(folder_or_path, str):
            folder = self.get_folder_by_path(folder_or_path, parent_folder)
        else:
            folder = folder_or_path or parent_folder or self._root_folder

        return folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_folders(self, 
                    folder_or_path: Union[str, SharepointFolder] = None, 
                    parent_folder: SharepointFolder = None) -> SharepointFolderCollection:
        """
        Recupera pastas

        :param folder_or_path: caminho da pasta ou objeto pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :returns: coleção de pastas do sharepoint
        """
        folder = self.get_folder(folder_or_path, parent_folder)
        return folder.folders.get().paged(100).execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def download_folder_as_zip(self, 
                               output_filename: str,
                               folder_or_path: Union[str, SharepointFolder] = None, 
                               parent_folder: SharepointFolder = None,
                               enable_progress: bool = False):
        """
        Efetua download da pasta do sharepoint como arquivo zip

        :param output_filename: caminho completo do arquivo zip de saída
        :param folder_or_path: caminho da pasta ou objeto pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :param enable_progress: quando True, apresenta barra de progresso
        """
        folder = self.get_folder(folder_or_path, parent_folder)
        items_to_download = self.list_item_names(folder_or_path, parent_folder, recursive=True)

        dirname = os.path.dirname(output_filename)
        if dirname != '' and not os.path.exists(dirname):
            os.makedirs(dirname, exist_ok=True)

        with tqdm(total=len(items_to_download), disable=not enable_progress) as pbar:
            with zipfile.ZipFile(output_filename, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for type_, item_path in items_to_download:
                    if type_=='file':
                        info = zipfile.ZipInfo(item_path)
                        info.compress_type = zipfile.ZIP_DEFLATED
                        file = self.get_file(item_path, parent_folder=folder)
                        zf.writestr(info, file.read())
                    else:
                        info = zipfile.ZipInfo(item_path + '/')
                        info.compress_type = zipfile.ZIP_DEFLATED
                        zf.writestr(info, '')
                    pbar.update()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def create_folder(self, 
                      folder_name: str, 
                      parent_folder: SharepointFolder = None, 
                      overwrite: bool = True) -> SharepointFolder:
        """
        Cria pasta

        :param folder_name: nome da pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :param overwrite: quando True, sobrescreve caso exista
        :returns: objeto da pasta criada
        """
        if parent_folder is None:
            parent_folder = self._root_folder
        
        new_folder = parent_folder.folders.add_using_path(folder_name, overwrite).execute_query()
        return new_folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def create_folders(self, 
                       folder_name: str, 
                       parent_folder: SharepointFolder = None, 
                       overwrite: bool = True) -> SharepointFolder:
        """
        Cria pastas em cascata

        :param folder_name: nome da pasta
        :param parent_folder: objeto de pasta ao qual o path inicia, por padrão root_folder
        :param overwrite: quando True, sobrescreve caso exista
        :returns: objeto da última pasta criada
        """
        if parent_folder is None:
            parent_folder = self._root_folder

        current_folder_name, *next_folders_name = folder_name.split('/')
        if current_folder_name != '':
            folder = self.create_folder(current_folder_name, parent_folder, overwrite=overwrite)
        else:
            folder = parent_folder

        if len(next_folders_name) > 0:
            new_folder = self.create_folders('/'.join(next_folders_name), folder, overwrite=overwrite)
        else:
            new_folder = folder

        return new_folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def rename_folder(self, 
                      folder: SharepointFolder, 
                      new_folder_name: str) -> SharepointFolder:
        """
        Renomeia pasta

        :param folder: objeto da pasta que se deseja renomear
        :param new_folder_name: novo nome da pasta
        :returns: objeto da pasta atualizado
        """
        parent_folder = folder.parent_folder
        folder.rename(new_folder_name).execute_query()
        return parent_folder.folders.get_by_path(new_folder_name).get().execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def delete_folder(self, 
                      folder: SharepointFolder) -> SharepointFolder:
        """
        Exclui pasta

        :param folder: objeto da pasta que se deseja renomear
        :returns: objeto da pasta excluído
        """
        folder.delete_object().execute_query()
        folder.properties['Exists'] = False
        return folder
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_file_by_path(self, 
                         path: str, 
                         parent_folder: SharepointFolder = None) -> SharepointFile:
        """
        Recupera arquivo através do seu path

        :param path: caminho relativo do arquivo
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: objeto do arquivo
        """
        *folder_names, filename = path.split('/')
        folder_path = '/'.join(folder_names)
        folder = self.get_folder_by_path(folder_path, parent_folder)
        return folder.files.get_by_url(filename).get().execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_file(self, 
                 path: str, 
                 parent_folder: SharepointFolder = None) -> SharepointFile:
        """
        Recupera arquivo (semelhante ao método get_file_by_path)

        :param path: caminho relativo do arquivo
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: objeto do arquivo
        """
        return self.get_file_by_path(path, parent_folder)
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_files(self, 
                  folder_or_path: Union[str, SharepointFolder] = None,
                  parent_folder: SharepointFolder = None) -> SharepointFileCollection:
        """
        Recupera arquivos

        :param folder_or_path: path ou objeto da pasta onde se deseja listar os arquivos
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: coleção de arquivos
        """
        folder = self.get_folder(folder_or_path, parent_folder)
        return folder.files.get().paged(100).execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def get_file_metadata(self, 
                          file_or_path: Union[str, SharepointFile], 
                          parent_folder: SharepointFolder = None) -> SharepointListItem:
        """
        Recupera metadados do arquivo

        :param file_or_path: path ou objeto do arquivo 
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: objeto de lista com os metadados do arquivo
        """
        if isinstance(file_or_path, str):
            file = self.get_file(file_or_path, parent_folder)
        else:
            file = file_or_path

        return file.listItemAllFields.expand(['File']).get().execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def update_file_metadata(self, 
                             file_or_path: Union[str, SharepointFile], 
                             parent_folder: SharepointFolder = None,
                             data: Dict[str, Any] = None) -> SharepointListItem:
        """
        Atualiza metadados do arquivo

        :param file_or_path: path ou objeto do arquivo
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param data: dicionário com os dados a serem atualizados (podem conter campos internos ou não)
        :returns: item modificado
        """
        if isinstance(file_or_path, str):
            file = self.get_file(file_or_path, parent_folder)
        else:
            file = file_or_path

        item = file.listItemAllFields
        data = data or {}

        # para cada item do dicionário
        for k, v in data.items():
            # codifica o campo
            field = self.field_name_encoder(k)
            field_type = self.get_field_type(field)

            type_mapper = {
                'Boolean': bool,
                'Number': float
            }

            # converte valor de acordo com seu tipo
            converter = type_mapper.get(field_type, None)
            if converter is not None:
                v = converter(v)

            # atualiza propriedade
            item.set_property(field, v)

        # atualiza item
        return item.update().execute_query()
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def download_file(self,
                      file_or_path: Union[str, SharepointFile],
                      parent_folder: SharepointFolder = None,
                      output_dir: str = None,
                      output_filename: str = None,
                      overwrite: bool = True) -> str:
        """
        Efetua download do arquivo

        :param file_or_path: path ou objeto do arquivo
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param output_dir: diretório onde se deseja salvar o arquivo
        :param output_filename: nome do arquivo de saída
        :param overwrite: sobrescreve arquivo quando True
        :returns: caminho completo do arquivo baixado
        """
        if isinstance(file_or_path, str):
            file = self.get_file(file_or_path, parent_folder)
        else:
            file = file_or_path

        filename = output_filename or ''
        if filename == '':
            filename = file.name

        dirname = output_dir or ''
        if dirname != '' and not os.path.exists(dirname):
            os.makedirs(dirname, exist_ok=True)

        download_filename = os.path.join(dirname, filename)
        if overwrite or not os.path.exists(download_filename):
            with open(download_filename, 'wb') as f:
                file.download(f).execute_query()

        return download_filename
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def upload_file(self, 
                    filename: str, 
                    folder_or_path: Union[str, SharepointFolder], 
                    parent_folder: SharepointFolder = None) -> SharepointFile:
        """
        Efetua upload de arquivo

        :param filename: caminho completo do arquivo
        :param folder_or_path: path ou objeto da pasta que se deseja fazer o upload
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: objeto do arquivo que teve upload
        """
        folder = self.get_folder(folder_or_path, parent_folder)

        with open(filename, 'rb') as f:
            content = f.read()

        file_uploaded = folder.upload_file(os.path.basename(filename), content).execute_query()
        return file_uploaded
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def upload_files(self, 
                     filenames: List[str], 
                     folder_or_path: Union[str, SharepointFolder], 
                     parent_folder: SharepointFolder = None) -> List[SharepointFile]:
        """
        Efetua upload de vários arquivos

        :param filenames: lista com o caminho completo dos arquivos
        :param folder_or_path: path ou objeto da pasta que se deseja fazer o upload
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :returns: lista com os objetos dos arquivos que tiveram upload
        """
        folder = self.get_folder(folder_or_path, parent_folder)

        files_uploaded = []
        for filename in filenames:
            file_uploaded = self.upload_file(filename, folder)
            files_uploaded.append(file_uploaded)

        return files_uploaded
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def upload_large_file(self, 
                          filename: str, 
                          folder_or_path: Union[str, SharepointFolder], 
                          parent_folder: SharepointFolder = None,
                          enable_progress: bool = False) -> SharepointFile:
        """
        Efetua upload de arquivo com controle de progresso

        :param filename: caminho completo do arquivo
        :param folder_or_path: path ou objeto da pasta que se deseja fazer o upload
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param enable_progress: quando True, habilita barra de progresso
        :returns: objeto do arquivo que teve upload
        """
        folder = self.get_folder(folder_or_path, parent_folder)

        chunck_size = 1024 ** 2
        file_size = os.path.getsize(filename)

        with tqdm(total=file_size, unit='iB', unit_scale=True, desc=os.path.basename(filename), disable=not enable_progress) as pbar:
            def update_progressbar(offset: float):
                if enable_progress:
                    increment = offset - pbar.n
                    pbar.update(increment)

            file_uploaded = folder.files.create_upload_session(filename, chunck_size, update_progressbar).execute_query()
            if enable_progress and pbar.n==0:
                pbar.update(file_size)

        return file_uploaded
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def upload_large_files(self, 
                           filenames: List[str], 
                           folder_or_path: Union[str, SharepointFolder], 
                           parent_folder: SharepointFolder = None,
                           enable_progress: bool = False) -> List[SharepointFile]:
        """
        Efetua upload de vários arquivos com controle de progresso

        :param filenames: lista com o caminho completo dos arquivos
        :param folder_or_path: path ou objeto da pasta que se deseja fazer o upload
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param enable_progress: quando True, habilita barra de progresso
        :returns: lista com os objetos dos arquivos que tiveram upload
        """
        folder = self.get_folder(folder_or_path, parent_folder)

        files_uploaded = []
        for filename in filenames:
            file_uploaded = self.upload_large_file(filename, folder, enable_progress=enable_progress)
            files_uploaded.append(file_uploaded)

        return files_uploaded
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def rename_file(self, 
                    file: SharepointFile, 
                    new_filename: str) -> SharepointFile:
        """
        Renomeia arquivo

        :param file: objeto do arquivo
        :param new_filename: novo nome do arquivo
        :returns: objeto do arquivo renomeado
        """
        file.rename(new_filename).execute_query()
        return file
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def delete_file(self, file: SharepointFile) -> SharepointFile:
        """
        Exclui arquivo

        :param file: objeto do arquivo
        :returns: objeto do arquivo excluido
        """
        file.delete_object().execute_query()
        return file
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def list_folder_names(self, 
                          folder_or_path: Union[str, SharepointFolder] = None, 
                          parent_folder: SharepointFolder = None,
                          recursive: bool = False, 
                          _path: str = '') -> List[str]:        
        """
        Recupera lista com os nomes das pastas

        :param folder_or_path: path ou objeto da pasta onde se deseja listar os arquivos
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param recursive: quando True, efetua a listagem de forma recursiva
        :param _path: parâmetro de controle interno
        :returns: lista com o path das pastas
        """
        folder_names = []

        folders = self.get_folders(folder_or_path, parent_folder)
        for folder in folders:
            if folder.name.lower() != 'forms':
                folder_names.append(_path + folder.name)
                if recursive:
                    folder_names.extend(self.list_folder_names(folder, recursive=recursive, _path=_path + folder.name + '/'))

        return sorted(folder_names)
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def list_file_names(self, 
                        folder_or_path: Union[str, SharepointFolder] = None, 
                        parent_folder: SharepointFolder = None,
                        recursive: bool = False, 
                        _path: str = '') -> List[str]:   
        """
        Recupera lista com os nomes dos arquivos

        :param folder_or_path: path ou objeto da pasta onde se deseja listar os arquivos
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param recursive: quando True, efetua a listagem de forma recursiva
        :param _path: parâmetro de controle interno
        :returns: lista com os paths dos arquivos
        """     
        file_names = []

        folder = self.get_folder(folder_or_path, parent_folder)

        folders = self.get_folders(folder)
        for f in folders:
            if f.name.lower() != 'forms':
                if recursive:
                    file_names.extend(self.list_file_names(f, recursive=recursive, _path=_path + f.name + '/'))

        files = folder.files.get().paged(100).execute_query()
        for file in files:
            file_names.append(_path + file.name)

        return sorted(file_names)
    
    @retry((ConnectionError, ClientRequestException), tries=3, delay=1)
    def list_item_names(self, 
                        folder_or_path: Union[str, SharepointFolder] = None, 
                        parent_folder: SharepointFolder = None,
                        recursive: bool = False, 
                        _path: str = '') -> List[Tuple[str, str]]: 
        """
        Recupera lista com os itens da pasta

        :param folder_or_path: path ou objeto da pasta onde se deseja listar os arquivos
        :param parent_folder: objeto da pasta onde o path se inicia, por padrão root_folder
        :param recursive: quando True, efetua a listagem de forma recursiva
        :param _path: parâmetro de controle interno
        :returns: lista com os paths dos itens da pasta
        """       
        item_names = []

        folder = self.get_folder(folder_or_path, parent_folder)

        folders = self.get_folders(folder)
        for f in folders:
            if f.name.lower() != 'forms':
                item_names.append(('folder', _path + f.name))
                if recursive:
                    item_names.extend(self.list_item_names(f, recursive=recursive, _path=_path + f.name + '/'))

        files = folder.files.get().paged(100).execute_query()
        for file in files:
            item_names.append(('file', _path + file.name))

        return sorted(item_names, key=lambda item: item[1])
    
    
