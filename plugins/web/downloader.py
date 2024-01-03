import os
import requests
from tqdm.auto import tqdm

class DownloadFailedException(Exception):
    """ Exceção para indicar falha no download """
    pass

class DownloaderPlugin:

    @classmethod
    def check_already_downloaded(cls, url: str, /, *, 
                                 output_dir: str = '', output_filename: str = None, delete_previous: bool = True) -> bool:
        # compõe nome do arquivo
        filename = os.path.join(output_dir, output_filename or url.split('/')[-1])

        if not os.path.isfile(filename):
            return False
        
        response = requests.head(url)
        new_size = int(response.headers.get('content-length', 0))
        old_size = os.path.getsize(filename)
        if new_size != old_size:
            if delete_previous:
                os.remove(filename)
            return False # tamanho diferentes

        return True # arquivos sao iguais

    @classmethod
    def download(cls, url: str, /, * , 
                 output_dir: str = '', output_filename: str = None, max_tries: int = 100, enable_progress: bool = False) -> bool:
        """
        Efetua download de arquivo

        :param url: url do arquivo que se quer baixar
        :param output_dir: diretório onde se quer baixar o arquivo (por padrão, a pasta de execução)
        :param output_filename: nome do arquivo de saída (por padrão, o arquivo indicado pela url)
        :param max_tries: número máximo de tentativas, caso o download seja interrompido no meio (padrão 100)
        :param enable_progress: indica se deve mostrar barra de progresso (padrão não)
        :returns: True quando o download foi feito
        """
        already_downloaded = cls.check_already_downloaded(url, output_dir=output_dir, output_filename=output_filename, delete_previous=True)
        if already_downloaded:
            return False
        
        response = requests.get(url, stream=True)

        # recupera tamanho do arquivo
        total_size_in_bytes = int(response.headers.get('Content-length', 0))
        block_size = 1024 
        
        # compõe nome do arquivo
        filename = os.path.join(output_dir, output_filename or url.split('/')[-1])
        
        with tqdm(total=total_size_in_bytes, unit='iB', unit_scale=True, disable=not enable_progress) as progress_bar:
            t = 0
            total_downloaded = 0
            while t < max_tries:  
                # alterna modo de abertura do arquivo, dependendo do número de tentativas
                if t == 0:
                    mode = 'wb'
                else:
                    mode = 'ab'
                    filesize = os.path.getsize(filename)
                    # indica que queremos a partir de tal posição do arquivo (resumindo download)
                    response = requests.get(url, headers={'Range': f'bytes={filesize}-'}, stream=True)
                    # vish... não aceita, vamos ter que baixar de novo
                    if response.status_code==416:
                        mode = 'wb'
                        total_downloaded = 0
                        progress_bar.reset()
                        response = requests.get(url, stream=True)
                
                with open(filename, mode) as file:
                    for data in response.iter_content(block_size):
                        total_downloaded += len(data)
                        progress_bar.update(len(data))
                        file.write(data)

                # verifica se baixou tudo
                if total_size_in_bytes != 0 and total_downloaded != total_size_in_bytes:
                    t += 1
                else:
                    break
            else:
                # chegou no limite de tentativas, apaga arquivo
                os.remove(filename)
        
        # verifica se baixou, do contrário, lança exceção
        if not os.path.exists(filename):
            raise DownloadFailedException(f'O arquivo "{os.path.basename(filename).lower()}" não pode ser baixado.')
        
        return True
