import os
import fitz
from utils.sort import sort_items

def join_pdf(dir_path, logger):
    """
    Faz a junção de todos os pdfs dentro de um diretório, respeitando uma certa ordenação: RIR > N > C > *
    
    :param dir_path: diretório a ser processado.
    """
    with os.scandir(dir_path) as entries:
        files = []
        for entry in entries:
            files.append(entry.name)
        files = sort_items(files)
        
        if files != []:
            logger.info(f' | Salvando {os.path.basename(dir_path)}...')
            pdf = fitz.open()
            for file in files:
                path = os.path.join(dir_path, os.path.basename(file))
                file = fitz.open(path)
                pdf.insertPDF(file)
                file.close()
            pdf.save(os.path.join(dir_path, f'GRD-{os.path.basename(dir_path)}_F.pdf'))
            pdf.close()