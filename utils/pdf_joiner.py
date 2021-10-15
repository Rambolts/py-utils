"""
pdf_joiner
    Aplicação para combinar dois ou mais arquivos PDF's em um só.
"""
__author__ = 'TI AG'
__version__ = '1.1.0'

# libs para menu e saída
import sys
import argparse

# lib para trabalhar com arquivos PDF
import fitz

# lib para barra de progresso
from tqdm.auto import tqdm

import warnings
warnings.filterwarnings('ignore')

def main(output, *pdfs, silent=False):
    """
    Rotina para combinar PDF's

    :param output: caminho do arquivo de saída
    :param pdfs: pdfs que se deseja combinar na sequencia
    """
    if not silent: print('[INFO] Combinando arquivos: ', end=' ')
    result = fitz.open()

    # verifica como irá iterar cada arquivo
    pdfs = args.pdfs if silent else tqdm(args.pdfs, desc='[INFO] Combinando arquivos: ', unit=' arquivos')
    for pdf in pdfs:
        # combina o pdf ao objeto writer
        result.insertPDF(fitz.open(pdf))

    if not silent: print('[INFO] Criando arquivo de saída...', end=' ')
    # escreve arquivo de saída
    result.save(output)
    if not silent: print('OK')

if __name__ == '__main__':
    # criando argument parser
    parser = argparse.ArgumentParser(description='Combina dois ou mais PDF\'s em um só.')

    # argumentos obrigatórios
    parser.add_argument('output', help='Nome do arquivo PDF de saída.')
    parser.add_argument('pdfs', nargs='+', help='Arquivos PDF na sequência que deverão ser combinados.')
    # argumentos opcionais
    parser.add_argument('-s', '--silent', action='store_true', default=False, help='Suprime mensagens de execução.')

    args = parser.parse_args()

    try:
        main(args.output, args.pdfs, silent=args.silent)
    except Exception as e:
        if not args.silent: print('FALHOU!', e, sep='\n')
        # falha
        sys.exit(1)
