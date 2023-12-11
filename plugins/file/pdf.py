import fitz
import PIL.Image

import os
import re
import pathlib
import zipfile
import warnings
from operator import itemgetter
from itertools import groupby
from enum import Enum
from typing import Optional, Any, Dict, List, Tuple, Iterator, Union, Callable

try:
    from plugins.com import excel, powerpoint, word
    IMPORT_PYWIN32COM_OK = True
except ModuleNotFoundError:
    IMPORT_PYWIN32COM_OK = False

class SplitMethod(str, Enum):
    EQUAL_MATCHES = 'EQUAL_MATCHES'
    RIGHT_APPEND = 'RIGHT_APPEND'
    LEFT_APPEND = 'LEFT_APPEND'

def pixmap_to_image(pix: fitz.Pixmap) -> PIL.Image.Image:
    """
    Converte Pixmap para Image
    
    :param pix: objeto Pixmap para ser convertido
    :returns: objeto Image
    """
    img = PIL.Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    return img

def convert_form_value_to_bool(value: str) -> bool:
    """
    Converte valor de formulário PDF para booleano.

    :param value: valor que se deseja converter
    :returns: valor convertido
    """
    if isinstance(value, bool):
        return value
    
    if value.lower() in ['no', 'off', 'false', 'none', 'null']:
        return False
    
    return bool(value)

class PdfRecipes:

    @classmethod
    def load_pdf(cls, filename: str, *, password: Optional[str] = None) -> fitz.Document:
        """
        Carrega arquivo PDF

        :param filename: caminho completo do arquivo PDF
        :param password: senha, caso esteja criptografado
        :returns: objeto do arquivo pdf
        """
        pdf = fitz.open(filename)

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
        return pdf
    
    @classmethod
    def read_page_lines(cls, page: fitz.Page, *, 
                        bbox: Optional[Tuple[int, int, int, int]] = None, 
                        columns: Optional[List[int]] = None) -> List[str]:
        """
        Retorna as linhas lidas de um documento
        
        :param page: objeto do tipo fitz.Page
        :param bbox: tupla contendo as coordenadas do retângulo de leitura [xmin, ymin, xmax, ymax]
        :param columns: lista com as coordenadas das colunas
        :returns: retorna uma lista de strings
        """
        if bbox is None:
            bbox = page.rect
            
        tab_rect = fitz.Rect(bbox).irect
        xmin, _, xmax, _ = tuple(tab_rect)

        if tab_rect.isEmpty or tab_rect.isInfinite:
            warnings.warn('Coordenadas incorretas!', UserWarning)
            return []

        if type(columns) is not list or columns == []:
            coltab = [tab_rect.x0, tab_rect.x1]
        else:
            coltab = sorted(columns)

        if xmin < min(coltab):
            coltab.insert(0, xmin)
        if xmax > coltab[-1]:
            coltab.append(xmax)

        words = page.get_text_words()

        if words == []:
            warnings.warn('A página não contém nenhum texto!', UserWarning)
            return []

        alltxt = []

        # pega as palavras contidas no retângulo e distribui entre as colunas
        for w in words:
            ir = fitz.Rect(w[:4]).irect  # retângulo da palavra
            if ir in tab_rect:
                cnr = 0  
                for i in range(1, len(coltab)):  
                    if ir.x0 < coltab[i]:  
                        cnr = i - 1
                        break
                alltxt.append([ir.x0, ir.y0, ir.x1, cnr, w[4]])

        if alltxt == []:
            return []

        alltxt.sort(key=itemgetter(1))  # ordena as palavras verticalmente

        spantab = [] 

        for _, row in groupby(alltxt, itemgetter(1)):
            schema = [''] * (len(coltab) - 1)
            for c, words in groupby(row, itemgetter(3)):
                entry = ' '.join([w[4] for w in words])
                schema[c] = entry
            spantab.append(schema)

        return spantab

    @classmethod
    def read_all_pages(cls, document: Union[str, fitz.Document], *, 
                       password: Optional[str] = None,
                       use_line_parser: Optional[bool] = False) -> List[str]:
        """
        Extrai o conteúdo de texto de todas as páginas

        :param document: nome do arquivo PDF ou o próprio objeto PDF
        :param password: senha, caso esteja criptografado
        :param use_line_parser: quando True, faz o parser linha a linha
        :returns: lista com o conteúdo de texto de cada página
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
        
        if use_line_parser:
            return ['\n'.join(cls.read_page_lines(p)) + '\n' for p in pdf]
        else:
            return [p.get_text('text') for p in pdf]
    
    @classmethod
    def read_all(cls, document: Union[str, fitz.Document], *, 
                 password: Optional[str] = None,
                 use_line_parser: Optional[bool] = False) -> str:
        """
        Extrai todo o conteúdo de texto do arquivo PDF

        :param document: nome do arquivo PDF ou o próprio objeto PDF
        :param password: senha, caso esteja criptografado
        :param use_line_parser: quando True, faz o parser linha a linha
        :returns: string com todo o texto
        """
        return ''.join(cls.read_all_pages(document, password=password, use_line_parser=use_line_parser))
    
    @classmethod
    def read_form_values(cls, 
                         document: Union[str, fitz.Document], *, 
                         return_page: Optional[bool] = False, 
                         password: Optional[str] = None) -> Dict[Union[str, int], Any]:
        """
        Extrai valores de formulário.

        :param document: nome do arquivo PDF ou o próprio objeto PDF
        :param return_page: quando True, retorna o índice da página como chave no dicionário resultante
        :returns: dicionário com os dados do formulário.
            - return_page=False:
                {
                    'nome do campo': valor
                }
            
            - return_page=True:
                {
                    0: {
                        'nome do campo': valor
                    }
                }
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
        result = {}
        
        for i, page in enumerate(pdf):
            page_values = {}
            
            for widget in page.widgets():
                if widget.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                    page_values[widget.field_name] = convert_form_value_to_bool(widget.field_value)
                else:
                    page_values[widget.field_name] = widget.field_value
                    
            if return_page:
                result[i] = page_values
            else:
                result.update(page_values)
                
        return result
    
    @classmethod
    def update_form_values(cls, 
                           document: Union[str, fitz.Document], 
                           values: Dict[Union[str, int], Any], *,
                           output_filename: Optional[str] = None, 
                           password: Optional[str] = None) -> fitz.Document:
        """
        Atualiza campos do formulário

        :param document: nome do arquivo PDF ou o próprio objeto PDF
        :param values: dicionário contendo os valores dos campos do formulário
        :param output_filename: nome do arquivo a ser salvo, caso informado
        :param password: senha do arquivo, caso esteja protegido
        :returns: objeto do documento pdf com os campos do formulário preenchidos
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
        values_per_page = all([isinstance(k, int) for k in values])
        
        for i, page in enumerate(pdf):
            if values_per_page:
                page_values = values.get(i, {})
            else:
                page_values = values
                
            for widget in page.widgets():
                for k, v in page_values.items():
                    if widget.field_name==k:
                        if widget.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                            widget.field_value = convert_form_value_to_bool(v)
                        elif widget.field_type in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
                            if str(v) not in widget.choice_values:
                                warnings.warn(f'"{v}" não está na lista de opções do campo {k}', UserWarning)
                            widget.field_value = str(v)
                        else:
                            widget.field_value = str(v)
                        widget.update()
                        
            page = pdf.reload_page(page)

        if output_filename:
            pdf.save(output_filename)
            
        return pdf
    
    @classmethod
    def extract_pages(cls, 
                      document: Union[str, fitz.Document], 
                      pages: Union[str, List[int]], *, 
                      output_filename: Optional[str] = None, 
                      password: Optional[str] = None) -> fitz.Document:
        """
        Extrai páginas de um PDF e salva em outro arquivo.

        :param document: nome do arquivo PDF ou o próprio objeto PDF
        :param pages: pode ser uma lista de inteiros com os índices das páginas (0-index) ou uma string contendo os ranges desejados (1-index)
        :param output_filename: nome do arquivo PDF de saída
        :param password: senha do arquivo de origem, caso esteja criptografado
        :returns: pdf resultante
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')

        if isinstance(pages, str):
            segments = [s.strip() for s in re.split(r'[,;]', pages)]
            page_indexes = []
            for segment in segments:
                if '-' in segment:
                    first, last = [s.strip() for s in segment.split('-', maxsplit=1)]
                    for i in range(int(first), int(last)+1):
                        page_indexes.append(i-1)
                else:
                    page_indexes.append(int(segment)-1)
        else:
            page_indexes = pages

        assert max(page_indexes) < len(pdf) and min(page_indexes) >= 0, 'Foram indicadas páginas inexistentes no arquivo de origem'

        result_pdf = fitz.open()
        for i in page_indexes:
            result_pdf.insert_pdf(pdf, from_page=i, to_page=i)

        if output_filename:
            result_pdf.save(output_filename)
        return result_pdf

    @classmethod
    def join(cls, input_files: List[Union[str, fitz.Document]], *, output_filename: Optional[str] = None) -> fitz.Document:
        """
        Junta vários PDF's em um só

        :param input_files: arquivos que se deseja combinar. A ordem é importante.
        :param output_filename: nome do arquivo de saída.
        :returns: arquivo PDF combinado
        """
        pdf = fitz.open()
        for file in input_files:
            assert isinstance(file, (str, fitz.Document)), 'Tipo inválido de documento'
            if isinstance(file, str):
                pdf.insert_pdf(fitz.open(file))
            else:
                pdf.insert_pdf(file)

        if output_filename:
            pdf.save(output_filename)
        return pdf

    @classmethod
    def join_images(cls, img_list: List[str], *, output_filename: Optional[str] = None) -> fitz.Document:
        """
        Cria PDF a partir de uma lista de imagens

        :param img_list: lista com os arquivos de imagem
        :param output_filename: nome do arquivo de saída
        :returns: objeto do pdf criado
        """
        pdf = fitz.open()

        for f in img_list:
            # abre arquivo
            img = fitz.open(f)
            # avalia dimensões
            rect = img[0].rect
            # converte em pdf
            pdfbytes = img.convert_to_pdf()
            img.close()

            # abre pdf com os bytes
            img_pdf = fitz.open('pdf', pdfbytes)
            # cria nova página
            page = pdf.new_page(width=rect.width, height=rect.height)
            # carrega imagem na página
            page.show_pdf_page(rect, img_pdf, 0)

        if output_filename:
            pdf.save(output_filename)
        return pdf

    @classmethod
    def join_zipped(cls, zip_filename: str, *, output_filename: Optional[str] = None) -> fitz.Document:
        """
        Combina todos os PDF's de um arquivo zip em ordem alfabética em um só.

        :param zip_filename: path completo do arquivo zip de entrada
        :param output_filename: path completo do arquivo pdf de saída
        :returns: True quando houver pelo menos uma página no arquivo de saída
        """
        result = fitz.open()
        with zipfile.ZipFile(zip_filename) as z:
            for f in sorted(z.filelist, key=lambda zf: zf.filename.lower()):
                if f.filename.lower().endswith('.pdf'):
                    pdf = fitz.open('type', stream=z.read(f.filename))
                    result.insert_pdf(pdf)

        if len(result) > 0 and output_filename:
            result.save(output_filename)

        return result

    @classmethod
    def split(cls, 
              document: Union[str, fitz.Document], 
              sep: str, *, 
              method: SplitMethod = SplitMethod.EQUAL_MATCHES,
              preprocessing: Callable[[str], str] = None, 
              password: Optional[str] = None) -> Iterator[Tuple[re.Match, fitz.Document]]:
        """
        Separa PDF de acordo com seu conteúdo.

        :param document: nome do arquivo ou objeto pdf de entrada
        :param sep: padrão regex que indica a separação
        :param method: modo de quebra do documento
        :param preprocessing: função de preprocessamento do conteúdo.
        :param password: senha do arquivo, caso esteja criptografado
        :yields: tupla contendo o match do regex e o objeto pymupdf do documento quebrado
        """
        def equal_matches(m1: re.Match, m2: re.Match) -> bool:
            if m1 is None and m2 is None:
                return True
            
            if (m1 is not None and m2 is None) or \
               (m1 is None and m2 is not None):
                return False
            
            groups_m1 = m1.groups()
            groups_m2 = m2.groups()
            if len(groups_m1) > 0 or len(groups_m2) > 0:
                return ''.join(map(str.lower, groups_m1)) == ''.join(map(str.lower, groups_m2))

            match_m1 = m1.string[m1.start():m1.end()].lower()
            match_m2 = m2.string[m2.start():m2.end()].lower()
            return match_m1 == match_m2
        
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
        
        result_pdf = fitz.open()
        first_match = None
        for p, page in enumerate(pdf):

            # caso o sep seja None, separamos todas as páginas
            if sep is None:
                result_pdf.insert_pdf(pdf, from_page=p, to_page=p)
                yield None, result_pdf
                result_pdf = fitz.open()
                continue

            text = page.get_text('text')
            if preprocessing:
                text = preprocessing(text)

            match = re.search(sep, text, flags=re.I)

            if method==SplitMethod.EQUAL_MATCHES:
                if not equal_matches(match, first_match):
                    if len(result_pdf):
                        yield first_match, result_pdf
                        result_pdf = fitz.open()

                    first_match = match

                result_pdf.insert_pdf(pdf, from_page=p, to_page=p)

            elif method==SplitMethod.RIGHT_APPEND:
                if match:
                    if len(result_pdf):
                        yield first_match, result_pdf
                        result_pdf = fitz.open()
                    
                    first_match = match

                result_pdf.insert_pdf(pdf, from_page=p, to_page=p)

            elif method==SplitMethod.LEFT_APPEND:
                result_pdf.insert_pdf(pdf, from_page=p, to_page=p)
                if match:
                    yield match, result_pdf
                    result_pdf = fitz.open()

        # envia o que sobrou
        if len(result_pdf):
            yield first_match, result_pdf
            
    @classmethod
    def decrypt(cls, input_filename: str, password: str, *, output_filename: Optional[str] = None) -> fitz.Document:
        """
        Remove criptografia de arquivo

        :param input_filename: nome do arquivo de entrada
        :param password: senha do arquivo de entrada
        :param output_filename: nome do arquivo de saída
        """
        pdf: fitz.Document = fitz.open(input_filename)
        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
        
        if output_filename:
            pdf.save(output_filename, encryption=fitz.PDF_ENCRYPT_NONE)
        return pdf
    
    @classmethod
    def encrypt(cls, document: Union[str, fitz.Document], output_filename: str, password: str) -> bool:
        """
        Criptografa arquivo

        :param document: nome do arquivo ou objeto pdf de entrada
        :param output_filename: nome do arquivo de saída
        :param password: senha do arquivo de entrada
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
            pdf.save(output_filename, encryption=fitz.PDF_ENCRYPT_KEEP)
        else:
            pdf.save(output_filename, encryption=fitz.PDF_ENCRYPT_AES_128, user_pw=password)

        return True
    
    @classmethod
    def compress(cls, document: Union[str, fitz.Document], output_filename: str, *, password: Optional[str] = None) -> bool:
        """
        Comprime o PDF removendo duplicidades e compactando imagens e dados.

        :param document: nome do arquivo ou objeto pdf de entrada
        :param output_filename: nome do arquivo de saída
        :param password: senha, caso esteja protegido
        :returns: True quando consegue salvar
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
        pdf.save(output_filename, garbage=4, deflate=True)
        return True

    @classmethod
    def get_image_from_page(cls, 
                            document: Union[str, fitz.Document], 
                            page_index: int, *,
                            image_filename: Optional[str] = None,
                            zoom_factor: Optional[float] = None,
                            dpi: Optional[int] = None,
                            password: Optional[str] = None) -> PIL.Image.Image:
        """
        Extrai imagem da página do pdf.

        :param document: caminho completo ou objeto do arquivo pdf
        :param page_index: indice da página requerida
        :param image_filename: se indicado, é o nome do arquivo a ser salvo
        :param zoom_factor: fator de zoom (1=100%)
        :param dpi: pontos por polegada
        :param password: senha, caso esteja criptografado
        :returns: objeto de imagem
        """
        assert isinstance(document, (str, fitz.Document)), 'Tipo inválido de documento'

        if isinstance(document, str):
            pdf = fitz.open(document)
        else:
            pdf = document

        if pdf.is_encrypted and password is not None:
            rc = pdf.authenticate(password)
            if not rc > 0:
                raise ValueError('Senha incorreta')
            
        kwargs = {}
        if zoom_factor and not dpi:
            kwargs['matrix'] = fitz.Matrix(zoom_factor, zoom_factor)
        elif dpi:
            kwargs['dpi'] = dpi

        page: fitz.Page = pdf.load_page(page_index)
        pix: fitz.Pixmap = page.get_pixmap(**kwargs)

        if image_filename:
            pix.save(image_filename)

        return pixmap_to_image(pix)

    @classmethod
    def from_docx(cls, doc_filename: str, pdf_filename: str):
        """
        Converte arquivo DOCX em PDF. Necessário ter o Word instalado.

        :param doc_filename: nome do arquivo DOCX de entrada
        :param pdf_filename: nome do arquivo PDF de saída
        """
        assert IMPORT_PYWIN32COM_OK, 'Os requisitos para essa funcionalidade não foram instalados/ configurados.'
        
        with word.application(kill_after=True) as wd:
            doc_filename = str(pathlib.Path(doc_filename).resolve())
            pdf_filename = str(pathlib.Path(pdf_filename).resolve())

            doc = wd.Documents.Open(doc_filename)

            if os.path.exists(pdf_filename):
                os.remove(pdf_filename)

            doc.SaveAs(pdf_filename, FileFormat=word.constants.wdFormatPDF)
            doc.Close()
        return

    @classmethod
    def from_pptx(cls, pp_filename: str, pdf_filename: str):
        """
        Converte arquivo PPTX em PDF. Necessário ter o PowerPoint instalado.

        :param pp_filename: nome do arquivo PPTX de entrada
        :param pdf_filename: nome do arquivo PDF de saída
        """
        assert IMPORT_PYWIN32COM_OK, 'Os requisitos para essa funcionalidade não foram instalados/ configurados.'
        
        with powerpoint.application(kill_after=True) as pp:
            pp_filename = str(pathlib.Path(pp_filename).resolve())
            pdf_filename = str(pathlib.Path(pdf_filename).resolve())

            ppt = pp.Presentations.Open(pp_filename, WithWindow=0)

            if os.path.exists(pdf_filename):
                os.remove(pdf_filename)

            ppt.SaveAs(pdf_filename, FileFormat=powerpoint.constants.ppSaveAsPDF)
            ppt.Close()
        return

    @classmethod
    def from_xlsx(cls, excel_filename: str, pdf_filename: str, sheetname: str = None):
        """
        Converte arquivo XLSX em PDF. Necessário ter o Excel instalado.

        :param excel_filename: nome do arquivo XLSX de entrada
        :param pdf_filename: nome do arquivo PDF de saída
        :param sheetname: nome da aba (quando não informado, usa a ativa)
        """
        assert IMPORT_PYWIN32COM_OK, 'Os requisitos para essa funcionalidade não foram instalados/ configurados.'

        with excel.application(kill_after=True) as xl:
            excel_filename = str(pathlib.Path(excel_filename).resolve())
            pdf_filename = str(pathlib.Path(pdf_filename).resolve())

            wb = xl.Workbooks.Open(excel_filename)

            if sheetname:
                ws = wb.Sheets(sheetname)
            else:
                ws = wb.ActiveSheet

            if os.path.exists(pdf_filename):
                os.remove(pdf_filename)

            ws.ExportAsFixedFormat(0, pdf_filename)
            wb.Close()
        return