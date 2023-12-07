import json
import hashlib
import pandas as pd
from typing import Dict, Any

def hash_str(s: str, *, ignore_case: bool = False) -> str:
    """
    Calcula hash MD5 em uma string.

    :param s: string de entrada
    :returns: string com o hash
    """
    if ignore_case:
        s = s.lower()

    return hashlib.md5(s.encode('utf-8')).hexdigest()

def hash_dict(d: Dict[str, Any], *, ignore_case: bool = False) -> str:
    """
    Calcula hash de um dicionário

    :param d: dicionário de entrada
    :returns: string com o hash
    """
    hash_algorithm = hashlib.md5()

    json_str = json.dumps(d, ensure_ascii=False, sort_keys=True)
    if ignore_case:
        json_str = json_str.lower()

    encoded = json_str.encode()
    hash_algorithm.update(encoded)
    return hash_algorithm.hexdigest()

def hash_row(row: pd.Series, *, ignore_case: bool = False) -> str:
    """
    Cria um hash a partir do conteúdo do DataFrame

    :param row: linha do DataFrame
    :returns: hash da linha
    """
    s = ''.join(tuple(row.fillna('').astype('str')))
    if ignore_case:
        s = s.lower()
        
    return hashlib.md5(s.encode('utf-8')).hexdigest()