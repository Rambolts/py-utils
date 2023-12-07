from typing import List
import pandas as pd
from pydantic import BaseModel

def to_lower_camel(s: str) -> str:
    """
    Converte uma string no formato snake_case para lowerCamelCase

    :param s: string de entrada (formato snake_case)
    :returns: string de saída (formato lowerCamelCase)
    """
    camel = ''.join(word.capitalize() for word in s.split('_'))
    return camel[0].lower() + camel[1:]

def convert_model_list_to_dataframe(model_list: List[BaseModel], **todict_kwargs) -> pd.DataFrame:
    """
    Converte uma lista de modelos Pydantic em DataFrame

    :param model_list: lista de modelos Pydantic
    :param todict_kwargs: parâmetros a serem passados ao método dict() do modelo
    :returns: DataFrame
    """
    return pd.DataFrame([item.dict(**todict_kwargs) for item in model_list])
