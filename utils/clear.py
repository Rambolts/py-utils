import os

def clear_dir(root_path):
    """
    Limpa um diretório

    :param root_path: diretório a ser limpo
    """
    for root, dirs, files in os.walk(root_path, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))