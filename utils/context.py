from ..data_source.sharepoint import SharepointDataSource

def conn(url: str, client_id, client_secret, logger):
    page = url.split('/')[-1]
    try:
        ctx = SharepointDataSource(url = url, client_id = client_id, client_secret = client_secret).connect()
        logger.info(f'\n\n######  Conexão com Sharepoint estabelecida, em {page} ######\n')
        return ctx
    except Exception as e:
        logger.error(f'\n\n###### !!! Falha na conexão com Sharepoint ({page}) !!! ######\n {e}')
        return