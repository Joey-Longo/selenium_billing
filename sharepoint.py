from office365.sharepoint.client_context import ClientContext, AuthenticationContext
import os
from pure_helper import sp_client_id, sp_secret_plog


def file_upload(file_path, folder_url):
    site_url = 'Sharepoint_site'
    app_principal = {
        'client_id': sp_client_id(),
        'client_secret': sp_secret_plog(),
    }

    context_auth = AuthenticationContext(url=site_url)
    context_auth.acquire_token_for_app(client_id=app_principal['client_id'], client_secret=app_principal['client_secret'])
    ctx = ClientContext(site_url, context_auth)
    with open(file_path, 'rb') as content_file:
        file_content = content_file.read()

    target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    name = os.path.basename(file_path)
    target_file = target_folder.upload_file(name, file_content)
    ctx.execute_query()
    print(f'File has been uploaded to: {target_file.serverRelativeUrl}')
