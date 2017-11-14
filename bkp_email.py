# -*- coding: utf-8 -*-

import subprocess
from datetime import date
import os
from apiclient.http import BatchHttpRequest
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient.http import MediaFileUpload
import httplib2

FILE_NAME = 'C:\\Users\\rafael.basics\\Documents\Arquivos do Outlook\\Meu Arquivo de Dados do Outlook(1).pst'


def zip(filename):
    """ Cria o arquivo de backup do outlook.
        Um arquivo .zip com a data que o arquivo foi gerado

        Retorna: Nome do arquivo de backup
    """
    filename_ziped = '{}_{}.zip'.format(filename[:-4], date.today().strftime('%Y%m%d'))
    print('Gerando arquivo {}...'.format(filename_ziped))

    FNULL = open(os.devnull, 'w')
    subprocess.call(['7z', 'a', '-y', filename_ziped, filename],
                                            stdout=FNULL, stderr=subprocess.STDOUT)
    return filename_ziped


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/drive-python-quickstart.json
    SCOPES = 'https://www.googleapis.com/auth/drive'
    CLIENT_SECRET_FILE = 'client_secret.json'
    APPLICATION_NAME = 'Bkp Email'

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-bkp-email.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, None)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_parent_id(parent_name):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    parents = parent_name.split('\\')
    last_id_parent = None

    for index, item in enumerate(parents):

        query = "name = '{0}' \
                and mimeType = 'application/vnd.google-apps.folder' \
                and trashed = false".format(parents[index])

        if (last_id_parent == None and index > len(parents) and index > 0):
            query += " and {} in parents".format(last_id_parent)

        results = drive_service.files().list(q=query,
                                pageSize=10,
                                fields="nextPageToken, files(id, name)").execute()

        items = results.get('files', [])

        if not items:
            raise ValueError('Pasta não encontrada, por favor criar uma pasta com o nome {}'.format(parent_name))
        else:
            last_id_parent = items[0]['id']

    return last_id_parent


def upload(filename):
    """Upload do arquivo de backup para o Google Drive

    """
    FOLDER = 'semparar\\emails-bkp'
    
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    file_metadata = {'name': filename.split('\\')[-1:],
                     'mimeType': 'application/zip',
                     'parents': [get_parent_id(FOLDER)]}
    media = MediaFileUpload(filename,
                            mimetype='application/zip',
                            chunksize=1024*1024,
                            resumable=True)
    file = drive_service.files().create(body=file_metadata,
                                    media_body=media,
                                    fields='id').execute()

    print('Arquivo {} enviado com sucesso ({})'.format(filename, file.get('id')))

    
def delete(filename):
    """Remove arquivo local gerado

    """
    os.remove(filename)

    print('Arquivo removido {}'.format(filename))
    

def main():
    """Cria e faz o upload do arquivo de backup do outlook
    """
    filename_bkp = zip(FILE_NAME)
    upload(filename_bkp)
    delete(filename_bkp)


if __name__ == '__main__':
    main()


