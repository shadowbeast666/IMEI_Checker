from shareplum import Office365, Site
from shareplum.site import Version

import json, os, requests

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
config_path = '\\'.join([ROOT_DIR, 'config.json'])


# read json file
with open(config_path) as config_file:
    config = json.load(config_file)
    config = config['share_point']

USERNAME = config['user']
PASSWORD = config['password']
SHAREPOINT_URL = config['url']
SHAREPOINT_SITE = config['site']
SHAREPOINT_DOC = config['doc_library']

class Doit:
    def auth(self):
        self.authcookie = Office365(SHAREPOINT_URL, username=USERNAME, password=PASSWORD).GetCookies()
        self.site = Site(SHAREPOINT_SITE, version=Version.v365, authcookie=self.authcookie)

        return self.site

    
    def connect_folder(self, folder_name):
        self.auth_site = self.auth()

        self.sharepoint_dir = '\\'.join([SHAREPOINT_DOC, folder_name])
        self.folder = self.auth_site.Folder(self.sharepoint_dir)

        return self.folder

    def download_file(self, file_name, folder_name):
        self._folder = self.connect_folder(folder_name)
        return self._folder.get_file(file_name)


file_name = 'SOP.xlsx'

# set the folder name
folder_name = ''

# get file
file  = Doit().download_file(file_name, folder_name)

# save file
with open(file_name, 'wb') as f:
    f.write(file)
    f.close()

