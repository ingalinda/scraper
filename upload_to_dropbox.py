
import sys
import dropbox

from dropbox.files import WriteMode
from dropbox.exceptions import ApiError, AuthError

TOKEN = 'Access-token'

LOCALFILE = '/home/Linda/myfiles/Guardian.xlsx'

BACKUPPATH = '/Guardian.xlsx'


def checkFileDetails():
    print("File list is : ")
    for entry in dbx.files_list_folder('').entries:
        print(entry.name)

if __name__ == '__main__':
    if (len(TOKEN) == 0):
        sys.exit("ERROR: Invalid access token.")
    print("Creating a Dropbox object...")
    dbx = dropbox.Dropbox(TOKEN)
    
    try:
        dbx.users_get_current_account()
    except AuthError as err:
        sys.exit(
            "ERROR: Invalid access token; try re-generating an access token from the app console on the web.")

    try:
        checkFileDetails()
    except Error as err:
        sys.exit("Error while checking file details")

print("File uploaded to dropbox")
