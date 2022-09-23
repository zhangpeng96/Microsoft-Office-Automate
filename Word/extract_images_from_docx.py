from zipfile import ZipFile, BadZipfile
from pprint import pprint
from os import path
from os import walk as walk

def extract_images(file_path):
    with ZipFile(file_path, 'r') as archive:
        try:
            # pprint( archive.infolist() )
            members = []
            for file in archive.infolist():
                if ( file.filename.startswith('word/media/') and file.file_size > 1024
                     and path.splitext(file.filename)[-1] in ('.jpg', '.jpeg', '.png') ):
                    archive.extract(file, 'D:\\docss\\{}'.format(path.basename(file_path)))
                    # members.append(file)
                    print(file)
                # print(file.file_size)
            # archive.extractall(path=extract_path, pwd=password)
        except BadZipfile:
            print("Invalid Zip file")

