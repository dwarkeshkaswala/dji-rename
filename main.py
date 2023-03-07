import os
import win32com.client
from pathlib import Path


def rename(path, org_file):
    sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
    file = Path(f'{path}/{org_file}')
    ns = sh.NameSpace(str(file.parent))
    name = ns.ParseName(file.name)
    tag = ns.GetDetailsOf(name,18)
    colon = "꞉"
    pipe = "⏐"
    time = ns.GetDetailsOf(name,3).replace(":", colon)
    filename = f"{time} {pipe} {file.name}"
    os.rename(f'{path}/{org_file}', f'{path}/{filename}')

    return filename, tag

def main():
    path = "D:/99/Drone"
    

    if not os.path.exists(f'{path}/Images/Raw'):
        os.makedirs(f'{path}/Images/Raw')
        
    if not os.path.exists(f'{path}/Images/HDR'):
        os.makedirs(f'{path}/Images/HDR')

    if not os.path.exists(f'{path}/Images/Pano'):
        os.makedirs(f'{path}/Images/Pano')

    if not os.path.exists(f'{path}/Images/Normal'):
        os.makedirs(f'{path}/Images/Normal')

    if not os.path.exists(f'{path}/Videos'):
        os.makedirs(f'{path}/Videos')

    lst = os.listdir(path)

    lst.remove('Images')
    lst.remove('Videos')

    # print((lst))

    for file in lst:
        new_filename, tag = rename(path, file)

        print(new_filename)

        try:

            if new_filename.endswith('.JPG'):
                if tag == 'hdr':
                    os.rename(f'{path}/{new_filename}', f'{path}/Images/HDR/{new_filename}')
                elif tag == 'pano':
                    os.rename(f'{path}/{new_filename}', f'{path}/Images/Pano/{new_filename}')
                elif tag == 'single':
                    os.rename(f'{path}/{new_filename}', f'{path}/Images/Normal/{new_filename}')
        
            elif new_filename.endswith('.MP4') or new_filename.endswith('.SRT'):
                os.rename(f'{path}/{new_filename}', f'{path}/Videos/{new_filename}')

            elif new_filename.endswith('.DNG'):
                os.rename(f'{path}/{new_filename}', f'{path}/Images/Raw/{new_filename}')
        except:
            print('Error')



if __name__ == '__main__':
    main()