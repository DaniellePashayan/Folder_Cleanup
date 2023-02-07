import os
import shutil
import pandas as pd
import re
from datetime import datetime as dt
import numpy as np

SOURCE_PATH = 'M:/CPP-Data/CBO BATCHES/'
DEST_PATH = 'M:/CPP-Data/ARCHIVE - CBO Batches/'

INV_REGEX = '(?<!\d)(([4-9]\d{7})|([0-1])\d{8})(?!\d)'
CORR_EXT = ['doc', 'docx', 'pdf', 'png', 'msg',
            'tif', 'csv', 'txt', 'jpg', 'xlsx', 'xls']
DELETE_EXT = ['html', 'css', 'js', 'db', 'lnk', 'download','js.download']
FILE_AGE = 365
FILE_SIZE = 500  # MB


def get_connection_status():
    if os.path.exists('M:/'):
        return True
    else:
        return False


def get_info(path):
    if os.path.isfile(path):
        filename = path.split('/')[-1]
        invoice = '' if re.search(INV_REGEX, filename) is None else re.search(INV_REGEX, filename).group()
        try:
            ext = (filename.split('.')[-1]).lower()
        except:
            ext = ''
        size = int(os.path.getsize(path) / 1000)  # KB
        cre_dt = dt.fromtimestamp(os.path.getmtime(path))
        age = (dt.today() - cre_dt).days
        return [invoice, filename, ext, path, cre_dt, age, size]


def main():
    #TODO check connection to MDrive/VPN

    for _, subdirs, __ in os.walk(SOURCE_PATH):
        for subdir in subdirs:
            for file in [x.path for x in os.scandir(SOURCE_PATH + subdir) if os.path.isfile(x)]:
                new_path = file.replace(SOURCE_PATH + subdir, SOURCE_PATH)
                shutil.move(file, new_path)
            shutil.rmtree(SOURCE_PATH + subdir)

    for _, __, files in os.walk(SOURCE_PATH):
        df = [get_info(SOURCE_PATH + file) for file in files]
        df = pd.DataFrame(df, columns = ['Invoice', 'FileName', 'Extension', 'Path', 'CreationDate', 'Age', 'Size'])

    df['CreMonth'] = df['CreationDate'].apply(lambda x: pd.to_datetime(x).strftime('%m'))
    df['CreYear'] = df['CreationDate'].apply(lambda x: pd.to_datetime(x).strftime('%Y'))
    df['InvalidInvoice'] = df['Invoice'].apply(lambda x: False if len(x) > 0 else True)
    df["InvalidExtension"] = df["Extension"].apply(lambda x: False if x in CORR_EXT else True)
    df["TooBig"] = df["Size"].apply(lambda x: True if (x > FILE_SIZE & FILE_AGE > 60) else False)
    df["TooOld"] = df["Age"].apply(lambda x: True if x > FILE_AGE else False)
    df["DeleteExtension"] = np.where((df['Extension'].isin(DELETE_EXT)) |
                                     (df['Extension'] == df['FileName']), True, False)
    df['Move'] = np.where(
        (((df['InvalidInvoice']) | (df['InvalidExtension']) | (df['TooOld'])) &
         (df['DeleteExtension'] == False)), True, False)
    df.to_excel('last_run.xlsx', index = None)

    df_delete = df.copy()
    df_delete = df_delete[df_delete.DeleteExtension]

    for index, row in df_delete.iterrows():
        path = SOURCE_PATH + row.FileName
        os.remove(path)
        df_delete = df_delete.drop(index = index)

    df_archive = df.copy()
    df_archive = df_archive.query('TooBig')
    df_archive = df_archive.sort_values(by = 'CreationDate', ascending = True)

    import zipfile
    for index, row in df_archive.iterrows():
        if os.path.exists(row.Path):
            month = row.CreMonth
            year = row.CreYear

            yearpath = DEST_PATH + year + '/'
            monthpath = DEST_PATH + year + '/' + month + '/'

            if not os.path.exists(yearpath):
                os.mkdir(DEST_PATH + year)
            if not os.path.exists(monthpath):
                os.mkdir(DEST_PATH + year + '/' + month)

            filename = row.FileName.replace('.' + row.Extension, '') + '.zip'
            src = SOURCE_PATH + filename
            dest = monthpath + filename

            try:
                with zipfile.ZipFile(src, 'w') as zf:
                    zf.write(row.Path)
                shutil.move(src, dest)
                os.remove(row.Path)
                df_archive = df_archive.drop(index = index)
            except FileNotFoundError:
                df_archive = df_archive.drop(index = index)

    df_move = df.copy()
    df_move = df_move.query('Move')

    for index, row in df_move.iterrows():
        if os.path.exists(row.Path):
            exists = os.path.exists(row.Path)
            month = row.CreMonth
            year = row.CreYear

            yearpath = DEST_PATH + year
            monthpath = DEST_PATH + year + '/' + month + '/'
            newpath = monthpath + row.FileName

            if not os.path.exists(DEST_PATH + year):
                os.mkdir(DEST_PATH + year)
            if not os.path.exists(DEST_PATH + year + '/' + month):
                os.mkdir(DEST_PATH + year + '/' + month)
            shutil.move(row.Path, newpath)
            df_move = df_move.drop(index = index)
        else:
            df.drop(index = index)


if __name__ == '__main__':
    main()
