#!/usr/bin/env python
# encoding: utf-8
"""Managing photo and video file names.

Copyright (c) aarongae. All rights reserved.
"""

import os
import re
import glob
import argparse
import calendar
from datetime import datetime, timezone
from PIL import Image


def main(directory, keep_name, rename_existing, include_subfolders, new_name, date_format, include_text):
    image_file_extensions = ['.jpg', '.jpeg', '.png', '.raw', '.nef', '.bmp', '.webp', '.svg', '.tif', '.tiff']
    video_file_extensions = ['.mov', '.mp4', '.webm', '.flv']
    text_file_extensions = ['.pdf'] if include_text else []

    rename_suggestions = []    

    filepaths = get_filepaths(directory, include_subfolders, image_file_extensions+video_file_extensions+text_file_extensions)

    if not filepaths:
        print("No files found in this directory.")
        return

    for filepath in filepaths:
        if check_iso_format(filepath) and not rename_existing:
            continue
        file_extension = os.path.splitext(filepath)[-1]
        if file_extension in image_file_extensions:
            date_time = get_date_taken(filepath)
        elif file_extension in video_file_extensions:
            date_time = get_media_created_date(filepath)
        elif file_extension in text_file_extensions:
            date_time = get_text_file_date(filepath) 

        if date_time is None:
            date_time = get_date_from_filename(filepath)

        name =  os.path.splitext(filepath.split(os.sep)[-1])[0] if keep_name else None

        if new_name is not None and date_time is not None:
            name = new_name

        if date_time:
            try:
                date_time = date_time.strftime(date_format)
            except: 
                print(f"{date_format} is not valid date format.")
                date_time = date_time.strftime("%Y-%m-%d_%H%M%S")
                
        if date_time or name:
            suggestion = '_'.join(filter(None, (date_time, name)))
            rename_suggestions.append((filepath, suggestion))

    if rename_suggestions:
        rename_files(filepath, rename_suggestions)
    else:
        print("No renaming suggestions found.")

def get_text_file_date(filepath):
    from PyPDF2 import PdfFileReader
    with open(filepath, 'rb') as f:
        pdf = PdfFileReader(f)
        info = pdf.getDocumentInfo()
        try:
            date_time = info['/CreationDate']
        except:
            return None
        try:
            date_time = datetime.strptime(date_time.replace("'", ""), "D:%Y%m%d%H%M%S%z")
        except:
            date_time = date_time.replace("Z", "+")
            date_time = datetime.strptime(date_time.replace("'", ""), "D:%Y%m%d%H%M%S%z")

        date_time = date_time.astimezone()

    return date_time


def get_filepaths(directory, include_subfolders, file_extensions):
    filepaths = []
    file_extensions += [ext.upper() for ext in file_extensions] # Add uppercase extensions for case insensitivity
    if include_subfolders:
        for format in file_extensions:
            filepaths.extend(glob.iglob(f'{directory}{os.sep}**{os.sep}*{format}', recursive=True))

    else:
        for format in file_extensions:
            filepaths.extend(glob.iglob(f'{directory}{os.sep}*{format}'))
    return filepaths

def rename_files(filepath, rename_suggestions):
    print("RENAMING SUGGESTIONS")
    for origin, suggestion in rename_suggestions:
        print(f"{origin}  --->  ..{os.sep}{suggestion}{os.path.splitext(origin)[-1]}")

    answer = input("Rename to suggestions? (y/n)")
    if answer.lower() in ["y","yes"]:
        for origin, suggestion in rename_suggestions:
            directory = os.path.dirname(filepath) # get path without file name extension
            extension = os.path.splitext(filepath)[-1]
            new_path = os.path.join(directory, suggestion) + extension

            counter = 1
            # Check for existing name
            while os.path.exists(new_path):
                new_path = os.path.join(directory, suggestion) + "_" + str(counter) + extension
                counter += 1
            try:
                os.rename(filepath, new_path)
            except:
                print(filepath + " could not be renamed.")
                    
        print(f"{len(rename_suggestions)} files renamed.")
    else:
        print("No files renamed.")
   
    
def get_date_taken(filepath):
    try:
        im = Image.open(filepath)
        date_time = im._getexif()[36867]
        #exif = im.getexif() new method since Python 3.7 does not work
        #date, time = exif.get(36867)
    except:
        return None

    date_time = datetime.strptime(date_time, '%Y:%m:%d %H:%M:%S')

    return date_time

def get_media_created_date(filepath):
    from win32com.propsys import propsys, pscon

    try:
        properties = propsys.SHGetPropertyStoreFromParsingName(filepath)
        dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()    
    except:
        return None
    
    # convert UNIX timestamp to local timezone
    dt_local = dt.astimezone(datetime.now(timezone.utc).astimezone().tzinfo)
    
    return dt_local

def get_date_from_filename(filepath):
    filename = os.path.splitext(filepath.split(os.sep)[-1])[0] # remove file extension

    dcim_prefixes = ['IMG-','IMG_','DCIM', 'WIN_','MOV-','MOV_','DSC0','DSC-','DSC_','DSCN','VID-','VID_']
    if filename[:4].upper() in dcim_prefixes:
        filename = filename[4:]
    try:
        name_split = filename.split("_")
    except:
        return None
    
    if len(name_split[0]) == 8:
        try:
            date = datetime.datetime.strptime(name_split[0], '%Y%m%d')
        except:
            return None

    elif len(name_split[0]) == 6:
        try:
            date = datetime.strptime(name_split[0], '%y%m%d')
        except:
            return None
    else:
         return None

    return date

def check_iso_format(filepath):
    filename = filepath.split(os.sep)[-1]

    try:
        date = filename[0:10]
    except:
        return None

    try:
        datetime.fromisoformat(date)
    except:
        return False
    return True

def find_dates(filename): 
    locale.setlocale(locale.LC_TIME, "")
    month_list = [months for months in calendar.month_name if months != ""]
    
    regex = rf"((19|20)\d{{2}})[-_.](1[0-2]|0[1-9])[-_.](3[01]|0[1-9]|[12][0-9])"
    m = re.search(regex, filename)
    if m:
        if m.start() == 0:
            date = filename[0:10].replace("_", "-")
            date = filename[0:10].replace(".", "-")
            return date + filename[10:]
        else:
            date = m.group().replace("_", "-")
            date = m.group().replace(".", "-")
            filename = f"{date}_{filename}"
            return filename

    # Search for month and insert into name if found
    for index, month in enumerate(month_list):
        if month in filename:
            filename = f"{index+1:02}-{filename}" 
    # Search for year and insert into name if found        
    m = re.search(r'((19|20)\d{2})', filename)
    if m:
        filename = f"{m.group()}-{filename}"
    return filename

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", help="Directory with image or video files")
    parser.add_argument("--new-name", dest="new_name", action = "store")
    parser.add_argument("--date-format", dest="date_format", action = "store")
    parser.add_argument('--keep-name', dest='keep_name', action='store_true')
    parser.add_argument('--rename-existing', dest='rename_existing', action='store_true')
    parser.add_argument('--include-subfolders', dest='include_subfolders', action='store_true')
    parser.add_argument('--include-text', dest='include_text', action='store_true')


    parser.set_defaults(keep_name = False, rename_existing = False, 
            include_subfolders = False, include_text = False, date_format = "%Y-%m-%d")
    args = parser.parse_args()
    main(args.directory, args.keep_name, args.rename_existing,
            args.include_subfolders, args.new_name, args.date_format, args.include_text)
