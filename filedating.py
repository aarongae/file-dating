#!/usr/bin/env python
# encoding: utf-8
"""A simple command line tool for automatically renaming image, video and text files.

Copyright (c) Aaron Gätje. All rights reserved.
"""
import os
import re
import glob
import argparse
import calendar
import locale
from datetime import datetime, timezone
from PIL import Image


def main(directory, keep_name, rename_existing, include_subfolders, new_name, date_format, include_text):
    IMAGE_FILE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.raw', '.nef', '.bmp', '.webp', '.svg', '.tif', '.tiff']
    VIDEO_FILE_EXTENSIONS = ['.mov', '.mp4', '.webm', '.flv']
    TEXT_FILE_EXTENSIONS = ['.pdf'] if include_text else []
    DCIM_PREFIXES = ['IMG-', 'IMG_', 'DCIM', 'WIN_', 'MOV-', 'MOV_', 'DSC0', 'DSC-', 'DSC_', 'DSCN', 'VID-', 'VID_']
    rename_suggestions = []

    filepaths = get_filepaths(directory, include_subfolders,
                              IMAGE_FILE_EXTENSIONS + VIDEO_FILE_EXTENSIONS + TEXT_FILE_EXTENSIONS)
    if not filepaths:
        print("No files found in this directory.")
        return

    for filepath in filepaths:
        if check_iso_format(filepath) and not rename_existing:
            continue
        file_extension = os.path.splitext(filepath)[-1]
        if file_extension in IMAGE_FILE_EXTENSIONS + [ext.upper() for ext in IMAGE_FILE_EXTENSIONS]:
            date_time = get_date_taken(filepath)
        elif file_extension in VIDEO_FILE_EXTENSIONS + [ext.upper() for ext in VIDEO_FILE_EXTENSIONS]:
            date_time = get_media_created_date(filepath)
        elif file_extension in TEXT_FILE_EXTENSIONS + [ext.upper() for ext in TEXT_FILE_EXTENSIONS]:
            date_time = get_text_file_date(filepath)

        if date_time is None:
            date_time = get_date_from_filename(filepath, DCIM_PREFIXES)
        name = os.path.splitext(filepath.split(os.sep)[-1])[0] if (
                    keep_name or (file_extension in TEXT_FILE_EXTENSIONS)) else None
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
        rename_files(rename_suggestions)
    else:
        print("No renaming suggestions found.")


def get_filepaths(directory, include_subfolders, file_extensions):
    filepaths = []
    if include_subfolders:
        for format in file_extensions:
            filepaths.extend(glob.iglob(f'{directory}{os.sep}**{os.sep}*{format}', recursive=True))
    else:
        for format in file_extensions:
            filepaths.extend(glob.iglob(f'{directory}{os.sep}*{format}'))
    return filepaths


def check_iso_format(filepath):
    filename = filepath.split(os.sep)[-1]
    try:
        date = filename[0:10]
        datetime.fromisoformat(date)
    except:
        return False
    return True


def get_date_taken(filepath):
    try:
        im = Image.open(filepath)
        exif = im.getexif()
        if exif.get(306):
            date_time = exif.get(306)
        else:
            return None
            # date_time = im._getexif()[36867]
        # exif = im.getexif() new method since Python 3.7 does not work
        # date, time = exif.get(36867)
    except Exception as e:
        print(e)
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
    date_time = dt.astimezone(datetime.now(timezone.utc).astimezone().tzinfo)
    return date_time


def get_date_from_filename(filepath, dcim_prefixes):
    filename = os.path.splitext(filepath.split(os.sep)[-1])[0]  # remove file extension
    if filename[:4].upper() in dcim_prefixes:
        filename = filename[4:]

    try:
        if filename[6] == "_" or filename[6] == "-":
            try:
                return datetime.strptime(filename[0:6], '%y%m%d')
            except:
                return None
    except:
        pass
    try:
        if filename[8] == "_" or filename[8] == "-":
            try:
                return datetime.strptime(filename[0:8], '%Y%m%d')
            except:
                return None
    except:
        return None


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
            filename = f"{index + 1:02}-{filename}"
            # Search for year and insert into name if found
    m = re.search(r'((19|20)\d{2})', filename)
    if m:
        filename = f"{m.group()}-{filename}"
    return filename


def rename_files(rename_suggestions):
    print("RENAMING SUGGESTIONS")
    for origin, suggestion in rename_suggestions:
        print(f'{origin:50} --> ..{os.sep}{suggestion}{os.path.splitext(origin)[-1]}')
    answer = input("Rename to suggestions? (y/n) ")

    if answer.lower() in ["y", "yes"]:
        rename_error_counter = 0
        for origin, suggestion in rename_suggestions:
            directory = os.path.dirname(origin)  # get path without file name extension
            extension = os.path.splitext(origin)[-1]
            new_path = os.path.join(directory, suggestion) + extension

            counter = 1
            # Check for existing name
            while os.path.exists(new_path):
                new_path = os.path.join(directory, suggestion) + "_" + str(counter) + extension
                counter += 1
            try:
                os.rename(origin, new_path)
            except Exception as e:
                print(e)
                rename_error_counter += 1
                print(origin + " could not be renamed.")

        print(f"{len(rename_suggestions) - rename_error_counter} files renamed.")
    else:
        print("No files renamed.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", help="Directory with image or video files")
    parser.add_argument("-n", "--new-name", dest="new_name", action="store")
    parser.add_argument("-f", "--date-format", dest="date_format", action="store")
    parser.add_argument("-k", '--keep-name', dest='keep_name', action='store_true')
    parser.add_argument("-e", '--rename-existing', dest='rename_existing', action='store_true')
    parser.add_argument("-s", '--include-subfolders', dest='include_subfolders', action='store_true')
    parser.add_argument("-te", '--include-text', dest='include_text', action='store_true')

    parser.set_defaults(keep_name=False, rename_existing=False,
                        include_subfolders=False, include_text=False, date_format="%Y-%m-%d")
    args = parser.parse_args()

    main(args.directory, args.keep_name, args.rename_existing,
         args.include_subfolders, args.new_name, args.date_format, args.include_text)
