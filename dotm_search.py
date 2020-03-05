#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# import sys
import re
import os
import argparse
import zipfile
import glob

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Andrew Belanger"


def scan_files(key, file_name):
    os.chdir(file_name)
    valid_files = []
    files = glob.glob("*.dotm")
    for file in files:
        current_path = os.getcwd()
        with zipfile.ZipFile(file, "r") as f:
            with f.open('word/document.xml', 'r') as read_file:
                for text in read_file.readlines():
                    if key.encode() in text:
                        match = text.find(key.encode())
                        pharse = text[match - 40: match + 40].decode()
                        valid_files.append(file)
                        print("Match Found in file " + current_path + file)
                        print("..." + str(pharse) + "...")
                        break
    print("Total dotm files searched: ", len(files))
    print("Total dotm files matched: ", len(valid_files))
    return valid_files


def main(args):
    print(args)
    scan_files(args.key, args.search_location)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="find the key in the given files")
    parser.add_argument("key", help="key to be searched for")
    parser.add_argument("search_location", default="./dotm_files/", help="the location to seach for the key in")
    args = parser.parse_args()
    main(args)
