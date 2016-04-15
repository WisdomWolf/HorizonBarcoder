#!/usr/bin/python
import sys

def myPrint(uglyString):
    print(str(uglyString.encode(sys.stdout.encoding, errors='replace'))).decode()

def safePrint(nString):
    try:
        print(nString)
    except UnicodeEncodeError:
        print('Error encoding text, trying saferPrint\n***\n')
        try:
            myPrint(nString)
        except UnicodeEncodeError:
            print('Safe Print failed')