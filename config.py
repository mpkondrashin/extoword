#!/usr/bin/env python
# -*- coding: utf-8 -*-
# cython: language_level=3
"""
    Convert RFP from Excel to Word.

    Configuration module
"""

import os
import errno
from collections import defaultdict
from functools import partial
import plistlib
from contextlib import suppress


__file_name = 'config.plist'
__folder = os.path.expanduser('~/Library/Application Support/ExToWord')
__file_path = os.path.join(__folder, __file_name)

with suppress(FileExistsError):
    os.mkdir(__folder)

__data = defaultdict(str)


def load():
    global __data
    try:
        with open(__file_path, 'rb') as fp:
            dd = partial(defaultdict, str)
            __data = plistlib.load(fp, fmt=plistlib.FMT_XML, dict_type=dd)
    except FileNotFoundError as e:
        if e.errno != errno.ENOENT:
            raise


def save():
    with open(__file_path, 'wb') as fp:
        plistlib.dump(__data, fp, fmt=plistlib.FMT_XML)


def get(key):
    load()
    return __data[key]


def set(key, value):
    __data[key] = value
    save()


if __name__ == '__main__':
    pass
