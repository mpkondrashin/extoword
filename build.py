#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    Generate executable
"""
import os
import sys
import platform

def exe():
    if platform.system() == 'Windows':
        return '.exe'
    return ''

def build():
    os.chdir(os.path.dirname(sys.argv[0]))
    src = [
        'xtow.py',
        'gui.py',
        'build.py',
        'version.py',
        'media/bullet-icon.gif'
    ]
    mtime = [os.path.getmtime(f) for f in src]
    latest_change = max(mtime)
    try:
        if os.path.getmtime('xtoword' + exe()) > latest_change:
            print('Skip build')
            return 1
    except FileNotFoundError:
        pass

    icon = 'icon/bullet.ico'
    if sys.platform == 'darwin':
        icon = 'icon/bullet.icns'

    # --noconsole ' \
    default_docx = 'venv/lib/python3.9/site-packages/docx/templates/default.docx'
    if platform.system() == "Windows":
        default_docx = 'venv/lib/site-packages/docx/templates/default.docx'
    command = 'pyinstaller --noconfirm --clean --onedir --onefile --specpath . ' \
              '--distpath . --workpath temp ' \
              '--add-data media/bullet-icon.gif{ps}media ' \
              '--add-data {default_docx}{ps}media ' \
              '--icon {icon} ' \
              '--name=xtoword gui.py'.format(ps=os.pathsep,
                                             default_docx=default_docx,
                                             icon=icon)

    print(command)
    return os.system(command)


if __name__ == '__main__':
    sys.exit(build())
