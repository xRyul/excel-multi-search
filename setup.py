from setuptools import setup

APP = ['main.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': ['wx', 'pandas', 'openpyxl'],
}

setup(
    app=APP,
    name='Excel Multi Search',
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
