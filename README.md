# Excel Multi Search

Search through multiple excel files across multiple folders.



https://github.com/xRyul/excel-multi-search/assets/47340038/c95eaba6-3d47-43d1-bbbc-559cfe403fe1


## How to install

a) Download the latest releae  

b) Setup `venv environemnt` -> clone the repo -> install all the dependencies `pip install -r requirements.txt` -> rrun it `python main.py`

## How to compile whole python project into a macOS app  

1. Install py2app by running the command `pip install py2app` in your terminal.
2. In the directory containing your Python script, create a `setup.py` file with the following contents:

```python
from setuptools import setup

APP = ['your_script.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': ['wx', 'pandas', 'openpyxl'],
}

setup(
    app=APP,
    name='Your Custom App Name',
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
```  

3. In your terminal, navigate to the directory containing your Python script and setup.py file, and run the command `python setup.py py2app`. This will create a standalone macOS application bundle in a new `dist` directory.

4. You can now run your Python app by opening the .app bundle in the `dist` directory.


```bash
pip install -r requirements.txt
```  
