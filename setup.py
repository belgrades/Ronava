from distutils.core import setup
import py2exe

Mydata_files = [('images', ['C:\\Users\\Crema\\PycharmProjects\\Ronava\\LogoRonava.png'])]

setup(
    windows=[{
        "script": "ronava.py",
         "icon_resources": [(1, "ronava.ico")]
    }],
    data_files=Mydata_files,
    options={"py2exe": {
                "includes": ["easygui", "os", "sys", "openpyxl"]
             }}
    )
