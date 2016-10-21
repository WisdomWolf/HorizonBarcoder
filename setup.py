from setuptools import setup
setup(name='horizon barcoder',
      version='0.3',
      py_modules=['HorizonBarcodePrepare'],
      install_requires=[
        "xlwt",
        "xlrd"
      ]
)
