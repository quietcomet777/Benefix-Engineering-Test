from setuptools import setup

setup(
   name='Benefix-Engineering-Test',
   version='1.0',
   description='Converts pdf data to excel file',
   author='Dan Molina',
   author_email='digzmol45@gmail.com',
   packages=['my_App'],  
   install_requires=['openpyxl', 'pdfminer'], #external packages as dependencies
)