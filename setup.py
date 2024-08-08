import sys
import os
from cx_Freeze import setup, Executable

# Lista de arquivos a serem incluídos no build
files = ['Dollar.ico',] 

config = Executable(
    script='app.py',
    icon='Dollar.ico'
)

setup(
    name='DollarReportBot',
    version='1.0',
    description='This code fetches the current USD to BRL exchange rate, creates a Word report with the data and a screenshot, and converts the report to PDF using LibreOffice.',
    author='Patrícia Jorge',
    options={'build_exe': {'include_files': files, 'include_msvcr': True}},
    executables=[config]
)
