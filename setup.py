import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_options = {
    'packages': [], 
    'excludes': [],
    'include_files': []
}

base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('gui_barcode_generator.py', base=base, target_name='BarcodeGenerator.exe')
]

setup(name='BarcodeGenerator',
      version = '1.0',
      description = 'Программа для генерации штрихкодов',
      options = {'build_exe': build_options},
      executables = executables)