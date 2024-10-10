# hook-pyarrow.py

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all submodules
hiddenimports = collect_submodules('pyarrow')

# Collect data files
datas = collect_data_files('pyarrow')
