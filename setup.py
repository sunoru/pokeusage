from distutils.core import setup
import py2exe

setup(
    console=['pokeusage/main.py'],
    options={
        'py2exe': {'bundle_files': 2}
    },
    zipfile=None,
)
