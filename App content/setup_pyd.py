from setuptools import setup
from Cython.Build import cythonize

# This setup file is now much more powerful.
# It uses glob patterns to find all your .py files.

setup(
    ext_modules = cythonize(
        [
            "main_app.py",       # Compile your main app
            "my_lib/*.py",       # Compile ALL .py files in the my_lib folder
            "Bot_module/*.py"  # Compile ALL .py files in the Bot_module folder
        ],
        compiler_directives={'language_level' : "3"} # Use Python 3 syntax
    )
)
