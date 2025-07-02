# setup.py - File per creare l'eseguibile con cx_Freeze

import sys
from cx_Freeze import setup, Executable

# Dipendenze da includere
build_exe_options = {
    "packages": [
        "tkinter", 
        "openpyxl", 
        "json", 
        "datetime", 
        "os"
    ],
    "excludes": [
        "matplotlib", 
        "numpy", 
        "pandas",
        "scipy",
        "PIL"
    ],
    "include_files": [
        # Se hai file aggiuntivi da includere, aggiungili qui
        # ("path/to/file", "destination/in/build")
    ],
    "optimize": 2,
}

# Informazioni sull'eseguibile
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Per nascondere la console su Windows

setup(
    name="GeneratoreQuotazioni",
    version="2.0",
    description="Generatore di Template Excel per Quotazioni Progetti",
    author="Il tuo nome",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "quotation_generator.py",
            base=base,
            target_name="GeneratoreQuotazioni.exe",
            icon=None  # Aggiungi il path dell'icona se ne hai una
        )
    ]
)

# ISTRUZIONI PER COMPILARE:
# 
# 1. Installa le dipendenze:
#    pip install cx_Freeze openpyxl
#
# 2. Salva questo file come setup.py nella stessa cartella del main
#
# 3. Compila con il comando:
#    python setup.py build
#
# 4. L'eseguibile sarà nella cartella build/
#
# ALTERNATIVA CON PYINSTALLER (più semplice):
#
# 1. Installa PyInstaller:
#    pip install pyinstaller openpyxl
#
# 2. Compila con:
#    pyinstaller --onefile --windowed --name="GeneratoreQuotazioni" quotation_generator.py
#
# 3. L'exe sarà nella cartella dist/
