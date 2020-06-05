from cx_Freeze import setup, Executable

base = None    

executables = [Executable("main.py", base=base)]

includeFiles = ['SISFIN.edp', 'classes\\sisfin\\__init__.py', 'credenciais.py']

packages = ["idna", "time", "pyodbc", "datetime", "getpass", "os"]
options = {
    'build_exe': {    
        'packages':packages,
        'include_files': includeFiles,
    },    
}

setup(
    name = "insert_saques_depositos_bb_base_sql",
    options = options,
    version = "1.0",
    description = 'Script que captura os dados do SISFIN referentes aos saques e depósitos do BB/BACEN e inclui no banco de dados.',
    executables = executables
)