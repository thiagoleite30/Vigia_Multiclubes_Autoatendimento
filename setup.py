"""import sys

from cx_Freeze import setup, Executable

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [
    Executable(script='main.py', base=base, icon='icon.ico')
]

buildOptions = {
    'packages': ['json', 'logging', 'os', 'shutil', 'socket', 'sys', 'time', 'datetime', 'pandas', 'psutil',
                 'TOPdeskPy', 'requests', 'venv'],
    'includes': ['datetime', 'winreg'],
    'include_files': ['configs.json', 'vcruntime140.dll']
}

setup(
    name='Vigia - Multiclubes Kiosk',
    version='1.3',
    author="Thiago Leite",
    description='Programa que monitora processo do Kiosk e abre chamado caso ele não esteja funcionando.',
    options=dict(build_exe=buildOptions),
    executables=executables
)
"""
import PyInstaller.__main__


PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--noconsole',
    '-iicon.ico',
    '-nVigia - Multiclubes Autoatendimento'
])
