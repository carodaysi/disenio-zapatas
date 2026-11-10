@echo off
rem Abre la interfaz grafica de diseno de zapatas con el entorno spyder-env.
rem Si Miniconda esta en otra ruta, corrija la linea siguiente.
set PYTHONW=D:\Programas_inst\miniconda\envs\spyder-env\pythonw.exe
cd /d "%~dp0"
start "" "%PYTHONW%" "%~dp0Interfaz_Zapatas.py"
