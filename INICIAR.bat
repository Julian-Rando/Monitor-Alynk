@echo off
REM --- CONFIGURACION ---
REM 1. Ruta donde está el programa de Python (ejecutor.py)
set PROGRAMA="C:\Users\USUARIO\Desktop\proyectoExcel\ejecutor.py"

REM 2. Ruta COMPLETA donde está su Excel de Finanzas
set EXCEL="C:\Users\USUARIO\Desktop\proyectoExcel\excel.xlsx"

REM --- EJECUCION ---
python %PROGRAMA% %EXCEL%