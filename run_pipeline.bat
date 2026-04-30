@echo off
cd /d "C:\proyectos\Dash_consultas_pbi"
set PYTHONIOENCODING=utf-8
"C:\proyectos\.venv\Scripts\python.exe" -u push_dash_pbi_consultas.py >> "C:\proyectos\Dash_consultas_pbi\logs\pipeline.log" 2>&1
