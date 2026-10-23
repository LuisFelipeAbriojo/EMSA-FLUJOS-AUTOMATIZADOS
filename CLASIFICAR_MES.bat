@echo off
cd /d "%~dp0"
if not exist modelo.pkl (
  echo Entrenando el modelo por primera vez...
  python entrenar.py historico.csv
)
python clasificar_mes.py LIBROS --salida SALIDA %*
pause
