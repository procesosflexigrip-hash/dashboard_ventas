@echo off
REM ============================================================
REM Actualiza el dashboard de Seguimiento de OT y lo sube a GitHub.
REM Pensado para correr 3 veces al dia via el Programador de tareas
REM de Windows (10:00, 12:00, 16:00).
REM ============================================================

REM --- AJUSTA ESTAS 3 RUTAS A TU COMPUTADORA ---
SET REPO_DIR=C:\Dashboards\seguimiento-ot
SET EXCEL_PATH=C:\Users\TU_USUARIO\OneDrive - Flexigrip\Seguimiento\SEGUIMIENTO_DE_OT1.xlsx
SET HTML_NAME=index.html
REM ------------------------------------------------

REM Archivo de log para poder revisar si algo fallo
SET LOG_FILE=%REPO_DIR%\actualizar_local.log

echo. >> "%LOG_FILE%"
echo ===== Ejecucion: %date% %time% ===== >> "%LOG_FILE%"

cd /d "%REPO_DIR%"
if errorlevel 1 (
    echo ERROR: no se pudo entrar a %REPO_DIR% >> "%LOG_FILE%"
    exit /b 1
)

REM Trae cualquier cambio remoto antes de generar (por si el repo
REM se edito desde otro lado, ej. GitHub web)
git pull >> "%LOG_FILE%" 2>&1

REM Regenera el HTML a partir del Excel local
python generar_dashboard_seguimiento.py --excel "%EXCEL_PATH%" --plantilla "%HTML_NAME%" --salida "%HTML_NAME%" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo ERROR: fallo generar_dashboard_seguimiento.py >> "%LOG_FILE%"
    exit /b 1
)

REM Commit y push solo si hubo cambios
git add "%HTML_NAME%" >> "%LOG_FILE%" 2>&1
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Actualizacion automatica: %date% %time%" >> "%LOG_FILE%" 2>&1
    git push >> "%LOG_FILE%" 2>&1
    echo Cambios subidos correctamente. >> "%LOG_FILE%"
) else (
    echo Sin cambios en los datos, no se hizo commit. >> "%LOG_FILE%"
)

echo ===== Fin ejecucion ===== >> "%LOG_FILE%"
