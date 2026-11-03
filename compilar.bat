@echo off
REM Compila INFOLOG REMITOS en un unico .exe portable (dist\INFOLOG REMITOS.exe).
REM Requiere el entorno virtual creado con: python -m venv venv  y  pip install -r requirements.txt

call venv\Scripts\activate.bat || exit /b 1

pyinstaller --noconfirm --clean --onefile --windowed ^
    --name "INFOLOG REMITOS" ^
    --icon "recursos\imagenes\icono_infolog.ico" ^
    --add-data "recursos;recursos" ^
    --collect-data customtkinter ^
    --hidden-import PIL._tkinter_finder ^
    main.py

echo.
echo Listo: dist\INFOLOG REMITOS.exe
pause
