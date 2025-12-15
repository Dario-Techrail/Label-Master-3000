@echo off
REM Launcher per Generatore Documenti Excel
REM Avvia l'applicazione Python

echo ========================================
echo Generatore Documenti Excel
echo ========================================
echo.
echo Avvio applicazione...
echo.

REM Prova ad avviare con python
python main.py

REM Se python non funziona, prova con py
if errorlevel 1 (
    py main.py
)

REM Pausa alla fine per vedere eventuali errori
if errorlevel 1 (
    echo.
    echo ERRORE: Impossibile avviare l'applicazione.
    echo Assicurati che Python sia installato.
    echo.
    pause
)
