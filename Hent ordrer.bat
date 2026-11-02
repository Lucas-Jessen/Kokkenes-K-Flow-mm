@echo off
rem Henter nye ordrekvitteringer fra Outlook ind i fanen Ordreliste. Kan køres med dobbeltklik eller af Opgavestyring.
cd /d "%~dp0"
python ordre_til_ark.py >> ordre_til_ark.log 2>&1
