@echo off
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico khipro_ime.py
