@echo off
if exist requirements.txt goto install
:start
python sender.py
goto exit
:install
pip install -r requirements.txt
del requirements.txt
goto start
:exit