@echo off
title Email sender
if exist requirements.txt goto install
:start
cls
python sender.py
goto exit
:install
pip install -r requirements.txt
del requirements.txt
goto start
:exit