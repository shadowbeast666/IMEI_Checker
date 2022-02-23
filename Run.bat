@echo off
python Download.py
ping -n 6 127.0.0.1 > nul
python MakeFile.py
ping -n 6 127.0.0.1 > nul
python EmailSender.py
