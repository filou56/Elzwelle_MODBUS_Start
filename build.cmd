copy ..\elzwelle_modbus_start.py .
\opt\miniconda3\Scripts\pyinstaller.exe elzwelle_modbus_start.py
copy \opt\miniconda3\Library\bin\libcrypto-3-x64.dll dist\elzwelle_modbus_start\_internal
copy \opt\miniconda3\Library\bin\libssh2.dll dist\elzwelle_modbus_start\_internal
