rem pyinstaller -y  -F adfc_rest2.py myLogger.py printHandler.py textHandler.py tourRest.py tourServer.py
rem pyinstaller -y -F --debug --log-level DEBUG -p  ..\lib\site-packages --add-data "fonts/*.ttf;fonts" --add-data "ADFC_LOGO.png;." adfc_gui.py myLogger.py printHandler.py textHandler.py rawHandler.py pdfHandler.py tourRest.py tourServer.py
rem pyinstaller -y -F                           -p  ..\lib\site-packages --add-data "_builtin_fonts/*.ttf;_builtin_fonts" --add-data "ADFC_MUENCHEN.png;." adfc_gui.py myLogger.py printHandler.py textHandler.py rawHandler.py pdfHandler.py tourRest.py tourServer.py
pyinstaller -y -F -n ebics3  -p  ..\lib\site-packages --add-data "credentials.json.encr;." gui.py
