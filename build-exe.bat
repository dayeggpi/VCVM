call pyinstaller --onefile --windowed --name "VCVM" ^
  --hidden-import pystray ^
  --hidden-import pystray._win32 ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageDraw ^
  --hidden-import pycaw ^
  --hidden-import pycaw.pycaw ^
  --hidden-import comtypes ^
  --hidden-import comtypes.client ^
  --collect-data pycaw ^
  --collect-binaries comtypes ^
  --add-data "icon.ico;." ^
  --add-data "icon_status_on.ico;." ^
  --add-data "icon_status_off.ico;." ^
  --icon "icon.ico" ^
  VCVM.py
  
echo Done!
echo The executable is in the dist folder.
pause