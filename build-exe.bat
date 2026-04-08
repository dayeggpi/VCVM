call pyinstaller --clean --noconfirm --onefile --windowed --name "VCVM" ^
  --hidden-import pystray ^
  --hidden-import pystray._win32 ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageDraw ^
  --hidden-import pycaw ^
  --hidden-import pycaw.pycaw ^
  --hidden-import comtypes ^
  --exclude-module IPython ^
  --exclude-module jupyter ^
  --exclude-module matplotlib ^
  --exclude-module numpy ^
  --exclude-module pandas ^
  --exclude-module pytest ^
  --exclude-module scipy ^
  --exclude-module tkinter ^
  --add-data "icon.ico;." ^
  --add-data "icon_status_on.ico;." ^
  --add-data "icon_status_off.ico;." ^
  --icon "icon.ico" ^
  VCVM.py
  
echo Done!
echo The executable is in the dist folder.
pause
