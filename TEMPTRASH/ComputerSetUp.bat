echo off
python compSetUp.bat
set /p username=ENTER YOUR USERNAME
cd C:/Users/%username%/downloads/
start Setup.Def.en-us_O365BusinessRetail_06daa574-a2fa-4d8c-afa9-df2b6d2f7541_TX_PR_Platform_def_b_32_
echo CLICK YES ON NEXT SCREEN
timeout 5
start Teams_windows_x64
timeout 2
cd C:/Users/%username%/appdata/roaming/certasun/admin - Documents/Templates/For new blank documents
copy Blank.potx C:/Users/%username%/appdata/Roaming/Microsoft/templates/Blank.potx
copy Normal.dotm C:/Users/%username%/appdata/Roaming/Microsoft/templates/Normal.dotm
copy NormalEmail.dotm C:/Users/%username%/appdata/Roaming/Microsoft/templates/NormalEmail.dotm
copy Book.xltm C:/Users/%username%/appdata/roaming/microsoft/excel/xlstart/Book.xltm
cd C:/Users/%username%/appdata/roaming/microsoft/excel/xlstart
start Book.xltm
