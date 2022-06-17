
REM Asigno los parametros de entrada a variables
set Archivo=%1
set IE=%2

REM Guardo la fecha actual en la variabl FECHA
for /f "tokens=1-3 delims=/- " %%a in ('date /t') do set FECHA=%%a/%%b/%%c

REM Quito el IEXXxxxxxxx y las comillas dek nombre del archivo base
set Archivo=%Archivo:IEXXxxxxxxx=% 
set Archivo=%Archivo:"=% 

REM Creo el nombre del TXT final.
set "Txt=%IE%%Archivo%"

REM Copio el archivo base con el nombre del IE correspondiente
copy nul "%Txt%"

REM Relleno el Txt creado

echo. >> "%Txt%"
echo        Nombre técnico:     R.Díaz >> "%Txt%"
echo        Fecha:              %FECHA% >> "%Txt%"
echo        Tema:               El archivo/objeto se encuentra en el proyecto original %3 >> "%Txt%"