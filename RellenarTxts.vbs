' Declaración de variables.
Dim ObjFSO, oShell, f, objFile, file
Dim IeNuevo, IeOriginal, Txt, Fecha

' Compruebo los parámetros de entrada del script
if WScript.Arguments.Count <> 2 then ' No tiene 2 parámetros de entrada

  ' Solicitar al usuario introducir los datos necesarios.
  IeNuevo = Inputbox("Introduce el IE del proyecto actual: ")
  ' Si se cancela termino la ejecución
  if IsEmpty(IeNuevo) then
    WScript.Quit
  end if

  IeOriginal = Inputbox("Introduce el IE del proyecto original: ")
  ' Si se cancela termino la ejecución
  if IsEmpty(IeOriginal) then
    WScript.Quit
  end if

else ' Tiene 2 parámetros de entrada

  ' Leer IEs de los parámetros de entrda del script si se ejecuta por consola introduciendo 2 parámetros de entrada
  IeNuevo    = Wscript.Arguments(0)
  IeOriginal = Wscript.Arguments(1)

end if

' Crear objeto para ejecutar comandos de sistema operativo
Set oShell = CreateObject ("WScript.Shell")
' Guardo los Txts encontrados en el archivo Txts. Espero a que el comando se ejecute para seguir con el script
oShell.run "cmd.exe /C dir IEXXxxxxxxx*.txt /b/s > Txts",0,1 

' Creo objeto para poder trabajar con archivos
Set ObjFSO = CreateObject("Scripting.FileSystemObject") 
Set f      = ObjFSO.OpenTextFile("Txts", 1, False)

' Obtengo fecha actual
Fecha = (FormatDateTime(Now(),2))

' Iterara por cada línea del archivo Txts (Por cada txt que empiece por IEXXxxxxxxx)  
Do Until f.AtEndOfStream

  ' Asigno en la variable file la ruta completa del Txt de la iteración
  file = f.ReadLine

  ' Reemplazo IEXXxxxxxxx por el IE correspondiente expresado en el parámetro de entrada
  Txt = Replace(file, "IEXXxxxxxxx", IeNuevo)

  ' Creo y relleno el Txt
  Set objFile = ObjFSO.CreateTextFile(Txt,True)
  objFile.WriteLine ("")
  objFile.WriteLine ("           Técnico:       R.Díaz")
  objFile.WriteLine ("           Fecha:         " & Fecha)
  objFile.WriteLine ("           Tema:          El archivo/objeto se encuentra en el proyecto original " & IeOriginal)
  objFile.Close

loop

' Cierro el archivo
f.close

' Compruebo como se ha ejecutado el script, si no tiene parámetros de entrada (entiendo que ha sido por doble click)
' informo de que los txts se han generado con éxito.
if WScript.Arguments.Count <> 2 then

' Informo de que los Txts se han generado con éxito.
Msgbox("Los Txts se han generado con exito")

end if

' Borro el archivo Txts
oShell.run "cmd.exe /C del Txts",0,1
