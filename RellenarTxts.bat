echo off

forfiles /M IEXXxxxxxxx_*.txt /s /c "cmd /c AuxBat.bat @file %1 %2"

cls

echo Los txts han sido generados