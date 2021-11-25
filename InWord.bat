@echo OFF
REM VERSION 0.1 -- Version de Prueba
REM Esta variable es el valor de "inw" en hexadecimal
set inw="69006E0077"


set wordVer=""
REM Recuperamos la version de word que utiliza
for /f "tokens=3" %%a in ('reg query HKCU\SOFTWARE\IN\in4\General ^| findstr "VersiWord"') do (
    if %%a == 0x0 ( set wordVer=7.0
        goto COMPROBARCOMPS
    ) 
    if %%a == 0x1 ( set wordVer=8.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x2 ( set wordVer=9.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x3 ( set wordVer=10.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x4 ( set wordVer=11.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x5 ( set wordVer=12.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x6 ( set wordVer=14.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x7 ( set wordVer=15.0
        goto COMPROBARCOMPS
    )
    if %%a == 0x8 ( set wordVer=16.0
        goto COMPROBARCOMPS
    )
)
exit

REM Comprobamos si hay algun complemento de inw deshabilitado
:COMPROBARCOMPS
for /f "tokens=*" %%a in ('reg query "HKCU\SOFTWARE\Microsoft\Office\%wordVer%\Word\Resiliency\disableditems" ^| find %inw%') do (
    goto COMPROBARPROCESOS
)
goto ABRIRWORD

REM Comprobamos si hay algun proceso de WINWORD abierto
:COMPROBARPROCESOS
for /f "tokens=*" %%a in ('tasklist /FI "IMAGENAME eq WINWORD.EXE" ^| find /N "WINWORD.EXE"') do (
    goto MENSAJECOMP
)
goto HABILITARCOMP

REM Hemos encontrado algun elemento deshabilitado y tiene Word abierto, preguntamos al ususario
:MENSAJECOMP
echo wscript.echo msgbox("Hay un problema con Word."+vbCr+"Guarde todos los documentos y pulse aceptar.",1+16,"Advertencia") >"%temp%\input.vbs"
for /f "tokens=* delims=" %%a in ('cscript //nologo "%temp%\input.vbs" "HOLA" "ADIOS"') do set YesNo=%%a
if "%YesNo%"=="2" (
    exit
)


REM Cerramos los procesos y habilitamos los complementos
:HABILITARCOMP
TASKKILL /f /t /im winword.exe
reg delete "HKCU\SOFTWARE\Microsoft\Office\%wordVer%\Word\Resiliency" /f

REM Abrimos word despues de comprobar que todo este correcto
:ABRIRWORD
REM Cogemos la ruta de instalacion de Word desde el registro
for /f "tokens=2*" %%a in ('reg query "HKCU\SOFTWARE\Microsoft\Office\%wordVer%\Word\Options" ^| find "PROGRAMDIR"') do (
    echo %%b
    start "" /MAX "%%bWINWORD.EXE"
)