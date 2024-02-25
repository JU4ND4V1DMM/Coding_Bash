'Autor Ju4n D4v1d M. - desarrollado en NG
'Scripting.vbs 09/11/2023 (EC204h)
'1011 1000 1111 0100 0101 1110 1000 0100
'Modificaciones introducidas bajo ejecución y acumulativo de procesos

Option Explicit

Sub wait(seconds_t)
    WScript.Sleep seconds_t * 1000
End Sub 

Dim options
Dim options_2
Dim options_3
Dim options_4
Dim options_5
Dim menu
Dim oShell
'Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject ("WScript.Shell")

menu = "Menu de finalizadores, elige una opcion:" & vbCrLf & _
        "" & vbCrLf & _
        "1. Aplicaciones moviles" & vbCrLf & _
        "2. Aplicaciones fijas" & vbCrLf & _
        "3. Aplicaciones generales" & vbCrLf & _
        "4. Salir" & vbCrLf & _
        ""

    options = InputBox(menu & vbCrLf & "Ingresa un numero segun tu necesidad:","FINALIZADOR - JDM4")

    Select Case options

        Case "1"

            Dim menu_2
            menu_2 = "Menu de finalizadores, elige una opcion:" & vbCrLf & _
                    "" & vbCrLf & _
                    "1. AC / AC Corporativo" & vbCrLf & _
                    "2. Gestor Corporativo" & vbCrLf & _
                    "3. Salir" & vbCrLf & _
                    ""

                options_2 = InputBox(menu_2 & vbCrLf & "Ingresa un numero segun la aplicacion:","FINALIZADOR - JDM4")

                Select Case options_2

                    Case "1"
                        oShell.Exec "taskkill /f /im ""ac administrador de clientes.exe"""
                        WScript.Echo "AC se ha finalizado con exito"

                    Case "2"
                        oShell.Exec "taskkill /f /im ""gestion_corporativos.exe"""
                        WScript.Echo "El Gestor Corporativo se ha finalizado con exito"

                    Case "3"
                        'Pause

                    Case Else
                        MsgBox "Esa no es una opcion, hay un problema entre la pantalla y la silla.",16,"Error de compilación"

                End Select 

        Case "2"

            Dim menu_3
            menu_3 = "Menu de finalizadores, elige una opcion:" & vbCrLf & _
                    "" & vbCrLf & _
                    "1. RR" & vbCrLf & _
                    "2. SGA" & vbCrLf & _
                    "3. CRM" & vbCrLf & _
                    "4. NGN" & vbCrLf & _
                    ""

                options_3 = InputBox(menu_3 & vbCrLf & "Ingresa un numero segun la aplicacion:","FINALIZADOR - JDM4")

                Select Case options_3

                    Case "1"
                        oShell.Exec "taskkill /f /im ""rr.hod"""
                        oShell.Exec "taskkill /f /im ""comando rr.hod"""
                        WScript.Echo "RR se ha finalizado con exito"

                    Case "2"
                        oShell.Exec "taskkill /f /im ""modcxc.exe"""
                        oShell.Exec "taskkill /f /im ""sga.exe"""
                        WScript.Echo "SGA se ha finalizado con exito"
                    
                    Case "3"
                        oShell.Exec "taskkill /f /im ""crm.exe"""
                        WScript.Echo "CRM se ha finalizado con exito"

                    Case "4"
                        oShell.Exec "taskkill /f /im ""javaw.exe"""
                        oShell.Exec "taskkill /f /im ""java.exe"""
                        WScript.Echo "NGN se ha finalizado con exito"

                    Case Else 
                        MsgBox "Esa no es una opcion, hay un problema entre la pantalla y la silla.",16,"Error de compilación"

                End Select

        Case "3"

            Dim menu_4
            menu_4 = "Menu de finalizadores, elige una opcion:" & vbCrLf & _
                    "" & vbCrLf & _
                    "1. CTI BAR" & vbCrLf & _
                    "2. Avaya" & vbCrLf & _
                    "3. Excel / LibreOffice" & vbCrLf & _
                    "4. Word / Libreoffice" & vbCrLf & _
                    "5. Genesys (WorkSpace)" & vbCrLf & _
                    "6. Chrome" & vbCrLf & _
                    "7. Edge" & vbCrLf & _
                    "8. Internet Explorer" & vbCrLf & _
                    "9. OutLook" & vbCrLf & _
                    "0. Teams (Aplicacion)" & vbCrLf & _
                    ""

                options_4 = InputBox(menu_4 & vbCrLf & "Ingresa un numero segun la aplicacion:","FINALIZADOR - JDM4")

                Select Case options_4

                    Case "1"                        
                        oShell.Exec "taskkill /f /im ""ctibar.exe"""
                        WScript.Echo "La CTIBAR se ha finalizado con exito"

                    Case "2"
                        oShell.Exec "taskkill /f /im ""onexagentui.exe"""
                        WScript.Echo "El avaya se ha finalizado con exito"
                    
                    Case "3"
                        oShell.Exec "taskkill /f /im ""scalc.exe"""
                        oShell.Exec "taskkill /f /im ""excel.exe"""
                        WScript.Echo "Excel y LibreOffice han finalizado con exito"

                    Case "4"
                        oShell.Exec "taskkill /f /im ""libreoffice.exe"""
                        oShell.Exec "taskkill /f /im ""winword.exe"""
                        WScript.Echo "Word y LibreOffice han finalizado con exito"
                    
                    Case "5"
                        oShell.Exec "taskkill /f /im ""interactionworkspace.exe"""
                        WScript.Echo "WorkSpace o Genesys, se ha finalizado con exito"
                    
                    Case "6"
                        oShell.Exec "taskkill /f /im ""chrome.exe"""
                        WScript.Echo "Chrome se ha finalizado con exito"

                    Case "7"
                        oShell.Exec "taskkill /f /im ""msedge.exe"""
                        WScript.Echo "Edge se ha finalizado con exito"

                    Case "8"
                        oShell.Exec "taskkill /f /im ""interactionworkspace.exe"""
                        WScript.Echo "Internet Explorer se ha finalizado con exito"

                    Case "9"
                        oShell.Exec "taskkill /f /im ""outlook.exe"""
                        WScript.Echo "OutLook se ha finalizado con exito"
                    Case "0"
                        oShell.Exec "taskkill /f /im ""teams.exe"""
                        WScript.Echo "Teams se ha finalizado con exito"

                    Case Else 
                        MsgBox "Esa no es una opcion, hay un problema entre la pantalla y la silla.",16,"Error de compilación"

                End Select

        Case "4"
        'Pause

        Case Else
        MsgBox "Esa no es una opcion, hay un problema entre la pantalla y la silla.",16,"Error de compilación"



    End Select

wait(3)
WScript.Echo "Nos vemos luego  :)"

'WScript.Echo "Antes de despedirnos, limiparé temporales para mayor optimización en tu PC."

'@echo off
'del /S /F /Q %Temp%
'rd /S /Q %temp%
'echo terminado
'pause
