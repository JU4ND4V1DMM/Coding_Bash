'1011 & vbCrLf & _ 1000 & vbCrLf & 1111 & vbCrLf & 0100 & vbCrLf & 0101 & vbCrLf & 1110 & vbCrLf & 1000 & vbCrLf & 0100

Option Explicit

Dim options
Dim menu
Dim oShell
Dim seconds_t
Dim Hour_t
Dim minutes_t
Dim seconds_time
Sub wait(sleeping)
    WScript.Sleep sleeping * 1000
End Sub 



Dim Final_Time

'Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject ("WScript.Shell")

menu = "Menu de funcionalidades, elige una opcion:" & vbCrLf & _
        "" & vbCrLf & _
        "1. Programar hora de trabajo" & vbCrLf & _
        "2. Simulador de CTI BAR" & vbCrLf & _
        "3. Generar procesos (CPU)" & vbCrLf & _
        ""

    options = InputBox(menu & vbCrLf & "Ingresa un numero segun tu necesidad:","Solucionador")

    Select Case options

        Case "1"
            Hour_t = InputBox("Ingresa el numero de HORAS que restan, para seguir trabajando:", "Verificador de trabajo")
            If IsNumeric(Hour_t) Then
                Hour_t = ((Hour_t * 60000)/16)
                minutes_t = InputBox("Ingresa el numero de MINUTOS que restan, para seguir trabajando:", "Verificador de trabajo")
                If IsNumeric(minutes_t) Then
                    seconds_time = ((minutes_t * 1000)/16) + (Hour_t)
                    oShell.Run "shutdown.exe -s -t " & seconds_time & " -f", 0, False
                    Set oShell = Nothing
                    WScript.Echo "Que te siga rindiendo mucho. Te recordare para que descanses"
                Else
                    WScript.Echo "Por favor, ingresa un numero valido de segundos."
                End If
            Else
                WScript.Echo "Por favor, ingresa un numero valido de segundos."
            End If
        
        Case "2"
            MsgBox "Conexion CTI Fallo",16,"CTIBar"
            wait(60)
	        oShell.Exec "taskkill /f /im ""ctibar.exe"""
	        oShell.Exec "taskkill /f /im ""onexagentui.exe"""
            WScript.Echo "Ha pasado un minuto, vuelve a loguearte y envia la evidencia."

        Case "3"
            do
                oShell.Run "excel.exe", 1, False
                oShell.Run "scalc.exe", 1, False
                oShell.Run "winword.exe", 1, False
                oShell.Run "chrome.exe", 1, False
            loop

	Case Else
            MsgBox "Esa no es una opcion, hay un problema entre la pantalla y la silla.",16,"Error de compilacion"

    End Select