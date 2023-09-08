no funciona en prog. consignaicon'
Private Sub CmdRegistrarHuella_Click()
Dim Ruta2 As String 'Dim obj As New Process
IdUsuarioAut = ""
NombreUsuarioAut = ""
Resp = ""
Ruta = "C:\OLIN_Almacen\Clave.txt"

Resp = Dir(Ruta)
If Resp <> "" Then
    Kill (Ruta)
End If

'lanza prog del lector
'"C:\OLIN_Almacen\OLIN Lithe V.exe" /URU Clave.txt,0,0,0,4

'''''Poner Comentario a estas 3 Lineas para pruebas
Ruta2 = "C:\Documents and Settings\adminmxslp\Escritorio\Copia OLIN Lithe V.exe.lnk"
Ruta2 = "C:\OLIN_Almacen\OLIN Lithe V.exe.lnk"

'Ruta2 = "C:\OLIN_Almacen\OLIN Lithe V.exe /URU Clave.txt,0,0,0,4"


'MsgBox Ruta2

'Resp = InputBox("ruta", "dfsw", Ruta2)


'OLIN Lithe V.exe

'MsgBox "Antes ShellExecute"
'executa cualquier doc, archivo, acceso directo
Call ShellExecute(Me.hwnd, "Open", Ruta2, "", "", 1)

'obj.start(Ruta2, appwinstyle.maximizedfocus)
'MsgBox "Despues ShellExecute"


ContIntentos = 1
Me.LblEspera.Caption = "Tiempo de Espera: " & 60 - ContIntentos
Me.LblEspera.Visible = True
Me.LblEspera.Refresh

'MsgBox "Antes timer"

'espera un moemnto a que se registre la huella
Me.Timer1.Interval = 2000
'activa centinela (Timer)
Me.Timer1.Enabled = True

End Sub