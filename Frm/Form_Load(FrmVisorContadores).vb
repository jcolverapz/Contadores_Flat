Private Sub Form_Load()
'Frm Visor Contadores

On Error GoTo ErrHandler

'Debug.Print "FrmVisorContadores: " & Now

Dim noTarjeta As Integer

' Set Progress Bar Color RGB(26, 188, 156)
  Call SendMessage(ProgressBar1.hwnd, PBM_SETBARCOLOR, 0, ByVal RGB(252, 255, 0)) ' Black
  ' Set Progress Bar Background Color
  Call SendMessage(ProgressBar1.hwnd, PBM_SETBKCOLOR, 0, ByVal RGB(0, 0, 0))   ' -
  ' Set Progress Bar Position

'No_Linea_EXE = Right("VisorContadoresL1", 1)
'No_Linea_EXE = Right("VisorContadoresL2", 1)
No_Linea_EXE = Right("VisorContadoresL3", 1)

  Select Case No_Linea_EXE
      Case 1
          noTarjeta = 0             ' Board number    'LINEA 1
          CodLinea = 111
          No_Linea = 1
          No_Linea_EXE = 1
          
            
      Case 2
          noTarjeta = 1              ' Board number    'LINEA 2
          CodLinea = 221
          No_Linea = 2
          No_Linea_EXE = 2
            
        Case 3
          noTarjeta = 0              ' Board number    'LINEA 3
          CodLinea = 331
          No_Linea = 3
          No_Linea_EXE = 3
            
      Case 4
          noTarjeta = 0              ' Board number    'LINEA 4
          CodLinea = 441
          No_Linea = 4
          No_Linea_EXE = 4
            
  End Select

    Me.Label1.Caption = "Linea: " & No_Linea

Linea = "Linea " & No_Linea

Fecha = Date
                    
IdOperacion = 0
                    
Call UltimoTicketxLinea(Fecha, CodLinea)
 
Call LstActualizaContadores

'Tiempo Muerto
Me.TmrColores.Enabled = True

Sensor = ""

'Colocacion en Pantalla
If No_Linea_EXE = 1 Then
    Me.Left = -30000
    Sensor = "EI"
End If

If No_Linea_EXE = 2 Then
    Me.Left = -15000
    Sensor = "ED"
End If

If No_Linea_EXE = 3 Then
    Me.Left = -25000
    Sensor = "ED"
End If

If No_Linea_EXE = 4 Then
    Me.Left = 100
    Sensor = "ED"
End If

Turno = 0

'En Pantalla
Me.TxtJobCard(2).Text = IdJC
Me.TxtTicketMP(2).Text = TicketGem

Me.TxtNoParte(2).Text = NoParte
Me.lblfechaUltimo.Caption = FechaUltimo
'Me.TxtTM(3).BackColor = RGB(255, 161, 98)
'Me.TxtTM(3).Alignment = 1

Me.Caption = Me.Caption & "   Ver. " & App.Major & "." & App.Minor & "." & App.Revision

Me.Shape1.BorderColor = &H8000000F

ColorAlerta = &H8000000F

ErrHandler:

End Sub




