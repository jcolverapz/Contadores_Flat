Public Sub MAxTM()

'On Error GoTo ErrHandler

'Conexion = IsWebConnected(MSG)
'If Conexion = False Then Exit Sub
'CNN.rsCmdMaxTM.Open
'If IsNumeric(CNN.rsCmdMaxTM!m) Then
'    IdTM = CNN.rsCmdMaxTM!m + 1
'Else
'    IdTM = 1
'End If
'CNN.rsCmdMaxTM.Close

If CNN.rsCmdTiempoMuerto.State = 1 Then
        CNN.rsCmdTiempoMuerto.Close
End If

CNN.CmdTiempoMuerto (-10) '(IdTM)
If CNN.rsCmdTiempoMuerto.EOF = True Then
    CNN.rsCmdTiempoMuerto.AddNew
 '       CNN.rsCmdTiempoMuerto!IdTM = IdTM
        CNN.rsCmdTiempoMuerto!Fec_Tiempo = Fecha
        CNN.rsCmdTiempoMuerto!IDParo = 999    ''''0
        
''          If UltimoConteo = "" Then
''                    CNN.rsCmdTiempoMuerto!hor_ini = Now
''          Else
''                    Resp = Hour(Mid(UltimoConteo, 14, 12))
''                    If Hour(Mid(UltimoConteo, 14, 12)) <> IdHorario Then
''                          CNN.rsCmdTiempoMuerto!hor_ini = Now
''                    Else
''                          CNN.rsCmdTiempoMuerto!hor_ini = UltimoConteo
''                    End If
''          End If
      
        CNN.rsCmdTiempoMuerto!hor_ini = UltimoConteo      '  - FraccionTiempo_2_5))   ' Now  'Time          '- (AlertaSupervisor / 3) ' UltimoConteo   'Right(UltimoConteo, 13)  '
        CNN.rsCmdTiempoMuerto!hor_fin = Now_      'Time
        CNN.rsCmdTiempoMuerto!Minutos = Format(((CNN.rsCmdTiempoMuerto!hor_fin - CNN.rsCmdTiempoMuerto!hor_ini) * 1440), "0.00")
        CNN.rsCmdTiempoMuerto!FechaCap = Date
        CNN.rsCmdTiempoMuerto!HoraCap = Time
        CNN.rsCmdTiempoMuerto!EmpleadoiId = "011-156"
        CNN.rsCmdTiempoMuerto!PC = PC
        CNN.rsCmdTiempoMuerto!AutorizaId = ""
        CNN.rsCmdTiempoMuerto!Status = "A"
        CNN.rsCmdTiempoMuerto!CodLinea = CodLinea
        CNN.rsCmdTiempoMuerto!Observ = ""
        CNN.rsCmdTiempoMuerto!Turno = Turno
        
        TurnoTM = Turno
        CNN.rsCmdTiempoMuerto!codopera = 0
        CNN.rsCmdTiempoMuerto!Item = 0
        CNN.rsCmdTiempoMuerto!codmaquina = 0
        CNN.rsCmdTiempoMuerto!Oma_Id = IdJC_TM
        
        CNN.rsCmdTiempoMuerto!IdHorario = IdHorario
        
        
        
    CNN.rsCmdTiempoMuerto.Update
    
    
     IdTM = CNN.rsCmdTiempoMuerto!IdTM
    
    
'     frmDigIn.LblVariables.Text = "PzsxMinuto:   " & PzsxMinuto & vbNewLine & frmDigIn.LblVariables.Text
'     frmDigIn.LblVariables.Text = "CicloxPz:   " & Format(CicloxPzs, "0.00") & vbNewLine & frmDigIn.LblVariables.Text
'     frmDigIn.LblVariables.Text = "TiempoCicloxPzsx2_5:   " & Format(TiempoCicloxPzsx2_5, "0.00") & vbNewLine & frmDigIn.LblVariables.Text
'     frmDigIn.LblVariables.Text = "FraccionTiempo_2_5:   " & Format(FraccionTiempo_2_5, "0.00000") & vbNewLine & frmDigIn.LblVariables.Text
'     frmDigIn.LblVariables.Text = "FECHA HORA " & Now & vbNewLine & frmDigIn.LblVariables.Text
'     frmDigIn.LblVariables.Text = "Ultimo Conteo: " & UltimoConteo & "   " & vbNewLine & frmDigIn.LblVariables.Text
    
   ' Me.FreLinea(2).Caption = DescLinea & "               TM *"
    
'    frmDigIn.LblVariables.Text = frmDigIn.LblVariables.Text & vbNewLine & "* New TM " & Now & vbNewLine
    
End If
CNN.rsCmdTiempoMuerto.Close
'ErrHandler:
'If Err.Number = -2147467259 Then
  ' Unload Me
  ' Me.lblactualiza.Caption = "Err: " & Format(Now, "dd/MM hh:mm")
 'Call IsWebConnected(MSG)
    
'End If

End Sub




