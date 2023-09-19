Private Sub TiempoMuerto()

On Error GoTo ErrHandler

 ' Tiempo Muerto - TM

'Tbl_DetOma Vs. Orden_Man Vs. TblHorariosHxH Vs. StdxLineaxHora
CNN.CmdStdxNoParte (Fecha), (CodLinea), (Turno)
If CNN.rsCmdStdxNoParte.EOF <> True Then
     IdJC_TM2_5 = CNN.rsCmdStdxNoParte!Oma_Id
     PzsxMinuto = (CNN.rsCmdStdxNoParte!pzsxhora / 60)
     CicloxPzs = (60 / PzsxMinuto)

     TiempoCicloxPzsx2_5 = (CicloxPzs * 2.5)

     FraccionTiempo_2_5 = (((TiempoCicloxPzsx2_5) / 1440) / 60)
'''     TmrColores.Interval = ((TiempoCicloxPzsx2_5 - 1) * 1000)
'     If Len(frmDigIn.LblVariables.Text) > 2000 Then
'          frmDigIn.LblVariables.Text = ""
'     End If

Else
          'busca para AT
          'CNN.CmdStdxNoParteAT (Fecha - 1), (CodLinea), (CodLinea), (Turno)
          CNN.CmdStdxNoParteAT (Fecha), (CodLinea), (Turno)
          If CNN.rsCmdStdxNoParteAT.EOF <> True Then
               IdJC_TM2_5 = CNN.rsCmdStdxNoParteAT!Oma_Id
               PzsxMinuto = (CNN.rsCmdStdxNoParteAT!pzsxhora / 60)
               CicloxPzs = (60 / PzsxMinuto)
               TiempoCicloxPzsx2_5 = (CicloxPzs * 2.5)
               FraccionTiempo_2_5 = (((TiempoCicloxPzsx2_5) / 1440) / 60)
         '''      TmrColores.Interval = ((TiempoCicloxPzsx2_5 - 1) * 1000)
'               If Len(frmDigIn.LblVariables.Text) > 2000 Then
'                    frmDigIn.LblVariables.Text = ""
'               End If
          Else
                    If CNN.rsCmdStdxNoParte.State = 1 Then
                         CNN.rsCmdStdxNoParte.Close
                    End If
                    If Weekday(Fecha) = 2 Then
                         CNN.CmdStdxNoParte (Fecha - 2), (CodLinea), (Turno)
                    Else
                         CNN.CmdStdxNoParte (Fecha), (CodLinea), (Turno)
                    End If
                    If CNN.rsCmdStdxNoParte.EOF <> True Then

                         IdJC_TM2_5 = CNN.rsCmdStdxNoParte!Oma_Id
                         PzsxMinuto = (CNN.rsCmdStdxNoParte!pzsxhora / 60)
                         CicloxPzs = (60 / PzsxMinuto)
                         TiempoCicloxPzsx2_5 = (CicloxPzs * 2.5)
                         FraccionTiempo_2_5 = (((TiempoCicloxPzsx2_5) / 1440) / 60)
'                       TmrColores.Interval = ((TiempoCicloxPzsx2_5 - 1) * 1000)
'                         If Len(frmDigIn.LblVariables.Text) > 2000 Then
'                              frmDigIn.LblVariables.Text = ""
'                         End If
'                             Else
'                         FraccionTiempo_2_5 = 0.00014 '(((TiempoCicloxPzsx2_5) / 1440) / 60)

                    End If
          End If
          CNN.rsCmdStdxNoParteAT.Close
End If
If CNN.rsCmdStdxNoParte.State = 1 Then
     CNN.rsCmdStdxNoParte.Close
End If

'frmDigIn.LblVariables.Text = frmDigIn.LblVariables.Text & vbNewLine & "NoParte: " & NoParte & vbNewLine

AlertaSupervisor = 0
AlertaGerenteProd = 0
AlertaDirector = 0

'''CNN.CmdTiempoAlertasParo ("Flat"), ("AlertaSupervisor")
'''If CNN.rsCmdTiempoAlertasParo.EOF <> True Then
'''    AlertaSupervisor = CNN.rsCmdTiempoAlertasParo!Mimutos / 1440
'''
''' ''  Alerta = CNN.rsCmdTiempoAlertasParo!Mimutos / 1440
'''
'''Else
'''    AlertaSupervisor = 0
'''End If
'''CNN.rsCmdTiempoAlertasParo.Close

AlertaSupervisor = (FraccionTiempo_2_5 * 2)

Call TurnoL
Call Horarios

'32:Corte Cuadros
'30: lavado
'31: impresion

IdOperacion = 30

'SQL = ""
'SQL = "SELECT TOP 10 Tbl_DetOma.IDENTITYCOL, Tbl_DetOma.Oma_Id,"
'SQL = SQL & "      Tbl_DetOma.IdHorario, Tbl_DetOma.HoraIni, Tbl_DetOma.HoraFin,"
'SQL = SQL & "      Tbl_DetOma.CodVidrio, Tbl_DetOma.PzsOK, Tbl_DetOma.FechaCap,"
'SQL = SQL & "      Tbl_DetOma.HoraCap, Tbl_DetOma.Observaciones, Tbl_DetOma.Item,"
'SQL = SQL & "      TblHorariosHxH.Turno, Tbl_DetOma.IdOperacion, lineas.codlinea,"
'SQL = SQL & "      lineas.descripcion"
'SQL = SQL & "  FROM Tbl_DetOma INNER JOIN"
'SQL = SQL & "      TblHorariosHxH ON"
'SQL = SQL & "      Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario INNER JOIN"
'SQL = SQL & "      Orden_Man ON"
'SQL = SQL & "      Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN"
'SQL = SQL & "      lineas ON Orden_Man.Codlinea = lineas.codlinea"

SQL = ""
SQL = "SELECT        TOP (10) Tbl_DetOma.Oma_Id, Tbl_DetOma.IdHorario, Tbl_DetOma.HoraIni, Tbl_DetOma.HoraFin, Tbl_DetOma.CodVidrio, Tbl_DetOma.PzsOK, Tbl_DetOma.FechaCap, Tbl_DetOma.HoraCap, Tbl_DetOma.Observaciones,"
SQL = SQL & "  Tbl_DetOma.Item , TblHorariosHxH.Turno, Tbl_DetOma.IdOperacion, Tbl_DetOma.CodLinea"
SQL = SQL & "  FROM            Tbl_DetOma INNER JOIN"
SQL = SQL & "                           TblHorariosHxH ON Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario"

'SQL = SQL & "  WHERE (Tbl_DetOma.FechaCap = CONVERT(DATETIME,"
SQL = SQL & "  WHERE "
'SQL = SQL & "   "
SQL = SQL & "  Tbl_DetOma.IdOperacion=" & IdOperacion & "AND Tbl_DetOma.codlinea=" & CodLinea
SQL = SQL & "  ORDER BY Tbl_DetOma.IDENTITYCOL DESC, Tbl_DetOma.HoraFin DESC"

If CNN.rsCmdUltimoConteoxJC.State = 1 Then
      CNN.rsCmdUltimoConteoxJC.Close
End If

  'Debug.Print "UltimoConteo" & ": " & UltimoConteo & ":" & Now

UltimoConteo = ""

'Tbl_detOma vs. tblHorariosHxH vx. Orden_Man Vs. Lineas
CNN.rsCmdUltimoConteoxJC.Open SQL, CNN.CNN
If CNN.rsCmdUltimoConteoxJC.EOF <> True Then

      IdJC_TM = CNN.rsCmdUltimoConteoxJC!Oma_Id
      'HoraUltimoConteo = CNN.rsCmdUltimoConteoxJC!HoraFin
    UltimoConteo = CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & CNN.rsCmdUltimoConteoxJC!HoraFin)
    'UltimoConteo = CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & "10:20:00")
    FrmVisorContadores3Lineas.Label21.Caption = "Ultimo Conteo: " & UltimoConteo & vbNewLine & "  Turno: " & Turno

    '   If (Now >= CDate(CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & (CNN.rsCmdUltimoConteoxJC!horafin + AlertaSupervisor)))) Then '  0.003  = 5 minutos
  'Debug.Print "UltimoConteo" & ": " & UltimoConteo & ":" & Now
      'LastConteo = CDate((CDbl(CDate(UltimoConteo)) + 0.002))
      If Now >= CDate(CDate((CDbl(CDate(UltimoConteo)) + AlertaSupervisor))) Then '  0.003  = 5 minutos

'   Resp = CSng(Now)
  '     If CSng(Now) >= CSng(CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & CNN.rsCmdUltimoConteoxJC!horafin)) + AlertaSupervisor Then    '  0.003  = 5 minutos
               ColorAlerta = &HFFFF&         ''Amarillo

               'Call SendAlert
                  ''If UltimoConteo <= (CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & CNN.rsCmdUltimoConteoxJC!horafin + AlertaSupervisor)) Then
                  Resp = (CDbl(CDate(CNN.rsCmdUltimoConteoxJC!FechaCap & " " & CNN.rsCmdUltimoConteoxJC!HoraFin) + AlertaSupervisor))

                  Dim FechaCompara As Date
                  If CDbl(Time) > 0 And CDbl(Time) < 0.302083 Then    'Es tercer turno
                         FechaCompara = Fecha + 1
                  Else
                         FechaCompara = Fecha
                  End If

                  If Resp < CDbl(CDate(FechaCompara & " " & Time)) Then   'Now Then

                        'Si hay el tiempo Muerto
                        If IdTM = 0 Then
                              Call MAxTM
                        Else
                              CNN.CmdTiempoMuerto (IdTM)
                              If CNN.rsCmdTiempoMuerto.EOF <> True Then

                                    ' Se termino el turno genera nuevo TM
                                      If CInt(CNN.rsCmdTiempoMuerto!Turno) <> Turno Then
                                          CNN.rsCmdTiempoMuerto.Close
                                          CNN.rsCmdUltimoConteoxJC.Close

                                          UltimoConteo = Now

                                          Call MAxTM

                                          Exit Sub

                                      Else
                                                ' Se termino el Horario nuevo TM
                                                If CNN.rsCmdTiempoMuerto!IdHorario <> IdHorario Then
                                                   CNN.rsCmdTiempoMuerto.Close
                                                   UltimoConteo = Now
                                                    Call MAxTM
                                                     If CNN.rsCmdTiempoMuerto.State = 1 Then
                                                            CNN.rsCmdTiempoMuerto.Close
                                                    End If
                                                    If CNN.rsCmdUltimoConteoxJC.State = 1 Then
                                                            CNN.rsCmdUltimoConteoxJC.Close
                                                    End If
                                                    Exit Sub

                                                Else
                                                        CNN.rsCmdTiempoMuerto!hor_fin = Now              'Time
                                                        CNN.rsCmdTiempoMuerto!Minutos = Format(((CNN.rsCmdTiempoMuerto!hor_fin - CNN.rsCmdTiempoMuerto!hor_ini) * 1440), "0.00")
'                                                        frmDigIn.LblVariables.Text = frmDigIn.LblVariables.Text & vbNewLine & "Update TM Mins: " & CNN.rsCmdTiempoMuerto!Minutos & "   " & Now & vbNewLine

                                                        If (CNN.rsCmdTiempoMuerto!Minutos >= 2 And CNN.rsCmdTiempoMuerto!IDParo = 999) Then
                                                                 CNN.rsCmdTiempoMuerto!IDParo = 0
                                                        End If

                                                        CNN.rsCmdTiempoMuerto.Update

                                                     'Me.FreLinea(2).Caption = DescLinea & "                TM ."
                                                End If
                                      End If
                              End If
                              CNN.rsCmdTiempoMuerto.Close
                        End If
                  End If
       End If
Else
      ''' Sino la encuentra busca en el programa
      '  (dbo.Schedule_Lineas.CodLinea = ?) AND (dbo.Schedule_Lineas.Fecha = ?) AND (dbo.Schedule_Lineas.IdHora = ?)
End If
CNN.rsCmdUltimoConteoxJC.Close

ErrHandler:
If Err.Number = -2147467259 Then
   'Unload Me
   ' Me.lblactualiza.Caption = "Err: " & Format(Now, "dd/MM hh:mm")
 Call IsWebConnected(MSG)

End If

End Sub


