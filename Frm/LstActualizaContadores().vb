Public Sub LstActualizaContadores()
On Error GoTo ErrHandler
'Debug.Print "LstActualizaContadores: " & Now

'Call LimpiarTXT
'Fecha = CDate("09/05/2017") '' Date
'IdJC = 1061570
Fecha = Date

ColorAlerta = &H8000000F '' &H80FF80

Call TurnoL

' Textbox(2)
ContLinea = 2

'En Pantalla: Contadores
    Me.TxtCorte(ContLinea).Text = 0
    Me.TxtCorte2(ContLinea).Text = 0
    Me.TxtCanteado(ContLinea).Text = 0
    Me.TxtCanteado2(ContLinea).Text = 0
    Me.TxtLavadora(ContLinea).Text = 0
    Me.TxtImpresora(ContLinea).Text = 0
' 
'Lineas Vs Maquina Vs. Operaciones Vs.Orden_Man
SQL = " SELECT lineas.descripcion AS Linea, Maquina.Descricion AS Maquina,"
SQL = SQL & "     operaciones.descripcion AS Operacion,"
SQL = SQL & "     Orden_Man.Oma_Id AS JobCard, Tbl_DetOma.CodVidrio AS Vidrio,"
SQL = SQL & "     Tbl_DetOma.Ticketgem AS Ticket, Tbl_DetOma.PxH,"
SQL = SQL & "     SUM(Tbl_DetOma.PzsOK) AS PzsOK, Tbl_DetOma.PzsScrap,"
SQL = SQL & "     lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,"
SQL = SQL & "     Orden_Man.Codlinea, TblHorariosHxH.Turno,"
SQL = SQL & "     Orden_Man.Oma_pza_prog AS goal, Tbl_DetOma.IDENTITYCOL"
SQL = SQL & " FROM Tbl_DetOma INNER JOIN"
SQL = SQL & "     Orden_Man ON"
SQL = SQL & "     Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN"
SQL = SQL & "     lineas ON Orden_Man.Codlinea = lineas.codlinea INNER JOIN"
SQL = SQL & "     operaciones ON"
SQL = SQL & "     Tbl_DetOma.IdOperacion = operaciones.codopera INNER JOIN"
SQL = SQL & "     Maquina ON"
SQL = SQL & "     Tbl_DetOma.IdMaquina = Maquina.codmaquina INNER JOIN"
SQL = SQL & "     TblHorariosHxH ON"
SQL = SQL & "     Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario"
SQL = SQL & " GROUP BY lineas.descripcion, Maquina.Descricion,"
SQL = SQL & "     operaciones.descripcion, Orden_Man.Oma_Id, Tbl_DetOma.CodVidrio,"
SQL = SQL & "     Tbl_DetOma.Ticketgem, Tbl_DetOma.PxH, Tbl_DetOma.PzsScrap,"
SQL = SQL & "     lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,"
SQL = SQL & "     Orden_Man.Codlinea, TblHorariosHxH.Turno,"
SQL = SQL & "     Orden_Man.Oma_pza_prog"
SQL = SQL & " HAVING (Orden_Man.Oma_Id = " & IdJC & ")" 
SQL = SQL & "    AND (lineas.descripcion LIKE N'%Linea " & No_Linea_EXE & "%')"
SQL = SQL & " ORDER BY Tbl_DetOma.IDENTITYCOL"

CNN.rsCmd_DetOma.Open SQL

    Resp = CNN.rsCmd_DetOma.RecordCount
    
    NoOperacionLstView = 0

        If CNN.rsCmd_DetOma.EOF <> True Then
            

            'Setear a cero los textbox
           ' Me.TxtStd.Text = 672
            Me.TxtPzsOK(ContLinea).Text = 0
            Me.TxtScrap(ContLinea).Text = 0
            Me.TxtTM(ContLinea).Text = 0
            Me.TxtOEE(ContLinea).Text = 0 '"85%"
            'Me.TxtStdxHr.Text = 0
            Me.TxtGoal(2).Text = CNN.rsCmd_DetOma!goal

            ' Actualiza Generales x Linea
            'Tbl_DetOma Vs Orden_Man
            
            CNN.CmdTotalesOkSxScrap (CodLinea), (Turno), (Fecha), (IdJC)          
            If CNN.rsCmdTotalesOkSxScrap.EOF <> True Then
               Do While CNN.rsCmdTotalesOkSxScrap.EOF <> True
                         ''   Me.TxtPzsOK(ContLinea).Text = (CNN.rsCmdTotalesOkSxScrap!PzsOK)
                          '  Me.TxtScrap(ContLinea).Text = Me.TxtScrap(ContLinea).Text + (CNN.rsCmdTotalesOkSxScrap!PzsScrap)
                          'Tiempo Muerto
                            Me.TxtTM(ContLinea).Text = Format((Me.TxtTM(ContLinea).Text + CNN.rsCmdTotalesOkSxScrap!MinsTM), "0.0")
                            Me.TxtTM(ContLinea).BackColor = vbWhite

                           'Goal
                           Me.TxtGoal(ContLinea).Text = CNN.rsCmdTotalesOkSxScrap!goal
                        
                            If Me.TxtPzsOK(ContLinea).Text < CNN.rsCmdTotalesOkSxScrap!goal Then
                                    Me.TxtGoal(ContLinea).BackColor = vbYellow
                            End If
                    CNN.rsCmdTotalesOkSxScrap.MoveNext
               Loop

               Me.TxtOEE(ContLinea).Text = 0 '"85%"
            Else
                Me.TxtPzsOK(ContLinea).Text = 0
                Me.TxtScrap(ContLinea).Text = 0
              '  Me.TxtTM(ContLinea).Text = 0
                Me.TxtOEE(ContLinea).Text = 0 
            End If
            CNN.rsCmdTotalesOkSxScrap.Close

            ''''            If Me.TxtPzsOK(ContLinea).Text > 20 Then
            ''''                    CNN.CmdStdxHr (CodLinea), (Turno), (Fecha), (IdJC)
            ''''                    If CNN.rsCmdStdxHr.EOF <> True Then
            ''''                     ''       Resp = CSng((CNN.rsCmdStdxHr!horafin - CNN.rsCmdStdxHr!horaini) * 1440)
            ''''                            Me.TxtStdxHr.Text = Int(CNN.rsCmdStdxHr!PzsOK / (((CNN.rsCmdStdxHr!horafin - CNN.rsCmdStdxHr!horaini) * 1440) / 60))
            ''''                    ''        Me.TxtOEE(ContLinea).Text = "%"
            ''''                    End If
            ''''                    CNN.rsCmdStdxHr.Close
            ''''             End If

            Call LimpiarTXT

            Do While CNN.rsCmd_DetOma.EOF <> True

                ' Busca la descripcion de la linea
                CNN.CmdBuscaLineas (CNN.rsCmd_DetOma!CodLinea)
                If CNN.rsCmdBuscaLineas.EOF <> True Then
                    Me.FreLinea(ContLinea).Caption = CNN.rsCmdBuscaLineas!descripcion
                    DescLinea = CNN.rsCmdBuscaLineas!descripcion
                Else
                    Me.FreLinea(ContLinea).Caption = "Sin Linea"
                End If
                CNN.rsCmdBuscaLineas.Close

             '   If NoOperacionLstView = 0 Then      'Para el primer no de parte

                    ' Busca OF
                    CNN.CmdOMA (CNN.rsCmd_DetOma!JobCard)
                    If CNN.rsCmdOMA.EOF <> True Then
                    
                        'Busca el No.Parte
                        CNN.CmdBuscaOF (CNN.rsCmdOMA!OF)
                        If CNN.rsCmdBuscaOF.EOF <> True Then
                            NoParte = CNN.rsCmdBuscaOF!nodeparte
                        Else
                            NoParte = CNN.rsCmdOMA!OF
                        End If
                        CNN.rsCmdBuscaOF.Close
                    End If
                    CNN.rsCmdOMA.Close

                    'En Pantalla
                    Me.TxtJobCard(ContLinea).Text = CNN.rsCmd_DetOma!JobCard
                    Me.TxtTicketMP(ContLinea).Text = TicketGem
                    Me.TxtNoParte(ContLinea).Text = NoParte
                    
                    'Estandar de Programa
                    CNN.CmdStdPrograma (CNN.rsCmd_DetOma!JobCard)


                    If CNN.rsCmdStdPrograma.EOF <> True Then
                         Me.TxtStd.Text = CNN.rsCmdStdPrograma!Pzs
                    Else
                        Me.TxtStd.Text = 0
                    End If
                    CNN.rsCmdStdPrograma.Close
           

                
                If CNN.rsCmd_DetOma!Operacion = "Canteado 2" Then
                    NoParte = NoParte
                End If



                'Resp = CNN.rsCmd_DetOma.RecordCount
                'Actualiza Cuenta por { operacion }
                Select Case CNN.rsCmd_DetOma!Operacion

                    Case "Corte Lineal"
                        'If Val(Me.TxtCorte(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                          '  Me.TxtCorte(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                            'If CNN.rsCmd_DetOma!PzsOK = 1 Then
                              '  Call LimpiarTXT
                            'End If
                        'Else
                         '   Me.TxtCorte(ContLinea).BackColor = &HFFFFFF      ' Blanco
                       ' End If
                        Me.TxtCorte(ContLinea).Text = CLng(Me.TxtCorte(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        Me.TxtCorte(ContLinea).BackColor = &HBFFFBF
                        
                        
                     Case "Corte Cuadros"
                       ' If Val(Me.TxtCorte2(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                           ' Me.TxtCorte2(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                        'Else
                          '  Me.TxtCorte2(ContLinea).BackColor = &HFFFFFF      ' Blanco
                        'End If
                        Me.TxtCorte2(ContLinea).Text = CLng(Me.TxtCorte2(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        Me.TxtCorte2(ContLinea).BackColor = &HBFFFBF
 
                    Case "Canteado 1"
                        'If Val(Me.TxtCanteado(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                           ' Me.TxtCanteado(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                        'Else
                          '  Me.TxtCanteado(ContLinea).BackColor = &HFFFFFF      ' Blanco
                        'End If
                        Me.TxtCanteado(ContLinea).Text = CLng(Me.TxtCanteado(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        'Me.TxtCanteado(ContLinea).BackColor = &HFFFF00

                    Case "Canteado 2"
                        'If Val(Me.TxtCanteado2(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                            'Me.TxtCanteado2(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                        'Else
                          '  Me.TxtCanteado2(ContLinea).BackColor = &HFFFFFF      ' Blanco
                        'End If
                        Me.TxtCanteado2(ContLinea).Text = CLng(Me.TxtCanteado2(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        'Me.TxtCanteado2(ContLinea).BackColor = &HFFFF00

                    Case "Lavado"
                       ' If Val(Me.TxtLavadora(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                           ' Me.TxtLavadora(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                        'Else
                          '  Me.TxtLavadora(ContLinea).BackColor = &HFFFFFF      ' Blanco
                        'End If

                        Me.TxtLavadora(ContLinea).Text = CLng(Me.TxtLavadora(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                       Me.TxtLavadora(ContLinea).BackColor = &HBFFFBF
 
                    Case "Impresion 1"
                        'If Val(Me.TxtImpresora(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                        '    Me.TxtImpresora(ContLinea).BackColor = &HFFFF00    '&H80FF80  ' Verde
                        'Else
                          '  Me.TxtImpresora(ContLinea).BackColor = &HFFFFFF      ' Blanco
                        'End If

                        Me.TxtImpresora(ContLinea).Text = CLng(Me.TxtImpresora(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        Me.TxtImpresora(ContLinea).BackColor = &HBFFFBF
 
                        ' Lbl Piezas Ok
                        Me.TxtPzsOK(ContLinea).Text = CLng(Me.TxtPzsOK(ContLinea).Text) + CNN.rsCmd_DetOma!PzsOK
                        Me.TxtPzsOK(ContLinea).BackColor = &H80FF80   ' Verde


                         If Me.TxtPzsOK(ContLinea).Text > 20 Then
                               CNN.CmdStdxHr (CodLinea), (Turno), (Fecha), (CNN.rsCmd_DetOma!JobCard) '(IdJC)
                               If CNN.rsCmdStdxHr.EOF <> True Then
                                       Resp = CSng((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440)

                                       If (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60) <> 0 Then
                                                
                                                
                                                Me.TxtStdxHr.Text = CLng(CNN.rsCmdStdxHr!PzsOK / (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60))
                                                Me.TxtOEE(ContLinea).Text = 0 ' "85%"
                                       End If
                               End If
                               CNN.rsCmdStdxHr.Close
                        End If
                '                    Case "Impresion 2"
                '                        If Val(Me.TxtImpresora2(ContLinea).Text) <> CNN.rsCmd_DetOma!PzsOK Then
                '                            Me.TxtImpresora2(ContLinea).BackColor = &H80FF80  ' Verde
                '                        Else
                '                            Me.TxtImpresora2(ContLinea).BackColor = &HFFFFFF      ' Blanco
                '                        End If
                '                        Me.TxtImpresora2(ContLinea).Text = CNN.rsCmd_DetOma!PzsOK
                '
                '                        Me.TxtPzsOK(ContLinea).Text = CNN.rsCmd_DetOma!PzsOK
                '                        Me.TxtPzsOK(ContLinea).BackColor = &H80FF80   ' Verde
                '
                '                         If Me.TxtPzsOK(ContLinea).Text > 10 Then
                '                               CNN.CmdStdxHr (CodLinea), (Turno), (Fecha), (IdJC)
                '                               If CNN.rsCmdStdxHr.EOF <> True Then
                '                                       Resp = CSng((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440)
                '
                '                                       If (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60) <> 0 Then
                '                                                Me.TxtStdxHr.Text = Int(CNN.rsCmdStdxHr!PzsOK / (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60))
                '                                                Me.TxtOEE(ContLinea).Text = 0 ' "85%"
                '                                       End If
                '                               End If
                '                               CNN.rsCmdStdxHr.Close
                '                        End If
                '

                End Select

                CNN.rsCmd_DetOma.MoveNext

                NoOperacionLstView = NoOperacionLstView + 1

            Loop

        End If
        CNN.rsCmd_DetOma.Close

        Me.Label1.Caption = "Ultima Actualizacion: " & Time

'Revisar
'Actualiza Scrap   Diferencia ( Corte Cuadros - Impresion )
'
'Tiempo Muerto

''Actualiza Conteos TM, Pzs Prod Ultima Etapa

'WHERE (Tbl_DetOma.IdOperacion = 31) AND      (Tbl_DetOma.FechaCap = CONVERT(DATETIME, '2018-11-07 00:00:00',    102)) AND (TblHorariosHxH.Turno = 1)

'No encontro datos??????

'Piezas por Turno
SQL = "  SELECT SUM(Tbl_DetOma.PzsOK) AS T_PzsOK, TblHorariosHxH.Turno,"
SQL = SQL & "      Tbl_DetOma.FechaCap"
SQL = SQL & "  FROM Tbl_DetOma INNER JOIN"
SQL = SQL & "      TblHorariosHxH ON"
SQL = SQL & "      Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario INNER JOIN"
SQL = SQL & "      Orden_Man ON"
SQL = SQL & "      Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN"
SQL = SQL & "      lineas ON Orden_Man.Codlinea = lineas.codlinea"
SQL = SQL & "  WHERE (Tbl_DetOma.IdOperacion = 1) AND"
SQL = SQL & "    (Tbl_DetOma.FechaCap = CONVERT(DATETIME, '" & Year(Fecha) & "-" & Month(Fecha) & "-" & Day(Fecha) & " 00:00:00',"
SQL = SQL & "    102)) AND (TblHorariosHxH.Turno = " & Turno & ") AND"
SQL = SQL & "      (lineas.descripcion LIKE N'%Linea " & No_Linea & "%')"
SQL = SQL & "  GROUP BY Tbl_DetOma.IdOperacion, TblHorariosHxH.Turno,"
SQL = SQL & "      Tbl_DetOma.FechaCap"
SQL = SQL & "  ORDER BY Tbl_DetOma.FechaCap DESC"

'"Pzs x Turno:
CNN.rsCmdPzsProdTurno.Open SQL
If CNN.rsCmdPzsProdTurno.EOF <> True Then
    Me.LblPzsOK.Caption = "Pzs x Turno: " & CNN.rsCmdPzsProdTurno!t_pzsok

           CNN.CmdStdxHr (CodLinea), (Fecha), (IdJC), (Turno)
           If CNN.rsCmdStdxHr.EOF <> True Then
                   Resp = CSng((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440)

                   If (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60) <> 0 Then
                            Me.TxtStdxHr.Text = CLng(CNN.rsCmdStdxHr!PzsOK / (((CNN.rsCmdStdxHr!HoraFin - CNN.rsCmdStdxHr!HoraIni) * 1440) / 60))
                          '  Me.TxtOEE(ContLinea).Text = "85%"
                   End If
           End If
           CNN.rsCmdStdxHr.Close

Else
    Me.LblPzsOK.Caption = "Pzs x Turno: 0"
End If
CNN.rsCmdPzsProdTurno.Close

'Tiempo Muerto
SQL = " SELECT Reg_TiempoMuerto.Fec_Tiempo,"
SQL = SQL & "     SUM(Reg_TiempoMuerto.Minutos) AS T, Reg_TiempoMuerto.Status,"
SQL = SQL & "     Reg_TiempoMuerto.Turno"
SQL = SQL & " FROM Reg_TiempoMuerto INNER JOIN"
SQL = SQL & "     lineas ON Reg_TiempoMuerto.CodLinea = lineas.codlinea"
SQL = SQL & " WHERE  (Reg_TiempoMuerto.Fec_Tiempo = CONVERT(DATETIME, '" & Year(Fecha) & "-" & Month(Fecha) & "-" & Day(Fecha) & " 00:00:00', 102))   AND "
SQL = SQL & "     (Reg_TiempoMuerto.Status = N'A') AND"
SQL = SQL & "     (Reg_TiempoMuerto.Turno = " & Turno & ") AND (lineas.descripcion LIKE N'%" & No_Linea & "%')"
SQL = SQL & " GROUP BY Reg_TiempoMuerto.Fec_Tiempo, Reg_TiempoMuerto.Status,"
SQL = SQL & "     Reg_TiempoMuerto.Turno"

  
'En Pantalla
CNN.rsCmdTotalTM.Open SQL
If CNN.rsCmdTotalTM.EOF <> True Then

            If CNN.rsCmdTotalTM!t > 0 Then
            
                Me.TxtTM(2).Text = Format(CNN.rsCmdTotalTM!t, "0.0")  ''Format(DisturbanceMinutos, "0.0")
                Me.TxtTM(2).BackColor = RGB(255, 165, 0)
            Else
                Me.TxtTM(2).Text = "0"
                Me.TxtTM(2).BackColor = &HBFFFBF
            End If

Else
       
End If
CNN.rsCmdTotalTM.Close

       
        'Scrap
        CNN.CmdScrap IdJC
        
        'If CNN.rsCmdScrap.EOF = False Then
            Do While CNN.rsCmdScrap.EOF = False
            
            Select Case CNN.rsCmdScrap!IdOperacion
            Case 1
            
                            If CorteLineal = True Then
                                'PzCorte = PzCorte + (PxH / 2)
                                PzCorte = (CNN.rsCmdScrap!PzsOK / 2)
                            Else
                               'PzCorte = PzCorte + (PxH)
                                PzCorte = CNN.rsCmdScrap!PzsOK
                            End If
                            
            Case 31
                PzImpresion = CNN.rsCmdScrap!PzsOK
            End Select
            
            CNN.rsCmdScrap.MoveNext
            
            Loop
         CNN.rsCmdScrap.Close
         
         Scrap = (PzCorte * PxH) - PzImpresion
         
         If Scrap <= 0 Then
            Me.TxtScrap(2).BackColor = &HBFFFBF
         Else
            Me.TxtScrap(2).Text = Scrap
            Me.TxtScrap(2).BackColor = RGB(255, 165, 0)
      End If
      
'SQL = " SELECT TOP 3 IDENTITYCOL, IdMDA_Status, codlinea, Fecha, Hora, Turno,  IdHorario, HoraInicio, HoraTermino, Minutos, EstadoAnterior,"
'SQL = SQL & "     EstadoActual , TimeStamp From Tbl_MDA_Status "
'SQL = SQL & "     WHERE (codlinea = N'" & Mid(CodLinea, 1, 1) & "') AND (Fecha = N'" & Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha) & "') AND (Turno = " & Turno & ")"
'SQL = SQL & "     ORDER BY IDENTITYCOL DESC"
'

'SQL = " SELECT TOP 3 IDENTITYCOL, IdMDA_Status, codlinea, Fecha, Hora, Turno,  IdHorario, HoraInicio, HoraTermino, Minutos, EstadoAnterior,"
'SQL = SQL & "     EstadoActual , TimeStamp From Tbl_MDA_Status "
'SQL = SQL & "     WHERE (codlinea = N'" & CodLinea & "') AND (Fecha = N'" & Day(Fecha) & "/" & Month(Fecha) & "/" & Year(Fecha) & "') AND (Turno = " & Turno & ")"
'SQL = SQL & "     ORDER BY IDENTITYCOL DESC"


'EstadoLinea = ""
'''(codlinea = ?) AND (Fecha = ?) AND (Turno = ?)
'CNN.CmdStatusLineaMDA (CodLinea), (Fecha), (Turno)

'MsgBox SQL

'Estados
'CNN.rsCmdStatusLineaMDA.Open SQL
'
'If CNN.rsCmdStatusLineaMDA.EOF <> True Then
'          If CNN.rsCmdStatusLineaMDA!EstadoActual = "Failure" Then
'               Me.ShpEstado.BackColor = vbRed
'               EstadoLinea = "Falla"
'          End If
'          If CNN.rsCmdStatusLineaMDA!EstadoActual = "Hold" Then
'               Me.ShpEstado.BackColor = vbYellow
'               EstadoLinea = " Paro"
'          End If
'          If CNN.rsCmdStatusLineaMDA!EstadoActual = "Run" Then
'               Me.ShpEstado.BackColor = vbGreen
'               EstadoLinea = "Corriendo"
'          End If
'          If CNN.rsCmdStatusLineaMDA!EstadoActual = "Off" Then
'               Me.ShpEstado.BackColor = &H808080
'               EstadoLinea = "Apagado"
'          End If
'          MinutosEstadoLinea = Format(CNN.rsCmdStatusLineaMDA!Minutos, "00.0")
'          Me.LblEstado.Caption = "Estado Actual de la Linea: " & MinutosEstadoLinea & " Mins  " & EstadoLinea & "."
'Else
'          EstadoLinea = ""

        '  If CDbl(UltimoConteoEstado) <> 0 Then

'                If ((Now - UltimoConteoEstado) * 1440) > 0.001 Then
'                     If ((UltimoConteoEstado - UltimoParoEstado) * 1440) < 2 Then
'                           Me.ShpEstado.BackColor = vbRed
'                           Me.LblEstado.Caption = "Estado Actual de la Linea: " & Abs(CInt((UltimoConteoEstado - CDate(Fecha & " " & Right(UltimoParoEstado, 13))) * 1440)) & " Minutos en Paro."
'                           MinutosEstadoLinea = Format(Abs(CInt((UltimoConteoEstado - UltimoParoEstado) * 1440)), "00.0")
'                           EstadoLinea = "Paro"
'                     Else
'                         Resp = Right(UltimoParoEstado, 13)
'                           Me.ShpEstado.BackColor = vbGreen
'                           Me.LblEstado.Caption = "Estado Actual de la Linea: " & (CInt((UltimoConteoEstado - CDate(Fecha & " " & Right(UltimoParoEstado, 13))) * 1440)) & " Minutos Activa"
'                           MinutosEstadoLinea = Format(CInt((UltimoConteoEstado - CDate(Fecha & " " & Right(UltimoParoEstado, 13))) * 1440), "00.0")
'                           EstadoLinea = "Activa"
'                     End If
'               End If


              ' Me.ShpEstado.BackColor = vbMagenta
              ' Me.LblEstado.Caption = "PC MDA apagada por favor incie sesion."
              ' MinutosEstadoLinea = Format(Abs(CInt((UltimoConteoEstado - UltimoParoEstado) * 1440)), "00.0")
               'EstadoLinea = "Off"
             '  Me.LblEstado.Refresh

        '  End If

'End If
'CNN.rsCmdStatusLineaMDA.Close
'Me.LblEstado.Refresh



ErrHandler:





End Sub


