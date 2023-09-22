Public Sub UpdateConteos(Cont As Integer, CodLinea, Tot_Pz, Ticket)   ' IdJobCard As Long, NoCajaAzul As Integer)

'On Error GoTo ErrHandler
 
Dim TotPz   As Long

EmpleadoId = "011-056"

'Me.LblTarjeta.Caption = NoCajaAzul

TotPz = Tot_Pz
  
Fecha = Date
 
Call TurnoL


Select Case Cont 'Cont = IdOperacion
    Case 1
        IdMaquina = 201     'Cortadora Lineal
        IdOperacion = 1
        
    Case 2
        IdMaquina = 201     'Corte Cuadros
        IdOperacion = 32
    Case 3
        'IdMaquina = 205     'Canteado
       ' IdOperacion = 2
    Case 4
       ' IdMaquina = 206     'Canteado 2
       ' IdOperacion = 33
    Case 5
        IdMaquina = 207     'Lavado
        IdOperacion = 30
    Case 6
      
        IdMaquina = 208     'Impresion 1  'Aqui vamos a descontar  en la alimentacion manual los registros dobles
        IdOperacion = 31
            
    Case 7
        IdMaquina = 216     ''Impresion 2
        IdOperacion = 34
End Select

NoRegistro = 0


'Resp = CSng(Time)
'Resp = Mid(Time, 1, 8)

''' ???????
If CSng(Time) > 0.5 Then
    SQL = "SELECT IdHorario, HoraInicio, HoraTermino, Turno From TblHorariosHxH WHERE (HoraInicio <= '1/1/1900 " & Mid(Time, 1, 8) & " PM') AND (HoraTermino >= '1/1/1900 " & Mid(Time, 1, 8) & " PM')"
Else
    SQL = "SELECT IdHorario, HoraInicio, HoraTermino, Turno From TblHorariosHxH WHERE (HoraInicio <= '1/1/1900 " & Mid(Time, 1, 8) & " AM') AND (HoraTermino >= '1/1/1900 " & Mid(Time, 1, 8) & " AM')"
End If


'SQL = "SELECT IdHorario, HoraInicio, HoraTermino, Turno From TblHorariosHxH WHERE (HoraInicio <= '1/1/1900 " & Time & "') AND (HoraTermino >= '1/1/1900 " & Time & "')"

CNN.rsCmdIdHoraHxH.Open SQL, CNN.CNN
If CNN.rsCmdIdHoraHxH.EOF <> True Then
        IdHorario = CNN.rsCmdIdHoraHxH!IdHorario
Else
        'MsgBox "No hay Horario definido."
End If
CNN.rsCmdIdHoraHxH.Close

EsAcumulado = 0
'''

''Tbl:
CNN.CmdBuscaPKL_Programa Ticket, CodLinea        '''''', (Fecha)           '(IdHora),
If CNN.rsCmdBuscaPKL_Programa.EOF <> True Then
    Cantidad = CNN.rsCmdBuscaPKL_Programa!Cantidad
    'CNN.rsCmdInsertDetOM!CodVidrio = CNN.rsCmdBuscaPKL_Programa!Codigo

    CNN.CmdBuscaOF (CNN.rsCmdBuscaPKL_Programa!OF)

    If CNN.rsCmdBuscaOF.EOF <> True Then
        NoParte = CNN.rsCmdBuscaOF!nodeparte
    Else
        NoParte = CNN.rsCmdBuscaPKL_Programa!OF
    End If
    CNN.rsCmdBuscaOF.Close

    'Piezas por hoja
    '
    OF = CNN.rsCmdBuscaPKL_Programa!OF

    'Tbl: Orden_Man Vs Lite Vs Det_PackLst
    CNN.CmdPxH (Ticket), (OF), (IdJC)
    If CNN.rsCmdPxH.EOF <> True Then
        PxH = CNN.rsCmdPxH!nopzas
    Else

        CNN.CmdPxHat (OF), (Ticket)

            If CNN.rsCmdPxHAT.EOF <> True Then
                PxH = CNN.rsCmdPxHAT!nopzas
            Else
                PxH = 0
            End If
            CNN.rsCmdPxHAT.Close
    End If
            CNN.rsCmdPxH.Close
Else
    CNN.rsCmdBuscaPKL_Programa.Close

'    CNN.rsCmdUpdateConteos.Close
    'MsgBox "El Ticket NO esta programado para esta hora.", vbCritical, MSG
'Exit Sub
CNN.rsCmdBuscaPKL_Programa.Close

End If

'''
'Busca el Ultimo Ticket
'CNN.CmdUltimoTicketxLinea (Fecha), ("%" & No_Linea & "%")                    '("%3%")   '(CodLinea)    ', (Turno)
'Call UltimoTicketxLinea(Fecha, CodLinea)

 'En Pantalla
    'Me.TxtJobCard(2).Text = IdJC
   ' Me.TxtTicketMP(2).Text = Ticket
    
   ' Me.TxtNoParte(2).Text = NoParte
   ' Me.lblfechaUltimo.Caption = FechaUltimo
       
'Ultimo Conteo por Linea
'Tbl:
' CNN.CmdUltimoTicketxLinea Fecha, CodLinea
' If CNN.rsCmdUltimoTicketxLinea.EOF <> True Then
'    If CNN.rsCmdUltimoTicketxLinea!IdHorario = -1 Then
'            If CSng(Time) > 0.5 Then
'                SQL = "SELECT IdHorario, HoraInicio, HoraTermino, Turno From TblHorariosHxH WHERE (HoraInicio <= '1/1/1900 " & Mid(CNN.rsCmdUltimoTicketxLinea!HoraIni, 1, 8) & " PM') AND (HoraTermino >= '1/1/1900 " & Mid(CNN.rsCmdUltimoTicketxLinea!HoraIni, 1, 8) & " PM')"
'            Else
'                SQL = "SELECT IdHorario, HoraInicio, HoraTermino, Turno From TblHorariosHxH WHERE (HoraInicio <= '1/1/1900 " & Mid(CNN.rsCmdUltimoTicketxLinea!HoraIni, 1, 8) & " AM') AND (HoraTermino >= '1/1/1900 " & Mid(CNN.rsCmdUltimoTicketxLinea!HoraIni, 1, 8) & " AM')"
'            End If
'            CNN.rsCmdIdHoraHxH.Open SQL, CNN.CNN
'            If CNN.rsCmdIdHoraHxH.EOF <> True Then
'                    IdHorario = CNN.rsCmdIdHoraHxH!IdHorario
'            Else
'                   ' MsgBox "No hay Horario definido."
'            End If
'            CNN.rsCmdIdHoraHxH.Close
'            '
'            CNN.CmdBuscaDetOmaUpdate (IdJC), (TicketGem)
'            If CNN.rsCmdBuscaDetOmaUpdate.EOF <> True Then
'                        CNN.rsCmdBuscaDetOmaUpdate!IdHorario = IdHorario
'                CNN.rsCmdBuscaDetOmaUpdate.Update
'            Else
'                MsgBox " ", vbInformation, "Sin Scaneos"
'            End If
'            CNN.rsCmdBuscaDetOmaUpdate.Close
'    End If
'Else
' '   MsgBox "No hay contenedores scaneado en Corte.", vbInformation, "Sin Scaneos"
'    'CNN.rsCmdUltimoTicketxLinea.Close
''    Exit Sub
'End If
'CNN.rsCmdUltimoTicketxLinea.Close

'Busca si existe un registro con esos datos
'Actualiza los contadores por IdOperacion
'Tbl: Tbl_DetOma
CNN.CmdUpdateConteos (IdJC), (IdOperacion), (IdMaquina), (Ticket), (IdHorario)
'CNN.rsCmdUpdateConteos.Close

If CNN.rsCmdUpdateConteos.EOF <> True Then
    
   ' CNN.CmdPxH (Ticket), (OF), (IdJC)
    
       ' If CNN.rsCmdPxH.EOF <> True Then
           ' PxH = CNN.rsCmdPxH!nopzas
       ' Else
           ' CNN.CmdPxHat (OF), (IdJC)   ' (TicketGem)
           ' If CNN.rsCmdPxHAT.EOF <> True Then
            '    PxH = CNN.rsCmdPxHAT!nopzas
          '  Else
          '      PxH = 0
          '  End If
          '  CNN.rsCmdPxHAT.Close
      '  End If
        'CNN.rsCmdPxH.Close
        
        'Corte Lineal
        If IdOperacion = 1 Then
                'busca si hay corte Lineal o no
                CNN.CmdBuscaNoParte (OF)
                If CNN.rsCmdBuscaNoParte.EOF <> True Then
                    If CNN.rsCmdBuscaNoParte!CorteLineal = "N" Then
                        CorteLineal = False
                    Else
                        CorteLineal = True
                    End If
                Else
                    CNN.CmdBuscaVMPS (OF)
                    If CNN.rsCmdBuscaVMPS.EOF <> True Then
                        If CNN.rsCmdBuscaVMPS!CorteLineal = "N" Then
                            CorteLineal = False
                        Else
                            CorteLineal = True
                        End If
                    Else
                        CorteLineal = True
                    End If
                    CNN.rsCmdBuscaVMPS.Close
                End If
                CNN.rsCmdBuscaNoParte.Close
        End If
        
            CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + TotPz
               
            CNN.rsCmdUpdateConteos!HoraFin = Time
            'Ultimo Conteo
            UltimoConteo = CNN.rsCmdUpdateConteos!HoraFin
            
            CNN.rsCmdUpdateConteos.Update
            
                'CorteLineal
                
                            'If CorteLineal = False Then
                             '   CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + (PxH)
                            'Else
                             '   CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + (PxH / 2)
                            'End If
                            
                'rsCmdUpdateConteos :
                
'                If CorteLineal = False Then
'
'                Else
'                    CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + TotPz
'                End If
               
        'If IdOperacion = 1 Then
        
                
        'ElseIf IdOperacion = 32 Then
            'If (PxH Mod 2) = 1 Then     ' Es par  o es Non
               ' CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + TotPz
           ' Else
                'CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + TotPz ' (PxH / 2)
           ' End If
           ' CNN.rsCmdUpdateConteos!HoraFin = Time
            'UltimoConteo = CNN.rsCmdUpdateConteos!HoraFin
            'CNN.rsCmdUpdateConteos.Update
            
        'ElseIf IdOperacion = 30 Then
            'CNN.rsCmdUpdateConteos!PzsOK = CNN.rsCmdUpdateConteos!PzsOK + TotPz
            'CNN.rsCmdUpdateConteos!HoraFin = Time
            'UltimoConteo = CNN.rsCmdUpdateConteos!HoraFin
            'CNN.rsCmdUpdateConteos.Update
        'End If
'                  If EsAcumulado = "0" And IdOperacion = 31 Then
'
'''Guarda el dato de los conteos para el templado
'                              Call Guarda_Entrada_Templado

'               '''' actualiza conteos de entrada para la JC que se sube el material
'                              SQL = "              SELECT Tbl_DetOma.FechaCap, Orden_Man.Oma_Id,"
'                              SQL = SQL & "      Orden_Man.Oma_Tipo, Orden_Man.[OF], Orden_Man.Codlinea,"
'                              SQL = SQL & "      Orden_Man.Sensores, Orden_Man.En_Linea, Orden_Man.Oma_Status,"
'                              SQL = SQL & "      Tbl_DetOma.IdHorario, Tbl_DetOma.TicketProv,"
'                              SQL = SQL & "      Tbl_DetOma.IDENTITYCOL , Tbl_DetOma.PzsOK"
'                              SQL = SQL & "  FROM Orden_Man INNER JOIN"
'                              SQL = SQL & "      Tbl_DetOma ON Orden_Man.Oma_Id = Tbl_DetOma.Oma_Id"
'                              SQL = SQL & "  WHERE (Orden_Man.Oma_Tipo = N'Tem') AND"
'                              SQL = SQL & "      (Orden_Man.Sensores LIKE N'%" & Sensor & "%') AND"
'                              SQL = SQL & "      (Tbl_DetOma.TicketProv = N'TempEnt') AND"
'                              SQL = SQL & "     (Tbl_DetOma.FechaCap = CONVERT(DATETIME, '" & Year(Fecha) & "-" & Month(Fecha) & "-" & Day(Fecha) & " 00:00:00',102)) "
'                              SQL = SQL & "  AND (Tbl_DetOma.IdHorario = " & IdHorario & ") AND  (Orden_Man.Oma_Status = N'Activa') AND (Orden_Man.En_Linea = N'')"
'
'                             ' MsgBox SQL
'                              CNN.rsCmdBuscaJcProgramadaMismoCarril.Open SQL, CNN.CNN
'                              If CNN.rsCmdBuscaJcProgramadaMismoCarril.EOF <> True Then
'                                    Resp = CNN.rsCmdBuscaJcProgramadaMismoCarril!Oma_Id
'                                    NoRegistro = CNN.rsCmdBuscaJcProgramadaMismoCarril!Item
'                                    '''Busca el regsitro al que le voy a descontar las piezas de la impresion ''' Debe estar dobleteado el conteo
'                                    CNN.rsCmdBuscaItem.Open "UPDATE Tbl_DetOma SET PzsOK = PzsOK - 1,  HoraFin = CONVERT(DATETIME, '1899-12-30 " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "', 102) WHERE (IDENTITYCOL = " & NoRegistro & " )"
'                              Else
'                                    NoRegistro = 0
'                              End If
'                              CNN.rsCmdBuscaJcProgramadaMismoCarril.Close
'
'                              Me.LblVariables.Caption = Me.LblVariables.Caption & " Sensor: " & Sensor & " IDJC: - " & Resp & " NoRegistro: " & NoRegistro
'
'                  End If
            
        
        IdTM = 0
        'Me.TxtTM(2).BackColor = vbWhite
        
     
Else
        'Nueva Operacion
        'Tbl:
        CNN.CmdInsertDetOM (IdJC), (IdOperacion), (Ticket), (IdHorario), (IdMaquina)
        If CNN.rsCmdInsertDetOM.EOF = True Then
            CNN.rsCmdInsertDetOM.AddNew
                CNN.rsCmdInsertDetOM!Oma_Id = IdJC
                CNN.rsCmdInsertDetOM!IdOperacion = IdOperacion
                CNN.rsCmdInsertDetOM!IdMaquina = IdMaquina
                CNN.rsCmdInsertDetOM!TicketGem = Ticket
                CNN.rsCmdInsertDetOM!HoraIni = Time
                CNN.rsCmdInsertDetOM!HoraFin = Time
                
                'Busca PackingList
                CNN.CmdBuscaPKL_Programa Ticket, CodLinea
                
                If CNN.rsCmdBuscaPKL_Programa.EOF <> True Then
                    Cantidad = CNN.rsCmdBuscaPKL_Programa!Cantidad
                    CNN.rsCmdInsertDetOM!CodVidrio = CNN.rsCmdBuscaPKL_Programa!Codigo
                    
                    CNN.CmdBuscaOF (CNN.rsCmdBuscaPKL_Programa!OF)
                    If CNN.rsCmdBuscaOF.EOF <> True Then
                        NoParte = CNN.rsCmdBuscaOF!nodeparte
                    Else
                        NoParte = CNN.rsCmdBuscaPKL_Programa!OF
                    End If
                    CNN.rsCmdBuscaOF.Close
                    ''''Me.LblLinea.Caption = "Linea: " & CNN.rsCmdBuscaPKL_Programa!descripcion
                    
                    
                    'Piezas por hoja
                    '
                    OF = CNN.rsCmdBuscaPKL_Programa!OF
                    
                    'Orden_Man Vs Lite Vs Det_PackLst
                    CNN.CmdPxH (Ticket), (OF), (IdJC)
                    If CNN.rsCmdPxH.EOF <> True Then
                        PxH = CNN.rsCmdPxH!nopzas
                    Else
                        CNN.CmdPxHat (OF), (Ticket)
                        If CNN.rsCmdPxHAT.EOF <> True Then
                            PxH = CNN.rsCmdPxHAT!nopzas
                        Else
                            PxH = 0
                        End If
                        CNN.rsCmdPxHAT.Close
                    End If
                    CNN.rsCmdPxH.Close
                Else
                        CNN.rsCmdBuscaPKL_Programa.Close
                  '      CNN.rsCmdInsertDetOM.Close
                        CNN.rsCmdUpdateConteos.Close
                        MsgBox "El Ticket NO esta programado para esta hora.", vbCritical, MSG
                        'Exit Sub
                End If
                CNN.rsCmdBuscaPKL_Programa.Close
                
                CNN.rsCmdInsertDetOM!PxH = PxH
                CNN.rsCmdInsertDetOM!PzsPT = PxH * Cantidad
                CNN.rsCmdInsertDetOM!LaminasMP = Cantidad
                
                'If IdOperacion = 1 Then
                CNN.rsCmdInsertDetOM!PzsOK = TotPz                 '(PxH / 2)
                'ElseIf IdOperacion = 32 Then
                           ' If (PxH Mod 2) = 1 Then     ' Es par  o es Non
                               ' CNN.rsCmdInsertDetOM!PzsOK = TotPz
                           ' Else
                              '  CNN.rsCmdInsertDetOM!PzsOK = TotPz ' (PxH / 2)
                            'End If
                ' ElseIf IdOperacion = 30 Then
                      'CNN.rsCmdInsertDetOM!PzsOK = TotPz
                ' End If
                
                CNN.rsCmdInsertDetOM!PzsScrap = 0
                CNN.rsCmdInsertDetOM!MinsTM = 0
                CNN.rsCmdInsertDetOM!EmpleadoId = EmpleadoId '011-056"
                CNN.rsCmdInsertDetOM!FechaCap = Fecha 'Date
                CNN.rsCmdInsertDetOM!HoraCap = Time
                CNN.rsCmdInsertDetOM!Observaciones = Observaciones
                CNN.rsCmdInsertDetOM!IdHorario = IdHorario
                CNN.rsCmdInsertDetOM!CodLinea = CodLinea
            CNN.rsCmdInsertDetOM.Update
        Else
            'MsgBox "Ya hay una captura para esta operacion y este Ticket.", vbCritical, MSG
        End If
        CNN.rsCmdInsertDetOM.Close
End If
CNN.rsCmdUpdateConteos.Close

'ErrHandler:
'If Err.Number = -2147467259 Then
    
    'Call IsWebConnected(MSG)
    
'End If


End Sub




