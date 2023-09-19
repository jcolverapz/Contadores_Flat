Private Sub Form_Load()
Me.LstReq.MousePointer = 14
Me.Label3.Caption = Me.Label3.Caption & " del " & Format(F_Ini, "dddd, d mmm yyyy") & " al " & Format(F_Fin, "dddd, d mmm yyyy") & " "
Me.LstReq.ListItems.Clear
Columna = 10
MaxCol = (F_FinR - F_IniR) * 2 + 10
Columna = 1
Renglon = 1
ii = 1

For j = 0 To 500
    VDifV(j) = 0
    VDifVMPS(j) = 0
Next j
Band = False
SumPzs = 0

Me.LstReq.ListItems.Add (ii)
Me.LstReq.ListItems.Item(ii).SubItems(12) = "*"
Do While Renglon <= MaxRenglon
    'Busca Flat
    OF = FrmPed.FgdPedidos.TextMatrix(Renglon, 3)
    NoParte = FrmPed.FgdPedidos.TextMatrix(Renglon, 4)
    If NoParte = "240350106" Then '  Or NoParte = "240350605" Or NoParte = "240350613" Then
        NoParte = NoParte
    End If
    If OF = "134701900" Then '  Or NoParte = "240350605" Or NoParte = "240350613" Then
        NoParte = NoParte
    End If
    CNN.CmdBuscaFlat (OF)
    If CNN.rsCmdBuscaFlat.EOF = True Then
        Columna = MaxCol
    Else
        CantidadPzs = 0
        Do While Columna <= MaxCol
            BandPrimero = False
            If Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)) < 0 And Columna > 9 Then
                'Cantidad = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                NombreCorto = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
                Dia = Weekday(CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1)))
                ' 6 dias de anticipo para la produccion
                F_Prod = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                F_Ped = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                If NombreCorto = "Electrolux" Then
                    Select Case Dia
                        Case 1 'Domingo
                            F_Prod = CDate(CDbl(F_Prod) - 10)
                        Case 2 'Lunes
                            F_Prod = CDate(CDbl(F_Prod) - 10)
                        Case 3 'Martes
                            F_Prod = CDate(CDbl(F_Prod) - 8)
                        Case 4 'Miercoles
                            F_Prod = CDate(CDbl(F_Prod) - 8)
                        Case 5 'Jueves
                            F_Prod = CDate(CDbl(F_Prod) - 8)
                        Case 6 'Viernes
                            F_Prod = CDate(CDbl(F_Prod) - 8)
                        Case 7 'Sabado
                            F_Prod = CDate(CDbl(F_Prod) - 8)
                    End Select
                Else
                    Select Case Dia
                        Case 1 'Domingo
                            F_Prod = CDate(CDbl(F_Prod) - 3)
                        Case 2 'Lunes
                            F_Prod = CDate(CDbl(F_Prod) - 4)
                        Case 3 'Martes
                            F_Prod = CDate(CDbl(F_Prod) - 3)
                        Case 4 'Miercoles
                            F_Prod = CDate(CDbl(F_Prod) - 2)
                        Case 5 'Jueves
                            F_Prod = CDate(CDbl(F_Prod) - 2)
                        Case 6 'Viernes
                            F_Prod = CDate(CDbl(F_Prod) - 2)
                        Case 7 'Sabado
                            F_Prod = CDate(CDbl(F_Prod) - 2)
                    End Select
                End If
                If NoParte = "134701900" Then ''Or NoParte = "240350605" Or NoParte = "240350613" Then  'If NoParte = "2179259" Or NoParte = "2216108" Or NoParte = "4890JL1002U" Then
                    NoParte = NoParte
                End If
                SumPzs = 0      '' Fecha = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                    If OF = "M21238" Then
                        OF = OF
                    End If
                    
                    CNN.CmdPed2 (F_Ped), (OF)
                         If CNN.rsCmdPed2.EOF <> True Then
                             Do While CNN.rsCmdPed2.EOF <> True
                                 If ((F_Prod >= F_Ini) And (F_Prod <= F_Fin)) Then
                                     'If VDifV(Renglon) = 0 Then
                                      ' es   Dif <= Pedido
                                     Me.LstReq.ListItems.Item(ii).SubItems(1) = ii
                                     Me.LstReq.ListItems.Item(ii).SubItems(2) = FrmPed.FgdPedidos.TextMatrix(Renglon, 1)
                                     Me.LstReq.ListItems.Item(ii).SubItems(3) = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
                                     Me.LstReq.ListItems.Item(ii).SubItems(4) = OF
                                     Me.LstReq.ListItems.Item(ii).SubItems(5) = NoParte
                                     Me.LstReq.ListItems.Item(ii).SubItems(6) = CNN.rsCmdBuscaFlat!Descripcion
                                     Me.LstReq.ListItems.Item(ii).SubItems(12) = Me.LstReq.ListItems.Item(ii).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
                                     
                                     If OF = "M21539" Or OF = "M21487" Then
                                        OF = OF
                                     End If
                                     
                                     CNN.CmdLiteV (OF)   'Busca PxH
                                     If CNN.rsCmdLiteV.EOF <> True Then
                                         PxH = CNN.rsCmdLiteV!NoPzas
                                         If CNN.rsCmdLiteV!NoPzas > 0 Then
                                             Cantidad = CantidadPzs / PxH
                                             Codigo = CNN.rsCmdLiteV!CodVidrio
                                             Me.LstReq.ListItems.Item(ii).SubItems(9) = Codigo
                                             Me.LstReq.ListItems.Item(ii).SubItems(10) = PxH
                                             Me.LstReq.ListItems.Item(ii).SubItems(11) = ""
                                         Else
                                             Me.LstReq.ListItems.Item(ii).SubItems(9) = "Sin Def."
                                             Me.LstReq.ListItems.Item(ii).SubItems(10) = 0
                                             CNN.rsCmdCorreo.Open
                                                 CNN.rsCmdCorreo.AddNew
                                                     CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                     CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                     CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
                                                     CNN.rsCmdCorreo!CCO = ""
                                                     CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Def. en Lite--> " & NoParte '& " <<Pruebas TI>>"
                                                     CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice la Materia Prima para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                     CNN.rsCmdCorreo!aplicacion = "JC"
                                                     CNN.rsCmdCorreo!Prioridad = 1
                                                 CNN.rsCmdCorreo.Update
                                             CNN.rsCmdCorreo.Close
                                         End If
                                     End If
                                     CNN.rsCmdLiteV.Close
                                     Me.LstReq.ListItems.Item(ii).SubItems(8) = ""
                                     CNN.CmdStdPackFlat (OF)
                                        If CNN.rsCmdStdPackFlat.EOF <> True Then
                                            Do While CNN.rsCmdStdPackFlat.EOF <> True
                                                Me.LstReq.ListItems.Item(ii).SubItems(8) = Me.LstReq.ListItems.Item(ii).SubItems(8) & CNN.rsCmdStdPackFlat!Cantidad & ", "
                                                CNN.rsCmdStdPackFlat.MoveNext
                                            Loop
                                        Else
                                            Me.LstReq.ListItems.Item(ii).SubItems(8) = "Sin Def"
                                            CNN.rsCmdCorreo.Open      'manda correos
                                                CNN.rsCmdCorreo.AddNew
                                                    CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
                                                    CNN.rsCmdCorreo!CCO = ""
                                                    CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Standar Pack --> " & NoParte '& " <<Pruebas TI>>"
                                                    CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice el Standar Pack para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                    CNN.rsCmdCorreo!aplicacion = "JC"
                                                    CNN.rsCmdCorreo!Prioridad = 1
                                                    CNN.rsCmdCorreo!enviado = "S"
                                                CNN.rsCmdCorreo.Update
                                            CNN.rsCmdCorreo.Close
                                        End If
                                        CNN.rsCmdStdPackFlat.Close
                                        ' entonces es el primero
                                     If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
                                        VDifV(Renglon) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                        CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
                                        Me.LstReq.ListItems.Item(ii).SubItems(7) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
                                     Else
                                        If Me.LstReq.ListItems.Item(ii).SubItems(7) = "" Then
                                            Me.LstReq.ListItems.Item(ii).SubItems(7) = 0
                                        End If
                                         CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                         Me.LstReq.ListItems.Item(ii).SubItems(7) = CLng(Me.LstReq.ListItems.Item(ii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                         '''Me.LstReq.ListItems.Item(ii).SubItems(12) = Me.LstReq.ListItems.Item(ii).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
                                     End If
                                 End If
                                 CNN.rsCmdPed2.MoveNext
                             Loop
                         Else
                             Band = True
                         End If
                         CNN.rsCmdPed2.Close
            End If
            Columna = Columna + 2
        Loop
        Columna = 1
    End If
    CNN.rsCmdBuscaFlat.Close
    'Limpia cadena en grid
    If Band = True And ii <> 0 Then
        If Me.LstReq.ListItems.Item(ii).SubItems(12) = "*" Then
            'Me.LstReq.ListItems.Item(ii).SubItems(7) = "*"
            Me.LstReq.ListItems.Remove (ii)
            ii = ii - 1
        End If
    End If
    Band = False
    
    
    Columna = 1
    
    
    If ii <> 0 Then
        If Me.LstReq.ListItems.Item(ii).SubItems(7) = "" Then
            Me.LstReq.ListItems.Item(ii).SubItems(7) = 0
        End If
        If Me.LstReq.ListItems.Item(ii).SubItems(7) = 0 Then
            Me.LstReq.ListItems.Remove (ii)
            ii = ii - 1
        End If

    End If
    If ii >= 1 Then
        If (Left(Me.LstReq.ListItems.Item(ii).SubItems(12), 1) <> "*" And Len(Me.LstReq.ListItems.Item(ii).SubItems(12)) > 1) Then
            Me.LstReq.ListItems.Item(ii).SubItems(12) = "*" & Me.LstReq.ListItems.Item(ii).SubItems(12)
        End If
    End If


    Renglon = Renglon + 1
    ii = ii + 1
    Me.LstReq.ListItems.Add (ii)

    SumPzs = 0
Loop
Columna = Columna
m = Renglon - 1







''''Desde aqui modificacion para Vidrio Circular agregado en almacen de traspaso
'''' Copia de seguridad mas abajo

MaxCol = (F_FinR - F_IniR) * 2 + 10
Columna = 1
Renglon = 1
iii = 1
Maxiii = iii
Me.LstReqAT.ListItems.Clear

''''Vidrio MPS
Me.LstReqAT.ListItems.Add (iii)
Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*"

FrmFechasJC.PBar5.Max = MaxRenglon
Do While Renglon <= MaxRenglon
    'Busca Vidrio MPS
    OF = FrmPed.FgdPedidos.TextMatrix(Renglon, 3)
    NoParte = FrmPed.FgdPedidos.TextMatrix(Renglon, 4)
    If Renglon = 62 Then
        OF = OF
    End If
    
    If NoParte = "134701900" Then     'Or NoParte = "240350605" Or NoParte = "240350613" Then
        NoParte = NoParte
    End If
       
    '''AHT72913601   1403 y 1404    E21386
    '''AHT72913602   1403 y 1405    E21387
    '''AHT33603801  2096 y 1096     E21319
     
    If OF = "E21386" Or OF = "E21387" Or OF = "E21319" Then
        OF = OF
    End If
    
    If OF = "E88906" Then
        OF = OF
    End If
    
    CNN.CmdBuscaV_MPS (OF)
    If CNN.rsCmdBuscaV_MPS.EOF = True Then
        Columna = MaxCol
    Else
        NoVidriosMPS = 1
        Do While CNN.rsCmdBuscaV_MPS.EOF <> True
            Resp = CNN.rsCmdBuscaV_MPS.RecordCount
            CantidadPzs = 0
            Do While Columna <= MaxCol
                BandPrimero = False
                If Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)) < 0 And Columna > 9 Then
                    NombreCorto = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
                    Dia = Weekday(CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1)))
                    ' 6 dias de anticipo para la produccion
                    F_Prod = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                    F_Ped = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                    If NombreCorto = "Electrolux" Then
                        Select Case Dia
                            Case 1 'Domingo
                                F_Prod = CDate(CDbl(F_Prod) - 10)
                            Case 2 'Lunes
                                F_Prod = CDate(CDbl(F_Prod) - 10)
                            Case 3 'Martes
                                F_Prod = CDate(CDbl(F_Prod) - 8)
                            Case 4 'Miercoles
                                F_Prod = CDate(CDbl(F_Prod) - 8)
                            Case 5 'Jueves
                                F_Prod = CDate(CDbl(F_Prod) - 8)
                            Case 6 'Viernes
                                F_Prod = CDate(CDbl(F_Prod) - 8)
                            Case 7 'Sabado
                                F_Prod = CDate(CDbl(F_Prod) - 8)
                        End Select
                    Else
                        Select Case Dia
                            Case 1 'Domingo
                                F_Prod = CDate(CDbl(F_Prod) - 3)
                            Case 2 'Lunes
                                F_Prod = CDate(CDbl(F_Prod) - 4)
                            Case 3 'Martes
                                F_Prod = CDate(CDbl(F_Prod) - 3)
                            Case 4 'Miercoles
                                F_Prod = CDate(CDbl(F_Prod) - 2)
                            Case 5 'Jueves
                                F_Prod = CDate(CDbl(F_Prod) - 2)
                            Case 6 'Viernes
                                F_Prod = CDate(CDbl(F_Prod) - 2)
                            Case 7 'Sabado
                                F_Prod = CDate(CDbl(F_Prod) - 2)
                        End Select
                    End If
                    If NoParte = "242235502" Then ''Or NoParte = "240350605" Or NoParte = "240350613" Then  'If NoParte = "2179259" Or NoParte = "2216108" Or NoParte = "4890JL1002U" Then
                        NoParte = NoParte
                    End If
                    SumPzs = 0      '' Fecha = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
                    CNN.Cmd_IsLaundry (OF)
                    If CNN.rsCmd_IsLaundry.EOF = True Then
                        CNN.CmdPed2 (F_Ped), (OF)
                        If CNN.rsCmdPed2.EOF <> True Then
                            Do While CNN.rsCmdPed2.EOF <> True
                                If ((F_Prod >= F_Ini) And (F_Prod <= F_Fin)) Then
                                    n = 1
                                    Band = False
                                    Do While Len(Me.LstReqAT.ListItems.Item(n).SubItems(5)) <> 0
                                       If CNN.rsCmdBuscaV_MPS!Codigo = Me.LstReqAT.ListItems.Item(n).SubItems(5) Then
                                         Band = True
                                         Exit Do
                                       End If
                                       n = n + 1
                                    Loop
                                    If Band = True Then     'Ya existe el Vidio MPS en la lista
                                        Me.LstReqAT.ListItems.Item(n).SubItems(12) = Me.LstReqAT.ListItems.Item(n).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
                                           ' entonces es el primero
                                        If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
                                           VDifV(n) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                           CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
                                           Me.LstReqAT.ListItems.Item(n).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
                                        Else
                                           If Me.LstReqAT.ListItems.Item(n).SubItems(7) = "" Then
                                               Me.LstReqAT.ListItems.Item(n).SubItems(7) = 0
                                           End If
                                            CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                            Me.LstReqAT.ListItems.Item(n).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                        End If
                                    Else    'Es nuevo no existe en la lista
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(1) = iii
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(2) = FrmPed.FgdPedidos.TextMatrix(Renglon, 1)
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(3) = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(4) = CNN.rsCmdBuscaV_MPS!Codigo  'NoParte 'OF
                                        ''Me.LstReqAT.ListItems.Item(iii).SubItems(5) = NoParte
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(5) = CNN.rsCmdBuscaV_MPS!Codigo
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(6) = CNN.rsCmdBuscaV_MPS!Descripcion
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(12) = Me.LstReqAT.ListItems.Item(iii).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
                                        If Me.LstReqAT.ListItems.Item(iii).SubItems(14) <> NoParte & ", " Then
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(14) = Me.LstReqAT.ListItems.Item(iii).SubItems(14) & NoParte & ", "
                                        End If
                                        
                                        CNN.CmdLiteV (OF)   'Busca PxH
                                        If CNN.rsCmdLiteV.EOF <> True Then
                                            PxH = CNN.rsCmdLiteV!NoPzas
                                            
                                          '  Resp = CNN.rsCmdLiteV!t
                                            
                                            If CNN.rsCmdLiteV!NoPzas > 0 Then
                                                Cantidad = CantidadPzs / PxH
                                                Codigo = CNN.rsCmdLiteV!CodVidrio
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(9) = Codigo
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(10) = PxH
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(11) = ""
                                            Else
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(9) = "Sin Def."
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(10) = 0
                                                CNN.rsCmdCorreo.Open
                                                    CNN.rsCmdCorreo.AddNew
                                                        CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                        CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                        CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" '
                                                        CNN.rsCmdCorreo!CCO = "helios.ireta @"
                                                        CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Def. en Lite--> " & NoParte '& " <<Pruebas TI>>"
                                                        CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice la Materia Prima para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                        CNN.rsCmdCorreo!aplicacion = "JC"
                                                        CNN.rsCmdCorreo!Prioridad = 1
                                                    CNN.rsCmdCorreo.Update
                                                CNN.rsCmdCorreo.Close
                                            End If
                                        End If
                                        CNN.rsCmdLiteV.Close
                                        
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(8) = ""
                                        
                                        '''AHT72913601   1403 y 1404    E21386
                                        '''AHT72913602   1403 y 1405    E21387
                                        '''AHT33603801  2096 y 1096     E21319
                                        
                                        
                                        
                                       If OF = "E88906" Then
                                            OF = OF
                                        End If
                                        
                                        CNN.CmdStdPackFlat (OF)
                                        If CNN.rsCmdStdPackFlat.EOF <> True Then
                                            Do While CNN.rsCmdStdPackFlat.EOF <> True
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(8) = Me.LstReqAT.ListItems.Item(iii).SubItems(8) & CNN.rsCmdStdPackFlat!Cantidad & ", "
                                                CNN.rsCmdStdPackFlat.MoveNext
                                            Loop
                                        Else
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(8) = "Sin Def"
                                            CNN.rsCmdCorreo.Open      'manda correos
                                                CNN.rsCmdCorreo.AddNew
                                                    CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
                                                    CNN.rsCmdCorreo!CCO = ""
                                                    CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Standar Pack --> " & NoParte '& " <<Pruebas TI>>"
                                                    CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice el Standar Pack para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                    CNN.rsCmdCorreo!aplicacion = "JC"
                                                    CNN.rsCmdCorreo!Prioridad = 1
                                                    CNN.rsCmdCorreo!enviado = "N"
                                                CNN.rsCmdCorreo.Update
                                            CNN.rsCmdCorreo.Close
                                        End If
                                        CNN.rsCmdStdPackFlat.Close
                                        
                                        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
                                        End If
                                           ' entonces es el primero
                                        If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
                                           VDifV(Renglon) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                           CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
                                           Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
                                        Else
                                            CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                        End If
                                    End If
                                End If
                                CNN.rsCmdPed2.MoveNext
                            Loop
                        Else
                            Band = True
                        End If
                        CNN.rsCmdPed2.Close
                    Else
                        '''Is Laundry
                        CNN.CmdPed3 (F_Ped), (OF)
                        If CNN.rsCmdPed3.EOF <> True Then
                            Do While CNN.rsCmdPed3.EOF <> True
                                If ((F_Prod >= F_Ini) And (F_Prod <= F_Fin)) Then
                                    n = 1
                                    Band = False
                                    Do While Len(Me.LstReqAT.ListItems.Item(n).SubItems(5)) <> 0
                                       If CNN.rsCmdBuscaV_MPS!Codigo = Me.LstReqAT.ListItems.Item(n).SubItems(5) Then
                                         Band = True
                                         Exit Do
                                       End If
                                       n = n + 1
                                    Loop
                                    If Band = True Then     'Ya existe el Vidio MPS en la lista
                                        Me.LstReqAT.ListItems.Item(n).SubItems(12) = Me.LstReqAT.ListItems.Item(n).SubItems(12) & CNN.rsCmdPed3!IdPedido & "*"
                                           ' entonces es el primero
                                        If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
                                           VDifV(n) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                           CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
                                           Me.LstReqAT.ListItems.Item(n).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
                                        Else
                                           If Me.LstReqAT.ListItems.Item(n).SubItems(7) = "" Then
                                               Me.LstReqAT.ListItems.Item(n).SubItems(7) = 0
                                           End If
                                            CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                            Me.LstReqAT.ListItems.Item(n).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                        End If
                                    Else    'Es nuevo no existe en la lista
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(1) = iii
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(2) = FrmPed.FgdPedidos.TextMatrix(Renglon, 1)
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(3) = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(4) = CNN.rsCmdBuscaV_MPS!Codigo  'NoParte 'OF
                                        ''Me.LstReqAT.ListItems.Item(iii).SubItems(5) = NoParte
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(5) = CNN.rsCmdBuscaV_MPS!Codigo
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(6) = CNN.rsCmdBuscaV_MPS!Descripcion
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(12) = Me.LstReqAT.ListItems.Item(iii).SubItems(12) & CNN.rsCmdPed3!IdPedido & "*"
                                        
                                        If Me.LstReqAT.ListItems.Item(iii).SubItems(14) <> (NoParte & ", ") Then
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(14) = Me.LstReqAT.ListItems.Item(iii).SubItems(14) & NoParte & ", "
                                        End If
                                        
                                        
                                        CNN.CmdLiteV (OF)   'Busca PxH
                                        If CNN.rsCmdLiteV.EOF <> True Then
                                            PxH = CNN.rsCmdLiteV!NoPzas
                                            If CNN.rsCmdLiteV!NoPzas > 0 Then
                                                Cantidad = CantidadPzs / PxH
                                                Codigo = CNN.rsCmdLiteV!CodVidrio
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(9) = Codigo
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(10) = PxH
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(11) = ""
                                            Else
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(9) = "Sin Def."
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(10) = 0
                                                CNN.rsCmdCorreo.Open
                                                    CNN.rsCmdCorreo.AddNew
                                                        CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                        CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                        CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
                                                        CNN.rsCmdCorreo!CCO = ""
                                                        CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Def. en Lite--> " & NoParte '& " <<Pruebas TI>>"
                                                        CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice la Materia Prima para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                        CNN.rsCmdCorreo!aplicacion = "JC"
                                                        CNN.rsCmdCorreo!Prioridad = 1
                                                    CNN.rsCmdCorreo.Update
                                                CNN.rsCmdCorreo.Close
                                            End If
                                        End If
                                        CNN.rsCmdLiteV.Close
                                        
                                        Me.LstReqAT.ListItems.Item(iii).SubItems(8) = ""
                                        
                                       If OF = "e21359" Or OF = "M21397" Then
                                            OF = OF
                                        End If
                                        
                                        CNN.CmdStdPackFlat (OF)
                                        If CNN.rsCmdStdPackFlat.EOF <> True Then
                                            Do While CNN.rsCmdStdPackFlat.EOF <> True
                                                Me.LstReqAT.ListItems.Item(iii).SubItems(8) = Me.LstReqAT.ListItems.Item(iii).SubItems(8) & CNN.rsCmdStdPackFlat!Cantidad & ", "
                                                CNN.rsCmdStdPackFlat.MoveNext
                                            Loop
                                        Else
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(8) = "Sin Def"
                                            CNN.rsCmdCorreo.Open      'manda correos
                                                CNN.rsCmdCorreo.AddNew
                                                    CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
                                                    CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
                                                    CNN.rsCmdCorreo!CCO = ""
                                                    CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Standar Pack --> " & NoParte '& " <<Pruebas TI>>"
                                                    CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice el Standar Pack para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
                                                    CNN.rsCmdCorreo!aplicacion = "JC"
                                                    CNN.rsCmdCorreo!Prioridad = 1
                                                    CNN.rsCmdCorreo!enviado = "N"
                                                CNN.rsCmdCorreo.Update
                                            CNN.rsCmdCorreo.Close
                                        End If
                                        CNN.rsCmdStdPackFlat.Close
                                        
                                        
                                        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
                                        End If
                                           ' entonces es el primero
                                        If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
                                           VDifV(Renglon) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                           CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
                                           Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
                                        Else
                                            CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
                                        End If
                                    End If
                                End If
                                CNN.rsCmdPed3.MoveNext
                            Loop
                        Else
                            Band = True
                        End If
                        CNN.rsCmdPed3.Close
                        
                    End If
                    CNN.rsCmd_IsLaundry.Close
                End If
                Columna = Columna + 2
            Loop
    ''        Columna = 1
            CNN.rsCmdBuscaV_MPS.MoveNext
            
            NoVidriosMPS = NoVidriosMPS + 1
            
            If NoVidriosMPS <= CNN.rsCmdBuscaV_MPS.RecordCount And CNN.rsCmdBuscaV_MPS.RecordCount > 1 Then
                Columna = 1
                
                'Limpia cadena en grid
                If Band = True And iii <> 0 Then
                    If Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" Then
                        'Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "*"
                        Me.LstReqAT.ListItems.Remove (iii)
                        iii = iii - 1
                    End If
                End If
                Band = False
            
                If iii <> 0 Then
                    If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
                        Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
                    End If
                    If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0 Then
                        Me.LstReqAT.ListItems.Remove (iii)
                        iii = iii - 1
                    End If
                End If
                If iii >= 1 Then
                    If (Left(Me.LstReqAT.ListItems.Item(iii).SubItems(12), 1) <> "*" And Len(Me.LstReqAT.ListItems.Item(iii).SubItems(12)) > 1) Then
                        Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" & Me.LstReqAT.ListItems.Item(iii).SubItems(12)
                    End If
                End If
                iii = iii + 1
                Me.LstReqAT.ListItems.Add (iii)
                
            End If
        Loop
    End If
    Columna = 1
    CNN.rsCmdBuscaV_MPS.Close
    'Limpia cadena en grid
    If Band = True And iii <> 0 Then
        If Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" Then
            'Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "*"
            Me.LstReqAT.ListItems.Remove (iii)
            iii = iii - 1
        End If
    End If
    Band = False

    If iii <> 0 Then
        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
        End If
        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0 Then
            Me.LstReqAT.ListItems.Remove (iii)
            iii = iii - 1
        End If
    End If
    If iii >= 1 Then
        If (Left(Me.LstReqAT.ListItems.Item(iii).SubItems(12), 1) <> "*" And Len(Me.LstReqAT.ListItems.Item(iii).SubItems(12)) > 1) Then
            Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" & Me.LstReqAT.ListItems.Item(iii).SubItems(12)
        End If
    End If
    Renglon = Renglon + 1
    iii = iii + 1
    Me.LstReqAT.ListItems.Add (iii)
    SumPzs = 0
    FrmFechasJC.PBar5.Value = Renglon - 1
Loop
Columna = Columna




''''Copia de seguridad
'''''''''MaxCol = (F_FinR - F_IniR) * 2 + 10
'''''''''Columna = 1
'''''''''Renglon = 1
'''''''''iii = 1
'''''''''Maxiii = iii
'''''''''Me.LstReqAT.ListItems.Clear
'''''''''
'''''''''''''Vidrio MPS
'''''''''Me.LstReqAT.ListItems.Add (iii)
'''''''''Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*"
'''''''''Do While Renglon <= MaxRenglon
'''''''''    'Busca Vidrio MPS
'''''''''    OF = FrmPed.FgdPedidos.TextMatrix(Renglon, 3)
'''''''''    NoParte = FrmPed.FgdPedidos.TextMatrix(Renglon, 4)
'''''''''    If Renglon = 62 Then
'''''''''        OF = OF
'''''''''    End If
'''''''''
'''''''''    If NoParte = "240350106" Then     'Or NoParte = "240350605" Or NoParte = "240350613" Then
'''''''''        NoParte = NoParte
'''''''''    End If
'''''''''
'''''''''    CNN.CmdBuscaV_MPS (OF)
'''''''''    If CNN.rsCmdBuscaV_MPS.EOF = True Then
'''''''''        Columna = MaxCol
'''''''''    Else
'''''''''        CantidadPzs = 0
'''''''''        Do While Columna <= MaxCol
'''''''''            BandPrimero = False
'''''''''            If Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)) < 0 And Columna > 9 Then
'''''''''                NombreCorto = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
'''''''''                Dia = Weekday(CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1)))
'''''''''                ' 6 dias de anticipo para la produccion
'''''''''                F_Prod = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
'''''''''                F_Ped = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
'''''''''                If NombreCorto = "Electrolux" Then
'''''''''                    Select Case Dia
'''''''''                        Case 1 'Domingo
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 10)
'''''''''                        Case 2 'Lunes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 10)
'''''''''                        Case 3 'Martes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 8)
'''''''''                        Case 4 'Miercoles
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 8)
'''''''''                        Case 5 'Jueves
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 8)
'''''''''                        Case 6 'Viernes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 8)
'''''''''                        Case 7 'Sabado
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 8)
'''''''''                    End Select
'''''''''                Else
'''''''''                    Select Case Dia
'''''''''                        Case 1 'Domingo
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 3)
'''''''''                        Case 2 'Lunes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 4)
'''''''''                        Case 3 'Martes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 3)
'''''''''                        Case 4 'Miercoles
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 2)
'''''''''                        Case 5 'Jueves
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 2)
'''''''''                        Case 6 'Viernes
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 2)
'''''''''                        Case 7 'Sabado
'''''''''                            F_Prod = CDate(CDbl(F_Prod) - 2)
'''''''''                    End Select
'''''''''                End If
'''''''''                If NoParte = "240350106" Then ''Or NoParte = "240350605" Or NoParte = "240350613" Then  'If NoParte = "2179259" Or NoParte = "2216108" Or NoParte = "4890JL1002U" Then
'''''''''                    NoParte = NoParte
'''''''''                End If
'''''''''                SumPzs = 0      '' Fecha = CDate(FrmPed.FgdPedidos.TextMatrix(0, Columna - 1))
'''''''''                CNN.CmdPed2 (F_Ped), (OF)
'''''''''                If CNN.rsCmdPed2.EOF <> True Then
'''''''''                    Do While CNN.rsCmdPed2.EOF <> True
'''''''''                        If ((F_Prod >= F_Ini) And (F_Prod <= F_Fin)) Then
'''''''''                            n = 1
'''''''''                            Band = False
'''''''''                            Do While Len(Me.LstReqAT.ListItems.Item(n).SubItems(5)) <> 0
'''''''''                               If CNN.rsCmdBuscaV_MPS!Codigo = Me.LstReqAT.ListItems.Item(n).SubItems(5) Then
'''''''''                                 Band = True
'''''''''                                 Exit Do
'''''''''                               End If
'''''''''                               n = n + 1
'''''''''                            Loop
'''''''''                            If Band = True Then     'Ya existe el Vidio MPS en la lista
'''''''''                                Me.LstReqAT.ListItems.Item(n).SubItems(12) = Me.LstReqAT.ListItems.Item(n).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
'''''''''                                   ' entonces es el primero
'''''''''                                If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
'''''''''                                   VDifV(n) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                   CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
'''''''''                                   Me.LstReqAT.ListItems.Item(n).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
'''''''''                                Else
'''''''''                                   If Me.LstReqAT.ListItems.Item(n).SubItems(7) = "" Then
'''''''''                                       Me.LstReqAT.ListItems.Item(n).SubItems(7) = 0
'''''''''                                   End If
'''''''''                                    CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                    Me.LstReqAT.ListItems.Item(n).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(n).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                End If
'''''''''                            Else    'Es nuevo no existe en la lista
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(1) = iii
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(2) = FrmPed.FgdPedidos.TextMatrix(Renglon, 1)
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(3) = FrmPed.FgdPedidos.TextMatrix(Renglon, 2)
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(4) = CNN.rsCmdBuscaV_MPS!Codigo  'NoParte 'OF
'''''''''                                ''Me.LstReqAT.ListItems.Item(iii).SubItems(5) = NoParte
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(5) = CNN.rsCmdBuscaV_MPS!Codigo
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(6) = CNN.rsCmdBuscaV_MPS!Descripcion
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(12) = Me.LstReqAT.ListItems.Item(iii).SubItems(12) & CNN.rsCmdPed2!IdPedido & "*"
'''''''''                                CNN.CmdLiteV (OF)   'Busca PxH
'''''''''                                If CNN.rsCmdLiteV.EOF <> True Then
'''''''''                                    PxH = CNN.rsCmdLiteV!NoPzas
'''''''''                                    If CNN.rsCmdLiteV!NoPzas > 0 Then
'''''''''                                        Cantidad = CantidadPzs / PxH
'''''''''                                        Codigo = CNN.rsCmdLiteV!codvidrio
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(9) = Codigo
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(10) = PxH
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(11) = ""
'''''''''                                    Else
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(9) = "Sin Def."
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(10) = 0
'''''''''                                        CNN.rsCmdCorreo.Open
'''''''''                                            CNN.rsCmdCorreo.AddNew
'''''''''                                                CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
'''''''''                                                CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
'''''''''                                                CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
'''''''''                                                CNN.rsCmdCorreo!CCO = ""
'''''''''                                                CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Def. en Lite--> " & NoParte '& " <<Pruebas TI>>"
'''''''''                                                CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice la Materia Prima para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
'''''''''                                                CNN.rsCmdCorreo!aplicacion = "JC"
'''''''''                                                CNN.rsCmdCorreo!Prioridad = 1
'''''''''                                            CNN.rsCmdCorreo.Update
'''''''''                                        CNN.rsCmdCorreo.Close
'''''''''                                    End If
'''''''''                                End If
'''''''''                                CNN.rsCmdLiteV.Close
'''''''''
'''''''''                                Me.LstReqAT.ListItems.Item(iii).SubItems(8) = ""
'''''''''
'''''''''                               If OF = "e21359" Or OF = "M21267" Then
'''''''''                                    OF = OF
'''''''''                                End If
'''''''''
'''''''''                                CNN.CmdStdPackFlat (OF)
'''''''''                                If CNN.rsCmdStdPackFlat.EOF <> True Then
'''''''''                                    Do While CNN.rsCmdStdPackFlat.EOF <> True
'''''''''                                        Me.LstReqAT.ListItems.Item(iii).SubItems(8) = Me.LstReqAT.ListItems.Item(iii).SubItems(8) & CNN.rsCmdStdPackFlat!Cantidad & ", "
'''''''''                                        CNN.rsCmdStdPackFlat.MoveNext
'''''''''                                    Loop
'''''''''                                Else
'''''''''                                    Me.LstReqAT.ListItems.Item(iii).SubItems(8) = "Sin Def"
'''''''''                                    CNN.rsCmdCorreo.Open      'manda correos
'''''''''                                        CNN.rsCmdCorreo.AddNew
'''''''''                                            CNN.rsCmdCorreo!De = "alejandro.gallegos@schott.com"
'''''''''                                            CNN.rsCmdCorreo!Para = "alejandro.gallegos@schott.com"
'''''''''                                            CNN.rsCmdCorreo!CC = "blas.reyes@schott.com, helios.ireta@schott.com" 'zahira.perez@schott.com, fernando.sotomayor@schott.com"
'''''''''                                            CNN.rsCmdCorreo!CCO = ""
'''''''''                                            CNN.rsCmdCorreo!Titulo = "Producto Terminado sin Standar Pack --> " & NoParte '& " <<Pruebas TI>>"
'''''''''                                            CNN.rsCmdCorreo!Mensaje = MSG & "<br><br>A la brevedad actualice el Standar Pack para este numero de parte : " & NoParte & "<br>Por esta razón NO será incluido en la elaboracion de Job Cards " & Date & "  <br><br>Gracias."
'''''''''                                            CNN.rsCmdCorreo!aplicacion = "JC"
'''''''''                                            CNN.rsCmdCorreo!Prioridad = 1
'''''''''                                            CNN.rsCmdCorreo!enviado = "N"
'''''''''                                        CNN.rsCmdCorreo.Update
'''''''''                                    CNN.rsCmdCorreo.Close
'''''''''                                End If
'''''''''                                CNN.rsCmdStdPackFlat.Close
'''''''''
'''''''''
'''''''''                                If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
'''''''''                                    Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
'''''''''                                End If
'''''''''                                   ' entonces es el primero
'''''''''                                If (Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) <= Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))) Then
'''''''''                                   VDifV(Renglon) = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                   CantidadPzs = Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna)))
'''''''''                                   Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CDbl(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna))) 'CantidadPzs
'''''''''                                Else
'''''''''                                    CantidadPzs = CantidadPzs + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                    Me.LstReqAT.ListItems.Item(iii).SubItems(7) = CLng(Me.LstReqAT.ListItems.Item(iii).SubItems(7)) + Abs(Val(FrmPed.FgdPedidos.TextMatrix(Renglon, Columna - 1)))
'''''''''                                End If
'''''''''                            End If
'''''''''                        End If
'''''''''                        CNN.rsCmdPed2.MoveNext
'''''''''                    Loop
'''''''''                Else
'''''''''                    Band = True
'''''''''                End If
'''''''''                CNN.rsCmdPed2.Close
'''''''''            End If
'''''''''            Columna = Columna + 2
'''''''''        Loop
'''''''''''        Columna = 1
'''''''''    End If
'''''''''    Columna = 1
'''''''''    CNN.rsCmdBuscaV_MPS.Close
'''''''''    'Limpia cadena en grid
'''''''''    If Band = True And iii <> 0 Then
'''''''''        If Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" Then
'''''''''            'Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "*"
'''''''''            Me.LstReqAT.ListItems.Remove (iii)
'''''''''            iii = iii - 1
'''''''''        End If
'''''''''    End If
'''''''''    Band = False
'''''''''
'''''''''    If iii <> 0 Then
'''''''''        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = "" Then
'''''''''            Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0
'''''''''        End If
'''''''''        If Me.LstReqAT.ListItems.Item(iii).SubItems(7) = 0 Then
'''''''''            Me.LstReqAT.ListItems.Remove (iii)
'''''''''            iii = iii - 1
'''''''''        End If
'''''''''
'''''''''    End If
'''''''''    If iii >= 1 Then
'''''''''        If (Left(Me.LstReqAT.ListItems.Item(iii).SubItems(12), 1) <> "*" And Len(Me.LstReqAT.ListItems.Item(iii).SubItems(12)) > 1) Then
'''''''''            Me.LstReqAT.ListItems.Item(iii).SubItems(12) = "*" & Me.LstReqAT.ListItems.Item(iii).SubItems(12)
'''''''''        End If
'''''''''    End If
'''''''''    Renglon = Renglon + 1
'''''''''    iii = iii + 1
'''''''''    Me.LstReqAT.ListItems.Add (iii)
'''''''''
'''''''''    SumPzs = 0
'''''''''Loop
'''''''''Columna = Columna















'Temporalmente se Actualizan las tablas
'''CNN.rsCmdUpdateReq.Open "UPDATE Vidrio SET mat_asig = 0"
'''CNN.rsCmdUpdateReq.Open "UPDATE det_packlist SET status = 'Activo', idjc = 0 where status = 'Reservado'"
'''
'''
'''CNN.rsCmdUpdateReq.Open "Update Orden_Man Set Oma_Status = 'Cancelada'"
Call LLenarJC
End Sub


