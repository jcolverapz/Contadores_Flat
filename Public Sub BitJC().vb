Public Sub BitJC()

    CNN.CmdBitJC (IdJC)
    If CNN.rsCmdBitJC.EOF = True Then
        CNN.rsCmdBitJC.AddNew
        ' CNN.rsCmdBitJC!IdBitJC = 0
        CNN.rsCmdBitJC!Oma_Id = IdJC
        CNN.rsCmdBitJC!Oma_Status = Status
        CNN.rsCmdBitJC!Fecha = Date
        CNN.rsCmdBitJC!Hora = Time
        CNN.rsCmdBitJC!PC = PC
        CNN.rsCmdBitJC!IdUsuario = EmpleadoId
        CNN.rsCmdBitJC.Update
    End If
    CNN.rsCmdBitJC.Close


End Sub

Public Sub AsignaJCPed()
Cant2 = 0
PzsFaltantes = SumPzs
If Tipo = "F" Then
    If Len(Me.LstReq.SelectedItem.ListSubItems(12).Text) <> 0 Then
        Cadena = Me.LstReq.SelectedItem.ListSubItems(12).Text
        i = 1
        Do While i < Len(Cadena)
            Cadena2 = i
            Resp = InStr((i + 1), Cadena, "*")
            IdPedido = CLng(Mid(Cadena, (i + 1), ((Resp - 1) - i)))
            i = Resp
            'RUTA:
            CNN.CmdEnc_Pedidos (IdPedido)
            If CNN.rsCmdEnc_Pedidos.EOF <> True Then
                    CNN.cmdDet_Pedidos (IdPedido), (1)
                    If CNN.rscmdDet_Pedidos.EOF <> True Then
                        Cant2 = Cant2 + CNN.rscmdDet_Pedidos!detp_Cant
                    End If
                    If PzsFaltantes > 0 Then
                        If CNN.rscmdDet_Pedidos!detp_cantpen >= PzsFaltantes Then
                            CNN.rscmdDet_Pedidos!detp_cantpen = (CNN.rscmdDet_Pedidos!detp_cantpen - PzsFaltantes)
                            PzsFaltantes = 0
                            i = Len(Cadena)
                        Else
                            PzsFaltantes = PzsFaltantes - CNN.rscmdDet_Pedidos!detp_cantpen
                            CNN.rscmdDet_Pedidos!detp_cantpen = 0 'CNN.rscmdDet_Pedidos!detp_Cant
                        End If
                    Else
                        i = Len(Cadena)
                    End If
                    CNN.rscmdDet_Pedidos.Update
                    CNN.rscmdDet_Pedidos.Close
                    CNN.rsCmdEnc_Pedidos!IdJC = IdJC
                    CNN.rsCmdEnc_Pedidos.Update
            Else
                MsgBox "No existe el pedido", vbInformation, MSG
            End If
            CNN.rsCmdEnc_Pedidos.Close
        Loop
    Else
        MsgBox "No hay pedidos", vbExclamation, MSG
    End If
End If
If Tipo = "AT" Then
    If Len(Me.LstReqAT.SelectedItem.ListSubItems(12).Text) <> 0 Then
        Cadena = Me.LstReqAT.SelectedItem.ListSubItems(12).Text
        i = 1
        Do While i < Len(Cadena)
            Cadena2 = i
            Resp = InStr((i + 1), Cadena, "*")
            IdPedido = CLng(Mid(Cadena, (i + 1), ((Resp - 1) - i)))
            i = Resp
            CNN.CmdEnc_Pedidos (IdPedido)
            If CNN.rsCmdEnc_Pedidos.EOF <> True Then
                CNN.cmdDet_Pedidos (IdPedido), (1)
                    If CNN.rscmdDet_Pedidos.EOF <> True Then
                        Cant2 = Cant2 + CNN.rscmdDet_Pedidos!detp_Cant
                    End If
                    If PzsFaltantes > 0 Then
                        If CNN.rscmdDet_Pedidos!detp_cantpen >= PzsFaltantes Then
                            CNN.rscmdDet_Pedidos!detp_cantpen = (CNN.rscmdDet_Pedidos!detp_cantpen - PzsFaltantes)
                            PzsFaltantes = 0
                            i = Len(Cadena)
                        Else
                            PzsFaltantes = PzsFaltantes - CNN.rscmdDet_Pedidos!detp_cantpen
                            CNN.rscmdDet_Pedidos!detp_cantpen = 0 'CNN.rscmdDet_Pedidos!detp_Cant
                        End If
                    Else
                        i = Len(Cadena)
                    End If
                    CNN.rscmdDet_Pedidos.Update
                    CNN.rscmdDet_Pedidos.Close
                    CNN.rsCmdEnc_Pedidos!IdJC = IdJC
                    CNN.rsCmdEnc_Pedidos.Update
            Else
                MsgBox "No existe el pedido", vbInformation, MSG
            End If
            CNN.rsCmdEnc_Pedidos.Close
        Loop
    Else
        MsgBox "No ha pedidos", vbExclamation, MSG
    End If
End If



''''''''''''''''''''''''''''''''
''''' Respaldo Codigo

'''Cant2 = 0
'''If Tipo = "F" Then
'''    If Len(Me.LstReq.SelectedItem.ListSubItems(12).Text) <> 0 Then
'''        Cadena = Me.LstReq.SelectedItem.ListSubItems(12).Text
'''        i = 1
'''        Do While i < Len(Cadena)
'''            Cadena2 = i
'''            Resp = InStr((i + 1), Cadena, "*")
'''            IdPedido = CLng(Mid(Cadena, (i + 1), ((Resp - 1) - i)))
'''            i = Resp
'''            CNN.CmdEnc_Pedidos (IdPedido)
'''            If CNN.rsCmdEnc_Pedidos.EOF <> True Then
'''                If Cant2 < SumPzs Then
'''                    CNN.cmdDet_Pedidos (IdPedido), (1)
'''                    If CNN.rscmdDet_Pedidos.EOF <> True Then
'''                        Cant2 = Cant2 + CNN.rscmdDet_Pedidos!detp_Cant
'''                    End If
'''                    CNN.rscmdDet_Pedidos.Close
'''                    'Actualiza aqui
'''                        CNN.rsCmdEnc_Pedidos!IdJC = IdJC
'''                    CNN.rsCmdEnc_Pedidos.Update
'''                Else
'''                    If Cant2 >= SumPzs Then
'''                        i = Len(Cadena)
'''                    End If
'''                End If
'''            Else
'''                MsgBox "No existe el pedido", vbInformation, MSG
'''            End If
'''            CNN.rsCmdEnc_Pedidos.Close
'''        Loop
'''    Else
'''        MsgBox "No hay pedidos", vbExclamation, MSG
'''    End If
'''End If
'''If Tipo = "AT" Then
'''    If Len(Me.LstReqAT.SelectedItem.ListSubItems(12).Text) <> 0 Then
'''        Cadena = Me.LstReqAT.SelectedItem.ListSubItems(12).Text
'''        i = 1
'''        Do While i < Len(Cadena)
'''            Cadena2 = i
'''            Resp = InStr((i + 1), Cadena, "*")
'''            IdPedido = CLng(Mid(Cadena, (i + 1), ((Resp - 1) - i)))
'''            i = Resp
'''            CNN.CmdEnc_Pedidos (IdPedido)
'''            If CNN.rsCmdEnc_Pedidos.EOF <> True Then
'''
'''            If Cant2 < SumPzs Then
'''                CNN.cmdDet_Pedidos (IdPedido), (1)
'''                If CNN.rscmdDet_Pedidos.EOF <> True Then
'''                    Cant2 = Cant2 + CNN.rscmdDet_Pedidos!detp_Cant
'''                End If
'''                CNN.rscmdDet_Pedidos.Close
'''                If IsLaundry = True Then
'''                    CNN.rsCmdEnc_Pedidos!IdJCcorte = IdJC
'''                Else
'''                    CNN.rsCmdEnc_Pedidos!IdJC = IdJC
'''                End If
'''
'''                CNN.rsCmdEnc_Pedidos.Update
'''            Else
'''                If Cant2 >= SumPzs Then
'''                    i = Len(Cadena)
'''                End If
'''            End If
'''            Else
'''                MsgBox "No existe el pedido", vbInformation, MSG
'''            End If
'''            CNN.rsCmdEnc_Pedidos.Close
'''        Loop
'''    Else
'''        MsgBox "No ha pedidos", vbExclamation, MSG
'''    End If
'''End If




End Sub
