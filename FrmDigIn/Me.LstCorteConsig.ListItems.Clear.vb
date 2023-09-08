Me.LstCorteConsig.ListItems.Clear
Me.LstDetReq.ListItems.Clear
Me.LstSalidas.ListItems.Clear
IdProv = Me.CboProv.BoundText
If IdProv <> 0 Then
    'Busca arts con entradas sin matar.
    Me.LstCorteConsig.ListItems.Clear
    i = 1
    CNN.CmdArtConsig (IdProv)
    Band = False
    Do While CNN.rsCmdArtConsig.EOF <> True
        Resp = CNN.rsCmdArtConsig.RecordCount
        If i = 1 Then
            Me.LstCorteConsig.ListItems.Add (i)
                Me.LstCorteConsig.ListItems.Item(i).SubItems(1) = CNN.rsCmdArtConsig!Ref_Id
                Me.LstCorteConsig.ListItems.Item(i).SubItems(2) = CNN.rsCmdArtConsig!ref_desc
                Me.LstCorteConsig.ListItems.Item(i).SubItems(3) = CNN.rsCmdArtConsig!ind_RzSoc
                Me.LstCorteConsig.ListItems.Item(i).SubItems(4) = CNN.rsCmdArtConsig!Dcon_Precio
                Me.LstCorteConsig.ListItems.Item(i).SubItems(5) = CNN.rsCmdArtConsig!Moneda
            i = i + 1
        Else
            j = 1
            Band = False
            Do While j < i And Band = False
                If CNN.rsCmdArtConsig!Ref_Id = CDbl(Me.LstCorteConsig.ListItems.Item(j).SubItems(1)) Then
                    Band = True
                Else
                    j = j + 1
                End If
            Loop
            If Band = False Then
                Me.LstCorteConsig.ListItems.Add (i)
                    Me.LstCorteConsig.ListItems.Item(i).SubItems(1) = CNN.rsCmdArtConsig!Ref_Id
                    Me.LstCorteConsig.ListItems.Item(i).SubItems(2) = CNN.rsCmdArtConsig!ref_desc
                    Me.LstCorteConsig.ListItems.Item(i).SubItems(3) = CNN.rsCmdArtConsig!ind_RzSoc
                    Me.LstCorteConsig.ListItems.Item(i).SubItems(4) = CNN.rsCmdArtConsig!Dcon_Precio
                    Me.LstCorteConsig.ListItems.Item(i).SubItems(5) = CNN.rsCmdArtConsig!Moneda
                i = i + 1
            End If
        End If
        CNN.rsCmdArtConsig.MoveNext
    Loop
    CNN.rsCmdArtConsig.Close
Else
    MsgBox "No existe el proveedor, verifique los datos.", vbInformation, MSG
    Exit Sub
End If
