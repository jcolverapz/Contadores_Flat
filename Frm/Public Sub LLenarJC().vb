Public Sub LLenarJC()
Me.LstJC.ListItems.Clear
i = 1
'Para Flat
'RUTA:Orden_Man vs Clientes vs Especificaciones
CNN.CmdOma ("Activa")
Do While CNN.rsCmdOma.EOF <> True
    Resp = CNN.rsCmdOma.RecordCount
    Me.LstJC.ListItems.Add (i)
        Me.LstJC.ListItems.Item(i).SubItems(1) = CNN.rsCmdOma!Oma_Id
        IdJC = CNN.rsCmdOma!Oma_Id
        Me.LstJC.ListItems.Item(i).SubItems(2) = CNN.rsCmdOma!OF
        Me.LstJC.ListItems.Item(i).SubItems(3) = CNN.rsCmdOma!nodeparte
        Me.LstJC.ListItems.Item(i).SubItems(4) = CNN.rsCmdOma!Oma_pza_prog
        Me.LstJC.ListItems.Item(i).SubItems(5) = CNN.rsCmdOma!oma_NoCajas
'''        CNN.CmdOma_detV (IdJC), (1)
'''        If CNN.rsCmdOma_detV.EOF <> True Then
'''            Me.LstJC.ListItems.Item(i).SubItems(6) = CNN.rsCmdOma_detV!Codigo
'''        Else
'''            Me.LstJC.ListItems.Item(i).SubItems(6) = "Sin Asignación"
'''        End If
'''        CNN.rsCmdOma_detV.Close

        Me.LstJC.ListItems.Item(i).SubItems(6) = CNN.rsCmdOma!CodVidrio
        
        Me.LstJC.ListItems.Item(i).SubItems(7) = CNN.rsCmdOma!Oma_Status
        Me.LstJC.ListItems.Item(i).SubItems(8) = CNN.rsCmdOma!oma_observ
    i = i + 1
    CNN.rsCmdOma.MoveNext
Loop
CNN.rsCmdOma.Close


'Para V MPS
CNN.CmdOmaAT ("Activa")
Do While CNN.rsCmdOmaAT.EOF <> True
    Resp = CNN.rsCmdOmaAT.RecordCount
    Me.LstJC.ListItems.Add (i)
        Me.LstJC.ListItems.Item(i).SubItems(1) = CNN.rsCmdOmaAT!Oma_Id
        IdJC = CNN.rsCmdOmaAT!Oma_Id
        Me.LstJC.ListItems.Item(i).SubItems(2) = CNN.rsCmdOmaAT!OF
        Me.LstJC.ListItems.Item(i).SubItems(3) = CNN.rsCmdOmaAT!OF
        Me.LstJC.ListItems.Item(i).SubItems(4) = CNN.rsCmdOmaAT!Oma_pza_prog
        Me.LstJC.ListItems.Item(i).SubItems(5) = CNN.rsCmdOmaAT!oma_NoCajas
'''        CNN.CmdOma_detV (IdJC), (1)
'''        If CNN.rsCmdOma_detV.EOF <> True Then
'''            Me.LstJC.ListItems.Item(i).SubItems(6) = CNN.rsCmdOma_detV!Codigo
'''        Else
'''            Me.LstJC.ListItems.Item(i).SubItems(6) = "Sin Asignación"
'''        End If
'''        CNN.rsCmdOma_detV.Close
        Me.LstJC.ListItems.Item(i).SubItems(6) = CNN.rsCmdOmaAT!CodVidrio

        Me.LstJC.ListItems.Item(i).SubItems(7) = CNN.rsCmdOmaAT!Oma_Status
        Me.LstJC.ListItems.Item(i).SubItems(8) = CNN.rsCmdOmaAT!oma_observ
    i = i + 1
    CNN.rsCmdOmaAT.MoveNext
Loop
CNN.rsCmdOmaAT.Close
End Sub

