Private Sub TimerLstConteos_Timer()

Dim ii As Integer
Dim i As Integer
Dim Cont As Integer
Dim Contador(6) As Integer

On Error GoTo ErrHandler

Call UltimoTicketxLinea(Fecha, CodLinea)
 
If CNN.rsCmdConteosEliminarxOperacion.State = 1 Then
CNN.rsCmdConteosEliminarxOperacion.Close
End If

SQL = "SELECT Id, codlinea, NoPuerto, JobCard FROM TblConteos Where TicketGem = " & TicketGem

CNN.rsCmdConteosEliminarxOperacion.Open SQL

Do While CNN.rsCmdConteosEliminarxOperacion.EOF = False

    Select Case CNN.rsCmdConteosEliminarxOperacion!NoPuerto

        Case 1
        Contador(1) = Contador(1) + 1
        Case 2
        Contador(2) = Contador(2) + 1
        Case 5
        Contador(5) = Contador(5) + 1
        Case 6
        Contador(6) = Contador(6) + 1

    End Select

    CNN.rsCmdConteosEliminarxOperacion.Delete

    CNN.rsCmdConteosEliminarxOperacion.MoveNext

Loop

CNN.rsCmdConteosEliminarxOperacion.Close

For i = 1 To 6
    If Contador(i) > 0 Then
    Call FrmVisorContadores3Lineas.UpdateConteos(i, CodLinea, Contador(i))
    End If
Next i

'Debug.Print "Conteos: " & Now

Call LstActualizaContadores

ErrHandler:
'sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
'GoLogTheError sMsg

'CNN.rsCmdConteosEliminarxOperacion.Open "SELECT TOP " & TotPzs & " Id, codlinea, NoPuerto FROM TblConteos WHERE (codlinea = " & NoCajaAzul & ") AND (NoPuerto = " & i & ")"

End Sub

