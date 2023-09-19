Public Sub UltimoTicketxLinea(Fecha, CodLinea)
On Error GoTo ErrHandler

CNN.CmdEtiquetas CodLinea

If CNN.rsCmdEtiquetas.EOF <> True Then
    
    TicketGem = CNN.rsCmdEtiquetas!TicketGem         
    IdJC = CNN.rsCmdEtiquetas!Oma_Id
   ' OF = CNN.rsCmdEtiquetas!OF
    CodLinea = CNN.rsCmdEtiquetas!CodLinea
    'EsAcumulado = InStr(1, CNN.rsCmdUltimoTicketxLinea!descripcion, "Acumulado")
    FechaUltimo = CNN.rsCmdEtiquetas!FechaHora  
     
CNN.rsCmdEtiquetas.Close
End If
ErrHandler:

End Sub
