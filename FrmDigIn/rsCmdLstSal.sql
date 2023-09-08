rsCmdLstSal
SELECT TOP 200 Enc_Ind_Salidas.idsalida, Enc_Ind_Salidas.ind_fechasal, 
    Enc_Ind_Salidas.Codturno, EMPLEADO.EMP_NOMBRE AS Entrego, 
    EMPLEADO1.EMP_NOMBRE AS Solicita, Dpto.Dpto_Descripcion, 
    Enc_Ind_Salidas.Consignacion, Enc_Ind_Salidas.Status, 
    Enc_Ind_Salidas.dptoid
FROM Enc_Ind_Salidas INNER JOIN
    EMPLEADO ON 
    Enc_Ind_Salidas.Empleadoid = EMPLEADO.EMPLEADOID INNER JOIN
    EMPLEADO EMPLEADO1 ON 
    Enc_Ind_Salidas.SolicitaID = EMPLEADO1.EMPLEADOID INNER JOIN
    Dpto ON Enc_Ind_Salidas.dptoid = Dpto.DeptoId
WHERE (Enc_Ind_Salidas.Consignacion = 'S')
ORDER BY Enc_Ind_Salidas.idsalida DESC