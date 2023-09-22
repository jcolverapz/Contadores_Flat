SELECT lineas.descripcion AS Linea, Maquina.Descricion AS Maquina,
    operaciones.descripcion AS Operacion,
    Orden_Man.Oma_Id AS JobCard, Tbl_DetOma.CodVidrio AS Vidrio,
    Tbl_DetOma.Ticketgem AS Ticket, Tbl_DetOma.PxH,
    SUM(Tbl_DetOma.PzsOK) AS PzsOK, Tbl_DetOma.PzsScrap,
    lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,
    Orden_Man.Codlinea, TblHorariosHxH.Turno,
    Orden_Man.Oma_pza_prog AS goal
FROM Tbl_DetOma INNER JOIN
    Orden_Man ON
    Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN
    lineas ON Orden_Man.Codlinea = lineas.codlinea INNER JOIN
    operaciones ON
    Tbl_DetOma.IdOperacion = operaciones.codopera INNER JOIN
    Maquina ON
    Tbl_DetOma.IdMaquina = Maquina.codmaquina INNER JOIN
    TblHorariosHxH ON
    Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario
GROUP BY lineas.descripcion, Maquina.Descricion,
    operaciones.descripcion, Orden_Man.Oma_Id, Tbl_DetOma.CodVidrio,
    Tbl_DetOma.Ticketgem, Tbl_DetOma.PxH, Tbl_DetOma.PzsScrap,
    lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,
    Orden_Man.Codlinea, TblHorariosHxH.Turno,
    Orden_Man.Oma_pza_prog
HAVING (Orden_Man.Oma_Id =  & IdJC & )