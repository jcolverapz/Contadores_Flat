SELECT        TOP (10) Tbl_DetOma.Oma_Id, Tbl_DetOma.IdHorario, Tbl_DetOma.HoraIni, Tbl_DetOma.HoraFin, Tbl_DetOma.CodVidrio, Tbl_DetOma.PzsOK, Tbl_DetOma.FechaCap, Tbl_DetOma.HoraCap, Tbl_DetOma.Observaciones, 
                         Tbl_DetOma.Item, TblHorariosHxH.Turno, Tbl_DetOma.IdOperacion, lineas.descripcion, lineas.codlinea
FROM            Tbl_DetOma INNER JOIN
                         TblHorariosHxH ON Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario INNER JOIN
                         Orden_Man ON Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN
                         lineas ON Orden_Man.Codlinea = lineas.codlinea
WHERE        (Tbl_DetOma.FechaCap = CONVERT(DATETIME, '2023-07-21 00:00:00', 102)) AND (lineas.codlinea = N'331') AND (TblHorariosHxH.Turno = 2)