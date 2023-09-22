SELECT        TOP (10) Tbl_DetOma.Oma_Id, Tbl_DetOma.IdHorario, Tbl_DetOma.HoraIni, Tbl_DetOma.HoraFin, Tbl_DetOma.CodVidrio, Tbl_DetOma.PzsOK, Tbl_DetOma.FechaCap, Tbl_DetOma.HoraCap, Tbl_DetOma.Observaciones,
  Tbl_DetOma.Item , TblHorariosHxH.Turno, Tbl_DetOma.IdOperacion, Tbl_DetOma.CodLinea
  FROM            Tbl_DetOma INNER JOIN
                           TblHorariosHxH ON Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario

  WHERE 

Tbl_DetOma.codlinea= & Codlinea

  ORDER BY Tbl_DetOma.IDENTITYCOL DESC, Tbl_DetOma.HoraFin DESC