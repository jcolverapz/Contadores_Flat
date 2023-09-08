SQL = " SELECT lineas.descripcion AS Linea, Maquina.Descricion AS Maquina,"
SQL = SQL & "     operaciones.descripcion AS Operacion,"
SQL = SQL & "     Orden_Man.Oma_Id AS JobCard, Tbl_DetOma.CodVidrio AS Vidrio,"
SQL = SQL & "     Tbl_DetOma.Ticketgem AS Ticket, Tbl_DetOma.PxH,"
SQL = SQL & "     SUM(Tbl_DetOma.PzsOK) AS PzsOK, Tbl_DetOma.PzsScrap,"
SQL = SQL & "     lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,"
SQL = SQL & "     Orden_Man.Codlinea, TblHorariosHxH.Turno,"
SQL = SQL & "     Orden_Man.Oma_pza_prog AS goal,"
SQL = SQL & "     Tbl_DetOma.IDENTITYCOL"
SQL = SQL & " FROM Tbl_DetOma INNER JOIN"
SQL = SQL & "     Orden_Man ON"
SQL = SQL & "     Tbl_DetOma.Oma_Id = Orden_Man.Oma_Id INNER JOIN"
SQL = SQL & "     lineas ON Orden_Man.Codlinea = lineas.codlinea INNER JOIN"
SQL = SQL & "     operaciones ON"
SQL = SQL & "     Tbl_DetOma.IdOperacion = operaciones.codopera INNER JOIN"
SQL = SQL & "     Maquina ON"
SQL = SQL & "     Tbl_DetOma.IdMaquina = Maquina.codmaquina INNER JOIN"
SQL = SQL & "     TblHorariosHxH ON"
SQL = SQL & "     Tbl_DetOma.IdHorario = TblHorariosHxH.IdHorario"
SQL = SQL & " GROUP BY lineas.descripcion, Maquina.Descricion,"
SQL = SQL & "     operaciones.descripcion, Orden_Man.Oma_Id, Tbl_DetOma.CodVidrio,"
SQL = SQL & "     Tbl_DetOma.Ticketgem, Tbl_DetOma.PxH, Tbl_DetOma.PzsScrap,"
SQL = SQL & "     lineas.codlinea, Tbl_DetOma.FechaCap, operaciones.codopera,"
SQL = SQL & "     Orden_Man.Codlinea, TblHorariosHxH.Turno,"
SQL = SQL & "     Orden_Man.Oma_pza_prog , Tbl_DetOma.IDENTITYCOL"
SQL = SQL & " HAVING (SUM(Tbl_DetOma.PzsOK) <> 0) AND (TblHorariosHxH.Turno = " & Turno & ")"
SQL = SQL & "     AND (Tbl_DetOma.FechaCap = '" & Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha) & "') AND "
SQL = SQL & "    (Tbl_DetOma.codlinea = N'" & CodLinea & "'"