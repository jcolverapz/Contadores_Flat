    Public Sub Guardar(ByVal Parametro As Object)

        'Dim conn As ADODB.Connection
        'Dim rs As ADODB.Recordset

        'dgvListaDatos.EndEdit()
        'Me.dgvListaDatos.Update()

        Dim sqlstring As String
        Dim cmd As New OleDbCommand
        Dim i, q As Integer
        Dim mifecha As String
        Dim test As String

        ' test = Me.cboTipo.SelectedItem.Value



        ''sqlstring = "insert into Enc_gastos (descripcion) values ('" & txtDescripcion.Text & "')"
        '' sqlstring = "insert into Enc_gastos (descripcion) values ('1')"
        'sqlstring = "insert into Tbl_Historial (tot) values ('" & i & "')"

        'cmd.Connection = conexion
        'cmd.CommandText = sqlstring

        'i = cmd.ExecuteNonQuery



        ' mifecha = FormatDateTime(Now, DateFormat.ShortDate)
        ' mifecha= dtpFecha.CustomFormat "dd/MM/yyyy")

        'sqlstring = "insert into Det_gastos (tipo,periodo,cant, IdGastos,fecha) values ('" & 1 & "','" & txtDescripcion.Text & "','" & txtCantidad.Text & "', '" & Form2.IdEnc & "','" & dtpFecha.Value & "' )"
        'sqlstring = "Insert into Registros (Bit, Cuenta) values (@Bit, @cuenta)"
        ' sqlstring = "insert into Registros (Bit, Cuenta) values ('1','1')"
        ' sqlstring = "Insert into Det_gastos ( fecha) values (@fecha)"
        'dtpFecha.CustomFormat = "dd/MM/yyyy"
        ' sqlstring = "Insert into Det_gastos (fecha) values (@fecha)"

        'sqlstring = "insert into Tbl_Historial (tot) values ('" & i & "')"

        'conexion.Open()

        ' cmd.CommandText = sqlstring
        'cmd.Parameters.AddWithValue("tipo", cboTipo.Text)

        mifecha = Format(Now, "dd/mm/yyyy hh:mm:ss")

        Try
            conexion.Open()
            'sqlstring = "insert into Registros (descripcion) values ('1')"

            cmd.Connection = conexion

            'sqlstring = "insert into Tbl_Historial (cuenta) values (@cuenta)"
            sqlstring = "Update Tbl_Historial SET cuenta= @cuenta  Where ID=@id"
            'sqlstring = "Update Tbl_Historial set cuenta=@cuenta, fecha=@fecha Where bitn=@bitn"
            ' sqlstring = "Update Tbl_Historial set cuenta=@cuenta Where Id=@Id"
            'sqlstring = "Update Enc_gastos SET descripcion = @descripcion, frec=@frec, comentarios=@comentarios Where ID=@id"
            cmd.CommandText = sqlstring
            'cmd.CommandText = sqlstring
            cmd.CommandType = CommandType.StoredProcedure
            'i = 4

            For i = 1 To 6
                cmd.Parameters.Clear()

                cmd.Parameters.AddWithValue("cuenta", bitCounter(i))
                ' cmd.Parameters.AddWithValue("fecha", mifecha)
                cmd.Parameters.AddWithValue("Id", i)
                cmd.CommandType = CommandType.Text

                q = cmd.ExecuteNonQuery

            Next i

            If q > 0 Then
            Else
                '  MsgBox("No se guardo")
            End If

            Catch ex As Exception

            'MsgBox(ex.Message)
        Finally
                conexion.Close()
            End Try







    End Sub
