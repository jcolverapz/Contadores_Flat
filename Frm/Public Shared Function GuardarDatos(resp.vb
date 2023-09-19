    Public Shared Function GuardarDatos(respuestas, respCritica, IdAuditoria)
        Dim _comando As SqlCommand = MetodosDatos.CrearComando()
        'Dim _comando As SqlCommand = MetodosDatos.CrearComando()
        ' Dim agregar As SqlCommand = New SqlCommand("INSERT INTO [dbo].[Det_Auditoria]  ([IdAuditoria], [IdPregunta], [Resp]) VALUES (@IdAuditoria, @IdPregunta,@Resp_abierta, @Resp)", _comando.Connection)
        'Dim agregar As SqlCommand = New SqlCommand("INSERT INTO [dbo].[Det_Auditoria]  ([IdAuditoria]) VALUES (@IdAuditoria)", _comando.Connection)
        Dim _insertado As Boolean = False
        Dim mostrarMensaje As Boolean = False
        Dim stringvalues As String = ""
        Dim sql As String = ""

        '  _comando.Connection.Close()
        '  _comando.CommandType = CommandType.StoredProcedure
        'Dim tran As SqlTransaction
        ' tran = _comando.Connection.BeginTransaction
        Try
            ' Using tran = 
            _comando.CommandText = "INSERT INTO [dbo].[Det_Auditoria]  ([IdAuditoria], [IdPregunta], [Resp]) VALUES (@IdAuditoria, @IdPregunta, @Resp)"

            'insertar los datos
            'Dim fila As Integer
            _comando.Connection.Open()
            '_comando.Transaction = tran
            '_comando.ExecuteNonQuery()
            ' tran.Commit()
        Catch ex As Exception
            'tran.Rollback()
        End Try


        _comando.CommandType = CommandType.StoredProcedure


        Try

            For Each _resp As clsRespuesta In respCritica
                _comando.Parameters.Clear()


                _comando.Parameters.AddWithValue("@IdAuditoria", IdAuditoria)
                _comando.CommandType = CommandType.Text

                _comando.Parameters.AddWithValue("@IdPregunta", _resp.IdPregunta)
                '_comando.Parameters.AddWithValue("@Resp_abierta", "")
                _comando.Parameters.AddWithValue("@Resp", _resp.Respuesta)
                _comando.ExecuteNonQuery()

            Next

            For Each _resp As clsRespuesta In respuestas
                _comando.Parameters.Clear()


                _comando.Parameters.AddWithValue("@IdAuditoria", IdAuditoria)
                _comando.CommandType = CommandType.Text

                _comando.Parameters.AddWithValue("@IdPregunta", _resp.IdPregunta)
                '_comando.Parameters.AddWithValue("@Resp_abierta", "")
                _comando.Parameters.AddWithValue("@Resp", _resp.Respuesta)
                _comando.ExecuteNonQuery()

            Next

            'MsgBox("Guardado correctamente.", vbInformation, "Éxito")

        Catch ex As Exception
            MsgBox("no se guardo", MsgBoxStyle.OkOnly, "Éxito")


        End Try
        _comando.Connection.Close()

        '_comando.CommandText = _comando.CommandText & _comando.CommandText

        '_comando.Parameters("@IdAuditoria").Value = 1
        '_comando.Parameters("@IdPregunta").Value = resp.IdPregunta
        '_comando.Parameters("@Resp").Value = resp.Respuesta

        '_comando.Parameters.AddWithValue("@IdAuditoria", "some value")
        '_comando.Parameters.AddWithValue("@IdPregunta", "some value")
        '_comando.Parameters.AddWithValue("@Resp", "some value")

        ' _comando.CommandText = _comando.CommandText & "[" & propiedad.Nombre & "]"
        'stringvalues = stringvalues & "@" & propiedad.Nombre
        ' stringvalues = stringvalues & ", @" & propiedad.Nombre

        '  Console.WriteLine(resp.IdPregunta)

        'INSERT INTO Info (id, Cost, city)  
        'VALUES(1, 200, 'Pune'), (2, 150,'USA'), (3,345, 'France'); 
        ' _comando.ExecuteNonQuery()


        ' _comando.CommandText = _comando.CommandText & ") VALUES (" & stringValues & ")"
        '_comando.Connection.Open()
        '_comando.Connection.Close()


        '_comando.Parameters.Add(New SqlParameter("@IdPregunta", SqlDbType.Bit))
        '_comando.Parameters("@IdPregunta").Value = 1
        '_comando.Parameters.Add(New SqlParameter("@Resp", SqlDbType.Bit))
        '_comando.Parameters("@Resp").Value = 1

        ' _comando.Parameters("@Descripcion").Value = Convert.ToInt32(CallByName(_obj, Microsoft.VisualBasic.CallType.Get, Nothing))
        ' _comando.Parameters("@Descripcion").Value = Convert.ToString(_obj()

        'Try
        '    Dim objDataReader As SqlDataReader = _comando.ExecuteReader()
        '    If objDataReader.RecordsAffected = 0 Then
        '        MsgBox("Error al insertar el registro.", MsgBoxStyle.OkOnly, "Error")
        '        _insertado = False
        '    Else
        '        ' If mostrarMensaje = False Then
        '        'Else
        '        'End If
        '        _insertado = True
        '    End If
        'Catch ex As Exception
        '    MsgBox("Error al insertar el registro. " & ex.Message, MsgBoxStyle.OkOnly, "Error")
        '    _insertado = False
        'End Try
        Return _insertado


    End Function
