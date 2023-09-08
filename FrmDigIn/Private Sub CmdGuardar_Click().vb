Private Sub CmdGuardar_Click()
'Valida Datos
'Valida existencias y guarda encabezado y detalles de la salida.
'

ValText "No Folio", Me.TxtNoFolio.Text
If Band = False Then Exit Sub
ValText "Fecha", Me.DTP_Fecha.Value
If Band = False Then Exit Sub

ValText "Turno", Me.CboTurno.Text
If Band = False Then Exit Sub

ValText "Departamento", Me.CboDpto.Text
If Band = False Then Exit Sub
ValText "Linea", Me.CboLinea.Text
If Band = False Then Exit Sub
ValText "Maquina", Me.CboMaquina.Text
If Band = False Then Exit Sub
IdFolio = Me.TxtNoFolio.Text


Cadena4 = ""


CNN.rsCmdMaxSalRef.Open
If IsNumeric(CNN.rsCmdMaxSalRef!m) Then
    IdFolio = CNN.rsCmdMaxSalRef!m + 1
Else
    IdFolio = 1
End If
CNN.rsCmdMaxSalRef.Close
Me.TxtNoFolio.Text = IdFolio
Me.TxtNoFolio.Enabled = False



If BandEnc = False Then
    'Guarda encabezado
    CNN.CmdEncSal (IdFolio)
    If CNN.rsCmdEncSal.EOF = True Then
        CNN.rsCmdEncSal.AddNew
            CNN.rsCmdEncSal!idsalida = IdFolio
            CNN.rsCmdEncSal!ind_fechasal = Me.DTP_Fecha.Value
            CNN.rsCmdEncSal!codTurno = Me.CboTurno.Text
            CNN.rsCmdEncSal!EmpleadoId = EmpleadoId 'EntregoId 'EmpleadoId
            CNN.rsCmdEncSal!SolicitaId = IdUsuarioAut 'SolicitaId
            CNN.rsCmdEncSal!CodLinea = CodLinea
            CNN.rsCmdEncSal!CodMaquina = CodMaquina
            
            CNN.CmdBuscaDpto (Me.CboDpto.Text)
            If CNN.rsCmdBuscaDpto.EOF <> True Then
                DptoId = CNN.rsCmdBuscaDpto!deptoid
                CNN.rsCmdBuscaDpto.Close
            Else
                MsgBox "El departamento no existe", vbInformation, MSG
                CNN.rsCmdBuscaDpto.Close
                Exit Sub
            End If
            
            CNN.rsCmdEncSal!DptoId = DptoId
            CNN.rsCmdEncSal!tipo_vale = "S"
            CNN.rsCmdEncSal!Obser = "° Huella ° " & Trim(Me.TxtObserv.Text)
            CNN.rsCmdEncSal!AutorizoId = IdUsuarioAut 'AutorizaId
            CNN.rsCmdEncSal!Consignacion = Consignacion
        CNN.rsCmdEncSal.Update
        BandEnc = True
    Else
        MsgBox "El Folio ya existe.", vbInformation, MSG
        CNN.rsCmdEncSal.Close
        Exit Sub
    End If
    CNN.rsCmdEncSal.Close
End If

Cadena4 = ""

'Guarda detallle
Detalle = 1
BandReq = False
Do While Detalle <= (MaxRenglon)
    CNN.CmdDet_Sal (IdFolio), (Detalle)
    If CNN.rsCmdDet_Sal.EOF = True Then
        CNN.rsCmdDet_Sal.AddNew
            CNN.rsCmdDet_Sal!idsalida = IdFolio
            CNN.rsCmdDet_Sal!Item = Detalle
            CNN.rsCmdDet_Sal!Ref_Id = Me.LstDetalles.ListItems(Detalle).SubItems(2)
            CNN.rsCmdDet_Sal!Cantidad = Me.LstDetalles.ListItems(Detalle).SubItems(6)
            CNN.rsCmdDet_Sal!Status = "A"
            CNN.rsCmdDet_Sal!Consignacion = Consignacion
        CNN.rsCmdDet_Sal.Update
        
        
        'Actualiza inv
        CNN.CmdExist (Me.LstDetalles.ListItems(Detalle).SubItems(2))
        If CNN.rsCmdExist.EOF <> True Then
            If Consignacion = "N" Then
                CNN.rsCmdExist!ind_sal = CNN.rsCmdExist!ind_sal + CDbl(Me.LstDetalles.ListItems(Detalle).SubItems(6))
                CNN.rsCmdExist!ind_exi = (CNN.rsCmdExist!ind_exi_ini + CNN.rsCmdExist!ind_ent + CNN.rsCmdExist!ind_entsob) - CNN.rsCmdExist!ind_sal
            Else
                CNN.rsCmdExist!ind_salc = CNN.rsCmdExist!ind_salc + CDbl(Me.LstDetalles.ListItems(Detalle).SubItems(6))
                CNN.rsCmdExist!ind_exic = CNN.rsCmdExist!ind_exi_inic + CNN.rsCmdExist!ind_entc - CNN.rsCmdExist!ind_salc
            End If
            
           IdReq = 0
           DetReq = 0
            
            If Consignacion = "N" Then
                
                If CNN.rsCmdExist!CriticoenCero = "N" Then
                
                        If CNN.rsCmdExist!ref_Critico = 1 And CNN.rsCmdExist!ref_max <> 0 Then
                            If CNN.rsCmdExist!ind_exi <= CNN.rsCmdExist!ref_reorden Then
                                'If CNN.rsCmdExist!reqpr = "S" Then
                                 '  CNN.rsCmdExist!reqpr = "S"
                                    Articulo = CNN.rsCmdExist!Ref_Id
                                    DescArt = CNN.rsCmdExist!ref_desc
                                    
                                    'Busca ultima compra
                                    RefId = Me.LstDetalles.ListItems(Detalle).SubItems(2)
                                    CNN.CmdBuscaUltimacompra (RefId)
                                    If CNN.rsCmdBuscaUltimaCompra.EOF <> True Then
                                        DptoId = CNN.rsCmdBuscaUltimaCompra!deptoid
                                        TipoReq = CNN.rsCmdBuscaUltimaCompra!oci_tipo
                                    Else
                                        DptoId = "03"
                                        TipoReq = "N"
                                    End If
                                    CNN.rsCmdBuscaUltimaCompra.Close
                                    
                                    CNN.CmdUsuariosCompras (DptoId)
                                    If CNN.rsCmdUsuariosCompras.EOF <> True Then
                                        CorreoGerente = CNN.rsCmdUsuariosCompras!cc_Correo
                                    Else
                                        CorreoGerente = ""
                                    End If
                                    CNN.rsCmdUsuariosCompras.Close
                                    
                                    Dim fso, txtfile
                                    Dim j, m As Integer
                                   
                                    Para = CorreoGerente
                                     
                                    'Nombre = "Buyernet"
                                    'CorreoRem = "buyernet@gemtron.com.mx"
                                    
                                    De = "raul.chavez@schott.com"
                                        
                                        Asunto = "Notificación Refacciones Criticas " ''& "Test TI"
                                        
                                        If BandReq = False Then
                                            'Encabezado Req
                                            CNN.rsCmdMaxReq.Open
                                            If IsNull(CNN.rsCmdMaxReq!m) Then
                                                IdReq = 1
                                            Else
                                                IdReq = CNN.rsCmdMaxReq!m + 1
                                            End If
                                            CNN.rsCmdMaxReq.Close
                                            CNN.CmdEncReq (IdReq)
                                            If CNN.rsCmdEncReq.EOF = True Then
                                                CNN.rsCmdEncReq.AddNew
                                                    CNN.rsCmdEncReq!IdReq = IdReq
                                                    CNN.rsCmdEncReq!fecha_solicitud = Date
                                                    CNN.rsCmdEncReq!hora_solicitud = Time
                                                    CNN.rsCmdEncReq!fecha_req = Date + 1
                                                    'gerente
            '                                        CNN.CmdCorreoGerente (DptoId) '(DptoTemp)
            '                                        If CNN.rsCmdCorreoGerente.EOF <> True Then
            '                                            Para = CNN.rsCmdCorreoGerente!correo
            '                                            EmpleadoIdReq = CNN.rsCmdCorreoGerente!EmpleadoId
            '                                        End If
            '                                        CNN.rsCmdCorreoGerente.Close
                                                    CNN.rsCmdEncReq!EmpleadoId = SolicitaId
                                                        CNN.CmdBuscaDpto (Me.CboDpto.Text)
                                                        If CNN.rsCmdBuscaDpto.EOF <> True Then
                                                            DptoId = CNN.rsCmdBuscaDpto!deptoid
                                                            CNN.rsCmdBuscaDpto.Close
                                                        Else
                                                            MsgBox "El departamento no existe", vbInformation, MSG
                                                            CNN.rsCmdBuscaDpto.Close
                                                            Exit Sub
                                                        End If
                                                    CNN.rsCmdEncReq!DptoId = DptoId
                                                    CNN.rsCmdEncReq!uso_insumo = "Cubrir inventario (Automatico°)"
                                                    CNN.rsCmdEncReq!idEmpleado = EmpleadoId 'Comprador
                                                    CNN.rsCmdEncReq!req_status = "S"
                                                    CNN.rsCmdEncReq!req_tipo = TipoReq
                                                    CNN.rsCmdEncReq!id_oficina = "PSLP"
                                                CNN.rsCmdEncReq.Update
                                                BandReq = True
                                                DetReq = 1
                                            Else
                                                
                                            End If
                                            CNN.rsCmdEncReq.Close
                                        End If
                                        'Detalles Reqs
                                        CNN.CmdDetReq (IdReq), (DetReq)
                                        If CNN.rsCmdDetReq.EOF = True Then
                                            CNN.rsCmdDetReq.AddNew
                                                CNN.rsCmdDetReq!IdReq = IdReq
                                                CNN.rsCmdDetReq!Item = DetReq
                                                CNN.rsCmdDetReq!Ref_Id = Articulo
                                                CNN.rsCmdDetReq!Cantidad = (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                                CNN.rsCmdDetReq!cantidad_pend = (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                                CNN.rsCmdDetReq!det_status = "A"
                                                CNN.rsCmdDetReq!det_obs = ""
                                            CNN.rsCmdDetReq.Update
                                            Cadena4 = Cadena4 & "<br><br>No Art: " & Articulo
                                            Cadena4 = Cadena4 & "<br>Descripcion: " & DescArt
                                            Cadena4 = Cadena4 & "<br>Cantidad: " & (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                            Cadena4 = Cadena4 & "<br>UM: " & CNN.rsCmdExist!ref_um
                                            DetReq = DetReq + 1
                                        Else
                                            MsgBox "El detalle ya existe."
                                        End If
                                        CNN.rsCmdDetReq.Close
                                
                                'End If
                                
                            End If
                            
                        End If
                        
'''                        Para = "raul.chavez@schott.com; pilar.aguilar@schott.com, javier.barron@schott.com, alfonso.diaz@schott.com, manuel.lozano@schott.com,"
'''                        ConCopia = ""
'''                        'Para = "helios.ireta@schott.com"
'''                        If BandReq = True And Cadena4 <> "" Then
'''                            CNN.rsCmdCorreo.Open
'''                            CNN.rsCmdCorreo.AddNew
'''                                CNN.rsCmdCorreo!De = De 'Des
'''                                CNN.rsCmdCorreo!Para = Para '"helios.ireta@schott.com, blas.reyes@schott.com, " 'Para
'''                                CNN.rsCmdCorreo!CC = ConCopia & "manuel.lozano@schott.com, teresa.montemayor@schott.com, raul.chavez@schott.com "
'''                                CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, "
'''                                CNN.rsCmdCorreo!Titulo = "° MXAPD Req. Automatica Punto de Reorden No. " & IdReq
'''                                '<marquee style="background-color: #000080;" direction="right" loop="20" width="75%">
'''                                'Esto es una marquesina móvil
'''                                '</marquee>
'''                                '<MARQUEE BGCOLOR='#FFFF1D'>¡¡¡Importante Autorizar esta Requicisión!!!</MARQUEE>
'''
'''                                Cadena4 = "Los siguientes articulos llegaron al punto de reorden, se elaboro la Req. No. " & IdReq & Cadena4 & "<br>la cual esta pendiente por autorizar."
'''
'''                                CNN.rsCmdCorreo!Mensaje = Cadena4
'''                                CNN.rsCmdCorreo!Aplicacion = "Alm Ref"
'''                                CNN.rsCmdCorreo!Prioridad = 2
'''                            CNN.rsCmdCorreo.Update
'''                            CNN.rsCmdCorreo.Close
'''                        End If
                        
                        
                        
                        CNN.CmdBuscaDpto (Me.CboDpto.Text)
                        If CNN.rsCmdBuscaDpto.EOF <> True Then
                            DptoId = CNN.rsCmdBuscaDpto!deptoid
                            CNN.rsCmdBuscaDpto.Close
                        Else
                            MsgBox "El departamento no existe", vbInformation, MSG
                            CNN.rsCmdBuscaDpto.Close
                            Exit Sub
                        End If
                        
                        
                        
                        '''envia Correos
                        Para = ""
                        ConCopia = ""
                        If DptoId = "04" Or DptoId = "06" Or DptoId = "05" Or DptoId = "10" Then
                                Para = "antonio.marin@schott.com; "
                                
                                CNN.CmdBuscaCorreosMtto (107), (115), (119), (118)
                                Do While CNN.rsCmdBuscaCorreosMtto.EOF <> True
                                            Para = Para & CNN.rsCmdBuscaCorreosMtto!Emp_Usr_Correo & ";  "
                                        CNN.rsCmdBuscaCorreosMtto.MoveNext
                                Loop
                                CNN.rsCmdBuscaCorreosMtto.Close
                                
                        Else
                            CNN.CmdBuscaCorreoGte (DptoId)
                            Do While CNN.rsCmdBuscaCorreoGte.EOF <> True
                                    Para = Para & CNN.rsCmdBuscaCorreoGte!Emp_Usr_Correo & "; "
                                CNN.rsCmdBuscaCorreoGte.MoveNext
                            Loop
                            CNN.rsCmdBuscaCorreoGte.Close
                        End If
                        
                         If Para <> "" And DetReq <> 0 And IdReq <> 0 Then
                                Cadena4 = Cadena4 & "<br><br><br>Elaboro: " & Nombre & "<br>Fecha y Hora: " & Now
                                'envia correos
                                CNN.rsCmdCorreo.Open
                                CNN.rsCmdCorreo.AddNew
                                    CNN.rsCmdCorreo!De = De 'Des
                                    CNN.rsCmdCorreo!Para = Para '"helios.ireta@schott.com, blas.reyes@schott.com, " 'Para
                                    CNN.rsCmdCorreo!CC = ConCopia & "manuel.lozano@schott.com, teresa.montemayor@schott.com, raul.chavez@schott.com "
                                    CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, "
                                    CNN.rsCmdCorreo!Titulo = "° MXAPD Req. " & IdReq & " Automatica Max x Min."
                                    
                                    Cadena4 = "<br><br>Se Solicta la Autorizacion. <br>" & Cadena4
                                    
                                    CNN.rsCmdCorreo!Mensaje = Cadena4
                                    CNN.rsCmdCorreo!Aplicacion = "Alm Ref"
                                    CNN.rsCmdCorreo!Prioridad = 2
                                CNN.rsCmdCorreo.Update
                                CNN.rsCmdCorreo.Close
                        End If
                    
                    
                ElseIf CNN.rsCmdExist!CriticoenCero = "S" Then
                                        
                    '''If CNN.rsCmdExist!ref_Critico = 1 And CNN.rsCmdExist!ref_max = 0 And CNN.rsCmdExist!ind_exi = CNN.rsCmdExist!ref_reorden Then
                    
                    If CNN.rsCmdExist!ref_Critico = 1 And CNN.rsCmdExist!ind_exi = 0 Then
                    
                            Articulo = CNN.rsCmdExist!Ref_Id
                            DescArt = CNN.rsCmdExist!ref_desc
                            
                            'Busca ultima compra
                            RefId = Me.LstDetalles.ListItems(Detalle).SubItems(2)
                            CNN.CmdBuscaUltimacompra (RefId)
                            If CNN.rsCmdBuscaUltimaCompra.EOF <> True Then
                                DptoId = CNN.rsCmdBuscaUltimaCompra!deptoid
                                TipoReq = CNN.rsCmdBuscaUltimaCompra!oci_tipo
                            Else
                                DptoId = "03"
                                TipoReq = "N"
                            End If
                            CNN.rsCmdBuscaUltimaCompra.Close
                            
                            CNN.CmdUsuariosCompras (DptoId)
                            If CNN.rsCmdUsuariosCompras.EOF <> True Then
                                CorreoGerente = CNN.rsCmdUsuariosCompras!cc_Correo
                            Else
                                CorreoGerente = ""
                            End If
                            CNN.rsCmdUsuariosCompras.Close
                            
                            Para = CorreoGerente
                            
                            De = "raul.chavez@schott.com"
                            Asunto = "Notificación Refacciones Criticas " ''& "Test TI"
                            
                            If BandReq = False Then
                                'Encabezado Req
                                CNN.rsCmdMaxReq.Open
                                If IsNull(CNN.rsCmdMaxReq!m) Then
                                    IdReq = 1
                                Else
                                    IdReq = CNN.rsCmdMaxReq!m + 1
                                End If
                                CNN.rsCmdMaxReq.Close
                                
                                CNN.CmdEncReq (IdReq)
                                If CNN.rsCmdEncReq.EOF = True Then
                                    CNN.rsCmdEncReq.AddNew
                                        CNN.rsCmdEncReq!IdReq = IdReq
                                        CNN.rsCmdEncReq!fecha_solicitud = Date
                                        CNN.rsCmdEncReq!hora_solicitud = Time
                                        CNN.rsCmdEncReq!fecha_req = Date + 1
                                        
                                        CNN.rsCmdEncReq!EmpleadoId = SolicitaId
                                        CNN.rsCmdEncReq!DptoId = DptoId
                                        CNN.rsCmdEncReq!uso_insumo = "Cubrir inventario (Automatico°) Max y Critico en CERO"
                                        
                                        CNN.rsCmdEncReq!idEmpleado = EmpleadoId 'Comprador
                                        CNN.rsCmdEncReq!req_status = "S"
                                        CNN.rsCmdEncReq!req_tipo = TipoReq
                                        CNN.rsCmdEncReq!id_oficina = "PSLP"
                                    CNN.rsCmdEncReq.Update
                                    
                                    
                                    BandReq = True
                                    DetReq = 1
                                    
                                    
                                Else
                                    
                                End If
                                CNN.rsCmdEncReq.Close
                            End If
                            'Detalles Reqs
                            CNN.CmdDetReq (IdReq), (DetReq)
                            If CNN.rsCmdDetReq.EOF = True Then
                                CNN.rsCmdDetReq.AddNew
                                    CNN.rsCmdDetReq!IdReq = IdReq
                                    CNN.rsCmdDetReq!Item = DetReq
                                    CNN.rsCmdDetReq!Ref_Id = Articulo
                                    CNN.rsCmdDetReq!Cantidad = (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                    CNN.rsCmdDetReq!cantidad_pend = (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                    CNN.rsCmdDetReq!det_status = "A"
                                    CNN.rsCmdDetReq!det_obs = ""
                                CNN.rsCmdDetReq.Update
                                Cadena4 = Cadena4 & "<br><br>No Art: " & Articulo
                                Cadena4 = Cadena4 & "<br>Descripcion: " & DescArt
                                Cadena4 = Cadena4 & "<br>Cantidad: " & (CNN.rsCmdExist!ref_max - CNN.rsCmdExist!ind_exi)
                                Cadena4 = Cadena4 & "<br>UM: " & CNN.rsCmdExist!ref_um
                                DetReq = DetReq + 1
                            Else
                                MsgBox "El detalle ya existe."
                            End If
                            CNN.rsCmdDetReq.Close
                            
                            
                              '''envia Correos
                            Para = ""
                            If DptoId = "04" Or DptoId = "06" Or DptoId = "05" Or DptoId = "10" Then
                                    Para = "antonio.marin@schott.com; "
                                    CNN.CmdBuscaCorreosMtto (107), (115), (119), (118)
                                    Do While CNN.rsCmdBuscaCorreosMtto.EOF <> True
                                                Para = Para & "; " & CNN.rsCmdBuscaCorreosMtto!Emp_Usr_Correo
                                            CNN.rsCmdBuscaCorreosMtto.MoveNext
                                    Loop
                                    CNN.rsCmdBuscaCorreosMtto.Close
                                    
                            Else
                                CNN.CmdBuscaCorreoGte (DptoId)
                                Do While CNN.rsCmdBuscaCorreoGte.EOF <> True
                                        Para = Para & CNN.rsCmdBuscaCorreoGte!Emp_Usr_Correo & "; "
                                    CNN.rsCmdBuscaCorreoGte.MoveNext
                                Loop
                                CNN.rsCmdBuscaCorreoGte.Close
                            End If
                            
                             If Para <> "" And DetReq <> 0 And IdReq <> 0 Then
                                    Cadena4 = Cadena4 & "<br><br><br>Elaboro: " & Nombre & "<br>Fecha y Hora: " & Now
                                    'envia correos
                                    CNN.rsCmdCorreo.Open
                                    CNN.rsCmdCorreo.AddNew
                                        CNN.rsCmdCorreo!De = De 'Des
                                        CNN.rsCmdCorreo!Para = Para '"helios.ireta@schott.com, blas.reyes@schott.com, " 'Para
                                        CNN.rsCmdCorreo!CC = ConCopia & "manuel.lozano@schott.com, teresa.montemayor@schott.com, raul.chavez@schott.com "
                                        CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, "
                                        CNN.rsCmdCorreo!Titulo = "° MXAPD Req. Automatica No. " & IdReq & " Max y Critico en CERO."
                                        
                                        Cadena4 = "<br><br>Se Solicta la Autorizacion. <br>" & Cadena4
                                        
                                        caden4 = "Los siguientes articulos llegaron al punto de reorden, se elaboro la Req. No. " & IdReq & Cadena4 & "<br>la cual esta pendiente por autorizar.<br><br><br>" & caden4
                                        
                                        CNN.rsCmdCorreo!Mensaje = Cadena4
                                        CNN.rsCmdCorreo!Aplicacion = "Alm Ref"
                                        CNN.rsCmdCorreo!Prioridad = 2
                                    CNN.rsCmdCorreo.Update
                                    CNN.rsCmdCorreo.Close
                            End If
                        
                            
                            
''''                            Para = "raul.chavez@schott.com; pilar.aguilar@schott.com, javier.barron@schott.com, alfonso.diaz@schott.com, manuel.lozano@schott.com,"
''''                            ConCopia = ""
''''                            'Para = "helios.ireta@schott.com"
''''                            If BandReq = True And Cadena4 <> "" Then
''''                                CNN.rsCmdCorreo.Open
''''                                CNN.rsCmdCorreo.AddNew
''''                                    CNN.rsCmdCorreo!De = De 'Des
''''                                    CNN.rsCmdCorreo!Para = Para '"helios.ireta@schott.com, blas.reyes@schott.com, " 'Para
''''                                    CNN.rsCmdCorreo!CC = ConCopia & "manuel.lozano@schott.com, teresa.montemayor@schott.com, raul.chavez@schott.com "
''''                                    CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, "
''''                                    CNN.rsCmdCorreo!Titulo = "° MXAPD Req. Automatica No. " & IdReq & " Max y Critico en CERO."
''''                                    Cadena4 = "Los siguientes articulos llegaron al punto de reorden, se elaboro la Req. No. " & IdReq & Cadena4 & "<br>la cual esta pendiente por autorizar."
''''                                    CNN.rsCmdCorreo!Mensaje = Cadena4
''''                                    CNN.rsCmdCorreo!Aplicacion = "Alm Ref"
''''                                    CNN.rsCmdCorreo!Prioridad = 2
''''                                CNN.rsCmdCorreo.Update
''''                                CNN.rsCmdCorreo.Close
''''                            End If
                            
                            
                            
                    End If
                End If
            
            
            End If
            
            
            CNN.rsCmdExist.Update
        Else
            MsgBox "El Articulo NO existe.", vbInformation, MSG
        End If
        CNN.rsCmdExist.Close
    Else
        MsgBox "El Detalle ya existe.", vbInformation, MSG
        CNN.rsCmdDet_Sal.Close
        Exit Sub
    End If
    CNN.rsCmdDet_Sal.Close
    Detalle = Detalle + 1
    i = i + 1
Loop


''Para = "raul.chavez@schott.com; pilar.aguilar@schott.com, javier.barron@schott.com, alfonso.diaz@schott.com, manuel.lozano@schott.com,"
''
''ConCopia = ""
''
'''Para = "helios.ireta@schott.com"
''If BandReq = True And Cadena4 <> "" Then
''    CNN.rsCmdCorreo.Open
''    CNN.rsCmdCorreo.AddNew
''        CNN.rsCmdCorreo!De = De 'Des
''        CNN.rsCmdCorreo!Para = Para '"helios.ireta@schott.com, blas.reyes@schott.com, " 'Para
''        CNN.rsCmdCorreo!CC = ConCopia & "manuel.lozano@schott.com, teresa.montemayor@schott.com, raul.chavez@schott.com "
''        CNN.rsCmdCorreo!CCO = "helios.ireta@schott.com, blas.reyes@schott.com, "
''        CNN.rsCmdCorreo!Titulo = "° MXAPD Req. Automatica Punto de Reorden No. " & IdReq
''        '<marquee style="background-color: #000080;" direction="right" loop="20" width="75%">
''        'Esto es una marquesina móvil
''        '</marquee>
''        '<MARQUEE BGCOLOR='#FFFF1D'>¡¡¡Importante Autorizar esta Requicisión!!!</MARQUEE>
''
''        Cadena4 = "Los siguientes articulos llegaron al punto de reorden, se elaboro la Req. No. " & IdReq & Cadena4 & "<br>la cual esta pendiente por autorizar."
''
''        CNN.rsCmdCorreo!Mensaje = Cadena4
''        CNN.rsCmdCorreo!Aplicacion = "Alm Ref"
''        CNN.rsCmdCorreo!Prioridad = 2
''    CNN.rsCmdCorreo.Update
''    CNN.rsCmdCorreo.Close
''End If


MsgBox "La información de guardo correctamente.", vbInformation, MSG


'     CNN.rsCmdTempAlmRefSal.Open
'     CNN.rsCmdTempAlmRefSal.AddNew
'            CNN.rsCmdTempAlmRefSal!no = 0
'            CNN.rsCmdTempAlmRefSal!TipoAlm = "Refacciones"
'            CNN.rsCmdTempAlmRefSal!Ref_Id = 0
'            CNN.rsCmdTempAlmRefSal!ref_desc = "Usuario: " & NombreUsuarioAut & ""
'            CNN.rsCmdTempAlmRefSal!Cantidad = 0
'            CNN.rsCmdTempAlmRefSal!NoParte = ""
'            CNN.rsCmdTempAlmRefSal!ref_um = ""
'            CNN.rsCmdTempAlmRefSal!Fecha = ""
'     CNN.rsCmdTempAlmRefSal.Update
'     CNN.rsCmdTempAlmRefSal.Close


'    FrmValeSalida.FreEtiqueta.Visible = True
'    FrmValeSalida.FreEtiqueta.Refresh
'    FrmValeSalida.Timer1.Enabled = True
'
'    FrmValeSalida.FreEtiqueta.Left = 60
'    FrmValeSalida.FreEtiqueta.Top = 1350
'
'
'    FrmValeSalida.LblUsuario.Caption = "Usuario: " & vbNewLine & NombreUsuarioAut
'    FrmValeSalida.LblUsuario.Refresh
'
'    'FrmValeSalida.LblMsg.Caption = "Usuario: " & NombreUsuarioAut & ""
'    FrmValeSalida.FreEtiqueta.Visible = True
'    FrmValeSalida.FreEtiqueta.Refresh
'    FrmValeSalida.Timer1.Enabled = True

      CNN.CmdBuscaEmpleadoxId (IdUsuarioAut)
      If CNN.rsCmdBuscaEmpleadoxId.EOF <> True Then
            FrmValeSalida.LblUsuario.Caption = "Usuario: " & vbNewLine & CNN.rsCmdBuscaEmpleadoxId!EMP_NOMBRE & vbNewLine & vbNewLine & "La información se guardo correctamente."
            FrmValeSalida.LblUsuario.Visible = True
            FrmValeSalida.LblUsuario.Refresh
            FrmValeSalida.FreMsg.Visible = True
            
            FrmValeSalida.ShpVerde.FillColor = vbGreen
            FrmValeSalida.ShpVerde.Visible = True
            FrmValeSalida.ShpVerde.Refresh
            FrmValeSalida.Timer3.Enabled = True
            
      Else
            FrmValeSalida.LblUsuario.Caption = "CONTACTE A SISTEMAS " & vbNewLine & vbNewLine & "Por favor."
            FrmValeSalida.ShpVerde.FillColor = vbRed
            FrmValeSalida.ShpVerde.Visible = True
            FrmValeSalida.ShpVerde.Refresh
            FrmValeSalida.FreMsg.Visible = True
            'FrmValeSalida.TimerVerde.Enabled = True
            
      End If
      CNN.rsCmdBuscaEmpleadoxId.Close
     
      
      FrmValeSalida.Timer1.Enabled = True
      Band = True





MsgBox "La información de guardo correctamente.", vbExclamation, MSG
BandFrm = False
'FrmRepValeSal.Show 1

Unload FrmValeSalida

Unload Me


'BandFrm = False
'FrmRepValeSal.Show 1
'Unload Me



End Sub



