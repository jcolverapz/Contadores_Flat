Private Sub tmrReadInputs_Timer()

    'original  200
    '1500  23/01/18

   Dim BitPort As Long, I As Long
   Dim FirstBit As Long, BitNum As Long, BitValue As Integer
   Dim BitPortName As String

        '    read the input bits from the ports and display

        '     Parameters:
        '       BoardNum    :the number used by CB.CFG to describe this board
        '       PortType    :must be FIRSTPORTA or AUXPORT
        '       BitNum&     :the number of the bit to read from the port
        '       BitValue&   :the value read from the port

        '    For boards whose first port is not FIRSTPORTA (such as the USB-ERB08
        '    and the USB-SSR08) offset the BitNum by FirstBit

   'BitPort& = FIRSTPORTA
   'If PortNum& > AUXPORT Then BitPort& = FIRSTPORTA

       For I& = 1 To 5

        'BitNum& = I&

        ULStat& = cbDBitIn(0, FIRSTPORTA, 1, BitValue%)
            lblShowBitVal(1).Caption = BitValue%


        If lblShowBitVal(I&).Caption <> "0" And BandContador(I&) = False Then


           ' lblShowBitVal(I&).BackColor = vbGreen
            'Contador = I&
            'CntSensores(I&)=

            lblShowBitVal(I&).Caption = lblShowBitVal(I&).Caption + 1

            BandContador(I&) = True

        Else

                If lblShowBitVal(I&).Caption = "0" Then
                       BandContador(I&) = False
                       lblShowBitVal(I&).BackColor = vbWhite
                End If

        End If

       Next I&


    'Debug.Print "lblShowBitVal(I&).Caption :" & lblShowBitVal(I&).Caption
    'If ULStat& <> 0 Then Stop
    'ULStat& = cbDeclareRevision(CURRENTREVNUM)
    'lblShowBitVal(I&).Caption = Format$(BitValue%, "0")
    'If lblShowBitVal(I&).Caption <> 0 And VContadores(I&) = False Then                ' con un vector de banderas




   'BitPortName$ = GetPortString(BitPort&)
  ' lblBitVal.Caption = BitPortName$ & ", bit " & Format(FirstBit&, "0") & _
   " - " & Format(FirstBit& + (NumBits - 1), "0") & " values:"

    'FrmVisorContadores3Lineas.TimerLstConteos.Enabled = True
End Sub

