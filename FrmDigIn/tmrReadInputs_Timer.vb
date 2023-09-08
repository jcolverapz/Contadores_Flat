
Private Sub tmrReadInputs_Timer()

    'original  200
    '1500  23/01/18


   Dim BitPort As Long, I As Long
   Dim FirstBit As Long, BitNum As Long, BitValue As Integer
   Dim BitPortName As String
   
   ' read the input bits from the ports and display
   
   '  Parameters:
   '    BoardNum    :the number used by CB.CFG to describe this board
   '    PortType    :must be FIRSTPORTA or AUXPORT
   '    BitNum&     :the number of the bit to read from the port
   '    BitValue&   :the value read from the port
   
   ' For boards whose first port is not FIRSTPORTA (such as the USB-ERB08
   ' and the USB-SSR08) offset the BitNum by FirstBit
   
   BitPort& = AUXPORT
   If PortNum& > AUXPORT Then BitPort& = FIRSTPORTA
   
   For I& = 0 To NumBits - 1
      BitNum& = I&
      ULStat& = cbDBitIn(BoardNum, BitPort&, FirstBit& + BitNum&, BitValue%)
      If ULStat& <> 0 Then Stop
      
      lblShowBitVal(I&).Caption = Format$(BitValue%, "0")
      
      
      'If lblShowBitVal(I&).Caption <> 0 And VContadores(I&) = False Then                ' con un vector de banderas
      
      If lblShowBitVal(I&).Caption <> "0" And VContadores(I&) = False Then
      
'      If lblShowBitVal(I&).Caption <> 0 And VContadores(I&) = False Then                  ' con un vector de banderas
            lblShowBitVal(I&).BackColor = vbGreen
            Contador = I&
            VContadores(I&) = True
            
'''            If I& = 1 Then
'''                Band_1 = True
'''            End If
                        
          '  NoCajaAzul = 331
            Call FrmVisorContadores3Lineas.UpdateConteos(Contador, NoCajaAzul)
      '      VContadores(I&) = False
        Else
            
                If lblShowBitVal(I&).Caption = "0" Then
                       VContadores(I&) = False
                       lblShowBitVal(I&).BackColor = vbWhite
                End If
           
      End If
      
   Next I&
   BitPortName$ = GetPortString(BitPort&)
   lblBitVal.Caption = BitPortName$ & ", bit " & Format(FirstBit&, "0") & _
   " - " & Format(FirstBit& + (NumBits - 1), "0") & " values:"


FrmVisorContadores3Lineas.TimerLstConteos.Enabled = True

End Sub

