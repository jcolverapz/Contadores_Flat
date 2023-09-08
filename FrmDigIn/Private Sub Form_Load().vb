Private Sub Form_Load()
Dim noTarjeta As Integer

'      If IsNumeric(Right(App.EXEName, 1)) Then
'          No_Linea_EXE = Right(App.EXEName, 1)
'      Else
'          MsgBox "El Nombre del exe no es correcto."
'          End
'      End If

    'No_Linea_EXE = Right("VisorContadoresL1", 1)
    'No_Linea_EXE = Right("VisorContadoresL2", 1)
    No_Linea_EXE = Right("VisorContadoresL3", 1)


''''VisorContadoresL1    --->  NOMBRE
''''VisorContadoresL2    --->  NOMBRE
''''VisorContadoresL3    --->  NOMBRE

'MsgBox No_Linea_EXE

    Select Case No_Linea_EXE
        Case 1
            noTarjeta = 0             ' Board number    'LINEA 1
            NoCajaAzul = 111
            No_Linea = 1
            No_Linea_EXE = 1

            '    Const BoardNum = 0              ' Board number    'LINEA 1
           ' BoardNum = 1

        Case 2
            noTarjeta = 1              ' Board number    'LINEA 2
            NoCajaAzul = 221
            No_Linea = 2
            No_Linea_EXE = 2

              'Const BoardNum = 1              ' Board number    'LINEA 2
           ' BoardNum = 1

         Case 3
            noTarjeta = 0              ' Board number    'LINEA 3
            NoCajaAzul = 331
            No_Linea = 3
            No_Linea_EXE = 3

             '  Const BoardNum = 0              ' Board number    'LINEA 3

           '  BoardNum = 0

        Case 4
            noTarjeta = 0              ' Board number    'LINEA 4
            NoCajaAzul = 441
            No_Linea = 4
            No_Linea_EXE = 4

             ' Board number    'LINEA 4

           ' BoardNum = 0

    End Select

    Me.Label1.Caption = "Linea: " & No_Linea

    Linea = "Linea " & No_Linea


   Dim ReportError As Long, HandleError As Long
   Dim I As Integer, PortName As String


   ' declare revision level of Universal Library
   ULStat& = cbDeclareRevision(CURRENTREVNUM)

   ' Initiate error handling
   '  activating error handling will trap errors like
   '  bad channel numbers and non-configured conditions.
   '  Parameters:
   '    PRINTALL    :all warnings and errors encountered will be printed
   '    DONTSTOP    :if an error is encountered, the program will not stop,
   '                 errors must be handled locally

   ReportError = DONTPRINT
   HandleError = DONTSTOP
   ULStat& = cbErrHandling(ReportError, HandleError)
   If ULStat& <> 0 Then Stop
   SetDigitalIODefaults ReportError, HandleError

   ' If cbErrHandling& is set for STOPALL or STOPFATAL during the program
   ' design stage, Visual Basic will be unloaded when an error is encountered.
   ' We suggest trapping errors locally until the program is ready for compiling
   ' to avoid losing unsaved data during program design.  This can be done by
   ' setting cbErrHandling options as above and checking the value of ULStat&
   ' after a call to the library. If it is not equal to 0, an error has occurred.

   'determine if digital port exists, its capabilities, etc
   PortType = PORTIN
   NumPorts = FindPortsOfType(BoardNum, PortType, ProgAbility, PortNum, NumBits, FirstBit)
   If NumBits > 8 Then NumBits = 8
   For I% = NumBits To 7
       lblShowBitVal(I%).Visible = False
       lblShowBitNum(I%).Visible = False
   Next I%

   If NumPorts < 1 Then
       lblInstruct.Caption = "Board " & Format(BoardNum, "0") & _
         " has no compatible digital ports."
   Else
       ' if programmable, set direction of port to input
       ' configure the first port for digital input
       '  Parameters:
       '    PortNum        :the input port
       '    Direction      :sets the port for input or output

       If ProgAbility = DigitalIO.PROGPORT Then
           Direction = DIGITALIN
           ULStat = cbDConfigPort(BoardNum, PortNum, Direction)
           If Not (ULStat = 0) Then Stop
       End If
       PortName = GetPortString(PortNum)
       lblInstruct.Caption = "You may change the value read by applying " & _
       "a TTL high or TTL low to digital inputs on " & PortName & _
       " on board " & Format(BoardNum, "0") & "."
       lblBitNum.Caption = "The first " & Format(NumBits, "0") & " bits are:"
       tmrReadInputs.Enabled = True
   End If



    For NoContador = 0 To NumBits - 1
        BandContador(NoContador) = False
    Next NoContador

    'Shell "C:\Legacy\VisorContadoresL1.exe", vbNormalFocus
End Sub

