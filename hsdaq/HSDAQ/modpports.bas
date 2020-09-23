Attribute VB_Name = "Module2"

Sub DS1621_Init(ByVal Address As Integer)
DS1621_Stop
DS1621_SndCmd &HAC, Address, &H2
DS1621_SndCmd &HEE, Address, -1
DS1621_Stop
End Sub

'
Function DS1621_ReadCnt(ByVal Address As Integer) As Integer
DS1621_Start
DS1621_SndCmd &HA8, Address, -1
rep = DS1621_TXByte(&H91 + Address) 'rep is the ack bit
DS1621_ReadCnt = DS1621_RXByte(1)
DS1621_Stop
End Function

'
'Look at datasheet of ds1621 for more info on DS1621_ReadSlp
Function DS1621_ReadSlp(ByVal Address As Integer) As Integer
DS1621_Start
DS1621_SndCmd &HA9, Address, -1
rep = DS1621_TXByte(&H91 + Address)
DS1621_ReadSlp = DS1621_RXByte(1)
End Function

Function DS1621_ReadTemp(ByVal Address As Integer, ByVal HighRes) As Double
    DS1621_SndCmd &HAA, Address, -1
    DS1621_Start
    rep = DS1621_TXByte(&H91 + Address)
    'reading first byte with acknowledge bit
    temperature_int = DS1621_RXByte(1)
    'reading second byte without acknowledge bit
    temperature_frac = DS1621_RXByte(0)
    'add frac to int which give "xx.x" format
    temperature_int = IIf(temperature_frac > 0, temperature_int + 0.5, temperature_int)
    'if high res selected
    If HighRes Then
        cnt = DS1621_ReadCnt(Address)
        slp = DS1621_ReadSlp(Address)
        'calculate high res formulas
        temperature_int = temperature_int - 0.25 + ((slp - cnt) / slp)
    End If
    DS1621_Stop
    DS1621_ReadTemp = temperature_int
    
End Function

'
'
'clock in a bit
Function DS1621_RXBit() As Integer
Dim retval As Integer
    'set data pin high to enable input
    Out settings_.parallelport, AllBitPort(settings_.DS1621PinDataO)
    'Clock hi
    Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) Or AllBitPort(settings_.DS1621PinDataO)
    'get parallel port data register byte
    datavalue% = Inp(settings_.parallelport + 1)
    'clock low
    Out settings_.parallelport, 0
    'set proper bit value to return
    'DS1621DataBit is the bit number corresponding to the input pin
    'used to read the ds1621. Settings windows available
    retval = IIf((datavalue% And (AllBitPort(settings_.DS1621DataPin))), 1, 0)
    DS1621_RXBit = retval
        
End Function

'
''******************************************
'this will call 8 times the DS1621_RXBit function
'and return and integer value
Function DS1621_RXByte(acknowledge As Integer) As Integer
   Dim i As Integer
   Dim retval As Integer
  'get 8 bit and convert to integer as we read
   For i = 7 To 0 Step -1
       retval = retval + DS1621_RXBit() * (2 ^ i)
   Next
  'if requested to send acknowledge
  If acknowledge Then DS1621_TXBit_0

  DS1621_RXByte = retval
End Function

Sub DS1621_SndCmd(ByVal cmds As Integer, ByVal Address As Integer, ByVal datas)

   DS1621_Start
   rep = DS1621_TXByte(&H90 + Address)
   rep = DS1621_TXByte(cmds)
   If datas > -1 Then rep = DS1621_TXByte(datas)

End Sub

Sub DS1621_Start()
 Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) Or AllBitPort(settings_.DS1621PinDataO) 'clock hi
 Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) 'data lo
End Sub

Sub DS1621_Stop()
Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) 'clock hi
Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) Or AllBitPort(settings_.DS1621PinDataO) 'data hi
End Sub

Sub DS1621_TXBit_0()
Out settings_.parallelport, 0 'data low
Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) 'clock high
Out settings_.parallelport, 0 'clok low
End Sub

Sub DS1621_TXBit_1()
Out settings_.parallelport, AllBitPort(settings_.DS1621PinDataO) 'data high
Out settings_.parallelport, AllBitPort(settings_.DS1621PinCLK) Or AllBitPort(settings_.DS1621PinDataO) 'clock high
Out settings_.parallelport, AllBitPort(settings_.DS1621PinDataO) 'clock low
End Sub

Function DS1621_TXByte(ByVal b As Integer) As Integer
  Dim i As Integer
  bn = ""
For i% = 7 To 0 Step -1
      If (b And (2 ^ i%)) Then
         DS1621_TXBit_1
         bn = bn & "1"
      Else
         DS1621_TXBit_0
         bn = bn & "0"
      End If
      DoEvents
Next i%
     DS1621_TXByte = DS1621_RXBit()
     
End Function

Sub Init74hc166()
 'ititialyse ic
 Out settings_.parallelport, AllBitPort(settings_.InPutPinPL) 'Enable parallel input
 Out settings_.parallelport, ALLDATAOFF           'all low
 Out settings_.parallelport, AllBitPort(settings_.InPutPinCLK) 'clock hi
 Out settings_.parallelport, AllBitPort(settings_.InPutPinPL) 'clock low and enable input
End Sub

'This function will read 8 bit from
'the parallel port status port pin 12 - S4 status bit 4
'The stupid 74hc166n give us inverted input which means
'LFS less significant bit first or eightth bit first if you prefer
'this is an I3C communication routine
Function Read_74HC166_Byte() As Integer
  Dim retval As Integer
  retval = 255 'invert status port value
  For x = 7 To 0 Step -1 'invert also for loop
    datavalue% = Inp(settings_.parallelport + 1) 'Read the status port
    'we are looking for bit 4 here which correspond to
    'the S4 status port pin which correspond to an input pin
    'thus if this bit is high then add this bit to retval
    If (datavalue% And (AllBitPort(settings_.InPutDataPin))) = 0 Then
      'calculation is inverted cuz LSB
       retval = retval - (2 ^ x)
    End If
    Out settings_.parallelport, AllBitPort(settings_.InPutPinCLK) Or AllBitPort(settings_.InPutPinPL) 'clock hi
    Out settings_.parallelport, AllBitPort(settings_.InPutPinPL) 'clock lo
  Next
  Read_74HC166_Byte = retval
End Function

Sub Set_74HC374_Output(ByVal DataPortV As Integer, ByVal ic As Integer)
'setting oputput data pins to be shifted
Out settings_.parallelport, DataPortV
'if first ic then clock low first else clock low second
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(IIf(ic = 0, 1, 0)))) 'clock low
'lock shifted value to 74hc374 output (clock back hi)
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)))
'close opened data pins
Out settings_.parallelport, 0
End Sub

Function TLC1543_GetDataV() As Long
Dim portdata As Integer
Dim pdata As String
Dim vdata As Double
portdata = 0
'CS low
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)))

readinbinary = ""

Do
    PortStatus = Inp(settings_.parallelport + 1)
    DoEvents
    'if end of conversion bit is high
    If PortStatus And AllBitPort(settings_.AInPutEOCPin) Then
        For i = 9 To 0 Step -1
            PortStatus = Inp(settings_.parallelport + 1)
            'if read bit is 1
            If PortStatus And AllBitPort(settings_.AInputDataPin) Then
              readinbinary = readinbinary & "1"
              portdata = portdata + (2 ^ i)
             Else
              readinbinary = readinbinary & "0"
            End If
            Out settings_.parallelport, AllBitPort(settings_.AInPutPinCLK) ' /* "CLK" high                       */
            Out settings_.parallelport, 0   ' /* "CLK" low
        Next
        Exit Do
     End If
Loop
'CS high
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)) Or AllBitPort(settings_.AInPutPinCS))

TLC1543_GetDataV = portdata
End Function

Sub TLC1543_SendAdd(ByVal add As Single)
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)))
bb = ""
For x = 0 To 9 'Clock in 4bit address + 6 more clocks
   If x < 4 Then
        If Mid(AIADDPORT(add), x + 1, 1) = "1" Then
          Out settings_.parallelport, AllBitPort(settings_.AInPutPinADD)  '  /* Yes, set "ADDR" pin high         */
          Out settings_.parallelport, AllBitPort(settings_.AInPutPinADD) Or AllBitPort(settings_.AInPutPinCLK)  ' /* "CLK" high                       */
          Out settings_.parallelport, AllBitPort(settings_.AInPutPinADD)  ' /* "CLK" low
          bb = bb & "1"

        Else
          
          GoTo SameElse
        End If
  Else
  
SameElse:
            bb = bb & "0"
          Out settings_.parallelport, 0  '  /* No or its 0, set "ADDR" pin low         */
          Out settings_.parallelport, AllBitPort(settings_.AInPutPinCLK) ' /* "CLK" high                       */
          Out settings_.parallelport, 0    ' /* "CLK" low
  End If
Next

Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)) Or AllBitPort(settings_.AInPutPinCS))

End Sub


