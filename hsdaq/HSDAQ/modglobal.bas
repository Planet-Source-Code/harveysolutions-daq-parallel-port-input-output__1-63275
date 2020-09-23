Attribute VB_Name = "Module1"

Option Explicit

Global Const ERROR_TEMPERATURE_NOT_READ = 1000
Global Const ERROR_OPENING_PORT = 1
Global Const NO_ERROR = 0
Global Const CONTROLP_OFF = 11
Global CRLF As String
Global Const ALLDATAOFF = 0
Public AllBitPort
Public AIADDPORT
Global readinbinary As String


Type settingstp
     parallelport As Integer
     servertcpport As Integer
     nboutputics As Integer
     nbinputics As Integer
     nbds1621 As Integer
     OutPutPinCLK(1) As Integer
     InPutPinCLK As Integer
     InPutPinPL As Integer
     InPutDataPin As Integer
     AInPutPinCLK As Integer
     AInPutPinADD As Integer
     AInPutPinCS As Integer
     AInPutEOCPin As Integer
     AInputDataPin As Integer
     DS1621DataPin As Integer
     DS1621PinDataO As Integer
     DS1621PinCLK As Integer
     CONTROPOFF As Integer
End Type


Global settings_ As settingstp
Global CurrentTimer As Integer
Global TimerCount As Integer



Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)


Function b2d(ByVal binv As String) As Long

Dim bit As Integer
Dim Value As Integer
Dim counter As Integer

For counter = Len(binv) To 1 Step -1
    bit = Val(Mid(binv, counter, 1))

   If bit = 1 Then
         Value = Value + 2 ^ (Len(binv) - counter)
   End If

Next
   b2d = Value
End Function



Function d2b(ByVal nval As Integer) As String
Dim x, retval
For x = 0 To 7
   If nval And (2 ^ x) Then
     retval = retval & "1"
   Else
     retval = retval & "0"
   End If
Next
d2b = retval
End Function




Function invert(ByVal valtemp As Integer) As Integer
'revert the 3 control port that are inverted
valtemp = valtemp Xor 1 'Bit 0 - Pin 1
valtemp = valtemp Xor 2 'Bit 1 - pin 14
valtemp = valtemp Xor 8 'Bit 3 - pin 16
invert = valtemp
End Function

Sub LoadSettings()
Dim filep As Long, n
filep = FreeFile

On Error GoTo errload

'ReDim Settings(0)
If UCase(Dir(App.Path & "\hsdaq.st")) = "HSDAQ.ST" Then
Open App.Path & "\HSDAQ.st" For Random As #filep

    Get #filep, , settings_
    Close #filep

   Exit Sub
errload:
    MsgBox "Unable to load settings file"
    Resume Next
    

Else
  MsgBox "File not found !" & App.Path & "\hsdaq.st. Using default"
  savedefaultset
End If
End Sub

Sub LogToFile(logstr)
Dim Filep1 As Long
Filep1 = FreeFile
If Dir(App.Path & "\RelayConfigLog") = "RelayConfigLog" Then
  Open App.Path & "\RelayConfigLog" For Append As #Filep1
Else
  Open App.Path & "\RelayConfigLog" For Output As #Filep1
  Print #Filep1, "************   Relay Log File   ****************"
End If

Print #Filep1, logstr
Close #Filep1
End Sub

Sub savedefaultset()
settings_.parallelport = 888
settings_.servertcpport = 45
settings_.nboutputics = 1
settings_.nbinputics = 4
settings_.nbds1621 = 1
settings_.OutPutPinCLK(0) = 1
settings_.OutPutPinCLK(1) = 4
settings_.InPutPinCLK = 2
settings_.InPutPinPL = 1
settings_.InPutDataPin = 16
settings_.AInPutPinCLK = 8
settings_.AInPutPinADD = 4
settings_.AInPutPinCS = 8
settings_.AInPutEOCPin = 64
settings_.AInputDataPin = 8
settings_.DS1621DataPin = 32
settings_.DS1621PinDataO = 1
settings_.DS1621PinCLK = 2
settings_.CONTROPOFF = 11
savesettings
End Sub

Sub savesettings()

If Dir(App.Path & "\hsdaq.st") = "HSDAQ.ST" Then Kill App.Path & "\hsdaq.st"

Dim filep
filep = FreeFile

Open App.Path & "\hsdaq.st" For Random As #filep
Put #filep, , settings_
Close #filep

End Sub



