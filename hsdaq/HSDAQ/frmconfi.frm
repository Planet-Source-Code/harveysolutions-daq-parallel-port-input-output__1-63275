VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Settings"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   7290
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6960
      TabIndex        =   43
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   42
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "Server settings(TODO)"
      Height          =   1575
      Left            =   5760
      TabIndex        =   39
      Top             =   240
      Width           =   2655
      Begin VB.CheckBox Check3 
         Caption         =   "Accept connections"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text16 
         Height          =   405
         Left            =   240
         TabIndex        =   40
         Text            =   "Text16"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "LPT Port"
      Height          =   1695
      Left            =   6360
      TabIndex        =   19
      Top             =   4560
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "LPT 3"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Tag             =   "956"
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LPT 2"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Tag             =   "632"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LPT 1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Tag             =   "888"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "74HC166"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Text            =   "4"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Text            =   "2"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Text            =   "1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Text            =   "4"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NB of ICs :"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Clock :"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "PL :"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data in :"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "74HC374"
      Height          =   1575
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Second IC"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "#2 Select :"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "#1 Select :"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DS1621"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Text            =   "1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Text            =   "2"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Text            =   "5"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NB of ICs :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Clock :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Data in :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data out :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "TLC1543"
      Height          =   2175
      Left            =   2880
      TabIndex        =   13
      Top             =   1800
      Width           =   2775
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Text            =   "8"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Text            =   "4"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Text            =   "8"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Text            =   "8"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Text            =   "4"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CS :"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Address :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data in :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "CLock :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "EOC :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Info"
      Height          =   2175
      Left            =   5760
      TabIndex        =   44
      Top             =   1800
      Width           =   2655
      Begin VB.Label Label6 
         Caption         =   "Pins values are used for configuration but bit values are used to access registers. "
         Height          =   660
         Left            =   165
         TabIndex        =   48
         Top             =   375
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Value :1  2  4  8  16  32  64  128"
         Height          =   255
         Left            =   75
         TabIndex        =   47
         Top             =   1725
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "Bit      :0  1  2  3   4    5    6      7"
         Height          =   255
         Left            =   75
         TabIndex        =   46
         Top             =   1485
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "Bit values are :"
         Height          =   255
         Left            =   75
         TabIndex        =   45
         Top             =   1185
         Width           =   1935
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8400
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2610
      Left            =   90
      Picture         =   "frmconfi.frx":0000
      Top             =   4065
      Width           =   5520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port
Sub Command1_Click()
Unload Me
End Sub

Sub Command2_Click()
settings_.parallelport = Val(Port)
settings_.servertcpport = Text16.Text
settings_.nboutputics = Check1.Value
settings_.nbinputics = Text14.Text
settings_.nbds1621 = Text15.Text
settings_.OutPutPinCLK(0) = Text2.Text - 1
settings_.OutPutPinCLK(1) = Text3.Text - 1
settings_.InPutPinCLK = Text6.Text - 1
settings_.InPutPinPL = Text5.Text - 1
settings_.InPutDataPin = Text4.Text - 1
settings_.AInPutPinCLK = Text10.Text - 1
settings_.AInPutPinADD = Text9.Text - 1
settings_.AInPutPinCS = Text8.Text - 1
settings_.AInPutEOCPin = Text1.Text - 1
settings_.AInputDataPin = Text7.Text - 1
settings_.DS1621DataPin = Text12.Text - 1
settings_.DS1621PinDataO = Text11.Text - 1
settings_.DS1621PinCLK = Text13.Text - 1
settings_.CONTROPOFF = 11
savesettings

End Sub

Sub Form_Load()
'MsgBox UBound(AllBitPort)

Port = settings_.parallelport
Select Case Port
    Case 888: Option1(0).Value = True
    Case 632: Option1(1).Value = True
    Case 956: Option1(2).Value = True
End Select
Check1.Value = settings_.nboutputics
Text1.Text = settings_.AInPutEOCPin + 1
Text2.Text = settings_.OutPutPinCLK(0) + 1
Text3.Text = settings_.OutPutPinCLK(1) + 1
Text4.Text = settings_.InPutDataPin + 1
Text5.Text = settings_.InPutPinPL + 1
Text6.Text = settings_.InPutPinCLK + 1
Text7.Text = settings_.AInputDataPin + 1
Text8.Text = settings_.AInPutPinCS + 1
Text9.Text = settings_.AInPutPinADD + 1
Text10.Text = settings_.AInPutPinCLK + 1
Text12.Text = settings_.DS1621DataPin + 1
Text11.Text = settings_.DS1621PinDataO + 1
Text13.Text = settings_.DS1621PinCLK + 1
Text14.Text = settings_.nbinputics
Text15.Text = settings_.nbds1621
Text16.Text = settings_.servertcpport

End Sub

Sub Option1_Click(Index As Integer)
  Port = Option1(Index).Tag
End Sub


