VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Harvey Solutions DAQ"
   ClientHeight    =   1020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data aquisition board from Carl Harvey. harveysolutions@t2u.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image4 
      Height          =   4095
      Left            =   0
      Picture         =   "frmmain.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.Image Image3 
      Height          =   5850
      Left            =   0
      Picture         =   "frmmain.frx":1F6E
      Top             =   0
      Visible         =   0   'False
      Width           =   6585
   End
   Begin VB.Image Image2 
      Height          =   4215
      Left            =   0
      Picture         =   "frmmain.frx":4E44
      Top             =   0
      Visible         =   0   'False
      Width           =   7230
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Picture         =   "frmmain.frx":790B
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Menu daq_ 
      Caption         =   "DAQ"
      Begin VB.Menu inithard_ 
         Caption         =   "Initiate hardware"
      End
      Begin VB.Menu settingsm_ 
         Caption         =   "Settings..."
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu quit_ 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu testing_ 
      Caption         =   "Testing"
      Begin VB.Menu pporttest_ 
         Caption         =   "Parallel port registers"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu digitoutmnu_ 
         Caption         =   "2 x 8 Bit Digital output (16)"
         Begin VB.Menu schematic_ 
            Caption         =   "Schematic"
         End
         Begin VB.Menu digitout_ 
            Caption         =   "Tester"
         End
      End
      Begin VB.Menu digitinmnu_ 
         Caption         =   "2 x 8 Bit Digital input (16)"
         Begin VB.Menu smeatic2_ 
            Caption         =   "Schematic"
         End
         Begin VB.Menu digitin_ 
            Caption         =   "Tester"
         End
      End
      Begin VB.Menu ainputmnu_ 
         Caption         =   "11 x 10 Bit Analog input"
         Begin VB.Menu schematic3_ 
            Caption         =   "Schematic"
         End
         Begin VB.Menu ainput_ 
            Caption         =   "Tester"
         End
      End
      Begin VB.Menu temptestmnu_ 
         Caption         =   "Temperature (DS1621)"
         Begin VB.Menu schematic4_ 
            Caption         =   "Schematic"
         End
         Begin VB.Menu temptest_ 
            Caption         =   "Tester"
         End
      End
   End
   Begin VB.Menu help_ 
      Caption         =   "Help"
      Begin VB.Menu helpmnu_ 
         Caption         =   "Help on DAQ"
      End
      Begin VB.Menu aboutme_ 
         Caption         =   "About me"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command3_Click()
End
End Sub

Private Sub ainput__Click()
Form6.Show
End Sub

Private Sub digitin__Click()
Form3.Show
End Sub

Private Sub digitout__Click()
Form4.Show
End Sub

Sub Form_Load()

'was needed for 16bit version :) no vbcrlf in vb3 lol
CRLF = Chr(13) & Chr(10)

LoadSettings

'bit values to access parallel port register
AllBitPort = Split("1,1,2,4,8,16,32,64,128,64,128,32,16,2,8,4,8", ",")

 'these are tlc1543 input addresses
AIADDPORT = Split("0000,0001,0010,0011,0100,0101,0110,0111,1000,1001,1010,1011", ",")

inithard__Click
End Sub

Private Sub inithard__Click()
'Initiation todo here...

'this ensure output are locked to the 74hc374 (basic initiation)
Out settings_.parallelport + 2, invert(AllBitPort(settings_.OutPutPinCLK(0)) Or AllBitPort(settings_.OutPutPinCLK(1)))

End Sub

Private Sub pporttest__Click()
Form7.Show
End Sub

Sub quit__Click()
End
End Sub

Private Sub schematic__Click()
Label1.Visible = False
Image2.Visible = True: Image1.Visible = False: Image3.Visible = False: Image4.Visible = False
Me.Width = Image2.Width
Me.Height = Image2.Height + 800
End Sub

Private Sub schematic3__Click()
Label1.Visible = False
Image4.Visible = True: Image1.Visible = False: Image2.Visible = False: Image3.Visible = False
Me.Width = Image4.Width
Me.Height = Image4.Height + 800
End Sub

Private Sub schematic4__Click()
Label1.Visible = False
Image3.Visible = True: Image1.Visible = False: Image2.Visible = False: Image4.Visible = False
Me.Width = Image3.Width
Me.Height = Image3.Height + 800
End Sub

Sub settingsm__Click()
Form2.Show 1
End Sub

Private Sub smeatic2__Click()
Label1.Visible = False
Image1.Visible = True: Image2.Visible = False: Image3.Visible = False: Image4.Visible = False
Me.Width = Image1.Width
Me.Height = Image1.Height + 800
End Sub

Private Sub temptest__Click()
Form5.Show
End Sub
