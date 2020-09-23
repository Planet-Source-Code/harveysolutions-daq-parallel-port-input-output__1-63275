VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "TLC1543"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form6"
   ScaleHeight     =   3825
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "TLC1543cn"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Option1 
         Caption         =   "11"
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "10"
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Read from TLC1543"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Label3"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Label3"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "Label3"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Choose output to read from :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim addressToRead As Integer



'Dim PortStatus As Integer

Sub Command1_Click()
Unload Me
End Sub

Private Sub Command5_Click()



'SENDING ADDRESSS
TLC1543_SendAdd addressToRead
      
portdata = TLC1543_GetDataV()
vdata = (portdata / 1024) * 5

Label3.Caption = "Binary read in         : " & readinbinary
Label4.Caption = "Numeric value        : " & portdata
Label5.Caption = "Voltage read          : " & vdata
Label6.Caption = "LM35 conversion    : " & Format((vdata - 4.27) * 100, "00.00") & " deg. celcius."



End Sub

Sub Form_Load()

 'ensure all port closed
 Out settings_.parallelport, 0
 'vbout settings_.parallelport + 2, CONTROLP_OFF
  addressToRead = 0
End Sub

Sub Option1_Click(Index As Integer)
addressToRead = Index
End Sub


