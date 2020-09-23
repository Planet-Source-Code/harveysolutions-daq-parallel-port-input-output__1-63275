VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "74HC374"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form4"
   ScaleHeight     =   2520
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "74HC374cn"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "Send to second IC"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Send to first IC"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8"
         Height          =   375
         Index           =   7
         Left            =   4080
         TabIndex        =   11
         Tag             =   "128"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7"
         Height          =   375
         Index           =   6
         Left            =   3600
         TabIndex        =   10
         Tag             =   "64"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   9
         Tag             =   "32"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5"
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   8
         Tag             =   "16"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4"
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   7
         Tag             =   "8"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   6
         Tag             =   "4"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2"
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Tag             =   "2"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Tag             =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Reset 74hc374n"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Write to 74hc374n"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Value sent :"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Output to raise (High)"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'** Variables **
Dim DataPortValue As Integer


Sub Check1_Click(Index As Integer)
DataPortValue = DataPortValue Xor Check1(Index).Tag
End Sub

Sub Command2_Click()
Unload Me
End Sub

'Write to 74hc374n sub
Sub Command5_Click()

Set_74HC374_Output DataPortValue, IIf(Option1(0).Value, 0, 1)
'show sent data
Label4.Caption = DataPortValue
Out settings_.parallelport, 0
End Sub

Sub Command6_Click()
For i = 0 To 7
  Check1(i).Value = 0
Next

Set_74HC374_Output 0, IIf(Option1(0).Value, 0, 1)
End Sub

Sub Form_Load()
  DataPortStatus = 0
  Label4.Caption = DataPortStatus
  
  Option1(1).Enabled = settings_.nboutputics
  
End Sub


