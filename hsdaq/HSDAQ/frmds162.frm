VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "DS1621"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form5"
   ScaleHeight     =   3000
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "DS1621"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Read high resolution temperature"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Read temperature"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command3_Click()
Label1.Caption = "Temperature: " & DS1621_ReadTemp(Val(Text2.Text), IIf(Check1.Value = 1, 1, 0))
End Sub

Sub Command4_Click()
Unload Me
End Sub

Sub Form_Load()
DS1621_Init 0
End Sub

Private Sub Text2_Change()
If Val(Text2.Text) < 0 Or Val(Text2.Text) > 7 Then
  MsgBox "Minimum 0, Maximum 7 !", vbOKOnly, "Error"
  Text2.Text = 0
End If
End Sub
