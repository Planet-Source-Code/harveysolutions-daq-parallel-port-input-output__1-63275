VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "74HC166"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form3"
   ScaleHeight     =   2265
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "74HC166n"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Text            =   "4"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Read from 74HC166n"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Number of IC to read :"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   5775
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Example of reading from 1 to 8 74HC166n IC in series."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nbics As Integer

'
Sub Command1_Click()
 
 Dim retval As Integer
 'remove previously loaded labels
 For x = nbics - 1 To 1 Step -1
    Unload Label1(x)
 Next
 'get nb of ic to read
 nbics = Val(Text1.Text)
 
 Init74hc166
 
 Label1(0).Visible = True
 'read all specefyied ics
 For x = 1 To nbics
    retval = Read_74HC166_Byte() 'read status port
    If x > 1 Then 'if its not the loaded label 0 then load go for load
        Load Label1(x - 1)
        Label1(x - 1).Top = Label1(x - 2).Top + 400
        Label1(x - 1).Left = Label1(x - 1).Left
        Label1(x - 1).Visible = True
    End If
 
    'show read value in integer and binary format
    Label1(x - 1).Caption = "IC #" & x & " value read - Integer->" & retval & " - Binary->" & d2b(retval)

 Next

 'ajusting view
 Me.Height = 2700 + (nbics) * 400 '400 for button height
 Frame1.Height = 1425 + (nbics - 1) * 400
 Command2.Top = Me.Height - 875 'move close button

End Sub

Sub Command2_Click()
Unload Me
End Sub

Sub Form_Load()
  nbics = 0
End Sub


Private Sub Text1_Change()
If Val(Text1.Text) < 1 Or Val(Text1.Text) > 4 Then
  MsgBox "Minimum 1, Maximum 4 !", vbOKOnly, "Error"
  Text1.Text = 1
End If
End Sub
