VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Parallel port registers testing"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form7"
   ScaleHeight     =   6855
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Inport - Outport"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   5475
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   3060
         TabIndex        =   52
         Top             =   1710
         Width           =   2355
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   54
            Top             =   450
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   900
            TabIndex        =   53
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label6 
            Caption         =   "Integer :"
            Height          =   285
            Left            =   90
            TabIndex        =   56
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Binary :"
            Height          =   285
            Left            =   90
            TabIndex        =   55
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Height          =   825
         Left            =   3060
         TabIndex        =   40
         Top             =   2895
         Width           =   2355
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   41
            Top             =   450
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   288
            Left            =   960
            TabIndex        =   42
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Integer :"
            Height          =   285
            Left            =   90
            TabIndex        =   44
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Binary :"
            Height          =   285
            Left            =   90
            TabIndex        =   43
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   825
         Left            =   3060
         TabIndex        =   35
         Top             =   540
         Width           =   2355
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   37
            Top             =   450
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   870
            TabIndex        =   36
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "Integer :"
            Height          =   285
            Left            =   90
            TabIndex        =   39
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Binary :"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Write to D0 - D7"
         Height          =   345
         Left            =   90
         TabIndex        =   34
         Top             =   990
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1320
         TabIndex        =   33
         Text            =   "0"
         Top             =   630
         Width           =   852
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Read S3 - S7"
         Height          =   345
         Left            =   90
         TabIndex        =   32
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1350
         TabIndex        =   31
         Text            =   "0"
         Top             =   2970
         Width           =   852
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Write to C0 - C5"
         Height          =   345
         Left            =   90
         TabIndex        =   30
         Top             =   3330
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Revert inverted pins"
         Height          =   195
         Left            =   3390
         TabIndex        =   29
         Top             =   2655
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   75
         X2              =   5385
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bit 7 is inverted in view"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3450
         TabIndex        =   51
         Top             =   1470
         Width           =   2265
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   5400
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8 Bit Register values"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3540
         TabIndex        =   50
         Top             =   270
         Width           =   1875
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   5400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Data to send"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   49
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Data to send"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   48
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Status Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   46
         Top             =   1440
         Width           =   5325
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Control Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   45
         Tag             =   " "
         Top             =   2610
         Width           =   5325
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Data Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   47
         Top             =   255
         Width           =   5325
      End
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   372
      Left            =   4320
      TabIndex        =   27
      Top             =   6480
      Width           =   1212
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   2610
      Left            =   0
      Picture         =   "frmports.frx":0000
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   364
      TabIndex        =   0
      Top             =   3840
      Width           =   5520
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   5010
         TabIndex        =   26
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   3765
         TabIndex        =   25
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   2865
         TabIndex        =   24
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   22
         Left            =   885
         TabIndex        =   23
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   21
         Left            =   1245
         TabIndex        =   22
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "23"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   20
         Left            =   1605
         TabIndex        =   21
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   1965
         TabIndex        =   20
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   18
         Left            =   2325
         TabIndex        =   19
         Top             =   1275
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   2685
         TabIndex        =   18
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   3045
         TabIndex        =   17
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   3405
         TabIndex        =   16
         Top             =   1275
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   4125
         TabIndex        =   15
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   30
         Left            =   4485
         TabIndex        =   14
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   4845
         TabIndex        =   13
         Top             =   1275
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   31
         Left            =   705
         TabIndex        =   12
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   32
         Left            =   1065
         TabIndex        =   11
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   33
         Left            =   1785
         TabIndex        =   10
         Top             =   915
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   34
         Left            =   1425
         TabIndex        =   9
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   2145
         TabIndex        =   8
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   2505
         TabIndex        =   7
         Top             =   915
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   3225
         TabIndex        =   6
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   3585
         TabIndex        =   5
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   3945
         TabIndex        =   4
         Top             =   915
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   4305
         TabIndex        =   3
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4665
         TabIndex        =   2
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Female connector"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   18
         Left            =   2955
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   17
         Left            =   3315
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   24
         Left            =   795
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   23
         Left            =   1155
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   22
         Left            =   1515
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   21
         Left            =   1875
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   20
         Left            =   2235
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   19
         Left            =   2595
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   12
         Left            =   3675
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   11
         Left            =   4035
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   30
         Left            =   4395
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   10
         Left            =   4755
         Shape           =   3  'Circle
         Top             =   1245
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   31
         Left            =   615
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   32
         Left            =   975
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   33
         Left            =   1695
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   34
         Left            =   1335
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   8
         Left            =   2055
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   7
         Left            =   2415
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   6
         Left            =   2775
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   5
         Left            =   3135
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   4
         Left            =   3495
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   3
         Left            =   3855
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   2
         Left            =   4215
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   1
         Left            =   4575
         Shape           =   3  'Circle
         Top             =   885
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   9
         Left            =   4935
         Shape           =   3  'Circle
         Top             =   870
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
Private Sub Command1_Click()
Out settings_.parallelport, Val(Text1.Text)
ShowDataStat
End Sub

Private Sub Command2_Click()
ShowStatusStat
End Sub

Private Sub Command3_Click()
Out settings_.parallelport + 2, IIf(Check1.Value = 1, invert(Val(Text3.Text)), Val(Text3.Text))
ShowControlStat
End Sub


Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
ShowDataStat
ShowStatusStat
ShowControlStat
End Sub

Private Sub SetPinColor(ValueP, Pin, BcolorOn, BcolorOff, ind)
   If ValueP And (2 ^ Pin) Then
      Shape1(ind).BackColor = BcolorOn
      Label13(ind).ForeColor = &HFFFFFF
   Else
      Shape1(ind).BackColor = BcolorOff
      Label13(ind).ForeColor = &H80000008
   End If
End Sub

Private Sub ShowControlStat()
Dim CONTROL, x, ind

CONTROL = Inp(settings_.parallelport + 2)

Text5.Text = Inp(settings_.parallelport + 2)
Text9.Text = d2b(Text5.Text)

'force invert to show reality
CONTROL = CONTROL Xor 1
CONTROL = CONTROL Xor 2
CONTROL = CONTROL Xor 8
ind = 9
For x = 0 To 3
  SetPinColor CONTROL, x, &HFF0000, &HFFC0C0, ind
  ind = ind + 1
Next
End Sub

Private Sub ShowDataStat()
Dim DATAP, x
DATAP = Inp(settings_.parallelport)
For x = 0 To 7
   SetPinColor DATAP, x, &HFF&, &HC0E0FF, x + 1
Next
Text6.Text = DATAP
Text4.Text = d2b(DATAP)
End Sub

Private Sub ShowStatusStat()
Dim STATUS, x, ind
STATUS = Inp(settings_.parallelport + 1)

Text2.Text = STATUS
Text7.Text = d2b(STATUS)

'force invert to show reality
STATUS = STATUS Xor 128
ind = 30
For x = 3 To 7
   SetPinColor STATUS, x, &HC000&, &HC0FFC0, ind
   ind = ind + 1
Next
End Sub


