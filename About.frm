VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   5790
   ClientLeft      =   2790
   ClientTop       =   2265
   ClientWidth     =   4485
   ClipControls    =   0   'False
   Height          =   6300
   Icon            =   "About.frx":0000
   Left            =   2730
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleMode       =   0  'User
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   Top             =   1815
   Width           =   4605
   Begin VB.Frame Frame2 
      Height          =   5610
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   4245
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Geoffrey"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   820
         TabIndex        =   4
         Top             =   840
         Width           =   2595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BitMe 1.15"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         TabIndex        =   3
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   $"About.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   3855
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload frmAbout
    End If

End Sub
Private Sub Form_Load()
    
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2

End Sub
Private Sub Form_Unload(Cancel As Integer)

    SaveSetting "BitMe", "Settings", "Splash", "No"
    frmBitMe.Visible = True

End Sub
