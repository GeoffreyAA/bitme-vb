VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   5775
   ClientLeft      =   3300
   ClientTop       =   2265
   ClientWidth     =   4455
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frameAbout 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Frame frameLine 
         Height          =   135
         Left            =   240
         TabIndex        =   2
         Top             =   1380
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Author: Geoffrey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1455
         TabIndex        =   4
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "BitMe 1.16"
         BeginProperty Font 
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
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   $"About.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   3735
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload frmAbout
    End If
    
End Sub
Private Sub Form_Load()

    Call frmBitMe.p_centerWindow(frmAbout)

End Sub
Private Sub Form_Unload(Cancel As Integer)

    SaveSetting "BitMe", "Settings", "SeenSplash", "1"
    frmBitMe.Show
    
End Sub
