VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBitMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitMe [by Pancreas]"
   ClientHeight    =   5175
   ClientLeft      =   1605
   ClientTop       =   2235
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Icon            =   "BitMe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5175
   ScaleWidth      =   8175
   Begin VB.CheckBox chkCSize 
      Caption         =   " &Custom "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame frameCSize 
      Caption         =   " Custom File Size / Date Rate "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   51
      Top             =   4080
      Width           =   7935
      Begin VB.ComboBox comboMedia 
         Height          =   315
         ItemData        =   "BitMe.frx":57E2
         Left            =   240
         List            =   "BitMe.frx":57FE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton btnAccept 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox editCustomSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   16
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label labelCInfo 
         AutoSize        =   -1  'True
         Caption         =   "MB"
         Height          =   195
         Left            =   4080
         TabIndex        =   52
         Top             =   525
         Width           =   240
      End
   End
   Begin VB.Frame frameDRate 
      Caption         =   " Data Rate "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5040
      TabIndex        =   26
      Top             =   2040
      Width           =   3015
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Mbps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   50
         Top             =   1395
         Width           =   660
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   49
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   48
         Top             =   360
         Width           =   615
      End
      Begin VB.Label labelMbps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   930
         TabIndex        =   29
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label labelKbit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1305
         TabIndex        =   28
         Top             =   870
         Width           =   165
      End
      Begin VB.Label labelKBps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1305
         TabIndex        =   27
         Top             =   360
         Width           =   165
      End
   End
   Begin VB.Frame frameFSize 
      Caption         =   " Final File Size "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5040
      TabIndex        =   14
      Top             =   2040
      Width           =   3015
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "GB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   47
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "MB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   46
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "KB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   45
         Top             =   360
         Width           =   360
      End
      Begin VB.Label labelGB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label labelMB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   540
      End
      Begin VB.Label labelKB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame frameTime 
      Caption         =   " Duration "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   4815
      Begin ComctlLib.Slider sliderTime 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         _Version        =   327682
         Min             =   1
         Max             =   540
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   60
         Value           =   1
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "9 Hrs"
         Height          =   195
         Left            =   4170
         TabIndex        =   44
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "1 Min"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2345
         TabIndex        =   42
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[           ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2002
         TabIndex        =   41
         Top             =   1540
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3795
         TabIndex        =   40
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1995
         TabIndex        =   39
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   38
         Top             =   360
         Width           =   510
      End
      Begin VB.Label labelTClock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0   00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2145
         TabIndex        =   22
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label labelTSec 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   3555
         TabIndex        =   21
         Top             =   360
         Width           =   135
      End
      Begin VB.Label labelTMin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   1755
         TabIndex        =   20
         Top             =   360
         Width           =   135
      End
      Begin VB.Label labelTHrs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame frameAudio 
      Caption         =   " Audio Bitrate "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      Begin ComctlLib.Slider sliderAudio 
         Height          =   630
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   1
         Max             =   18
         TickStyle       =   2
         TickFrequency   =   18
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "640 kbps"
         Height          =   195
         Left            =   2100
         TabIndex        =   37
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "32 kbps"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2325
         TabIndex        =   35
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   680
         TabIndex        =   34
         Top             =   360
         Width           =   420
      End
      Begin VB.Label labelA2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2095
         TabIndex        =   18
         Top             =   360
         Width           =   135
      End
      Begin VB.Label labelA1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame frameVideo 
      Caption         =   " Video Bitrate "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Width           =   3495
      Begin ComctlLib.Slider sliderVideo 
         Height          =   630
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   50
         Max             =   25000
         TickStyle       =   2
         TickFrequency   =   25000
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "25 Mbps"
         Height          =   195
         Left            =   2610
         TabIndex        =   33
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0 kbps"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2790
         TabIndex        =   31
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   915
         TabIndex        =   30
         Top             =   360
         Width           =   420
      End
      Begin VB.Label labelV2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2385
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.Label labelV1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   " Options "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox chkDRate 
         Caption         =   " &Rate "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   750
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   " &Audio "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chkVideo 
         Caption         =   " &Video "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmBitMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Video As Integer
Private Audio As Single

Private Minutes As Integer
Private Seconds As Integer

Private FileSize As Double

Private Media(0 To 7) As Integer
Private Cut As String
Private Function calcAVIOverhead(ByVal ForceVideo As Boolean) As Long

    ' Video Enabled
    If ForceVideo Or Video > 0 Then
    
        If Audio > 0 Then
            calcAVIOverhead = Minutes * 581913
            Exit Function
        Else
            calcAVIOverhead = Minutes * 293794
            Exit Function
        End If
        
    Else ' Video Disabled
    
        calcAVIOverhead = 0
    
    End If
    
End Function
Private Sub changeCombo()

    editCustomSize.Text = Media(comboMedia.ListIndex)
    
End Sub
Public Sub p_centerWindow(xForm As Object)

    xForm.Left = (Screen.Width - xForm.Width) / 2
    xForm.Top = (Screen.Height - xForm.Height) / 2

End Sub
Private Sub i_toggleCheck(ByVal Path As Integer)

    Select Case Path

    '---------------------------------------------
       
        Case 1 ' Video CheckBox
        
            If chkVideo Then
                Call i_dragSlider(1)
                sliderVideo.Enabled = True
                labelV1.Enabled = True
                labelV2.Enabled = True
                chkCSize.Enabled = True
            Else
                Video = 0
                sliderVideo.Enabled = False
                labelV1.Enabled = False
                labelV2.Enabled = False
                chkCSize.Enabled = False
            End If
            
            Call i_toggleCheck(4)
            
    '---------------------------------------------
    
        Case 2 ' Audio CheckBox

            If chkAudio Then
                Call i_dragSlider(2)
                sliderAudio.Enabled = True
                labelA1.Enabled = True
                labelA2.Enabled = True
            Else
                Audio = 0
                sliderAudio.Enabled = False
                labelA1.Enabled = False
                labelA2.Enabled = False
            End If
            
    '---------------------------------------------
    
        Case 3 ' Data Rate CheckBox
    
            If chkDRate Then
                frameFSize.Visible = False
                frameDRate.Visible = True
                frameCSize.Caption = " Custom Data Rate "
                labelCInfo = "kBps"
            Else
                frameFSize.Visible = True
                frameDRate.Visible = False
                frameCSize.Caption = " Custom File Size "
                labelCInfo = "MB"
            End If
            
    '---------------------------------------------
    
        Case 4
            If chkVideo And chkCSize Then
                If Not IsNumeric(editCustomSize.Text) Then
                    Call changeCombo
                End If
                editCustomSize.Text = Trim(editCustomSize.Text)
                frameCSize.Enabled = True
                frameCSize.Visible = True
                frmBitMe.Height = 5685
                Dim Pos As Integer
                Pos = Screen.Height - frmBitMe.Height
                If frmBitMe.Top > Pos Then
                    frmBitMe.Top = Pos
                End If
            Else
                frameCSize.Enabled = False
                frameCSize.Visible = False
                frmBitMe.Height = 4580
            End If
            
    '---------------------------------------------
    
    End Select
    
    Call calcFileSizes
    
End Sub
Private Sub i_dragSlider(ByVal Path As Integer)

    Select Case Path
    
    '---------------------------------------------------------------------
        
        Case 1 ' Video Slider
        
            Video = sliderVideo.Value
            
            labelV1 = Video                         ' Video in kbps
            labelV2 = Format(Video / 8, "0.0")      ' Video in kBps
            
    '---------------------------------------------------------------------
    
        Case 2 ' Audio Slider
        
            Select Case sliderAudio.Value
                Case 0 To 4
                    Audio = sliderAudio.Value * 8 + 32
                Case 5 To 8
                    Audio = sliderAudio.Value * 16
                Case 9 To 12
                    Audio = (sliderAudio.Value - 4) * 32
                Case Else
                    Audio = (sliderAudio.Value - 8) * 64
            End Select
            
            labelA1 = Audio                         ' Audio in kbps
            labelA2 = Audio \ 8                     ' Audio in kBps
            
    '---------------------------------------------------------------------
            
        Case 3 ' Time Slider
        
            Minutes = sliderTime.Value
            Seconds = Minutes * 60
            
            labelTHrs = Minutes \ 60
            labelTMin = Minutes
            labelTSec = Seconds
            
            ' ----- Digital Clock -----
            Dim CMin As Integer
            CMin = Minutes Mod 60
    
            If CMin > 9 Then
                labelTClock = labelTHrs & "   " & CMin
            Else
                labelTClock = labelTHrs & "   0" & CMin
            End If
            
    '---------------------------------------------------------------------
            
    End Select

    Call calcFileSizes
    
End Sub
Private Sub calcFileSizes()
    
    ' File Size [bits] = (Video [bits] + Audio [bits]) * Seconds + Overhead [bits]
    ' Divide FS [bits] by 8192 --> FS in KB
    
    FileSize = (Video + Audio) * 1000 * Seconds + calcAVIOverhead(False)
    
    If chkDRate = 0 Then
        
        labelKB = Format(FileSize / 8192, Cut)
        labelMB = Format(FileSize / 8192 / 1024, Cut)
        labelGB = Format(FileSize / 8192 / 1024 / 1024, Cut)
        
    Else
    
        If Seconds > 0 Then ' Prevent Division-by-Zero
            labelKBps = Format(FileSize / Seconds / 8192, "Fixed")
            labelKbit = CInt(FileSize / Seconds / 1000)
            labelMbps = Format(FileSize / Seconds / 1000 / 1000, "Fixed")
        Else
            labelKBps = Format(0, "Fixed")
            labelKbit = 0
            labelMbps = Format(0, "Fixed")
        End If
        
    End If

End Sub
Private Sub InitProgram()

    sliderVideo.Value = 2000
    sliderAudio.Value = 10
    sliderTime.Value = 120
    
    comboMedia.ListIndex = 2
    
    chkVideo = 1
    chkAudio = 1
    chkDRate = 0
    chkCSize = 0
    
    Call i_toggleCheck(1)
    Call i_toggleCheck(2)
    Call i_toggleCheck(3)
    
    Call p_centerWindow(frmBitMe)
    
End Sub
Private Sub selAllText(xEdit As Object)

    xEdit.SelStart = 0
    xEdit.SelLength = Len(xEdit.Text)

End Sub
Private Sub btnAccept_Click()

    Call getVideoBitrate
    
End Sub

Private Sub btnAccept_GotFocus()

    frameCSize.Font.bold = True

End Sub


Private Sub btnAccept_LostFocus()

    frameCSize.Font.bold = False

End Sub


Private Sub chkAudio_Click()

    Call i_toggleCheck(2)

End Sub

Private Sub chkAudio_GotFocus()

    frameOptions.Font.bold = True

End Sub


Private Sub chkAudio_LostFocus()

    frameOptions.Font.bold = False

End Sub


Private Sub chkCSize_Click()

    Call i_toggleCheck(4)

End Sub
Private Sub chkDRate_Click()

    Call i_toggleCheck(3)

End Sub

Private Sub chkDRate_GotFocus()

    frameOptions.Font.bold = True

End Sub


Private Sub chkDRate_LostFocus()

    frameOptions.Font.bold = False

End Sub


Private Sub chkVideo_Click()

    Call i_toggleCheck(1)

End Sub

Private Sub chkVideo_GotFocus()

    frameOptions.Font.bold = True

End Sub


Private Sub chkVideo_LostFocus()

    frameOptions.Font.bold = False

End Sub

Private Sub comboMedia_Click()

    Call changeCombo
    
End Sub
Private Sub comboMedia_GotFocus()

    frameCSize.Font.bold = True

End Sub


Private Sub comboMedia_LostFocus()

    frameCSize.Font.bold = False

End Sub


Private Sub editCustomSize_GotFocus()

    frameCSize.Font.bold = True
    Call selAllText(editCustomSize)

End Sub
Private Sub editCustomSize_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Call getVideoBitrate
    End If

End Sub
Private Sub editCustomSize_LostFocus()

    frameCSize.Font.bold = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        ' [V] -- Video Bitrate
        Case vbKeyV
            If chkVideo Then
                chkVideo = 0
            Else
                chkVideo = 1
            End If
            
        ' [A] -- Audio Bitrate
        Case vbKeyA
            If chkAudio Then
                chkAudio = 0
            Else
                chkAudio = 1
            End If
            
        ' [R] -- Data Rate
        Case vbKeyR
            If chkDRate Then
                chkDRate = 0
            Else
                chkDRate = 1
            End If
            
        ' [C] -- Custom
        Case vbKeyC
            If chkCSize.Enabled Then
                If chkCSize Then
                    chkCSize = 0
                Else
                    chkCSize = 1
                End If
            End If
            
        ' [G] -- Digit Grouping
        Case vbKeyG
            If Cut = "Standard" Then
                Cut = "Fixed"
            Else
                Cut = "Standard"
            End If
            
            Call calcFileSizes
            
        ' [F5] -- Refresh BitMe
        Case vbKeyF5
            Call InitProgram
            
        ' [ESC] -- Close BitMe
        Case vbKeyEscape
            Unload frmBitMe
            
    End Select

End Sub
Private Sub Form_Load()
    
    If GetSetting("BitMe", "Settings", "SeenSplash", "0") = "0" Then
        frmAbout.Show
        frmBitMe.Hide
    End If
    
    ' Setup Media Sizes Array
    Media(0) = 350
    Media(1) = 650
    Media(2) = 700
    Media(3) = 880
    Media(4) = 1100
    Media(5) = 1466
    Media(6) = 2200
    Media(7) = 4400
    
    Cut = "Standard"
    
    Call InitProgram
    
End Sub
Private Sub sliderAudio_Change()

    Call i_dragSlider(2)

End Sub

Private Sub sliderAudio_GotFocus()

    frameAudio.Font.bold = True

End Sub


Private Sub sliderAudio_LostFocus()

    frameAudio.Font.bold = False

End Sub


Private Sub sliderAudio_Scroll()

    Call i_dragSlider(2)

End Sub
Private Sub sliderTime_Change()

    Call i_dragSlider(3)

End Sub

Private Sub sliderTime_GotFocus()

    frameTime.Font.bold = True

End Sub


Private Sub sliderTime_LostFocus()

    frameTime.Font.bold = False

End Sub


Private Sub sliderTime_Scroll()

    Call i_dragSlider(3)

End Sub

Private Sub sliderVideo_Change()

    Call i_dragSlider(1)

End Sub
Private Sub sliderVideo_GotFocus()

    frameVideo.Font.bold = True

End Sub

Private Sub sliderVideo_LostFocus()

    frameVideo.Font.bold = False

End Sub
Private Sub sliderVideo_Scroll()

    Call i_dragSlider(1)

End Sub



Private Function calcVideoBitrateFromFS(ByVal CustomMB As Double) As Integer

    Dim TotalBits As Double
    Dim VideoBits As Double
    Dim AudioBits As Double
    
    TotalBits = CustomMB * 1024 * 1024 * 8 - calcAVIOverhead(True)
    AudioBits = Audio * 1000 * Seconds
    
    VideoBits = (TotalBits - AudioBits) / (CDbl(Seconds) * 1000)

    If VideoBits < 0 Then
        VideoBits = 0
    End If
    
    calcVideoBitrateFromFS = CInt(VideoBits)
    
End Function
Private Function calcVideoBitrateFromDR(ByVal CustomKBps As Double) As Integer

    calcVideoBitrateFromDR = calcVideoBitrateFromFS(CustomKBps * Seconds / 1024)

End Function
Private Sub getVideoBitrate()

    On Error GoTo ErrorHandler
    
    Dim CustomFS As Double
    Dim VideoBitrate As Integer
    
    If IsNumeric(editCustomSize.Text) Then
        CustomFS = editCustomSize.Text
    Else
        GoTo ErrorHandler
    End If
    
    If chkDRate Then
        VideoBitrate = calcVideoBitrateFromDR(CustomFS)
    Else
        VideoBitrate = calcVideoBitrateFromFS(CustomFS)
    End If
    
    ' Does it fall within Slider bounds?
    If sliderVideo.Min < VideoBitrate And VideoBitrate <= sliderVideo.Max Then
        sliderVideo.Value = VideoBitrate
        Call i_dragSlider(1)
    Else
        GoTo ErrorHandler
    End If
    
    editCustomSize.Text = Trim(editCustomSize.Text)
    Call selAllText(editCustomSize)
    editCustomSize.SetFocus
    
Exit Sub

' ------------------------------------------

ErrorHandler:
    editCustomSize.Text = "Invalid"
    Call selAllText(editCustomSize)
    editCustomSize.SetFocus
    
End Sub
