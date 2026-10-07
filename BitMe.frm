VERSION 4.00
Begin VB.Form frmBitMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitMe [by Pancreas]"
   ClientHeight    =   4095
   ClientLeft      =   1215
   ClientTop       =   1935
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Height          =   4605
   Icon            =   "BitMe.frx":0000
   KeyPreview      =   -1  'True
   Left            =   1155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "BitMe.frx":57E2
   ScaleHeight     =   4095
   ScaleWidth      =   8175
   Top             =   1485
   Width           =   8295
   Begin VB.Frame RateFrame 
      Caption         =   " Data Rate "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Mbps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   46
         Top             =   1395
         Width           =   660
      End
      Begin VB.Label Mbps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   45
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   40
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Kbit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   39
         Top             =   870
         Width           =   165
      End
      Begin VB.Label KBps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   38
         Top             =   360
         Width           =   165
      End
   End
   Begin VB.Frame OptionsFrame 
      Caption         =   " Options "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   17
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox Check3 
         Caption         =   " &Rate "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   36
         Top             =   1320
         Width           =   750
      End
      Begin VB.CheckBox Check2 
         Caption         =   " &Audio "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   " &Video "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame SizeFrame 
      Caption         =   " Final File Size "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   13
      Top             =   2040
      Width           =   3015
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "GB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   28
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "MB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   27
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "kB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   26
         Top             =   360
         Width           =   330
      End
      Begin VB.Label GB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   16
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label MB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
      Begin VB.Label KB 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame TimeFrame 
      Caption         =   " Duration "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   8
      Top             =   2040
      Width           =   4815
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[           ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   44
         Top             =   1540
         Width           =   825
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   43
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label Clock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0   00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   42
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   35
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   34
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   33
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "9 Hrs"
         Height          =   195
         Left            =   4170
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "1 Min"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   390
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   30
         Min             =   1
         Max             =   540
         SelStart        =   120
         TickStyle       =   2
         TickFrequency   =   60
         Value           =   120
      End
      Begin VB.Label Sec 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Min 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Hrs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   9
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame AudioFrame 
      Caption         =   " Audio Bitrate "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   32
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   31
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "640 kbps"
         Height          =   195
         Left            =   2100
         TabIndex        =   23
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "32 kbps"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   7
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Audio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   630
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   1
         Max             =   18
         SelStart        =   10
         TickStyle       =   2
         TickFrequency   =   18
         Value           =   10
      End
   End
   Begin VB.Frame VideoFrame 
      Caption         =   " Video Bitrate "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "kBps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   30
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "kbps"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   915
         TabIndex        =   29
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "25 Mbps"
         Height          =   195
         Left            =   2610
         TabIndex        =   21
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "0 kbps"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   3
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Video 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         TabIndex        =   2
         Top             =   360
         Width           =   135
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   630
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   50
         Max             =   25000
         SelStart        =   1500
         TickStyle       =   2
         TickFrequency   =   25000
         Value           =   1500
      End
   End
End
Attribute VB_Name = "frmBitMe"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public Cut As String
Private Sub Check1_Click()

    Call i_toggleCheck(1)
    
End Sub
Private Sub Check1_GotFocus()

    OptionsFrame.Font.bold = True

End Sub
Private Sub Check1_LostFocus()

    OptionsFrame.Font.bold = False

End Sub


Private Sub Check2_Click()
    
    Call i_toggleCheck(2)
    
End Sub
Private Sub Check2_GotFocus()

    OptionsFrame.Font.bold = True
    
End Sub


Private Sub Check2_LostFocus()

    OptionsFrame.Font.bold = False

End Sub


Private Sub Check3_Click()
    
    Call i_toggleCheck(3)
    
End Sub
Private Sub Check3_GotFocus()

    OptionsFrame.Font.bold = True

End Sub


Private Sub Check3_LostFocus()

    OptionsFrame.Font.bold = False

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
            
        ' <V> - Video Bitrate
        Case vbKeyV
            If Check1 Then
                Check1 = 0
            Else
                Check1 = 1
            End If
            
        ' <A> - Audio Bitrate
        Case vbKeyA
            If Check2 Then
                Check2 = 0
            Else
                Check2 = 1
            End If
            
        ' <R> - Data Rate
        Case vbKeyR
            If Check3 Then
                Check3 = 0
            Else
                Check3 = 1
            End If
            
        ' <G> - Group Digits
        Case vbKeyG
            If Cut = "Standard" Then
                Cut = "Fixed"
            Else
                Cut = "Standard"
            End If
            
            Call c_whatCalc
            
        ' <F5> - Refresh BitMe
        Case vbKeyF5
            Unload frmBitMe
            Load frmBitMe
            
        ' <ESC> - Close BitMe
        Case vbKeyEscape
            Unload frmBitMe
            
    End Select
            
End Sub
Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
        
    Dim ShowSplash As String
    Let ShowSplash = GetSetting("BitMe", "Settings", "Splash", "Yes")
    
    If ShowSplash = "Yes" Then
        frmBitMe.Visible = False
        frmAbout.Visible = True
    Else
        frmBitMe.Visible = True
        frmAbout.Visible = False
        Unload frmAbout
    End If
    
    Cut = "Standard"
    
    ' Initialize Values
    For i = 3 To 1 Step -1
    
        Call i_dragSlider(i)
        Call i_toggleCheck(i)
    
    Next i
    
End Sub

Private Sub Slider1_Change()

    Call i_dragSlider(1)

End Sub

Private Sub Slider1_GotFocus()

    VideoFrame.Font.bold = True

End Sub
Private Sub Slider1_LostFocus()

    VideoFrame.Font.bold = False

End Sub





Private Sub Slider1_Scroll()

    Call i_dragSlider(1)
    
End Sub
Private Sub Slider2_Change()

    Call i_dragSlider(2)
        
End Sub
Private Sub Slider2_GotFocus()

    AudioFrame.Font.bold = True

End Sub


Private Sub Slider2_LostFocus()

    AudioFrame.Font.bold = False

End Sub


Private Sub Slider2_Scroll()
    
    Call i_dragSlider(2)
    
End Sub
Private Sub Slider3_Change()

    Call i_dragSlider(3)

End Sub
Private Sub Slider3_GotFocus()
    
    TimeFrame.Font.bold = True
    
End Sub


Private Sub Slider3_LostFocus()

    TimeFrame.Font.bold = False
    
End Sub


Private Sub Slider3_Scroll()

    Call i_dragSlider(3)

End Sub
Public Sub c_whatCalc()

    If Check1 Then
    
        ' Video + Audio Enabled
        If Check2 Then
            Call c_fileSizes(Video, Audio)
            Exit Sub
            
        ' Only Video Enabled
        Else
            Call c_fileSizes(Video, 0)
            Exit Sub
            
        End If
        
    Else
    
        ' Only Audio Enabled
        If Check2 Then
            Call c_fileSizes(0, Audio)
            Exit Sub
            
        ' Video + Audio Disabled
        Else
            Call c_fileSizes(0, 0)
        End If
        
    End If
    
End Sub
Static Sub c_fileSizes(VideoBitrate As Long, AudioBitrate As Integer)
    
    Dim Overhead As Long
    Dim FileSize As Double
    
    ' AVI Overhead
    If Check1 And VideoBitrate > 0 Then
    
        If Check2 Then
            Overhead = Slider3.Value * 720896
        Else
            Overhead = Slider3.Value * 327680
        End If
        
    Else
        Overhead = 0
    End If
    
    ' File Size in Bits = (Video in Bits + Audio in Bits) * Seconds + Overhead in Bits
    FileSize = ((VideoBitrate + AudioBitrate) * 1000 * Sec + Overhead) / 8192
    
    If Check3 = 0 Then
    
        KB = Format(FileSize, Cut)
        MB = Format(FileSize / 1024, Cut)
        GB = Format(FileSize / 1048576, Cut)
        Exit Sub
        
    Else
        
        KBps = CInt(FileSize / Sec)
        Kbit = KBps * 8
        Mbps = Format(Kbit / 1000, Cut)

    End If

End Sub
Public Sub i_toggleCheck(ByVal WhichCheck As Integer)

    Select Case WhichCheck

    '---------------------------------------------
       
        Case 1 ' Video CheckBox
        
            If Check1 Then
                Slider1.Enabled = True
                Video.Enabled = True
                Label2.Enabled = True
            Else
                Slider1.Enabled = False
                Video.Enabled = False
                Label2.Enabled = False
            End If
            
    '---------------------------------------------
    
        Case 2 ' Audio CheckBox
    
            If Check2 Then
                Slider2.Enabled = True
                Audio.Enabled = True
                Label4.Enabled = True
            Else
                Slider2.Enabled = False
                Audio.Enabled = False
                Label4.Enabled = False
            End If
            
    '---------------------------------------------
    
        Case 3 ' Data Rate CheckBox
    
            If Check3 Then
                SizeFrame.Visible = False
                RateFrame.Visible = True
            Else
                SizeFrame.Visible = True
                RateFrame.Visible = False
            End If
            
    '---------------------------------------------
            
    End Select
    
    Call c_whatCalc
    
End Sub
Public Sub i_dragSlider(ByVal WhichSlider As Integer)

    Select Case WhichSlider
        
    '---------------------------------------------------------------------
        
        Case 1 ' Video Slider
        
            Video = Slider1.Value                       ' Video in kbps
            Label2 = Format(Slider1.Value / 8, "0.0")   ' Video in KBps
            
    '---------------------------------------------------------------------
    
        Case 2 ' Audio Slider
        
            Select Case Slider2.Value
                Case 0 To 4
                    Audio = Slider2.Value * 8 + 32
                Case 5 To 8
                    Audio = Slider2.Value * 16
                Case 9 To 12
                    Audio = (Slider2.Value - 4) * 32
                Case Else
                    Audio = (Slider2.Value - 8) * 64
            End Select
    
            Label4 = Audio \ 8
            
    '---------------------------------------------------------------------
            
        Case 3 ' Time Slider
        
            Dim Minutes As Integer
            
            Hrs = Slider3.Value \ 60        ' Hours Label
            Min = Slider3.Value             ' Minutes Label
            Sec = Slider3.Value * 60        ' Seconds Label
        
            ' ---------------
            '  Clock Display
            ' ---------------
    
            Minutes = Slider3.Value Mod 60
    
            If Minutes < 10 Then
                Clock = Hrs & "   0" & Minutes
            Else
                Clock = Hrs & "   " & Minutes
            End If
            
    '---------------------------------------------------------------------
            
    End Select

    Call c_whatCalc
    
End Sub

