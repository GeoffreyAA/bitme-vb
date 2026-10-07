VERSION 4.00
Begin VB.Form frmBitMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitMe [by Pancreas]"
   ClientHeight    =   4095
   ClientLeft      =   1065
   ClientTop       =   2295
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Height          =   4605
   Icon            =   "BitMe.frx":0000
   KeyPreview      =   -1  'True
   Left            =   1005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "BitMe.frx":57E2
   ScaleHeight     =   4095
   ScaleWidth      =   8175
   Top             =   1845
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
         TabIndex        =   47
         Top             =   1400
         Width           =   660
      End
      Begin VB.Label Mbps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.66"
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
         TabIndex        =   46
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "KBps"
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
         Width           =   645
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
         Left            =   1570
         TabIndex        =   40
         Top             =   875
         Width           =   585
      End
      Begin VB.Label Kbit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1664"
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
         Left            =   855
         TabIndex        =   39
         Top             =   870
         Width           =   615
      End
      Begin VB.Label KBps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "208"
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
         Left            =   1005
         TabIndex        =   38
         Top             =   360
         Width           =   465
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
         Caption         =   "KB"
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
         Width           =   360
      End
      Begin VB.Label GB 
         AutoSize        =   -1  'True
         Caption         =   "1.43"
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
         Caption         =   "1462.57"
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
         Width           =   990
      End
      Begin VB.Label KB 
         AutoSize        =   -1  'True
         Caption         =   "1497669.38"
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
         Width           =   1440
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
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "00"
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
         Left            =   2445
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2"
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
         Width           =   135
      End
      Begin VB.Label Label26 
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
         Left            =   3800
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
         Left            =   2000
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
         AutoSize        =   -1  'True
         Caption         =   "9 Hrs"
         Height          =   195
         Left            =   4190
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
         Caption         =   "7200"
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
         Left            =   3195
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Min 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "120"
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
         Left            =   1515
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Hrs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2"
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
         AutoSize        =   -1  'True
         Caption         =   "KBps"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   360
         Width           =   480
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
         Caption         =   "24"
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
         Left            =   1965
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Audio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "192"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   375
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
      Begin VB.Label Cut 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   195
         Left            =   1560
         TabIndex        =   48
         Top             =   1440
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "KBps"
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
         Left            =   2760
         TabIndex        =   30
         Top             =   360
         Width           =   480
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
         AutoSize        =   -1  'True
         Caption         =   "25 Mbps"
         Height          =   195
         Left            =   2615
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
         Caption         =   "187.5"
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
         Left            =   2145
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Video 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1500"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   495
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

Private Sub AudioFrame_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Check1_Click()

    If Check1 = 1 Then
        Slider1.Enabled = True
        Video.Enabled = True
        Label2.Enabled = True
    Else
        Slider1.Enabled = False
        Video.Enabled = False
        Label2.Enabled = False
    End If
    
    ' Video + Audio Enabled
    If Check1 = 1 And Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
    End If
    
    ' Video Enabled
    If Check1 = 1 And Check2 = 0 Then
        Call fileCalc(Video, 0, "V")
    End If
    
    ' Audio Enabled
    If Check1 = 0 And Check2 = 1 Then
        Call fileCalc(0, Audio, 0)
    End If

    ' Video + Audio Disabled
    If Check1 = 0 And Check2 = 0 Then
        KB = Cut
        MB = Cut
        GB = Cut
        KBps = 0
        Kbit = 0
        Mbps = Cut
    End If
    
End Sub

Private Sub Check1_GotFocus()

    OptionsFrame.Font.Bold = True

End Sub


Private Sub Check1_LostFocus()

    OptionsFrame.Font.Bold = False

End Sub


Private Sub Check2_Click()
    
    If Check2 = 1 Then
        Slider2.Enabled = True
        Audio.Enabled = True
        Label4.Enabled = True
    Else
        Slider2.Enabled = False
        Audio.Enabled = False
        Label4.Enabled = False
    End If
    
    ' Video + Audio Enabled
    If Check1 = 1 And Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
    End If
    
    ' Video Enabled
    If Check1 = 1 And Check2 = 0 Then
        Call fileCalc(Video, 0, "V")
    End If
    
    ' Audio Enabled
    If Check1 = 0 And Check2 = 1 Then
        Call fileCalc(0, Audio, 0)
    End If
    
    ' Video + Audio Disabled
    If Check1 = 0 And Check2 = 0 Then
        KB = Cut
        MB = Cut
        GB = Cut
        KBps = 0
        Kbit = 0
        Mbps = Cut
    End If
    
End Sub


Private Sub Check2_GotFocus()

    OptionsFrame.Font.Bold = True
    
End Sub


Private Sub Check2_LostFocus()

    OptionsFrame.Font.Bold = False

End Sub


Private Sub Check3_Click()
    
    If Check3 = 0 Then
        SizeFrame.Visible = True
        RateFrame.Visible = False
    Else
        SizeFrame.Visible = False
        RateFrame.Visible = True
    End If

End Sub



Private Sub Check3_GotFocus()

    OptionsFrame.Font.Bold = True

End Sub


Private Sub Check3_LostFocus()

    OptionsFrame.Font.Bold = False

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape
            Unload frmBitMe
        Case vbKeyV
            If Check1 = 1 Then
                Check1 = 0
            Else
                Check1 = 1
            End If
        Case vbKeyA
            If Check2 = 1 Then
                Check2 = 0
            Else
                Check2 = 1
            End If
        Case vbKeyR
            If Check3 = 1 Then
                Check3 = 0
            Else
                Check3 = 1
            End If
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
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmBitMe
    
End Sub

Private Sub Slider1_Change()
 
    Select Case Slider1.Value
        Case 0
            Video = 0
            Label2 = "0.0"
        Case 25000
            Video = 25000
            Label2 = "3125.0"
    End Select
    
    ' Video + Audio Enabled
    If Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
        
    ' Video Enabled
    Else
        Call fileCalc(Video, 0, "V")
    End If
      
End Sub
Private Sub Slider1_GotFocus()

    VideoFrame.Font.Bold = True

End Sub


Private Sub Slider1_LostFocus()

    VideoFrame.Font.Bold = False

End Sub


Private Sub Slider1_Scroll()

    Video = Slider1.Value
    Label2 = Format(Slider1.Value / 8, "0.0")

    ' Video + Audio Enabled
    If Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
        
    ' Video Enabled
    Else
        Call fileCalc(Video, 0, "V")
    End If
    
End Sub
Private Sub Slider2_Change()

    Select Case Slider2.Value
        Case 0
            Audio = 32
            Label4 = 4
        Case 18
            Audio = 640
            Label4 = 80
    End Select
    
    ' Video + Audio Enabled
    If Check1 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
        
    ' Audio Enabled
    Else
        Call fileCalc(0, Audio, 0)
    End If
        
End Sub

Private Sub Slider2_GotFocus()

    AudioFrame.Font.Bold = True

End Sub


Private Sub Slider2_LostFocus()

    AudioFrame.Font.Bold = False

End Sub


Private Sub Slider2_Scroll()
    
    Select Case Slider2.Value
        Case 0
            Audio = 32
            Label4 = 4
        Case 1
            Audio = 40
            Label4 = 5
        Case 2
            Audio = 48
            Label4 = 6
        Case 3
            Audio = 56
            Label4 = 7
        Case 4
            Audio = 64
            Label4 = 8
        Case 5
            Audio = 80
            Label4 = 10
        Case 6
            Audio = 96
            Label4 = 12
        Case 7
            Audio = 112
            Label4 = 14
        Case 8
            Audio = 128
            Label4 = 16
        Case 9
            Audio = 160
            Label4 = 20
        Case 10
            Audio = 192
            Label4 = 24
        Case 11
            Audio = 224
            Label4 = 28
        Case 12
            Audio = 256
            Label4 = 32
        Case 13
            Audio = 320
            Label4 = 40
        Case 14
            Audio = 384
            Label4 = 48
        Case 15
            Audio = 448
            Label4 = 56
        Case 16
            Audio = 512
            Label4 = 64
        Case 17
            Audio = 576
            Label4 = 72
        Case 18
            Audio = 640
            Label4 = 80
    End Select
    
    ' Video + Audio Enabled
    If Check1 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
    
    ' Audio Enabled
    Else
        Call fileCalc(0, Audio, 0)
    End If
    
End Sub
Private Sub Slider3_Change()
    
    Select Case Slider3.Value
        Case 1
            Hrs = 0
            Min = 1
            Sec = 60
            label31 = 0
            Label32 = "01"
        Case 540
            Hrs = 9
            Min = 540
            Sec = 32400
            label31 = 9
            Label32 = "00"
    End Select
    
    ' Video + Audio Enabled
    If Check1 = 1 And Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
    End If
    
    ' Video Enabled
    If Check1 = 1 And Check2 = 0 Then
        Call fileCalc(Video, 0, "V")
    End If
    
    ' Audio Enabled
    If Check1 = 0 And Check2 = 1 Then
        Call fileCalc(0, Audio, 0)
    End If
    
    ' Video + Audio Disabled
    If Check1 = 0 And Check2 = 0 Then
        KB = Cut
        MB = Cut
        GB = Cut
    End If
    
End Sub
Private Sub Slider3_GotFocus()
    
    TimeFrame.Font.Bold = True
    
End Sub


Private Sub Slider3_LostFocus()

    TimeFrame.Font.Bold = False
    
End Sub


Private Sub Slider3_Scroll()
    
    Dim Minutes As Integer
    Minutes = Slider3.Value Mod 60
    Hrs = Fix(Slider3.Value / 60)
    Min = Slider3.Value
    Sec = Slider3.Value * 60
    label31 = Hrs
     
    Select Case Minutes
        Case 0 To 9
            Label32 = "0" & Minutes
        Case Else
            Label32 = Minutes
    End Select
    
   ' Video + Audio Enabled
    If Check1 = 1 And Check2 = 1 Then
        Call fileCalc(Video, Audio, "V_A")
    End If
    
    ' Video Enabled
    If Check1 = 1 And Check2 = 0 Then
        Call fileCalc(Video, 0, "V")
    End If
    
    ' Audio Enabled
    If Check1 = 0 And Check2 = 1 Then
        Call fileCalc(0, Audio, 0)
    End If
    
    ' Video + Audio Disabled
    If Check1 = 0 And Check2 = 0 Then
        KB = Cut
        MB = Cut
        GB = Cut
    End If
    
End Sub
Public Sub fileCalc(VideoBitrate, AudioBitrate, OverheadType)

    Select Case OverheadType
    ' Video + Audio Enabled
        Case "V_A"
            If Video > 0 Then
                Overhead = Slider3.Value * 720896
            Else
                Overhead = 0
            End If
    ' Video Enabled
        Case "V"
            If Video > 0 Then
                Overhead = Slider3.Value * 327680
            Else
                Overhead = 0
            End If
    ' Audio Enabled
        Case 0
            Overhead = 0
    End Select
            
    FileSize = ((VideoBitrate * 1000 + AudioBitrate * 1000) * Sec + Overhead) / 8192
    
    KB = Format(FileSize, Cut)
    MB = Format(FileSize / 1024, Cut)
    GB = Format(FileSize / 1048576, Cut)
   
    KBps = CLng(KB / Sec)
    Kbit = KBps * 8
    Mbps = Format(Kbit / 1000, Cut)

End Sub
