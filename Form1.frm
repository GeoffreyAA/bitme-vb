VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " BitMe (By Pancreas)"
   ClientHeight    =   4095
   ClientLeft      =   705
   ClientTop       =   1365
   ClientWidth     =   8175
   Height          =   4500
   Icon            =   "Form1.frx":0000
   Left            =   645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8175
   Top             =   1020
   Width           =   8295
   Begin VB.Frame Frame5 
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2040
      Width           =   3255
      Begin VB.Label Label21 
         Caption         =   "Gigabyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Megabyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Kilobyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "1483"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "1519200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Length/Time "
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
      TabIndex        =   11
      Top             =   2040
      Width           =   4575
      Begin VB.Label Label28 
         Caption         =   "0 Hrs"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "6 Hrs"
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   510
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "7200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "120"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   30
         Max             =   360
         SelStart        =   120
         TickStyle       =   2
         TickFrequency   =   60
         Value           =   120
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.Label Label26 
         Caption         =   "448 Kbps"
         Height          =   255
         Left            =   2320
         TabIndex        =   34
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "32 Kbps"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Kilobytes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Kilobits"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "24"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "192"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         Max             =   16
         SelStart        =   10
         TickStyle       =   2
         TickFrequency   =   16
         Value           =   10
      End
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.Label Label24 
         Caption         =   "10000 Kbps"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "0 Kbps"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Kilobytes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2150
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Kilobits"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   800
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "187"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1700
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   630
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   50
         Max             =   10000
         SelStart        =   1500
         TickStyle       =   2
         TickFrequency   =   10000
         Value           =   1500
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox Check3 
         Caption         =   "????"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Audio"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Video"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Check1_Click()

    If Check1 = 1 Then
    Slider1.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
    End If
    
    If Check1 = 0 Then
    Slider1.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False
    End If
    
End Sub

Private Sub Check2_Click()
    
    If Check2 = 1 Then
    Slider2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    End If
    
    If Check2 = 0 Then
    Slider2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    End If
    
End Sub


Private Sub Check3_Click()

    If Check3 = 1 Then
    Slider1.Enabled = True
    Slider2.Enabled = True
    Slider3.Enabled = True
    End If
    
    If Check3 = 0 Then
    Slider1.Enabled = False
    Slider2.Enabled = False
    Slider3.Enabled = False
    End If
    
End Sub

Private Sub Slider1_Scroll()

    If Slider1 = Slider1.Value Then
    Label1 = Slider1.Value
    Label2 = Int(Slider1.Value / 8)
    End If
    
    If Check2 = 1 Then
    Label8 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60))
    Label9 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024)
    Label10 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024 / 1024)
    End If
    
    If Check2 = 0 Then
    Label8 = Int((Slider1.Value / 8) * (Slider3.Value * 60))
    Label9 = Int((Slider1.Value / 8) * (Slider3.Value * 60) / 1024)
    Label10 = Int((Slider1.Value / 8) * (Slider3.Value * 60) / 1024 / 1024)
    End If

    If Label1 = 10000 Then
    Label1.Left = 140
    End If
    
    If Label1 < 10000 And Label1 > 999 Then
    Label1.Left = 240
    End If
    
    If Label1 < 1000 And Label1 > 99 Then
    Label1.Left = 340
    End If
    
    If Label1 < 100 And Label1 > 9 Then
    Label1.Left = 440
    End If
    
    If Label1 < 10 Then
    Label1.Left = 540
    End If
    
    If Label2 > 999 Then
    Label2.Left = 1600
    End If
    
    If Label2 < 1000 And Label2 > 99 Then
    Label2.Left = 1700
    End If
    
    If Label2 < 100 And Label2 > 9 Then
    Label2.Left = 1800
    End If
    
    If Label2 < 10 Then
    Label2.Left = 1900
    End If
    
End Sub


Private Sub Slider2_Scroll()

    If Slider2.Value = 1 Then
    Label3 = 32
    Label4 = 4
    End If
    
    If Slider2.Value = 2 Then
    Label3 = 40
    Label4 = 5
    End If
    
    If Slider2.Value = 3 Then
    Label3 = 48
    Label4 = 6
    End If
    
    If Slider2.Value = 4 Then
    Label3 = 56
    Label4 = 7
    End If
    
    If Slider2.Value = 5 Then
    Label3 = 64
    Label4 = 8
    End If
    
    If Slider2.Value = 6 Then
    Label3 = 80
    Label4 = 10
    End If
    
    If Slider2.Value = 7 Then
    Label3 = 96
    Label4 = 12
    End If
    
    If Slider2.Value = 8 Then
    Label3 = 112
    Label4 = 14
    End If
    
    If Slider2.Value = 9 Then
    Label3 = 128
    Label4 = 16
    End If
    
    If Slider2.Value = 10 Then
    Label3 = 160
    Label4 = 20
    End If
    
    If Slider2.Value = 11 Then
    Label3 = 192
    Label4 = 24
    End If
    
    If Slider2.Value = 12 Then
    Label3 = 224
    Label4 = 28
    End If
    
    If Slider2.Value = 13 Then
    Label3 = 256
    Label4 = 32
    End If
    
    If Slider2.Value = 14 Then
    Label3 = 320
    Label4 = 40
    End If

    If Slider2.Value = 15 Then
    Label3 = 384
    Label4 = 48
    End If

    If Slider2.Value = 16 Then
    Label3 = 448
    Label4 = 56
    End If

    If Check1 = 1 Then
    Label8 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60))
    Label9 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024)
    Label10 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024 / 1024)
    End If
    
    If Check1 = 0 Then
    Label8 = Int(Label4 * (Slider3.Value * 60))
    Label9 = Int(Label4 * (Slider3.Value * 60) / 1024)
    Label10 = Int(Label4 * (Slider3.Value * 60) / 1024 / 1024)
    End If
    
    If Label3 > 99 Then
    Label3.Left = 240
    End If
    
    If Label3 < 100 Then
    Label3.Left = 340
    End If
    
    If Label4 > 9 Then
    Label4.Left = 1800
    End If
    
    If Label4 < 10 Then
    Label4.Left = 1900
    End If
    
End Sub


Private Sub Slider3_Scroll()

    If Slider3 = Slider3.Value Then
    Label5 = Int(Slider3.Value / 60)
    Label6 = Slider3.Value
    Label7 = Slider3.Value * 60
    End If
    
    If Check1 = 1 And Check2 = 1 Then
    Label8 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60))
    Label9 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024)
    Label10 = Int((Slider1.Value / 8 + Label4) * (Slider3.Value * 60) / 1024 / 1024)
    End If
 
    If Check1 = 0 Then
    Label8 = Int(Label4 * (Slider3.Value * 60))
    Label9 = Int(Label4 * (Slider3.Value * 60) / 1024)
    Label10 = Int(Label4 * (Slider3.Value * 60) / 1024 / 1024)
    End If
    
    If Check2 = 0 Then
    Label8 = Int((Slider1.Value / 8) * (Slider3.Value * 60))
    Label9 = Int((Slider1.Value / 8) * (Slider3.Value * 60) / 1024)
    Label10 = Int((Slider1.Value / 8) * (Slider3.Value * 60) / 1024 / 1024)
    End If
    
    If Check1 = 0 And Check2 = 0 Then
    Label8 = 0
    Label9 = 0
    Label10 = 0
    End If
    
    If Label6 > 99 Then
    Label6.Left = 1440
    End If
    
    If Label6 < 100 And Label6 > 9 Then
    Label6.Left = 1540
    End If
    
    If Label6 < 10 Then
    Label6.Left = 1640
    End If

    If Label7 > 9999 Then
    Label7.Left = 2900
    End If
    
    If Label7 > 999 And Label7 < 9999 Then
    Label7.Left = 3000
    End If
    
    If Label7 > 99 And Label7 < 1000 Then
    Label7.Left = 3100
    End If
    
    If Label7 > 9 And Label7 < 100 Then
    Label7.Left = 3200
    End If
    
    If Label7 < 10 Then
    Label7.Left = 3300
    End If
    
    
    
End Sub
