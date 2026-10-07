VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " The Bitrate Calculator (By Pancreas)"
   ClientHeight    =   4095
   ClientLeft      =   675
   ClientTop       =   1575
   ClientWidth     =   8655
   FillStyle       =   0  'Solid
   Height          =   4500
   Icon            =   "Form1.frx":0000
   Left            =   615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8655
   Top             =   1230
   Width           =   8775
   Begin VB.Frame Frame6 
      Caption         =   " Options "
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1095
      Begin VB.CheckBox Check3 
         Caption         =   "????"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Video"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Audio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Final File Size "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4920
      TabIndex        =   19
      Top             =   2040
      Width           =   3615
      Begin VB.Label Label29 
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "1167"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "1195200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Kilobyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Megabyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Gigabyte(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Audio Bitrate "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.Label Label22 
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
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "32 Kbps"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "320 Kbps"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Kilobytes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         Max             =   16
         SelStart        =   11
         TickStyle       =   2
         TickFrequency   =   36
         Value           =   11
      End
      Begin VB.Label Label16 
         Caption         =   "Kilobits"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Length/Time "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
      Begin VB.Label Label26 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label25 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label24 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "0 Hrs"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "5 Hrs"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Seconds"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Minutes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Hours"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   30
         Max             =   300
         SelStart        =   120
         TickStyle       =   2
         TickFrequency   =   60
         Value           =   120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Video Bitrate "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Label21 
         Caption         =   "150"
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
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "1200"
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
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Kilobytes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Kilobits"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "10000 Kbps"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "0 Kbps"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   615
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   50
         Max             =   10000
         SelStart        =   1200
         TickStyle       =   2
         TickFrequency   =   10000
         Value           =   1200
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Check1_Click()
    
    If Check1 = 1 Then
    Slider1.Enabled = True
    Label18.Enabled = True
    Label21.Enabled = True
    End If
    
    If Check1 = 0 Then
    Slider1.Enabled = False
    Label18.Enabled = False
    Label21.Enabled = False
    End If

End Sub


Private Sub Check2_Click()
    
    If Check2 = 1 Then
    Slider2.Enabled = True
    Label22.Enabled = True
    Label23.Enabled = True
    End If
    
    If Check2 = 0 Then
    Slider2.Enabled = False
    Label22.Enabled = False
    Label23.Enabled = False
    End If
    
End Sub
Private Sub Command1_Click()
    
    Unload Form1
    
End Sub

Private Sub Command2_Click()
    
    Unload Form1
    
End Sub

Private Sub Command3_Click()

    Form3.Show
    
End Sub

Private Sub Command4_Click()

    Load Form4
    Form4.Show
    
End Sub

Private Sub Option1_Click()

    If Option1 = True Then
    Frame3.Enabled = True
    Frame7.Enabled = False
    End If

End Sub

Private Sub Option2_Click()

    If Option2 = True Then
    Frame3.Enabled = False
    Frame7.Enabled = True
    End If
    
End Sub

Private Sub Check3_Click()

    If Check3 = 1 Then
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    End If
    
    If Check3 = 0 Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    End If
    
End Sub

Private Sub Slider1_Scroll()
    
    If Slider1 = Slider1.Value Then
    Label18 = Slider1.Value
    Label21 = Int(Slider1.Value / 8)
    End If
    
    If Slider2.Enabled = True Then
    Label27 = Int(((Slider1.Value / 8) + Label23) * (slider3.Value * 60))
    Label28 = Int((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024)
    Label29 = Int(((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider2.Enabled = False Then
    Label27 = Int((Slider1.Value / 8) * (slider3.Value * 60))
    Label28 = Int(((Slider1.Value / 8) * (slider3.Value * 60)) / 1024)
    Label29 = Int((((Slider1.Value / 8) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
End Sub
Private Sub Slider2_Scroll()
    
    If Slider2 = 1 Then
    Label22 = " 32"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 2 Then
    Label22 = " 40"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 3 Then
    Label22 = " 48"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 4 Then
    Label22 = " 56"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 5 Then
    Label22 = " 64"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 6 Then
    Label22 = " 80"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 7 Then
    Label22 = " 96"
    Label23 = Label22 / 8
    End If
   
    If Slider2 = 8 Then
    Label22 = "112"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 9 Then
    Label22 = "128"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 10 Then
    Label22 = "160"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 11 Then
    Label22 = "192"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 12 Then
    Label22 = "224"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 13 Then
    Label22 = "256"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 14 Then
    Label22 = "320"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 15 Then
    Label22 = "384"
    Label23 = Label22 / 8
    End If
    
    If Slider2 = 16 Then
    Label22 = "448"
    Label23 = Label22 / 8
    End If
    
    If Slider1.Enabled = True Then
    Label27 = Int(((Slider1.Value / 8) + Label23) * (slider3.Value * 60))
    Label28 = Int((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024)
    Label29 = Int(((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider1.Enabled = False Then
    Label27 = Int((Label23) * (slider3.Value * 60))
    Label28 = Int(((Label23) * (slider3.Value * 60)) / 1024)
    Label29 = Int((((Label23) * (slider3.Value * 60)) / 1024) / 1024)
    End If

End Sub





    


Private Sub Slider3_Scroll()
    
    If slider3 = slider3.Value Then
    Label24 = Int(slider3.Value / 60)
    Label25 = slider3.Value
    Label26 = slider3.Value * 60
    End If
    
    If Slider1.Enabled And Slider2.Enabled = True Then
    Label27 = Int(((Slider1.Value / 8) + Label23) * (slider3.Value * 60))
    Label28 = Int((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024)
    Label29 = Int(((((Slider1.Value / 8) + Label23) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider1.Enabled = False And Slider2.Enabled = True Then
    Label27 = Int(Label23 * (slider3.Value * 60))
    Label28 = Int((Label23 * (slider3.Value * 60)) / 1024)
    Label29 = Int(((Label23 * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider1.Enabled = True And Slider2.Enabled = False Then
    Label27 = Int((Slider1.Value / 8) * (slider3.Value * 60))
    Label28 = Int(((Slider1.Value / 8) * (slider3.Value * 60)) / 1024)
    Label29 = Int((((Slider1.Value / 8) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider1.Enabled = False And Slider2.Enabled = False Then
    Label27 = 0
    Label28 = 0
    Label29 = 0
    End If
    
End Sub
Private Sub Slider6_Click()
End Sub


Private Sub Slider4_Scroll()

    If Slider4 = 1 Then
    Label30 = " 8000"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 8
    Label37 = "  8"
    End If
    
    If Slider4 = 2 Then
    Label30 = " 8000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 16
    Label37 = " 16"
    End If
    
    If Slider4 = 3 Then
    Label30 = " 8000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 16
    Label37 = " 16"
    End If
    
    If Slider4 = 4 Then
    Label30 = " 8000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 31
    Label37 = " 31"
    End If
    
    If Slider4 = 5 Then
    Label30 = "11025"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 11
    Label37 = " 11"
    End If
    
    If Slider4 = 6 Then
    Label30 = "11025"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 22
    Label37 = " 22"
    End If
    
    If Slider4 = 7 Then
    Label30 = "11025"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 22
    Label37 = " 22"
    End If
    
    If Slider4 = 8 Then
    Label30 = "11025"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 43
    Label37 = " 43"
    End If
    
    If Slider4 = 9 Then
    Label30 = "12000"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 12
    Label37 = " 12"
    End If
    
    If Slider4 = 10 Then
    Label30 = "12000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 23
    Label37 = " 23"
    End If
    
    If Slider4 = 11 Then
    Label30 = "12000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 23
    Label37 = " 23"
    End If
    
    If Slider4 = 12 Then
    Label30 = "12000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 47
    Label37 = " 47"
    End If
    
    If Slider4 = 13 Then
    Label30 = "16000"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 16
    Label37 = " 16"
    End If
    
    If Slider4 = 14 Then
    Label30 = "16000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 31
    Label37 = " 31"
    End If
    
    If Slider4 = 15 Then
    Label30 = "16000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 31
    Label37 = " 31"
    End If
    
    If Slider4 = 16 Then
    Label30 = "16000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 63
    Label37 = " 63"
    End If
    
    If Slider4 = 17 Then
    Label30 = "22050"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 22
    Label37 = " 22"
    End If
    
    If Slider4 = 18 Then
    Label30 = "22050"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 43
    Label37 = " 43"
    End If
    
    If Slider4 = 19 Then
    Label30 = "24000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 47
    Label37 = " 47"
    End If
    
    If Slider4 = 20 Then
    Label30 = "24000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 47
    Label37 = " 47"
    End If
    
    If Slider4 = 21 Then
    Label30 = "24000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 94
    Label37 = " 94"
    End If
    
    If Slider4 = 22 Then
    Label30 = "32000"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 31
    Label37 = " 31"
    End If
    
    If Slider4 = 23 Then
    Label30 = "32000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 63
    Label37 = " 63"
    End If
    
    If Slider4 = 24 Then
    Label30 = "32000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 63
    Label37 = " 63"
    End If
    
    If Slider4 = 25 Then
    Label30 = "32000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 125
    Label37 = "125"
    End If
    
    If Slider4 = 26 Then
    Label30 = "44100"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 43
    Label37 = " 43"
    End If
    
    If Slider4 = 27 Then
    Label30 = "44100"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 86
    Label37 = " 86"
    End If
    
    If Slider4 = 28 Then
    Label30 = "44100"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 86
    Label37 = " 86"
    End If
    
    If Slider4 = 29 Then
    Label30 = "44100"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 172
    Label37 = "172"
    End If
    
    If Slider4 = 30 Then
    Label30 = "48000"
    Label31 = " 8"
    Label32 = " Mono"
    Label33 = 47
    Label37 = " 47"
    End If
    
    If Slider4 = 31 Then
    Label30 = "48000"
    Label31 = " 8"
    Label32 = "Stereo"
    Label33 = 94
    Label37 = " 94"
    End If
        
    If Slider4 = 32 Then
    Label30 = "48000"
    Label31 = "16"
    Label32 = " Mono"
    Label33 = 94
    Label37 = " 94"
    End If
        
    If Slider4 = 33 Then
    Label30 = "48000"
    Label31 = "16"
    Label32 = "Stereo"
    Label33 = 188
    Label37 = "188"
    End If
    
    If Slider1.Enabled And Slider4.Enabled = True Then
    Label27 = Int(((Slider1.Value / 8) + Label33) * (slider3.Value * 60))
    Label28 = Int((((Slider1.Value / 8) + Label33) * (slider3.Value * 60)) / 1024)
    Label29 = Int(((((Slider1.Value / 8) + Label33) * (slider3.Value * 60)) / 1024) / 1024)
    End If
    
    If Slider1.Enabled = False Then
    Label27 = Int((Label33) * (slider3.Value * 60))
    Label28 = Int(((Label33) * (slider3.Value * 60)) / 1024)
    Label29 = Int((((Label33) * (slider3.Value * 60)) / 1024) / 1024)
    End If

End Sub
Private Sub Text1_Change()

End Sub


Private Sub Text3_Change()

End Sub


Private Sub Text4_Change()

End Sub


Private Sub Text5_Change()

End Sub


Private Sub Text6_Change()

End Sub


Private Sub Text7_Change()

End Sub


Private Sub Text10_Change()

End Sub

Private Sub Text8_Change()

End Sub


Private Sub Text9_Change()

End Sub


Private Sub Text11_Change()

End Sub


Private Sub Text14_Change()

End Sub


Private Sub Text12_Change()

End Sub


