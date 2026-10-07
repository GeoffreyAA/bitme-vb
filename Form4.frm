VERSION 4.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " DivX Quantizer Quality Ratio"
   ClientHeight    =   3735
   ClientLeft      =   645
   ClientTop       =   1815
   ClientWidth     =   7215
   Height          =   4140
   Icon            =   "Form4.frx":0000
   Left            =   585
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Top             =   1470
   Width           =   7335
   Begin VB.Frame Frame1 
      Caption         =   " DivX Quantizer Quality Ratio "
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Text            =   "2"
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   615
         Left            =   4440
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Text            =   "100%"
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Quantizer"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Quality"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   4440
         Picture         =   "Form4.frx":0ECA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   $"Form4.frx":7028
         Height          =   1575
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "2"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "31"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   495
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   630
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1111
         _Version        =   327682
         LargeChange     =   1
         Min             =   2
         Max             =   31
         SelStart        =   31
         TickStyle       =   2
         TickFrequency   =   31
         Value           =   31
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Unload Form4
    
End Sub


Private Sub Slider1_Scroll()

    If Slider1.Value = 31 Then
    Text1 = "100%"
    Text2 = 2
    End If
    
    If Slider1.Value = 30 Then
    Text1 = "97%"
    Text2 = 3
    End If
    
    If Slider1.Value = 29 Then
    Text1 = "93%"
    Text2 = 4
    End If
    
    If Slider1.Value = 28 Then
    Text1 = "90%"
    Text2 = 5
    End If
    
    If Slider1.Value = 27 Then
    Text1 = "86%"
    Text2 = 6
    End If
    
    If Slider1.Value = 26 Then
    Text1 = "83%"
    Text2 = 7
    End If
    
    If Slider1.Value = 25 Then
    Text1 = "79%"
    Text2 = 8
    End If
    
    If Slider1.Value = 24 Then
    Text1 = "76%"
    Text2 = 9
    End If
    
    If Slider1.Value = 23 Then
    Text1 = "72%"
    Text2 = 10
    End If
    
    If Slider1.Value = 22 Then
    Text1 = "69%"
    Text2 = 11
    End If
    
    If Slider1.Value = 21 Then
    Text1 = "66%"
    Text2 = 12
    End If
    
    If Slider1.Value = 20 Then
    Text1 = "62%"
    Text2 = 13
    End If
    
    If Slider1.Value = 19 Then
    Text1 = "59%"
    Text2 = 14
    End If
    
    If Slider1.Value = 18 Then
    Text1 = "55%"
    Text2 = 15
    End If
    
    If Slider1.Value = 17 Then
    Text1 = "52%"
    Text2 = 16
    End If
    
    If Slider1.Value = 16 Then
    Text1 = "48%"
    Text2 = 17
    End If
    
    If Slider1.Value = 15 Then
    Text1 = "45%"
    Text2 = 18
    End If
    
    If Slider1.Value = 14 Then
    Text1 = "41%"
    Text2 = 19
    End If
    
    If Slider1.Value = 13 Then
    Text1 = "38%"
    Text2 = 20
    End If
    
    If Slider1.Value = 12 Then
    Text1 = "34%"
    Text2 = 21
    End If
    
    If Slider1.Value = 11 Then
    Text1 = "31%"
    Text2 = 22
    End If
    
    If Slider1.Value = 10 Then
    Text1 = "28%"
    Text2 = 23
    End If
    
    If Slider1.Value = 9 Then
    Text1 = "24%"
    Text2 = 24
    End If
    
    If Slider1.Value = 8 Then
    Text1 = "21%"
    Text2 = 25
    End If
    
    If Slider1.Value = 7 Then
    Text1 = "17%"
    Text2 = 26
    End If
    
    If Slider1.Value = 6 Then
    Text1 = "14%"
    Text2 = 27
    End If
    
    If Slider1.Value = 5 Then
    Text1 = "10%"
    Text2 = 28
    End If
    
    If Slider1.Value = 4 Then
    Text1 = "7%"
    Text2 = 29
    End If
    
    If Slider1.Value = 3 Then
    Text1 = "3%"
    Text2 = 30
    End If
    
    If Slider1.Value = 2 Then
    Text1 = "0%"
    Text2 = 31
    End If
    
End Sub


