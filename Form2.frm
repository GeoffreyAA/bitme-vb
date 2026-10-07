VERSION 4.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Help Me !"
   ClientHeight    =   3615
   ClientLeft      =   1740
   ClientTop       =   1290
   ClientWidth     =   6255
   Height          =   4020
   Icon            =   "Form2.frx":0000
   Left            =   1680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Top             =   945
   Width           =   6375
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   $"Form2.frx":030A
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5535
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Form2
    
End Sub


Private Sub OLE1_Updated(Code As Integer)

End Sub


Private Sub DBList1_Click()

End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub DBGrid1_Click()

End Sub


Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub Command2_Click()
    
    Load Form3
    Form3.Show
    
End Sub


