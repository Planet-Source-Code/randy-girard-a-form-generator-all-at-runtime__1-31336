VERSION 5.00
Begin VB.Form frmMake 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   $"Form2.frx":0000
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "HELLO WORLD!!!!!!"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheControls As New Collection
Dim TheControlTypes As New Collection
Dim TheControlNames As New Collection

Private Sub Form_Unload(Cancel As Integer)
  DeleteForm Me.Tag
End Sub
