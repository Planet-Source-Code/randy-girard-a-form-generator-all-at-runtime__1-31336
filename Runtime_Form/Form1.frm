VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName3 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtCaption2 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtName2 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Form"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit Form"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unload Form"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label6 
      Caption         =   "Name"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "I used the caption property just as an example, of course any property's could be used."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "New Caption"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Caption"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim MaForm As New frmMake
  MaForm.Caption = txtCaption.Text
  AddForm MaForm, txtName.Text
  Dim MaForm1 As Form
  Set MaForm1 = MaForms(MaForms.Count)
  MaForm1.Show
  Command2.Enabled = True
  Command3.Enabled = True
  txtName2.Text = txtName.Text
  txtCaption2.Text = txtCaption.Text
  txtName3.Text = txtName.Text
End Sub

Private Sub Command2_Click()
  On Error Resume Next
  Dim MaForm1 As Form
  Set MaForm1 = GetForm(txtName2.Text)
  MaForm1.Caption = txtCaption2.Text
End Sub

Private Sub Command3_Click()
  On Error Resume Next
  Dim MaForm1 As Form
  Set MaForm1 = GetForm(txtName3.Text)
  Unload MaForm1
End Sub

Private Sub txtName_Change()
  Command1.Enabled = True
End Sub
