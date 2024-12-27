VERSION 5.00
Begin VB.Form frm_baseMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_baseMain.frx":0000
   ScaleHeight     =   5430
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.Label lbl_bienvenidos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BIENVENIDOS"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Width           =   4365
   End
End
Attribute VB_Name = "frm_baseMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Me.Left = 0
Me.Top = 0

Me.Width = MDIForm_Principal.ScaleWidth
Me.Height = MDIForm_Principal.ScaleHeight

lbl_bienvenidos.Caption = "BIENVENIDO " + GlobalCurrentUser.FullName
lbl_bienvenidos.Left = (MDIForm_Principal.ScaleWidth / 2) - (lbl_bienvenidos.Width / 2)
lbl_bienvenidos.Top = MDIForm_Principal.ScaleHeight / 4

End Sub

