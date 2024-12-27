VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm_Principal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "My Portal | Vision Board"
   ClientHeight    =   5235
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10440
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBarBase 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   661
      SimpleText      =   "visual"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Usuario"
            TextSave        =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_Perfil 
      Caption         =   "Perfil"
      Begin VB.Menu mnu_cerrarSesion 
         Caption         =   "Cerrar Sesiòn"
      End
      Begin VB.Menu mnu_miPerfil 
         Caption         =   "Mi perfil"
      End
   End
   Begin VB.Menu mnu_metas 
      Caption         =   "Metas"
   End
End
Attribute VB_Name = "MDIForm_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frm_baseMain.Visible = True
frm_baseMain.Width = MDIForm_Principal.ScaleWidth
frm_baseMain.Height = MDIForm_Principal.ScaleHeight
StatusBarBase.Panels(1).Text = "Nombre: " + CStr(GlobalCurrentUser.FullName)
    
End Sub

Private Sub MDIForm_Resize()
'frm_baseMain.Visible = True
frm_admin_mismetas.Width = Me.ScaleWidth
frm_admin_mismetas.Height = Me.ScaleHeight

frm_baseMain.Width = Me.ScaleWidth
frm_baseMain.Height = Me.ScaleHeight


End Sub

Private Sub mnu_cerrarSesion_Click()
Unload Me
frm_login.Show
End Sub

Private Sub mnu_metas_Click()
    If frm_baseMain.Visible Then
        frm_baseMain.Hide
        
        frm_admin_mismetas.Show
    Else
        frm_admin_mismetas.Hide
        frm_baseMain.Show
    End If
        
    frm_admin_mismetas.Left = 0
    frm_admin_mismetas.Top = 0

End Sub
