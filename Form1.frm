VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   0  'None
   Caption         =   "Inicio de Sesiòn"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelarLogin 
      BackColor       =   &H00C0C000&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd_entrarLogin 
      BackColor       =   &H00C0C000&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txt_password 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txt_userName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label titleLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Vision Board"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   240
      Picture         =   "Form1.frx":1084A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label lbl_userName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   2445
      TabIndex        =   1
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lbl_userName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2460
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Formulario Login
Option Explicit

Private Sub cmd_cancelarLogin_Click()
    Unload Me
End Sub

Private Sub cmd_entrarLogin_Click()
    
    Dim dbHelper As New dbHelper
    Dim userlogin As user
    Dim username As String
    Dim Password As String

    ' Obtener los valores ingresados por el usuario
    username = txt_userName.Text
    Password = txt_password.Text
    
    
    ' Verificar si los campos están vacíos
    If username = "" Or Password = "" Then
        MsgBox "Por favor, ingresa tu nombre de usuario y contraseña", vbExclamation
        Exit Sub
    End If

    ' Intentar conectar a la base de datos
    On Error GoTo ErrorHandler
    dbHelper.Connect

    ' Obtener el usuario por nombre de usuario
    Set userlogin = dbHelper.GetUserByUsername(username)

    If Not userlogin Is Nothing Then
        ' Verificar si la contraseña es correcta
        If userlogin.VerifyPassword(Password) Then
            
            Set GlobalCurrentUser = userlogin
            ' Si las credenciales son correctas, muestra la pantalla principal
            MsgBox "Inicio de sesión exitoso", vbInformation
            MDIForm_Principal.Show
            Me.Hide
        Else
            MsgBox "Contraseña incorrecta", vbCritical
        End If
    Else
        MsgBox "Usuario no encontrado", vbCritical
    End If

    dbHelper.CloseConnection
    Exit Sub

ErrorHandler:
    MsgBox "Error al conectar con la base de datos: " & Err.Description, vbCritical
    dbHelper.CloseConnection
End Sub

Private Sub Form_Load()
    txt_userName.Text = ""
    txt_password.Text = ""
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_userName.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txt_userName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_password.SetFocus
        KeyAscii = 0
    End If
End Sub

