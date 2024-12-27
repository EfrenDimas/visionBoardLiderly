VERSION 5.00
Begin VB.Form frm_detalleGoal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   240
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btn_atrasDetalle 
      Caption         =   "Atras"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.Frame fra_detalleGoal 
      Caption         =   "Detalle de Meta"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txt_descripcion_detalle 
         Height          =   855
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CommandButton btn_CompletarMeta 
         Caption         =   "Completar"
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Image ImgbackgroundDetalle 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label lbl_colorstatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   3960
         Width           =   285
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl_fecha_creacionDetalle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label lbl_fechaTermino 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo:"
         Height          =   195
         Index           =   1
         Left            =   765
         TabIndex        =   9
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de creacion:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label lbl_idUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Usuario:"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_tituloGoalDetalle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label lbl_tipoSuscriptor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha compromiso:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frm_detalleGoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_atrasDetalle_Click(Index As Integer)
Unload Me
End Sub

Private Sub Cbo_SucursalNewSuscriptor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cbo_tipoSusNewSuscriptor.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub cbo_tipoSusNewSuscriptor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_nombreNewSuscriptor.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub btn_CompletarMeta_Click()
    'Set goals = dbHelper.GetGoalsByUserID(GlobalCurrentUser.userID)
    Dim selectedGoalID As Integer
    'Dim dbHelper_insert As New dbHelper
    
    selectedGoalID = frm_admin_mismetas.dgGoals.Columns(0).Value
    UpdateGoalStatus selectedGoalID, "Completado"
    'Dim selectedGoalID As Integer
    selectedGoalID = frm_admin_mismetas.dgGoals.Columns(0).Value ' Suponiendo que GoalID está en la columna 0
    LoadGoalsIntoDataGrid
    CargarDetalleMeta selectedGoalID
    
End Sub

Private Sub Form_Load()
    Dim selectedGoalID As Integer
    selectedGoalID = frm_admin_mismetas.dgGoals.Columns(0).Value ' Suponiendo que GoalID está en la columna 0
    CargarDetalleMeta selectedGoalID
End Sub

Function GetGoalByID(ByVal GoalID As Integer) As goal
    
    Dim g As goal
    Set GetGoalByID = Nothing ' Inicializamos el valor de retorno como Nothing
    
    For Each g In goals
        If g.GoalID = GoalID Then
            Set GetGoalByID = g ' Devolvemos la meta encontrada
            Exit Function
        End If
    Next g
End Function

Private Sub UpdateGoalStatus(GoalID As Integer, newStatus As String)
    Dim dbHelper As New dbHelper
    Dim sql As String

    ' Conectar a la base de datos
    dbHelper.Connect

    ' Construir la consulta SQL para actualizar el campo 'Status'
    sql = "UPDATE Goals SET Status = '" & newStatus & "' WHERE GoalID = " & GoalID

    ' Ejecutar la consulta SQL
    dbHelper.ExecuteSQL sql
    MsgBox "Meta completada, ¡Felicidades!"
    ' Cerrar la conexión a la base de datos
    dbHelper.CloseConnection
End Sub

' Función independiente para cargar la información de la meta
Private Sub CargarDetalleMeta(selectedGoalID As Integer)
    Dim selectedGoal As goal
    Dim selectedImagePath As String
    Dim selectedImagePathDefaul As String
    Dim fullImagePath As String
    
    ' Buscar la meta en la colección global
    Set selectedGoal = GetGoalByID(selectedGoalID)

    ' Si se encontró la meta, mostrar los detalles en los controles del formulario
    If Not selectedGoal Is Nothing Then
        lbl_idUser.Caption = selectedGoal.userID
        lbl_tituloGoalDetalle.Caption = selectedGoal.Title
        txt_descripcion_detalle.Text = selectedGoal.Description
        lbl_fecha_creacionDetalle.Caption = selectedGoal.CreatedDate
        lbl_fechaTermino.Caption = selectedGoal.DueDate
        lbl_status.Caption = selectedGoal.Status
        
        selectedImagePath = frm_admin_mismetas.dgGoals.Columns(6).Value
        fullImagePath = App.Path & "\" & selectedImagePath
        selectedImagePathDefaul = App.Path & "\images\error.jpg"
        
        If Dir(fullImagePath) <> "" Then
            On Error GoTo ErrorHandler
            ' Si el archivo existe, cargar la imagen en el PictureBox
            ImgbackgroundDetalle.Picture = LoadPicture(fullImagePath)
            ImgbackgroundDetalle.Stretch = True
            On Error GoTo 0
        Else
            ' Si no existe, mostrar una imagen por defecto
            On Error GoTo ErrorHandler
            ImgbackgroundDetalle.Picture = LoadPicture(selectedImagePathDefaul)
            ImgbackgroundDetalle.Stretch = True
            On Error GoTo 0
        End If
        
        If selectedGoal.Status = "Pendiente" Then
            btn_CompletarMeta.Visible = True
            lbl_colorstatus.BackColor = RGB(255, 165, 0)
        Else
            btn_CompletarMeta.Visible = False
            lbl_colorstatus.BackColor = RGB(0, 255, 0)
        End If
        
    Else
        MsgBox "No se encontró la meta seleccionada", vbExclamation
    End If

Exit Sub

ErrorHandler:
    ' Manejador de errores, en caso de que haya problemas al cargar la imagen
    MsgBox "Error al cargar la imagen", vbExclamation
    Resume Next
End Sub

