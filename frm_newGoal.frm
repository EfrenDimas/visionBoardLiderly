VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_newGoal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
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
   Begin VB.Frame fra_newContrato 
      Caption         =   "Nueva Meta"
      Height          =   6030
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4470
      Begin VB.CommandButton btn_cargarImg_newGoal 
         Caption         =   "Cargar imagen"
         Height          =   360
         Left            =   2640
         TabIndex        =   13
         Top             =   4320
         Width           =   1350
      End
      Begin VB.TextBox txt_descripcion_newGoal 
         Height          =   2175
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPickerNewGoal 
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   181075969
         CurrentDate     =   45652
         MinDate         =   44562
      End
      Begin VB.ComboBox cbo_categoriaNewGoal 
         Height          =   315
         ItemData        =   "frm_newGoal.frx":0000
         Left            =   1800
         List            =   "frm_newGoal.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_title_newGoal 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton btn_CancelarNewContrato 
         Caption         =   "Cancelar"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton btn_agregarNewGoal 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   1
         Top             =   5160
         Width           =   855
      End
      Begin VB.Image img_newGoal 
         Height          =   1575
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl_descripcion_newGoal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label lbl_nameSuscriptorSelect 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2025
         TabIndex        =   7
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lbl_fechavencimiento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha limite de cumplimiento"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lbl_tipoSuscriptor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Index           =   4
         Left            =   540
         TabIndex        =   4
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   450
      End
      Begin VB.Label lbl_idNewSuscriptor 
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
         TabIndex        =   2
         Top             =   720
         Width           =   60
      End
   End
End
Attribute VB_Name = "frm_newGoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_agregarNewGoal_Click(Index As Integer)
    Dim dbHelper_insert As New dbHelper
    Dim sql As String
    Dim nuevaMeta As goal
    Set nuevaMeta = New goal
    
    If AreFieldsFilled() Then
        If CommonDialog1.FileName <> "" Then
            If DTPickerNewGoal.Value > Date Then
               
               nuevaMeta.Title = txt_title_newGoal.Text
               nuevaMeta.userID = CStr(GlobalCurrentUser.userID)
               nuevaMeta.Category = cbo_categoriaNewGoal.Text
               nuevaMeta.Description = txt_descripcion_newGoal.Text
               nuevaMeta.DueDate = DTPickerNewGoal.Value
               nuevaMeta.ImagePath = GetFileName(CommonDialog1.FileName)
               nuevaMeta.Status = "Pendiente"
               
               Dim dueDateStr As String
                dueDateStr = Format(nuevaMeta.DueDate, "mm/dd/yyyy")
                Call CopyImageToProjectFolder(CommonDialog1.FileName)
               On Error GoTo ErrorHandler
                dbHelper_insert.Connect
               
               sql = "INSERT INTO Goals (Title, UserID, Category, Description, DueDate, ImagePath, Status) " & _
                "VALUES ('" & nuevaMeta.Title & "', " & nuevaMeta.userID & ", '" & nuevaMeta.Category & "', '" & nuevaMeta.Description & "', '" & dueDateStr & "', '" & nuevaMeta.ImagePath & "', '" & nuevaMeta.Status & "')"
                'Debug.Print sql
                dbHelper_insert.ExecuteSQL sql
                Unload Me
               MsgBox "Se ha registrado su nueva meta!"
               LoadGoalsIntoDataGrid
                        ' Ocultar columnas según sea necesario (puedes personalizar esto)
                
                
               
               
            Else
                MsgBox "Selecciona una fecha mayor a la de hoy"
            End If
        Else
            MsgBox "carga una imagen"
        End If
    Else
        MsgBox "Debes llenar todos los campos"
    End If
    dbHelper_insert.CloseConnection
    Exit Sub
ErrorHandler:
    'MsgBox "Error al conectar con la base de datos: " & Err.Description, vbCritical
    dbHelper_insert.CloseConnection
End Sub

Private Sub btn_CancelarNewContrato_Click(Index As Integer)
Unload Me
End Sub

Private Sub btn_cargarImg_newGoal_Click()

    CommonDialog1.Filter = "Archivos de imagen (*.jpg; *.jpeg; *.png; *.gif)|*.jpg;*.jpeg;*.png;*.gif|Todos los archivos (*.*)|*.*"

    ' Configurar el cuadro de diálogo para abrir un archivo
    CommonDialog1.ShowOpen

    ' Verificar si se seleccionó un archivo
    If Len(CommonDialog1.FileName) > 0 Then
        ' Cargar la imagen seleccionada en el PictureBox
        img_newGoal.Picture = LoadPicture(CommonDialog1.FileName)
    End If
End Sub

Private Sub Form_Load()
fra_newContrato.Left = 0
fra_newContrato.Top = 0

End Sub

Function AreFieldsFilled() As Boolean
    Dim ctl As Control

    ' Asumimos que todos los TextBox deben estar llenos
    AreFieldsFilled = True
    
    ' Recorremos todos los controles del formulario
    For Each ctl In Me.Controls
        ' Verificamos si el control es un TextBox
        If TypeOf ctl Is TextBox Then
            ' Si el TextBox está vacío, retornamos False
            If ctl.Text = "" Then
                AreFieldsFilled = False
                Exit Function ' Salimos de la función, ya que encontramos un TextBox vacío
            End If
        End If
        
        ' Verificamos si el control es un ComboBox
        If TypeOf ctl Is ComboBox Then
            ' Comprobamos si el ComboBox está vacío (ningún ítem seleccionado)
            If ctl.ListIndex = -1 Then
                AreFieldsFilled = False
                Exit Function ' Si no se ha seleccionado ningún valor, salimos de la función
            End If
        End If
    Next ctl
    
    
End Function

Function GetFileName(ByVal fullPath As String) As String
    ' Utilizar InStrRev para encontrar la última barra invertida
    Dim lastSlashPos As Integer
    lastSlashPos = InStrRev(fullPath, "\")
    
    ' Retornar el nombre del archivo
    GetFileName = "images/" & Right(fullPath, Len(fullPath) - lastSlashPos)
End Function



Private Sub CopyImageToProjectFolder(imgPath As String)
    Dim newPath As String
    
    ' Verificar si la carpeta 'images' existe en el directorio del proyecto
    If Dir(App.Path & "\images", vbDirectory) = "" Then
        ' Si no existe, crearla
        MkDir App.Path & "\images"
    End If
    
    ' Construir la nueva ruta de destino
    newPath = App.Path & "\images\" & Mid(imgPath, InStrRev(imgPath, "\") + 1)
    
    ' Verificar si el archivo ya existe en la carpeta de destino
    If Dir(newPath) <> "" Then
        ' Si el archivo ya existe, mostrar mensaje de registro
        'MsgBox "La imagen ya está registrada en la carpeta de destino.", vbInformation
        Exit Sub
    End If
    
    ' Si el archivo no existe, copiar la imagen a la carpeta 'images'
    On Error GoTo ErrorHandler
    FileCopy imgPath, newPath
    
    'MsgBox "Imagen copiada con éxito a: " & newPath
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocurrió un error al copiar la imagen: " & Err.Description, vbCritical
End Sub

