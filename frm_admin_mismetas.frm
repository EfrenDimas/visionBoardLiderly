VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_admin_mismetas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_misMetas 
      Caption         =   "Mis Metas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      Begin VB.Label lbl_frase_metas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   """El éxito comienza con una decisión"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton btn_agregarGoal 
      BackColor       =   &H8000000A&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton btn_Detalle 
      BackColor       =   &H8000000A&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgGoals 
      Height          =   3735
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_titleGoal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8280
      TabIndex        =   5
      Top             =   5040
      Width           =   60
   End
   Begin VB.Image ImageGoal 
      Height          =   3735
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frm_admin_mismetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_agregarGoal_Click()
frm_newGoal.Show vbModal
End Sub


Private Sub btn_Detalle_Click()

frm_detalleGoal.Show vbModal

End Sub

Private Sub dgGoals_Scroll(Cancel As Integer)
    frm_admin_mismetas.dgGoals.Columns(0).Visible = False ' Esto oculta la columna de Descripción
    frm_admin_mismetas.dgGoals.Columns(frm_admin_mismetas.dgGoals.Columns.Count - 1).Visible = False ' Esto oculta la última columna (ImagePath)
    frm_admin_mismetas.dgGoals.Columns(3).Visible = False ' Esto oculta la columna de Descripción
    
End Sub

Private Sub dgGoals_SelChange(Cancel As Integer)
    Dim selectedGoalID As Integer
    Dim selectedImagePath As String
    Dim selectedImagePathDefaul As String
    Dim fullImagePath As String
    

    ' Obtener el GoalID de la fila seleccionada
    selectedGoalID = dgGoals.Columns(0).Value ' O usar el índice de columna si lo prefieres
    
    ' Obtener la ruta de la imagen asociada a la meta seleccionada (ruta relativa)
    selectedImagePath = dgGoals.Columns(6).Value
    'selectedImagePathDefaul =
    ' Completar la ruta de la imagen (hacerla absoluta)
    ' App.Path obtiene la ruta del directorio donde se ejecuta la aplicación
    fullImagePath = App.Path & "\" & selectedImagePath
    selectedImagePathDefaul = App.Path & "\images\error.jpg"
    
    lbl_titleGoal.Top = dgGoals.Height
    lbl_titleGoal.Left = ImageGoal.Width + (ImageGoal.Width / 1.5)
    'lbl_titleGoal.Width = Me.Width - (ImageGoal.Width / 2)
    'lbl_titleGoal.Height = fra_misMetas.Height
    lbl_titleGoal.Caption = dgGoals.Columns("Title").Value & " | " & dgGoals.Columns("Category").Value
    ' Verificar si el archivo existe
    If Dir(fullImagePath) <> "" Then
        On Error GoTo ErrorHandler
        ' Si el archivo existe, cargar la imagen en el PictureBox
        ImageGoal.Picture = LoadPicture(fullImagePath)
        ImageGoal.Stretch = True
        On Error GoTo 0
        'lbl_titleGoal.Alignment = 1
    Else
        'fullImagePath = App.Path & "\images\nofound.JPG"
        ' Si no existe, mostrar un mensaje
        'MsgBox "La imagen no se encuentra: " & fullImagePath, vbExclamation
        On Error GoTo ErrorHandler
        ImageGoal.Picture = LoadPicture(selectedImagePathDefaul)
        ImageGoal.Stretch = True
        On Error GoTo 0
        'lbl_titleGoal.Caption = ""
        
    End If
    
ErrorHandler:
    ' Si ocurre un error al cargar la imagen, vaciar el PictureBox o poner una imagen predeterminada
    'MsgBox "Error al cargar la imagen. Se utilizará la imagen por defecto."
    'ImageGoal.Picture = LoadPicture(selectedImagePathDefaul)
End Sub

Private Sub Form_Activate()
    fra_misMetas.Left = 0
    fra_misMetas.Top = 0
End Sub

Private Sub Form_Initialize()
    fra_misMetas.Left = 0
    fra_misMetas.Top = 0
End Sub

Private Sub Form_Load()
    fra_misMetas.Left = 0
    fra_misMetas.Top = 0
    
    
    
    LoadGoalsIntoDataGrid


    ' Ocultar columnas según sea necesario (puedes personalizar esto)
    frm_admin_mismetas.dgGoals.Columns(0).Visible = False ' Esto oculta la columna de Descripción
    frm_admin_mismetas.dgGoals.Columns(dgGoals.Columns.Count - 1).Visible = False ' Esto oculta la última columna (ImagePath)
    frm_admin_mismetas.dgGoals.Columns(3).Visible = False ' Esto oculta la columna de Descripción
        
        
End Sub

Private Sub Form_Resize()
Me.Left = 0
Me.Top = 0
Me.Width = MDIForm_Principal.ScaleWidth
Me.Height = MDIForm_Principal.ScaleHeight

fra_misMetas.Top = 0
fra_misMetas.Width = Me.Width

'posicion de datagrid y tamaño
dgGoals.Top = fra_misMetas.Height
dgGoals.Left = 0

dgGoals.Height = Me.Height - (fra_misMetas.Height * 2) 'Espacio para los buttons de users
dgGoals.Width = Me.Width / 2


LoadGoalsIntoDataGrid

'posicion y tamaño del picturebox
ImageGoal.Top = fra_misMetas.Height * 1.6
ImageGoal.Left = Me.Width / 1.9
ImageGoal.Height = Me.Height - (fra_misMetas.Height * 4)
ImageGoal.Width = Me.Width / 2.3


lbl_titleGoal.Top = dgGoals.Height
lbl_titleGoal.Left = ImageGoal.Width + (ImageGoal.Width / 1.5)
'lbl_titleGoal.Width = Me.Width / 2
lbl_titleGoal.Height = fra_misMetas.Height
lbl_titleGoal.Font.Size = ImageGoal.Width / 900
'posicion de bottones, agregar, actualizar y eliminar
btn_agregarGoal.Top = dgGoals.Height + (fra_misMetas.Height) + (fra_misMetas.Height / 2.7)
'btn_modificarContrato.Top = dgGoals.Height + (fra_misMetas.Height) + (fra_misMetas.Height / 2.7)
btn_Detalle.Top = dgGoals.Height + (fra_misMetas.Height) + (fra_misMetas.Height / 2.7)

btn_agregarGoal.Left = (Me.Width / 4) - (btn_agregarGoal.Width / 2)
'btn_modificarContrato.Left = (Me.Width / 4) * 2 - (btn_agregarContrato.Width / 2)
btn_Detalle.Left = (Me.Width / 4) * 3 - (btn_agregarGoal.Width / 2)


lbl_frase_metas.Top = fra_misMetas.Height / 3
lbl_frase_metas.Left = fra_misMetas.Width / 2 - (lbl_frase_metas.Width / 2)

AdjustColumnWidths frm_admin_mismetas.dgGoals
'Ajuste de columnas de datagrid
End Sub





