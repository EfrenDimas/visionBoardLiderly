Attribute VB_Name = "modGlobals"
Public GlobalCurrentUser As user
Public goals As Collection


Public Sub LoadGoalsIntoDataGrid()
    ' Definir las variables necesarias
    Dim dbHelper As New dbHelper
    'Dim goals As Collection
    Dim goal As goal
    Dim rs As ADODB.Recordset
    
    ' Limpiar el DataGrid antes de agregar nuevas metas
    Set frm_admin_mismetas.dgGoals.DataSource = Nothing
    
    ' Conectar a la base de datos
    dbHelper.Connect
    
    Set goals = New Collection
    ' Obtener las metas del usuario actualmente logueado (usando GlobalCurrentUser)
    Set goals = dbHelper.GetGoalsByUserID(GlobalCurrentUser.userID)
    
    ' Verificar si existen metas
    If goals.Count > 0 Then
        ' Crear un Recordset para enlazar con el DataGrid
        Set rs = New ADODB.Recordset
        rs.Fields.Append "GoalID", adInteger
        rs.Fields.Append "Title", adVarWChar, 255
        rs.Fields.Append "Category", adVarWChar, 255
        rs.Fields.Append "Description", adVarWChar, 255
        rs.Fields.Append "DueDate", adDate
        rs.Fields.Append "Status", adVarWChar, 255
        rs.Fields.Append "ImagePath", adVarWChar, 255 ' Agregar el campo de la imagen
        rs.Open
        
        ' Recorrer las metas y agregar los datos al Recordset
        For Each goal In goals
            rs.AddNew
            rs.Fields("GoalID").Value = goal.goalID
            rs.Fields("Title").Value = goal.Title
            rs.Fields("Category").Value = goal.Category
            rs.Fields("Description").Value = goal.Description
            rs.Fields("DueDate").Value = goal.DueDate
            rs.Fields("Status").Value = goal.Status
            rs.Fields("ImagePath").Value = goal.ImagePath ' Asignar la ruta de la imagen
        Next goal
        
        ' Asignar el Recordset al DataGrid
        Set frm_admin_mismetas.dgGoals.DataSource = rs
        'frm_admin_mismetas.dgGoals.Columns()
        frm_admin_mismetas.dgGoals.Columns(0).Visible = False ' Esto oculta la columna de Descripción
        frm_admin_mismetas.dgGoals.Columns(frm_admin_mismetas.dgGoals.Columns.Count - 1).Visible = False ' Esto oculta la última columna (ImagePath)
        frm_admin_mismetas.dgGoals.Columns(3).Visible = False ' Esto oculta la columna de Descripción
        AdjustColumnWidths frm_admin_mismetas.dgGoals

    Else
        ' Si no hay metas, puedes mostrar un mensaje o realizar alguna otra acción
        MsgBox "No tienes metas aún.", vbInformation
    End If

    ' Cerrar la conexión a la base de datos
    dbHelper.CloseConnection
End Sub

Public Sub AdjustColumnWidths(dg As DataGrid)
    Dim i As Integer
    
    ' Recorrer todas las columnas y ajustar el ancho
    For i = 0 To dg.Columns.Count - 1
        dg.Columns(i).Width = dg.Width / 4
    Next i
    frm_admin_mismetas.dgGoals.Columns(frm_admin_mismetas.dgGoals.Columns.Count - 1).Visible = False ' Esto oculta la última columna (ImagePath)
    frm_admin_mismetas.dgGoals.Columns(3).Visible = False ' Esto oculta la columna de Descripción
    frm_admin_mismetas.dgGoals.Columns(0).Visible = False ' Esto oculta la columna de Descripción
End Sub





