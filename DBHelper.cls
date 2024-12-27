VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase DBHelper: DBHelper.cls
' Esta clase se encarga de las operaciones con la base de datos SQL Server

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset


Public Sub Connect()
    On Error GoTo ErrorHandler ' Manejo de errores en esta función
    


    ' Verificar si la conexión ya está abierta
    If conn Is Nothing Then
        ' Si el objeto de conexión no está inicializado, crear una nueva instancia
        Set conn = New ADODB.Connection
    End If
    
    ' Verificar si la conexión ya está abierta
    If conn.State = adStateClosed Then
        Dim server As String
        Dim database As String
        
        server = Environ("serverTestLiderly")
        database = Environ("dbTestLiderly")
        If server = "" Or database = "" Then
            MsgBox "Faltan las variables de entorno necesarias para la conexión", vbCritical
            Exit Sub
        End If
        ' Si la conexión está cerrada, abrirla
        conn.Open "Provider=SQLOLEDB;Server=" & server & ";Database=" & database & ";Trusted_Connection=yes;" 'serverTestLiderly, dbTestLiderly
        'MsgBox "Conexión abierta exitosamente."
    Else
        ' Si la conexión ya está abierta, solo mostrar un mensaje
        'MsgBox "La conexión ya está abierta."
    End If
    
    Exit Sub
    
ErrorHandler:
    'MsgBox "Error al conectar con la base de datos: " & Err.Description
End Sub


' Obtener usuario por nombre de usuario
Public Function GetUserByUsername(username As String) As user
    On Error GoTo ErrorHandler ' Manejo de errores en esta función
    
    ' Ejecutar la consulta
    Set rs = conn.Execute("SELECT * FROM Users WHERE Username = '" & username & "'")
    
    ' Verificar si se encontró un usuario
    If Not rs.EOF Then
        Dim user As New user
        user.userID = rs!userID
        user.username = rs!username
        user.Password = rs!Password
        user.FullName = rs!FullName
        user.Email = rs!Email
        user.CreatedDate = rs!CreatedDate
        Set GetUserByUsername = user
    Else
        Set GetUserByUsername = Nothing
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "Error al obtener el usuario: " & Err.Description
    Set GetUserByUsername = Nothing
End Function

' Obtener metas de un usuario por UserID
Public Function GetGoalsByUserID(userID As Long) As Collection
    On Error GoTo ErrorHandler ' Manejo de errores en esta función
    
    ' Ejecutar la consulta
    Set rs = conn.Execute("SELECT * FROM Goals WHERE UserID = " & userID)
    
    ' Crear una colección para almacenar las metas
    Dim goals As New Collection
    
    ' Recorrer los resultados y agregar cada meta a la colección
    While Not rs.EOF
        Dim goalUnique As goal
        Set goalUnique = New goal
        goalUnique.GoalID = rs!GoalID
        goalUnique.userID = rs!userID
        goalUnique.Title = rs!Title
        goalUnique.Category = rs!Category
        goalUnique.Description = rs!Description
        goalUnique.DueDate = rs!DueDate
        goalUnique.ImagePath = rs!ImagePath
        goalUnique.CreatedDate = rs!CreatedDate
        goalUnique.Status = rs!Status
        rs.MoveNext
        goals.Add goalUnique
    Wend
    
    ' Retornar la colección de metas
    Set GetGoalsByUserID = goals
    
    Exit Function

ErrorHandler:
    MsgBox "Error al obtener las metas: " & Err.Description
    Set GetGoalsByUserID = Nothing
End Function
' Ejecutar una consulta SQL (como INSERT, UPDATE, DELETE)
Public Sub ExecuteSQL(sql As String)
    ' Inicializar el objeto Command
    Set cmd = New ADODB.Command
    
    ' Asignar el comando a la conexión
    Set cmd.ActiveConnection = conn
    
    ' Asignar la consulta SQL al objeto Command
    cmd.CommandText = sql
    
    ' Ejecutar la consulta SQL
    cmd.Execute
    
    ' Limpiar el objeto cmd
    Set cmd = Nothing
End Sub


Public Sub CloseConnection()
    On Error Resume Next ' Para evitar errores si la conexión ya está cerrada
    
    ' Verificar si la conexión está abierta antes de cerrarla
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
    End If
    
    Set conn = Nothing ' Liberar el objeto de conexión
End Sub

