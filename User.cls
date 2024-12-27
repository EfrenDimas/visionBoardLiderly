VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase User: User.cls
' Esta clase representa a un usuario

Dim m_UserID As Long
Dim m_Username As String
Dim m_Password As String
Dim m_FullName As String
Dim m_Email As String
Dim m_CreatedDate As Date
Dim m_LastLoginDate As Date

' Propiedades de la clase User
Public Property Get userID() As Long
    userID = m_UserID
End Property

Public Property Let userID(Value As Long)
    m_UserID = Value
End Property

Public Property Get username() As String
    username = m_Username
End Property

Public Property Let username(Value As String)
    m_Username = Value
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Password(Value As String)
    m_Password = Value
End Property

Public Property Get FullName() As String
    FullName = m_FullName
End Property

Public Property Let FullName(Value As String)
    m_FullName = Value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(Value As String)
    m_Email = Value
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = m_CreatedDate
End Property

Public Property Let CreatedDate(Value As Date)
    m_CreatedDate = Value
End Property

Public Property Get LastLoginDate() As Date
    LastLoginDate = m_LastLoginDate
End Property

Public Property Let LastLoginDate(Value As Date)
    m_LastLoginDate = Value
End Property

' Método para verificar la contraseña
Public Function VerifyPassword(inputPassword As String) As Boolean
    If m_Password = inputPassword Then
        VerifyPassword = True
    Else
        VerifyPassword = False
    End If
End Function

