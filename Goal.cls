VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Goal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase Goal: Goal.cls
' Esta clase representa una meta del usuario

Dim m_GoalID As Long
Dim m_UserID As Long
Dim m_Title As String
Dim m_Category As String
Dim m_Description As String
Dim m_DueDate As Date
Dim m_ImagePath As String
Dim m_CreatedDate As Date
Dim m_Status As String

' Propiedades de la clase Goal
Public Property Get GoalID() As Long
    GoalID = m_GoalID
End Property

Public Property Let GoalID(Value As Long)
    m_GoalID = Value
End Property

Public Property Get userID() As Long
    userID = m_UserID
End Property

Public Property Let userID(Value As Long)
    m_UserID = Value
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(Value As String)
    m_Title = Value
End Property

Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Let Category(Value As String)
    m_Category = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get DueDate() As Date
    DueDate = m_DueDate
End Property

Public Property Let DueDate(Value As Date)
    m_DueDate = Value
End Property

Public Property Get ImagePath() As String
    ImagePath = m_ImagePath
End Property

Public Property Let ImagePath(Value As String)
    m_ImagePath = Value
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = m_CreatedDate
End Property

Public Property Let CreatedDate(Value As Date)
    m_CreatedDate = Value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Status(Value As String)
    m_Status = Value
End Property

