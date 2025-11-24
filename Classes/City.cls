VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name As String
Private m_Code As String

Friend Sub New_(Name As String, PostalCode As String)
    m_Name = Name: m_Code = PostalCode
End Sub

Public Sub NewC(Other As City)
    With Other: m_Name = .Name: m_Code = .PostalCode: End With
End Sub

Public Sub SetParams(Optional Name, Optional PostalCode)
    If Not IsMissing(Name) Then m_Name = Name
    If Not IsMissing(PostalCode) Then m_Code = PostalCode
End Sub

Public Function Clone() As City
    Set Clone = Me 'there is only one of it's kind
End Function

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get PostalCode() As String
    PostalCode = m_Code
End Property

Public Property Get Ptr() As LongPtr
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    'Key = m_Code & " " & m_Name
    Key = Me.Ptr
End Property

Public Function IsEqual(Other As City) As Boolean
    IsEqual = m_Name = Other.Name
End Function

Public Function IsSame(Other As City) As Boolean
    If Me Is Other Then IsSame = Me.Ptr = Other.Ptr
End Function

Public Function ToStr() As String
    ToStr = m_Code & " " & m_Name
End Function

