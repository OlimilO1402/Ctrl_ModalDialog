VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name   As String
Private m_BirthD As Date
Private m_City   As City
Private m_Index  As Long

Private Sub Class_Initialize()
    'm_Name = "Max Mustermann"
    'm_BirthD = DateSerial(Year(Now), Month(Now), Day(Now))
    'Set m_City = MData.Cities_Add(MNew.City("Musterstadt", "00000"))
    'm_Index = MData.Persons_Count
End Sub

Friend Sub New_(BirthDay As Date, City As City, Index As Long, Name As String)
    m_BirthD = BirthDay: Set m_City = City: m_Index = Index: m_Name = Name
End Sub

Public Sub NewC(Other As Person)
    With Other
        m_BirthD = .BirthDay
        Set m_City = .City.Clone
        m_Name = .Name
        m_Index = .Index
    End With
End Sub

Public Sub NewV(Other)
    Dim OtherPerson As Person: Set OtherPerson = Other: NewC OtherPerson
End Sub

Public Function Clone() As Person
    Set Clone = New Person: Clone.NewC Me
End Function

Public Sub SetParams(Optional BirthDay, Optional City, Optional Index, Optional Name)
    If Not IsMissing(Name) Then m_Name = Name
    If Not IsMissing(BirthDay) Then m_BirthD = BirthDay
    If Not IsMissing(City) Then Set m_City = City
    If Not IsMissing(Index) Then m_Index = Index
End Sub

Public Property Get Ptr() As LongPtr
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    Key = Me.Ptr
End Property

Public Function IsEqual(Other As Person) As Boolean
    IsEqual = Me.BirthDay = Other.BirthDay: If Not IsEqual Then Exit Function
    IsEqual = Me.Name = Other.Name:         If Not IsEqual Then Exit Function
    IsEqual = City.IsEqual(Other.City):     If Not IsEqual Then Exit Function
End Function

Public Function IsSame(Other As Person) As Boolean
    If Me Is Other Then IsSame = Me.Ptr = Other.Ptr
End Function

Public Property Get BirthDay() As Date
    BirthDay = m_BirthD
End Property

Public Property Get City() As City
    Set City = m_City
End Property
'Public Property Set City(ByVal Value As City)
'    Set m_City = Value
'End Property

Public Property Get Index() As Long
    Index = m_Index
End Property
Public Sub IndexInc()
    m_Index = m_Index + 1
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Function ToStr() As String
    ToStr = Me.Name & "; " & m_BirthD & "; " & m_City.ToStr
End Function
