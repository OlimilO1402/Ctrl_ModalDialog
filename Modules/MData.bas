Attribute VB_Name = "MData"
Option Explicit

Private m_Persons As Collection
Private m_Cities  As Collection

Public Sub Init()
    Set m_Persons = New Collection
    Set m_Cities = New Collection
End Sub

' v ' ############################## ' v '    ListBox    ' v ' ############################## ' v '
Public Sub ListBox_Add(aLB As ListBox, ByVal Object As Object)
    aLB.AddItem Object.ToStr
    aLB.ItemData(aLB.ListCount - 1) = Object.Key
End Sub

Public Sub ListBox_Swap(aLB As ListBox, ByVal i1 As Long, ByVal i2 As Long)
    Dim lc As Long: lc = aLB.ListCount
    If i1 < 0 Or lc - 1 < i1 Then Exit Sub
    If i2 < 0 Or lc - 1 < i2 Then Exit Sub
    With aLB
        Dim tmp As String, tid As Long
              tmp = .List(i1):           tid = .ItemData(i1)
        .List(i1) = .List(i2): .ItemData(i1) = .ItemData(i2)
        .List(i2) = tmp:       .ItemData(i2) = tid
    End With
End Sub

Public Function Listbox_IsOutOfBounds(aLB As ListBox, ByVal i As Long) As Boolean
    'returns true if i is out of bounds
    Dim lc As Long: lc = aLB.ListCount
    Listbox_IsOutOfBounds = i < 0 Or lc - 1 < i
End Function

Public Function Listbox_IsOutOfBounds2(aLB As ListBox, ByVal i1 As Long, ByVal i2 As Long) As Boolean
    'returns true  if one  of i1, i2 is  out of bounds
    'returns false if both of i1, i2 are inside bounds
    Dim lc As Long: lc = aLB.ListCount
    If (0 <= i1 And i1 < lc) And (0 <= i2 And i2 < lc) Then Exit Function
    Listbox_IsOutOfBounds2 = True
End Function

Public Sub ListBox_Remove(aLB As ListBox, ByVal i As Long)
    With aLB
        If 0 <= i And i < .ListCount Then
            .RemoveItem i
            .ListIndex = i
        End If
    End With
End Sub

Public Sub ListBox_MoveUp(aLB As ListBox) ', ByVal i As Long)
    Dim i1 As Long: i1 = aLB.ListIndex
    Dim i2 As Long: i2 = i1 - 1
    If Listbox_IsOutOfBounds2(aLB, i1, i2) Then Exit Sub
    ListBox_Swap aLB, i1, i2
    aLB.ListIndex = i2
End Sub

Public Sub ListBox_MoveDown(aLB As ListBox) ', ByVal i As Long)
    Dim i1 As Long: i1 = aLB.ListIndex
    Dim i2 As Long: i2 = i1 + 1
    If Listbox_IsOutOfBounds2(aLB, i1, i2) Then Exit Sub
    ListBox_Swap aLB, i1, i2
    aLB.ListIndex = i2
End Sub

' v ############################## v '    Cities     ' v ############################## v '
Public Function Cities_Add(City As City) As City
    'if the City already exists then we just return the one that is already there
    Set Cities_Add = Col_AddOrGet(m_Cities, City)
End Function

Public Function Cities_AddLB(aLB As ListBox, City As City) As City
    'if the City already exists then we just return the one that is already there
    Set Cities_AddLB = Col_AddOrGet(m_Cities, City)
    ListBox_Add aLB, City
End Function

Public Function Cities_Count() As Long
    Cities_Count = m_Cities.Count
End Function

Public Function Cities_Contains(Key As String) As Boolean
    Cities_Contains = Col_Contains(m_Cities, Key)
End Function

Public Sub Cities_Remove(City As City)
    MPtr.Col_Remove m_Cities, City
    'und jetzt noch in allen Persons entfernen?
End Sub

Public Property Get Cities_ObjectFromListCtrl(ComboBoxOrListBox, i_out As Long) As City
    Set Cities_ObjectFromListCtrl = Col_ObjectFromListCtrl(m_Cities, ComboBoxOrListBox, i_out)
End Property

Public Property Get Cities_IndexFromObject(City As City) As Long ', i_out As Long) As City
    Cities_IndexFromObject = MPtr.Col_IndexFromObject(m_Cities, City)
End Property

Public Sub Cities_ToListCtrl(ComboBoxOrListBox)
    Col_ToListCtrl m_Cities, ComboBoxOrListBox, False, True
End Sub

Public Sub Cities_MoveUp(aLB As ListBox) ', ByVal Index As Long)
    Dim i As Long: i = aLB.ListIndex
    MPtr.Col_MoveUpKey m_Cities, i + 1
    ListBox_MoveUp aLB
End Sub

Public Sub Cities_MoveDown(aLB As ListBox) 'ByVal Index As Long)
    Dim i As Long: i = aLB.ListIndex
    MPtr.Col_MoveDownKey m_Cities, i + 1
    ListBox_MoveDown aLB
End Sub
' ^ ############################## ^ '    Cities     ' ^ ############################## ^ '

' v ############################## v '    Persons    ' v ############################## v '
Public Function Persons_Add(Person As Person) As Person
    Set Persons_Add = MPtr.Col_AddOrGet(m_Persons, Person) 'Persons.Add( Person)', CStr(Person.Key)
End Function

Public Function Persons_AddLB(aLB As ListBox, Person As Person) As Person
    Set Persons_AddLB = MPtr.Col_AddOrGet(m_Persons, Person)
    ListBox_Add aLB, Person
End Function

Public Function Persons_Count() As Long
    Persons_Count = m_Persons.Count
End Function

Public Function Persons_UsingCity(City As City) As Collection
    Dim v, p As Person, c As New Collection
    For Each v In m_Persons
        Set p = v
        If p.City.IsSame(City) Then
            c.Add p
        End If
    Next
    Set Persons_UsingCity = c
End Function

Public Function Persons_Contains(ByVal Key As String) As Boolean
    Persons_Contains = Col_Contains(m_Persons, Key)
End Function

Public Sub Persons_Remove(Person As Person)
    MPtr.Col_Remove m_Persons, Person
    'Dim p As Person
    'For Each p In Persons
    '    If p.IsSame(Person) Then
    '        If Persons_Contains(Person.Key) Then Persons.Remove Person.Key
    '    End If
    'Next
End Sub

Public Property Get Persons_ObjectFromListCtrl(ComboBoxOrListBox, i_out As Long) As Person
    Set Persons_ObjectFromListCtrl = Col_ObjectFromListCtrl(m_Persons, ComboBoxOrListBox, i_out)
End Property

Public Sub Persons_ToListCtrl(ComboBoxOrListBox)
    Col_ToListCtrl m_Persons, ComboBoxOrListBox, False, True
End Sub

Public Property Get Persons_Item(Index) As Person
    Set Persons_Item = m_Persons.Item(Index + 1)
End Property

Public Sub Persons_MoveUp(aLB As ListBox) ', ByVal Index As Long)
    Dim i As Long: i = aLB.ListIndex
    MPtr.Col_MoveUpKey m_Persons, i + 1
    ListBox_MoveUp aLB
End Sub

Public Sub Persons_MoveDown(aLB As ListBox) 'ByVal Index As Long)
    Dim i As Long: i = aLB.ListIndex
    MPtr.Col_MoveDownKey m_Persons, i + 1
    ListBox_MoveDown aLB
End Sub

' ^ ############################## ^ '    Persons     ' ^ ############################## ^ '

'Public Function Date_TryParse(ByVal s As String, d_out As Date) As Boolean
'Try: On Error GoTo Catch
'    d_out = CDate(s)
'    Date_TryParse = True
'Catch:
'End Function


