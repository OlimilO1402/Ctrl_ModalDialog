Attribute VB_Name = "MNew"
Option Explicit

Public Function Person(ByVal BirthDay As Date, ByVal City As City, ByVal Index As Long, ByVal Name As String) As Person
    Set Person = New Person: Person.New_ BirthDay, City, Index, Name
End Function

Public Function PersonDefault(City As City) As Person
    Set PersonDefault = MNew.Person(DateSerial(1980, 1, 1), City, 1, "Max Mustermann")
End Function

Public Function City(ByVal Name As String, ByVal PostalCode As String) As City
    Set City = New City: City.New_ Name, PostalCode
End Function

Public Function ModalDialog(aDialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
    Set ModalDialog = New ModalDialog: ModalDialog.New_ aDialog, BtnOK, BtnCancel
End Function

Public Function DBcrud(Col As Collection, ListBox As ListBox, _
                       BtnAdd As CommandButton, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, _
                       Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch) As DBcrud
    Set DBcrud = New DBcrud: DBcrud.New_ Col, ListBox, BtnAdd, BtnAddClone, BtnInsert, BtnInsertClone, BtnEdit, BtnDelete, BtnMoveUp, BtnMoveDown, BtnSortUp, BtnSortDown, BtnSearch
End Function

