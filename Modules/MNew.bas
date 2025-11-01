Attribute VB_Name = "MNew"
Option Explicit

Public Function Person(ByVal BirthDay As Date, ByVal City As City, ByVal Index As Long, ByVal Name As String) As Person
    Set Person = New Person: Person.New_ BirthDay, City, Index, Name
End Function

Public Function City(ByVal Name As String, ByVal PostalCode As String) As City
    Set City = New City: City.New_ Name, PostalCode
End Function

Public Function ModalDialog(aDialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
    Set ModalDialog = New ModalDialog: ModalDialog.New_ aDialog, BtnOK, BtnCancel
End Function

