VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBcrud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mBtnAdd    As CommandButton ' 1. [ + ]  Hinzufügen,  ein neues       Objekt Erstellen, Editieren, und am Ende der Liste anfügen
Attribute mBtnAdd.VB_VarHelpID = -1
Private WithEvents mBtnAddCln As CommandButton ' 2. [ ++]  Klonen,      das ausgewählte Objekt Kopieren , Editieren, und am Ende der Liste anfügen
Attribute mBtnAddCln.VB_VarHelpID = -1
Private WithEvents mBtnIns    As CommandButton ' 3. [ i ]  Einfügen,    ein neues       Objekt Erstellen, Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
Attribute mBtnIns.VB_VarHelpID = -1
Private WithEvents mBtnInsCln As CommandButton ' 4. [i++]  Kop+Einf,    das ausgewählte Objekt Kopieren , Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
Attribute mBtnInsCln.VB_VarHelpID = -1
Private WithEvents mBtnEdit   As CommandButton ' 5. [ / ]  Edit,        das ausgewählte Objekt Editieren, einen Modalen Dialog mit den Eigenschaften des Objekts anzeigen, OK oder Abbrechen
Attribute mBtnEdit.VB_VarHelpID = -1
Private WithEvents mBtnDel    As CommandButton ' 6. [ - ]  Löschen,     das ausgewählte Objekt Löschen vorher nachfragen
Attribute mBtnDel.VB_VarHelpID = -1
Private WithEvents mBtnMovUp  As CommandButton ' 7. [ ^ ]  Nach Oben,   das ausgewählte Objekt um eine Stelle nach oben  schieben
Attribute mBtnMovUp.VB_VarHelpID = -1
Private WithEvents mBtnMovDwn As CommandButton ' 8. [ v ]  Nach Unten,  das ausgewählte Objekt um eine Stelle nach unten schieben
Attribute mBtnMovDwn.VB_VarHelpID = -1
Private WithEvents mBtnSrtUp  As CommandButton ' 9. [ |v]  Aufsteigend, alle Objekte aufsteigend Sortieren
Attribute mBtnSrtUp.VB_VarHelpID = -1
Private WithEvents mBtnSrtDwn As CommandButton '10. [ |^]  Absteigend,  alle Objekte absteigend  Sortieren
Attribute mBtnSrtDwn.VB_VarHelpID = -1
Private WithEvents mBtnSearch As CommandButton '11. [°\ ]  Suchen,      ein Objekt suchen, einen Suchdialog anzeigen, Ergebnisliste anzeigen
Attribute mBtnSearch.VB_VarHelpID = -1
Private WithEvents mListBox   As ListBox 'Or ComboBox
Attribute mListBox.VB_VarHelpID = -1
Private m_List As Collection

Friend Sub New_(Col As Collection, ListBox, BtnAdd, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch)
    Set m_List = Col: Set mListBox = ListBox
    If Not IsMissing(BtnAdd) Then Set mBtnAdd = BtnAdd
    If Not IsMissing(BtnAddClone) Then Set mBtnAddCln = BtnAddClone
    If Not IsMissing(BtnInsert) Then Set mBtnIns = BtnInsert
    If Not IsMissing(BtnInsertClone) Then Set mBtnInsCln = BtnInsertClone
    If Not IsMissing(BtnEdit) Then Set mBtnEdit = BtnEdit
    If Not IsMissing(BtnDelete) Then Set mBtnDel = BtnDelete
    If Not IsMissing(BtnMoveUp) Then Set mBtnMovUp = BtnMoveUp
    If Not IsMissing(BtnMoveDown) Then Set mBtnMovDwn = BtnMoveDown
    If Not IsMissing(BtnSortUp) Then Set mBtnSrtUp = BtnSortUp
    If Not IsMissing(BtnSortDown) Then Set mBtnSrtDwn = BtnSortDown
    If Not IsMissing(BtnSearch) Then Set mBtnSearch = BtnSearch
End Sub


Private Sub mBtnAdd_Click()
    '
End Sub

Private Sub mBtnAddCln_Click()
    '
End Sub

Private Sub mBtnDel_Click()
    '
End Sub

Private Sub mBtnEdit_Click()
    '
End Sub

Private Sub mBtnIns_Click()
    '
End Sub

Private Sub mBtnInsCln_Click()
    '
End Sub

Private Sub mBtnMovDwn_Click()
    '
End Sub

Private Sub mBtnMovUp_Click()
    '
End Sub

Private Sub mBtnSearch_Click()
    '
End Sub

Private Sub mBtnSrtDwn_Click()
    '
End Sub

Private Sub mBtnSrtUp_Click()
    '
End Sub

Private Sub mListBox_DblClick()
    '
End Sub

Private Sub mListBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyAdd:      mBtnAdd.Value = True
    Case KeyCodeConstants.vbKeySubtract: mBtnDel.Value = True
    Case KeyCodeConstants.vbKeyDivide:   mBtnEdit.Value = True
    Case 220:                            mBtnMovUp.Value = True
    Case KeyCodeConstants.vbKeyV:        mBtnMovDwn.Value = True
    End Select
End Sub

