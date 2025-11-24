VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModalDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object As Object
Private m_Dialog As Object
Private WithEvents mBtnOK     As CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1

Private Sub Class_Initialize()
    m_Result = VbMsgBoxResult.vbCancel
End Sub

Friend Sub New_(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton)
    Set m_Dialog = Dialog: Set mBtnOK = BtnOK: Set mBtnCancel = BtnCancel
End Sub

'Copy to MNew.bas:
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel
'End Function
'example for using this class in a ButtonClick on a Form:
'Private Sub BtnCityAdd_Click()
'    Dim City As City: Set City = MNew.City("Musterstadt", "00000")
'    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
'    MData.Cities_AddLB LstCities, City
'End Sub

Public Function ShowDialog(Obj As Object, Owner As Form) As VbMsgBoxResult
    Set m_Object = Obj.Clone
    m_Dialog.UpdateView m_Object
    m_Dialog.Show vbModal, Owner
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Obj.NewC m_Object
End Function

Private Sub mBtnOK_Click()
    If Not m_Dialog.UpdateData(m_Object) Then Exit Sub
    m_Result = VbMsgBoxResult.vbOK
    Unload m_Dialog
End Sub

Private Sub mBtnCancel_Click()
    Unload m_Dialog
End Sub
