VERSION 5.00
Begin VB.Form FPerson 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit Person"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtName 
      Height          =   435
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtBirthDay 
      Height          =   435
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox CmbCity 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   570
   End
   Begin VB.Label LblBirthDay 
      AutoSize        =   -1  'True
      Caption         =   "Birthday:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
End
Attribute VB_Name = "FPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Person As Person

Private Sub Form_Load()
    MData.Cities_ToListCtrl CmbCity
End Sub

Public Sub UpdateView(Obj)
    Set m_Person = Obj
    If m_Person Is Nothing Then MsgBox "The Person does not exist": Exit Sub
    TxtName.Text = m_Person.Name
    TxtBirthDay.Text = m_Person.BirthDay
    If m_Person.City Is Nothing Then MsgBox "The City does not exist": Exit Sub
    CmbCity.ListIndex = MData.Cities_IndexFromObject(m_Person.City)
    CmbCity.Text = m_Person.City.ToStr
End Sub

Public Function UpdateData(Obj) As Boolean
Try: On Error GoTo Catch
    Dim bIsOK As Boolean
    Dim bd As Date:     bd = m_Person.BirthDay: TxtBirthDay.Text = MString.Date_TryParseValidate(TxtBirthDay.Text, "Birthday", "", bIsOK, bd): If Not bIsOK Then Exit Function
    Dim ic As Long: ic = -1
    Dim ct As City: Set ct = MData.Cities_ObjectFromListCtrl(CmbCity, ic)
    Dim ii As Long:     ii = m_Person.Index
    Dim nm As String:   nm = TxtName.Text
    m_Person.New_ bd, ct, ii, nm
    UpdateData = True
Catch:
End Function

Private Sub TxtBirthDay_Validate(Cancel As Boolean)
    Dim bd As Date: bd = m_Person.BirthDay
    Dim bIsOK As Boolean: TxtBirthDay.Text = MString.Date_TryParseValidate(TxtBirthDay.Text, "Birthday", "", bIsOK, bd)
    If bIsOK Then m_Person.SetParams BirthDay:=bd
    Cancel = Not bIsOK
End Sub
