VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
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
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnPersonMoveDown 
      Caption         =   "Down [v]"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonMoveUp 
      Caption         =   "Up [^]"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonDel 
      Caption         =   "Del [ - ]"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Add a new City"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonEdit 
      Caption         =   "Edit [ / ]"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonAdd 
      Caption         =   "Add [ + ]"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Add a new City"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityMoveDown 
      Caption         =   "Down [v]"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityMoveUp 
      Caption         =   "Up [^]"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityDel 
      Caption         =   "Del [ - ]"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Add a new City"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityEdit 
      Caption         =   "Edit [ / ]"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox LstPersons 
      Height          =   3120
      Left            =   1440
      TabIndex        =   13
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton BtnCityAdd 
      Caption         =   "Add [ + ]"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Add a new City"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox LstCities 
      Height          =   2610
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Persons:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cities:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    UpdateView
End Sub

Sub UpdateView()
    MData.Cities_ToListCtrl LstCities
    MData.Persons_ToListCtrl LstPersons
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = LstCities.Left
    Dim t As Single: t = LstCities.Top
    Dim W As Single: W = Me.ScaleWidth - l
    Dim H As Single: H = LstCities.Height
    If W > 0 And H > 0 Then LstCities.Move l, t, W, H
    t = LstPersons.Top: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then LstPersons.Move l, t, W, H
End Sub

' v ' ############################## ' v '    Cities     ' v ' ############################## ' v '
Private Sub BtnCityAdd_Click()
    Dim City As City: Set City = MNew.City("Musterstadt", "00000")
    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
    MData.Cities_AddLB LstCities, City
End Sub

Private Sub BtnCityEdit_Click()
    'If LstCities.ListCount = 0 Then MsgBox "First add a city!": Exit Sub
    Dim i As Long: i = -1
    Dim City As City: Set City = MData.Cities_ObjectFromListCtrl(LstCities, i)
    If City Is Nothing Then MsgBox "Select a city first!": Exit Sub
    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
    LstCities.List(i) = City.ToStr
End Sub

Private Sub BtnCityDel_Click()
    'If LstCities.ListCount = 0 Then Exit Sub
    Dim i As Long: i = -1
    Dim City As City: Set City = MData.Cities_ObjectFromListCtrl(LstCities, i)
    If i = 0 Then MsgBox "Can not delete the first city!": Exit Sub
    If City Is Nothing Then MsgBox "First select an element to delete": Exit Sub
    Dim CitiesInUse As Collection: Set CitiesInUse = MData.Persons_UsingCity(City)
    Dim c As Long: c = CitiesInUse.Count
    If c Then
        MsgBox "Can not delete the city: " & City.ToStr & "; " & c & " person" & IIf(c > 1, "s are", " is") & " living there: " & vbCrLf & MPtr.Col_ToStr(CitiesInUse)
        Exit Sub
    End If
    If MsgBox("Are you sure to delete the city: " & vbCrLf & City.ToStr, vbOKCancel) = vbCancel Then Exit Sub
    MData.Cities_Remove City
    LstCities.RemoveItem i
End Sub

Private Sub BtnCityMoveUp_Click()
    MData.Cities_MoveUp LstCities
End Sub

Private Sub BtnCityMoveDown_Click()
    MData.Cities_MoveDown LstCities
End Sub

Private Sub LstCities_DblClick()
    BtnCityEdit.Value = True
End Sub

Private Sub LstCities_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyAdd:      Me.BtnCityAdd.Value = True
    Case KeyCodeConstants.vbKeySubtract: Me.BtnCityDel.Value = True
    Case KeyCodeConstants.vbKeyDivide:   Me.BtnCityEdit.Value = True
    Case KeyCodeConstants.vbKeyReturn:   Me.BtnCityEdit.Value = True
    Case 220:                            Me.BtnCityMoveUp.Value = True
    Case KeyCodeConstants.vbKeyV:        Me.BtnCityMoveDown.Value = True
    End Select
End Sub
' ^ ' ############################## ' ^ '    Cities     ' ^ ' ############################## ' ^ '

'the ui-code for city and person look pretty much identical
'so the question is could we make a class to do the same

' v ' ############################## ' v '    Persons    ' v ' ############################## ' v '
Private Sub BtnPersonAdd_Click()
    Dim i As Long: i = 0 'here we set i to 0 befcause for the undefined person we always want the first city
    Dim City As City: Set City = MData.Cities_ObjectFromListCtrl(LstCities, i)
    Dim Person As Person: Set Person = MNew.PersonDefault(City)
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(Person, Me) = vbCancel Then Exit Sub
    MData.Persons_AddLB LstPersons, Person
End Sub

Private Sub BtnPersonEdit_Click()
    'If LstPersons.ListCount = 0 Then Exit Sub
    Dim i As Long: i = -1
    If LstPersons.ListCount = 0 Then MsgBox "First add a person!": Exit Sub
    Dim Person As Person: Set Person = MData.Persons_ObjectFromListCtrl(LstPersons, i)
    If i < 0 Then
        MsgBox "Select an object first!"
        Exit Sub
    End If
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(Person, Me) = vbCancel Then Exit Sub
    LstPersons.List(i) = Person.ToStr
End Sub

Private Sub BtnPersonDel_Click()
    'If LstPersons.ListCount = 0 Then Exit Sub
    Dim i As Long: i = -1
    Dim Person As Person: Set Person = MData.Persons_ObjectFromListCtrl(LstPersons, i)
    If Person Is Nothing Then MsgBox "First select an element to delete": Exit Sub
    If MsgBox("Arey you sure to delete the person: " & vbCrLf & Person.ToStr, vbOKCancel) = vbCancel Then Exit Sub
    MData.Persons_Remove Person
    LstPersons.RemoveItem i
End Sub

Private Sub BtnPersonMoveUp_Click()
    'If LstPersons.ListCount = 0 Then Exit Sub
    MData.Persons_MoveUp LstPersons
End Sub

Private Sub BtnPersonMoveDown_Click()
    'If LstPersons.ListCount = 0 Then Exit Sub
    MData.Persons_MoveDown LstPersons
End Sub

Private Sub LstPersons_DblClick()
    'If LstPersons.ListCount = 0 Then Exit Sub
    BtnPersonEdit_Click
End Sub

Private Sub LstPersons_KeyDown(KeyCode As Integer, Shift As Integer)
    'If LstPersons.ListCount = 0 Then Exit Sub
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyAdd:      Me.BtnPersonAdd.Value = True
    Case KeyCodeConstants.vbKeySubtract: Me.BtnPersonDel.Value = True
    Case KeyCodeConstants.vbKeyDivide:   Me.BtnPersonEdit.Value = True
    Case KeyCodeConstants.vbKeyUp Or KeyCodeConstants.vbKeyDown
                                     If MString.IsShift(KeyCode, Shift, KeyCodeConstants.vbKeyUp) Then
                                         Me.BtnPersonMoveUp.Value = True
                                     ElseIf MString.IsShift(KeyCode, Shift, KeyCodeConstants.vbKeyDown) Then
                                         Me.BtnPersonMoveDown.Value = True
                                     End If
    End Select
End Sub

' ^ ' ############################## ' ^ '    Persons    ' ^ ' ############################## ' ^ '
