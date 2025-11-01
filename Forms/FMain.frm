VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
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
   ScaleHeight     =   5295
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnPersonMoveDown 
      Caption         =   "Down [v]"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonMoveUp 
      Caption         =   "Up [^]"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonDel 
      Caption         =   "Del [ - ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Add a new City"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonEdit 
      Caption         =   "Edit [ / ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnPersonAdd 
      Caption         =   "Add [ + ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Add a new City"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityMoveDown 
      Caption         =   "Down [v]"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityMoveUp 
      Caption         =   "Up [^]"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityDel 
      Caption         =   "Del [ - ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Add a new City"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BtnCityEdit 
      Caption         =   "Edit [ / ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox LstPersons 
      Height          =   2100
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton BtnCityAdd 
      Caption         =   "Add [ + ]"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Add a new City"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox LstCities 
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Persons:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cities:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
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

' v ' ############################## ' v '    Cities     ' v ' ############################## ' v '
Private Sub BtnCityAdd_Click()
    Dim City As City: Set City = MNew.City("Musterstadt", "00000")
    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
    MData.Cities_AddLB LstCities, City
End Sub

Private Sub BtnCityEdit_Click()
    Dim i As Long: i = -1
    Dim City As City: Set City = MData.Cities_ObjectFromListCtrl(LstCities, i)
    If City Is Nothing Then MsgBox "Select a city first!": Exit Sub
    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
    LstCities.List(i) = City.ToStr
End Sub

Private Sub LstCities_DblClick()
    BtnCityEdit_Click
End Sub

Private Sub BtnCityDel_Click()
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
' ^ ' ############################## ' ^ '    Cities     ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    Persons    ' v ' ############################## ' v '
Private Sub BtnPersonAdd_Click()
    Dim i As Long: i = 0 'here we set i to 1 befcause for the undefined person we always want the first city
    Dim City As City: Set City = MData.Cities_ObjectFromListCtrl(LstCities, i)
    Dim Person As Person: Set Person = MNew.Person(DateSerial(1980, 1, 1), City, 1, "Max Mustermann")
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(Person, Me) = vbCancel Then Exit Sub
    MData.Persons_AddLB LstPersons, Person
End Sub

Private Sub BtnPersonEdit_Click()
    Dim i As Long: i = -1
    Dim Person As Person: Set Person = MData.Persons_ObjectFromListCtrl(LstPersons, i)
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(Person, Me) = vbCancel Then Exit Sub
    LstPersons.List(i) = Person.ToStr
End Sub

Private Sub LstPersons_DblClick()
    BtnPersonEdit_Click
End Sub

Private Sub BtnPersonDel_Click()
    Dim i As Long: i = -1
    Dim Person As Person: Set Person = MData.Persons_ObjectFromListCtrl(LstPersons, i)
    If Person Is Nothing Then MsgBox "First select an element to delete": Exit Sub
    If MsgBox("Arey you sure to delete the person: " & vbCrLf & Person.ToStr, vbOKCancel) = vbCancel Then Exit Sub
    MData.Persons_Remove Person
    LstPersons.RemoveItem i
End Sub

Private Sub BtnPersonMoveUp_Click()
    MData.Persons_MoveUp LstPersons
End Sub

Private Sub BtnPersonMoveDown_Click()
    MData.Persons_MoveDown LstPersons
End Sub

' ^ ' ############################## ' ^ '    Persons    ' ^ ' ############################## ' ^ '
