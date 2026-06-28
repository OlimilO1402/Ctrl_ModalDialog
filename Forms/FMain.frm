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
   Begin VB.CommandButton BtnPersonView 
      Caption         =   "View r.o."
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "just view data read only"
      Top             =   5040
      Width           =   1215
   End
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
Private WithEvents mCrudCities  As DBcrud
Attribute mCrudCities.VB_VarHelpID = -1
Private WithEvents mCrudPersons As DBcrud
Attribute mCrudPersons.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set mCrudCities = MNew.DBcrud(MApp.Cities, True, Me.LstCities, Me.BtnCityAdd, , , , Me.BtnCityEdit, Me.BtnCityDel, Me.BtnCityMoveUp, Me.BtnCityMoveDown)
    Set mCrudPersons = MNew.DBcrud(MApp.Persons, True, Me.LstPersons, Me.BtnPersonAdd, , , , Me.BtnPersonEdit, Me.BtnPersonDel, Me.BtnPersonMoveUp, Me.BtnPersonMoveDown)
    UpdateView
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

Private Sub BtnPersonView_Click()
    Dim i As Long: i = -1
    If LstPersons.ListCount = 0 Then MsgBox "First add a person!": Exit Sub
    Dim Person As Person: Set Person = MApp.Persons.ObjectFromListCtrl(LstPersons, i)
    If i < 0 Then
        MsgBox "Select an object first!"
        Exit Sub
    End If
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel, True).ShowDialog(Person, Me) = vbCancel Then Exit Sub
    LstPersons.List(i) = Person.ToStr
End Sub

Sub UpdateView()
    MApp.Cities.ToListCtrl LstCities, , True
    MApp.Persons.ToListCtrl LstPersons, , True
End Sub

Private Sub mCrudCities_OnEdit(Obj_inout As Object)
    If Obj_inout Is Nothing Then Set Obj_inout = New City
    Dim obj As City: Set obj = Obj_inout
    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel, False).ShowDialog(obj, Me) = vbCancel Then Exit Sub
End Sub

Private Sub mCrudPersons_OnEdit(Obj_inout As Object)
    If Obj_inout Is Nothing Then
        Dim City As City: Set City = MApp.Cities.ObjectFromListCtrl(LstCities, 0)
        Set Obj_inout = MNew.PersonDefault(City)
    End If
    Dim obj As Person: Set obj = Obj_inout
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(obj, Me) = vbCancel Then Exit Sub
End Sub
