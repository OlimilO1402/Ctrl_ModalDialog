VERSION 5.00
Begin VB.Form FCity 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Edit City"
   ClientHeight    =   1695
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
   ScaleHeight     =   1695
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtPostalCode 
      Alignment       =   2  'Zentriert
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtName 
      Height          =   435
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LblPostalCode 
      AutoSize        =   -1  'True
      Caption         =   "PostalCode:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "FCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_City As City

Public Sub UpdateView(Obj)
    Set m_City = Obj
    With m_City
        TxtName.Text = .Name
        TxtPostalCode.Text = .PostalCode
    End With
End Sub

Public Function UpdateData(Obj) As Boolean
    m_City.NewC MNew.City(TxtName.Text, TxtPostalCode.Text)
    UpdateData = True
End Function
