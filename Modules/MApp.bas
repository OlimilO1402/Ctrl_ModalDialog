Attribute VB_Name = "MApp"
Option Explicit
Public Persons As CCollection
Public Cities  As CCollection

Sub Main()
    Set Persons = MNew.CCollection(True)
    Set Cities = MNew.CCollection(True)
    
    Cities.Add MNew.City("Musterstadt", "00000")
    Cities.Add MNew.City("New York", "07008")
    Cities.Add MNew.City("Berlin", "10176")
    Cities.Add MNew.City("Hamburg", "20144")
    Cities.Add MNew.City("Paris", "70123")
    Cities.Add MNew.City("München", "80336")
    
    FMain.Show
End Sub
