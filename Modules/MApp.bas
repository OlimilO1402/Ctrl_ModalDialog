Attribute VB_Name = "MApp"
Option Explicit

Sub Main()
    MData.Init
    
    MData.Cities_Add MNew.City("Musterstadt", "00000")
    MData.Cities_Add MNew.City("New York", "07008")
    MData.Cities_Add MNew.City("Berlin", "10176")
    MData.Cities_Add MNew.City("Hamburg", "20144")
    MData.Cities_Add MNew.City("Paris", "70123")
    MData.Cities_Add MNew.City("München", "80336")
    
    FMain.Show
End Sub
