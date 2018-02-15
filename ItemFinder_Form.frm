VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ItemFinder_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3345
   OleObjectBlob   =   "ItemFinder_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ItemFinder_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**
'* handles logic of when button is clicked
'**
Private Sub Calculate_Button_Click()
    
    Dim itemID_Col As Range: Set itemID_Col = Range(ItemIDCol_RefEdit.Text)
    Dim itemProducts_Col As Range: Set itemProducts_Col = Range(ItemProducts_RefEdit.Text)
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    'Remember time when macro starts
    StartTime = Timer

    Call Searcher1.BeginSearch(itemID_Col, itemProducts_Col)
    
    SecondsElapsed = Round(Timer - StartTime, 0)
    
    MsgBox SecondsElapsed & " seconds"
    Unload Me
End Sub
