VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LookUpReplacer2_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LookUpReplacer2_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LookUpReplacer2_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SecondaryWB As Workbook
Dim SecondaryPrimeID_Range As Range
Dim SecondaryNames_Range As Range
Dim SecondaryIDValuesList() As Variant
Dim SecondaryNamesValuesList() As Variant
 
'**
'* initialize the form
'**
Private Sub UserForm_Initialize()
    Dim wb As Workbook
   
    For Each wb In Application.Workbooks
        Me.SecondaryBook_ComboBox.AddItem wb.name
    Next wb
   
    Me.SecondaryBook_ComboBox = ActiveWorkbook.name
End Sub
 
'**
'* change the activesheet if the combo box value changes
'**
Private Sub SecondaryBook_ComboBox_Change()
    Application.ScreenUpdating = True
    If Me.SecondaryBook_ComboBox <> "" Then Application.Workbooks(Me.SecondaryBook_ComboBox.Text).Activate
    Application.ScreenUpdating = False
End Sub
 
'**
'* handles when the complete button is clicked
'**
Private Sub CompleteButton_Click()
    Set SecondaryWB = Application.Workbooks(Me.SecondaryBook_ComboBox.Text)
    Set SecondaryPrimeID_Range = Range(Me.SecondaryIDRange_RefEdit.Text)
    Set SecondaryNames_Range = Range(Me.SecondaryNamesRange_RefEdit.Text)
   
    SecondaryIDValuesList = getSecondaryIDValues(SecondaryPrimeID_Range)
    SecondaryNamesValuesList = getSecondaryNamesValues(SecondaryNames_Range)
   
    Dim PrimaryIDList As Variant: PrimaryIDList = LookUpReplacer1_Form.getPrimaryValues
    Dim PrimaryWorkbookName As String: PrimaryWorkbookName = LookUpReplacer1_Form.getPrimaryWorkbookName
 
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    Call Searcher.BeginSearch(PrimaryWorkbookName, PrimaryIDList, SecondaryIDValuesList, SecondaryNamesValuesList)
    SecondsElapsed = Round(Timer - StartTime, 2)
   
    MsgBox SecondsElapsed & " seconds"
    Unload Me
End Sub
 
'**
'* gets the values from the names/items/products range
'**
Private Function getSecondaryNamesValues(ByRef SecondaryNames_Range As Range) As Variant
    Dim totalRows As Long: totalRows = Cells(Rows.Count, SecondaryNames_Range.Column).End(xlUp).Row
    getSecondaryNamesValues = Range(Chr(SecondaryNames_Range.Column + 64) & "1:" & Chr(SecondaryNames_Range.Column + 64) & totalRows).Value2
End Function
'**
'* get the values from the secondary prime id range
'**
Private Function getSecondaryIDValues(ByRef SecondaryPrimeID_Range As Range) As Variant
    Dim columnLetter As String: columnLetter = Chr(SecondaryPrimeID_Range.Column + 64)
    Dim totalRows As Long: totalRows = Cells(Rows.Count, SecondaryPrimeID_Range.Column).End(xlUp).Row
    getSecondaryIDValues = Range(columnLetter & "1:" & columnLetter & totalRows).Value2
End Function
