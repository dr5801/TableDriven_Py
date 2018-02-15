VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LookUpReplacer1_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LookUpReplacer1_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LookUpReplacer1_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrimaryWorkbook As Workbook
Dim PrimaryBookPrimaryID_Range As Range
Dim primaryID_List() As Variant
Dim PrimaryWorkbookName As String
 
'* initialize the userform *'
Private Sub UserForm_Initialize()
    Dim wb As Workbook
   
    '* get th name of all the workbooks in the combobox '*
    For Each wb In Application.Workbooks
        Me.PrimaryBook_ComboBox.AddItem wb.name
    Next wb
   
    Me.PrimaryBook_ComboBox = ActiveWorkbook.name
End Sub
 
'* change the active workbook to the one the user selected and isn't already active *'
Private Sub PrimaryBook_ComboBox_Change()
    Application.ScreenUpdating = True
    Dim wb As Workbook
    If Me.PrimaryBook_ComboBox <> "" Then
        Set wb = Workbooks(Me.PrimaryBook_ComboBox.Text)
        wb.Activate
    End If
        
    Application.ScreenUpdating = False
End Sub
 
'**
'* handle the process for when the continue button is clicked
'**
Private Sub ContinueButton_Click()
    Set PrimaryWorkbook = Application.Workbooks(Me.PrimaryBook_ComboBox.Text)
    Set PrimaryBookPrimaryID_Range = Range(Me.PrimaryIDRange_RefEdit.Text)
    primaryID_List = getPrimaryIDValues(PrimaryBookPrimaryID_Range)
    PrimaryWorkbookName = Me.PrimaryBook_ComboBox.Text
   
    Dim savedDataRanges() As Variant
    Dim rangesString() As String
    If InStr(Me.SavedDataRange_RefEdit.Text, ",") > 0 Then
        rangesString = Split(Me.SavedDataRange_RefEdit.Text, ",")
       
        Dim totalRows As Long: totalRows = Cells(Rows.Count, PrimaryBookPrimaryID_Range.Column).End(xlUp).Row
        ReDim savedDataRanges(0 To totalRows, 0 To UBound(rangesString, 1))
        Dim element As Variant
        Dim r As Long, c As Long
        c = 0
        For Each element In rangesString
            r = 0
            Dim myRange As Range: Set myRange = Range(element)
            Set myRange = myRange(1, myRange.Column)
            Dim arr() As Variant
            Dim columnLetter As String: columnLetter = Chr(myRange.Column + 64)
            arr = Range(columnLetter & "1:" & columnLetter & totalRows).Value2
           
            Dim value As Variant
            For Each value In arr
                savedDataRanges(r, c) = value
                r = r + 1
            Next value
            c = c + 1
        Next element
    Else
        Dim theRange As Range: Set theRange = Range(Me.SavedDataRange_RefEdit.Text)
        Dim endingColumnLetter As String: endingColumnLetter = Chr(theRange.Columns.Count + 64)
        Set theRange = theRange(1, theRange.Column)
        Dim totRows As Long: totRows = Cells(Rows.Count, theRange.Column).End(xlUp).Row
        Dim beginningColumnLetter As String: beginningColumnLetter = Chr(theRange.Column + 64)
        savedDataRanges = Range(beginningColumnLetter & "1:" & endingColumnLetter & totRows)
    End If
   
'    Dim myMultiDimensionalRange As Variant
'    myMultiDimensionalRange = Range(Me.SavedDataRange_RefEdit.Text).Value2
'    MsgBox myMultiDimensionalRange(1, 2)
'
'    MsgBox savedDataRanges(0, 1, 0)
'
'    Dim element As Variant
'    For Each element In rangesString
'        MsgBox element
'    Next element
'    Dim myMultiDimensionalRange As Variant
'    myMultiDimensionalRange = Range(Me.SavedDataRange_RefEdit.Text).Value2
'
   
    LookUpReplacer2_Form.Show
   
    Unload Me
End Sub
 
'**
'* return the list of values within the range
'**
Private Function getPrimaryIDValues(ByRef PrimaryBookPrimaryID_Range As Range) As Variant
    Dim columnLetter As String: columnLetter = Chr(PrimaryBookPrimaryID_Range.Column + 64)
    Dim totalRows As Long: totalRows = Cells(Rows.Count, PrimaryBookPrimaryID_Range.Column).End(xlUp).Row
    getPrimaryIDValues = Range(columnLetter & "1:" & columnLetter & totalRows).Value2
End Function
 
'**
'* returns the primary wokbook
'**
Public Function getPrimaryWorkbook() As Workbook
    Set getPrimaryWorkbook = PrimaryWorkbook
End Function
 
'**
'* returns the name of the primary workbook
'**
Public Function getPrimaryWorkbookName() As String
    getPrimaryWorkbookName = PrimaryWorkbookName
End Function
 
'**
'* returns the primary values
'**
Public Function getPrimaryValues() As Variant
    getPrimaryValues = primaryID_List
End Function
