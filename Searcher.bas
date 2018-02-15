Attribute VB_Name = "Searcher"
'**
'* begins the search
'**
Public Sub BeginSearch(ByVal PrimaryWorkbookName As String, ByRef PrimaryIDList As Variant, _
                        ByRef SecondaryIDList As Variant, ByRef secondaryNamesList As Variant)
                       
    Dim tempTable: Set tempTable = CreateObject("Scripting.Dictionary")
    Dim primaryTable: Set primaryTable = CreateObject("Scripting.Dictionary")
    Dim secondaryCountTable: Set secondaryCountTable = CreateObject("Scripting.Dictionary")
    Dim primaryCounttable: Set primaryCounttable = CreateObject("Scripting.Dictionary")
    Dim r As Long, c As Long, identification As String, name As String, i As Long
      
    '* put the values from the secondary info into a table *'
    For r = (LBound(SecondaryIDList, 1) + 1) To UBound(SecondaryIDList, 1)
        identification = SecondaryIDList(r, 1)
        name = secondaryNamesList(r, 1)
 
        If Not tempTable.Exists(identification) Then
            Dim secondTempNamesList() As String: ReDim secondTempNamesList(0)
            secondTempNamesList(0) = name
            tempTable.Add identification, secondTempNamesList
            secondaryCountTable.Add identification, 1
        Else
            If Exists(tempTable.Item(identification), name) = False Then
                tempTable.Item(identification) = AddToList(tempTable.Item(identification), name)
            End If
            secondaryCountTable.Item(identification) = secondaryCountTable.Item(identification)
        End If
    Next r
 
    '* add only the values that exist in the primary book to the primarytable *'
    Dim firstRan As Boolean: firstRan = False
    For r = (LBound(PrimaryIDList, 1) + 1) To UBound(PrimaryIDList, 1)
        identification = PrimaryIDList(r, 1)
 
        If tempTable.Exists(identification) Then
            primaryTable.Add identification, tempTable.Item(identification)
            primaryCounttable.Add identification, secondaryCountTable(identification)
        Else
            If Not primaryTable.Exists(identification) Then
                Dim primeList() As String: ReDim primeList(0)
                primeList(0) = "N/A"
                primaryTable.Add identification, primeList
                primaryCounttable.Add identification, 1
            Else
                primaryCounttable.Item(identification) = primaryCounttable.Item(identification) + 1
            End If
        End If
    Next r
 
    Call WriteOutput(PrimaryWorkbookName, PrimaryIDList, primaryTable, primaryCounttable, secondaryNamesList(1, 1))
End Sub
 
'**
'* writes the output
'**
Private Sub WriteOutput(ByVal PrimaryWorkbookName As String, ByRef PrimaryIDList As Variant, _
                        ByRef primaryTable, ByRef primaryCounttable, ByVal NameOfNamesCol As String)
    Dim primaryID_Header As String: primaryID_Header = PrimaryIDList(1, 1)
    Dim strKey As Variant
    Dim r As Long, c As Long
   
    Application.Workbooks(PrimaryWorkbookName).Activate
    If Not Evaluate("ISREF('" & "Output" & "'!A1)") Then
        Sheets.Add
        ActiveSheet.name = "Output"
    Else
        Sheets("Output").Cells.Clear
    End If
   
    Sheets("Output").Select
   
    With Sheets("Output")
        r = 1: c = 1
        Cells(r, c).value = primaryID_Header
        Cells(r, c + 1).value = "Total Number"
        Cells(r, c + 2).value = NameOfNamesCol
        r = r + 1
       
        For Each strKey In primaryTable
            Cells(r, c).value = strKey
            Cells(r, c + 1).value = primaryCounttable.Item(strKey)
           
            Dim ranOnce As Boolean: ranOnce = False
            Dim element As Variant
            Dim listOfNames() As String: listOfNames = primaryTable.Item(strKey)
            For Each element In listOfNames
                If ranOnce Then
                    Cells(r, c + 2).value = Cells(r, c + 2).value & ", " & element
                Else
                    Cells(r, c + 2).value = element
                    ranOnce = True
                End If
            Next element
            r = r + 1
        Next strKey
    End With
End Sub
 
'**
'* adds the name to the list of names if it didn't already exist
'**
Private Function AddToList(ByRef listOfNames, ByVal name As String) As Variant
    ReDim Preserve listOfNames(0 To (UBound(listOfNames) + 1))
    listOfNames(UBound(listOfNames)) = name
    AddToList = listOfNames
End Function
 
'**
'* check if the name exists in the list of names
'**
Private Function Exists(ByRef listOfNames, ByVal name As String) As Boolean
    Dim element As Variant
    Dim nameExists As Boolean: nameExists = False
    For Each element In listOfNames
        If element = name Then
            nameExists = True
            Exit For
        End If
    Next element
   
    Exists = nameExists
End Function
 
'**
'* macro to run the script
'**
Public Sub RunVLookupReplacer()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
   
    LookUpReplacer2_Form.Show
   
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

