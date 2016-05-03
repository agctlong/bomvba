Attribute VB_Name = "BOMFeatures"

'-- Buttons / Behaviors
Sub QBRefresh_click()
    Call RefreshLinkedData
    'Call WhoToBlame("M3")
End Sub

Sub Master_ClearFilters_click()
    Call ActiveSheetClearFilters
End Sub



Sub QBBOM_ClearFilters_click()
    Call ActiveSheetClearFilters
End Sub

Public Sub MasterQBRefresh_click()
    Call RefreshLinkedData
    Call WhoToBlame("H8")
End Sub

Public Sub SalesOrderCompare_Click()
    Worksheets("Master").Columns("L").Hidden = False
    Worksheets("Master").Columns("M").Hidden = False
    Worksheets("Master").Columns("N").Hidden = False

    Call WhoToBlame("H6")
    
    Call SalesOrderCompare
    MsgBox "Action Completed!"
End Sub

Public Sub QBPartValidation_click()
    Call WhoToBlame("H4")
    Worksheets("Master").Columns("L").Hidden = False
    Worksheets("Master").Columns("M").Hidden = True
    Worksheets("Master").Columns("N").Hidden = True
    
    
    '-- Check QB For part numbers
    Call QBPartValidation
    MsgBox "Action Completed!"
    
End Sub

Public Sub GenerateMasterBOM_Click()
    '-- Date validation for generating the Master BOM
    Call ClearWhoToBlame("H4") '-- Clear validation blame
    Call ClearWhoToBlame("H6") '-- Clear compare blame
    Call WhoToBlame("H2")
    
    Worksheets("Master").Columns("L").Hidden = True
    Worksheets("Master").Columns("M").Hidden = True
    Worksheets("Master").Columns("N").Hidden = True
    
    Call CopyWorksheets
    MsgBox "Action Completed!"
    
End Sub



'-- Common Functions
Private Sub SalesOrderCompare()
    If SalesOrderConnectionChange Then
        Call RefreshLinkedData
        Call SalesOrderCompareFormatsFormulas
    End If
End Sub

Private Sub ActiveSheetClearFilters()
    With ActiveWorkbook.ActiveSheet
        If .AutoFilterMode Then
            If .FilterMode Then .ShowAllData
        End If
    End With
End Sub


Private Sub SalesOrderCompareFormatsFormulas()
    Dim strSheetMaster As String _
        , strSheetQB As String
    
    Dim intMasterMatchColumn As Integer _
        , intMasterMatchRow As Integer _
        , strMasterMatchFormula As String


    '-- What sheets are we dealing with here?
        strSheetMaster = "Master"
        strSheetQB = "QBBOM"
        
        '-- Master Cell References
            intMasterColumn = 13
            strMasterColumn = "M"
            intMasterRowStart = 13
            intMasterRowEnd = 400
        
        '-- QB Cell References
            intQBColumn = 12
            strQBColumn = "L"
            intQBRowStart = 10
            intQBRowEnd = 400
        
            
        '-- SOURCE FORMULA [=C14&F14&D14] '== =IF(F14&D14<>"",C14&"_"&F14&"_"&D14&"_"&J14,"")
        'strMasterMatchFormula = "=C" & intMasterRowStart & "&F" & intMasterRowStart & "&D" & intMasterRowStart  '"=C14&F14&D14"
        strMasterMatchFormula = "=IF(A" & intMasterRowStart & "<>"""",C" & intMasterRowStart & "&""_""&F" & intMasterRowStart & "&""_""&D" & intMasterRowStart & "&""_""&J" & intMasterRowStart & ","""")"
        
        '-- SOURCE FORMULA [=IFERROR(IF(NOT(ISBLANK(D14)),MATCH(M14,QBBOM!$M$11:$M$28,0),""),"Not Found")]
        'strMasterQBLineFormula = "=IFERROR(IF(NOT(ISBLANK(D" & intMasterRowStart & ")),MATCH(" & strMasterColumn & intMasterRowStart & "," & strSheetQB & "!$" & strQBColumn & "$" & intQBRowStart & ":$" & strQBColumn & "$" & intQBRowEnd & ",0),""""),""Not Found"")"
        strMasterQBLineFormula = "=IFERROR(IF(AND(M" & intMasterRowStart & "<>"""",K" & intMasterRowStart & "<>""Deleted Items""),MATCH(M" & intMasterRowStart & ",QBBOM!$L$10:$L$399,0),""""),""Not Found"")"
        
        
        '-- SOURCE FORMULA [=B11&D11&F11]
        'strQBMatchFormula = "=B" & intQBRowStart & "&D" & intQBRowStart & "&F" & intQBRowStart
        strQBMatchFormula = "=IF(D" & intQBRowStart & "&F" & intQBRowStart & "<>"""",B" & intQBRowStart & "&""_""&D" & intQBRowStart & "&""_""&F" & intQBRowStart & "&""_""&K" & intQBRowStart & ","""")"
        
        '-- SOURCE FORMULA [=IFERROR(IF(NOT(ISBLANK(F11)),MATCH(M11,Master!$M$14:$M$33,0),""),"Not Found")]
        'strQBMasterLineFormula = "=IFERROR(IF(NOT(ISBLANK(F" & intQBRowStart & ")),MATCH(" & strQBColumn & intQBRowStart & ",Master!$" & strMasterColumn & "$" & intMasterRowStart & ":$" & strMasterColumn & "$" & intQBRowEnd & ",0),""""),""Not Found"")"
        strQBMasterLineFormula = "=IFERROR(IF(L" & intQBRowStart & "<>"""",MATCH(L" & intQBRowStart & ",Master!$M$13:$M$400,0),""""),""Not Found"")"
        
        
        '== Master Changes ==
            Worksheets(strSheetMaster).Activate
            '-- Concatenate values for easy comparison
                With Range(Cells(intMasterRowStart, intMasterColumn), Cells(intMasterRowEnd, intMasterColumn))
                    .value = strMasterMatchFormula
                End With
            '-- Apply matching formula
                With Range(Cells(intMasterRowStart, intMasterColumn + 1), Cells(intMasterRowEnd, intMasterColumn + 1))
                    .value = strMasterQBLineFormula
                End With
            '-- Add conditional formatting to indicate missing lines?
                
        
        '== QB Changes ==
            Worksheets(strSheetQB).Activate
            '-- Concatenate values for easy comparison
            With Range(Cells(intQBRowStart, intQBColumn), Cells(intQBRowEnd, intQBColumn))
                .value = strQBMatchFormula
            End With
            '-- Apply matching formula
            With Range(Cells(intQBRowStart, intQBColumn + 1), Cells(intQBRowEnd, intQBColumn + 1))
                .value = strQBMasterLineFormula
            End With
            
        '== Jump back to Master
            Worksheets(strSheetMaster).Activate
            
End Sub


Private Function SalesOrderConnectionChange()
    
    '-- Declare Variables
        Dim strConnectionString As String, strIntranetPath As String, intSalesOrderID As Long
    
    '-- Get Values
        strIntranetPath = GetIntranetBOMURL
        If IsNumeric(Range("SalesOrderID").value) Then
            intSalesOrderID = Range("SalesOrderID").value
        Else
            MsgBox "Please set the SalesOrder ID from QB"
            Range("SalesOrderID").Select
            SalesOrderConnectionChange = False
            Exit Function
        End If
        
        strConnectionString = strIntranetPath & "?salesorderid=" & intSalesOrderID & "&view=excel"
        Range("DataSourceURL").value = strConnectionString
        strConnectionString = "URL;" & strConnectionString
        
    'If Not itnSalesOrderID Is Empty And intSalesOrderID > 0 Then
    If intSalesOrderID > 0 Then
    
    '-- Check the SalesID and Update the Datasource?
        Dim oSh As Worksheet
        Set oSh = Worksheets("QBBOM")
        Dim watchVal
        
        For x = 1 To oSh.QueryTables.Count
            With oSh.QueryTables(x)
                If Not .WorkbookConnection Is Nothing Then
                    If .QueryType = xlWebQuery Then
                        'MsgBox strConnectionString
                        .Connection = strConnectionString
                    End If
                End If
            End With
        Next
    Else
        Range("SalesOrderID").Select
        MsgBox "Please assign the SalesOrder ID on the Intranet site", vbCritical, "Error"
    End If
    
    SalesOrderConnectionChange = True

End Function


Private Sub RefreshLinkedData()
    Workbooks(ThisWorkbook.name).RefreshAll
End Sub



Public Sub CopyWorksheets()
    
    '-- Prevent messy looking screen flickering
        Application.ScreenUpdating = False
    
    '-- Update the index as that is our source of worksheet names
        Call RefreshIndex
        Dim arrExcludeSheets
        arrExcludeSheets = GetExcludes()
        
    '-- Initialize key variables
        Dim intNumberofSheets As Integer
        Dim mainworkBook As Workbook
        Dim CurrentRange As Range
        Dim LastRow As Long
        Dim LastCol As Long
        Dim x, y As Integer
        
        Dim thisName As String
        Set mainworkBook = ActiveWorkbook
    
        
    '-- Selects "Master" Worksheet and clears all of the data
        With mainworkBook.Sheets("Master")
            .Select
            If .AutoFilterMode Then
                If .FilterMode Then .ShowAllData
            End If
        End With
        Range("A13:L1000").Clear
        
        
    
    '-- Goes to the Index and gets the name of all of the sheets
        mainworkBook.Sheets("Index").Select
        
        Dim rngDrawingList As Range
        Set rngDrawingList = Range("A5", Range("A5").End(xlDown))
        intNumberofSheets = rngDrawingList.Rows.Count
        'MsgBox intNumberofSheets
        With rngDrawingList
            For x = 1 To intNumberofSheets
                '.Cells(x, 1).Select
                thisName = .Cells(x, 1).value
                boolCopySheet = True
                
                'If True = False Then
                    '-- Check for Excluded Names
                        For y = 0 To UBound(arrExcludeSheets)
                            If LCase(thisName) = LCase(arrExcludeSheets(y)) Then
                                boolCopySheet = False
                            End If
                        Next
                    
                    '-- Is sheet excluded
                    If boolCopySheet Then
                        '-- The first sheet to be copied is the third, and starts at A10 1st = "Revision Log" 2nd = "Master"
                        If x = 2 Then
                            CopyAndPaste thisName, True
                        '-- Any other sheet will be pasted at the end of the data currently on the "Master" Sheet
                        Else
                            CopyAndPaste thisName, False
                        End If
                    End If
                'End If
            Next
        End With
    
    '-- Force Master Page to scroll back up to the top
        mainworkBook.Sheets("Master").Range("A13").Select
    
    '-- Prevent messy looking screen flickering
        Application.ScreenUpdating = True
        
End Sub '-- CopyWorksheets


Public Sub CopyAndPaste(WorkSheetName As String, First As Boolean)

    Set mainworkBook = ActiveWorkbook

    '-- Skips function if the sheet is blank
    If WorkSheetName = "" Then
        'Do Nothing
    Else
        With mainworkBook.Sheets(WorkSheetName)
            .Select
            If .AutoFilterMode Then
                If .FilterMode Then .ShowAllData
            End If
        End With
        
        'Set shtMaster = mainworkBook.Sheet("Master")
        
        '-- Gets the last column and row of the current work from the "Index" list
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        End With
        
        '-- Checks to make sure there is information on the sheet, and gives a message box if there isn't, and sheet isn't "Deleted Items"
        If LastRow = 1 And WorkSheetName <> "Deleted Items" Then
            MsgBox "Sheet: " & WorkSheetName & " is missing item lines or revisions."
            mainworkBook.Sheets(WorkSheetName).Select
            Exit Sub
        End If
        
        '----------------------------Copy Function For Header-----------------------------------------------------------
        mainworkBook.Sheets(WorkSheetName).Range("A1:J1").Copy
        mainworkBook.Sheets("Master").Select
        
        '-- Uses the Boolean passed in to see if it's the "First" Worksheet, Pastes all copied data starting at Cell A13
        If First Then
            mainworkBook.Sheets("Master").Range("A13").Select
            mainworkBook.Sheets("Master").Paste
            
            '-- Uses the address of the paste option to get the first and last cell of the new data
            CurrentRange = Selection.Address
            CurrentRangeRows = Selection.Rows.Count
            CurrentRangeSplit = Split(Replace(CurrentRange, ":", ""), "$")
            
         Else
            '-- Gets the last column and row of the current work from the "Index" list
            With ActiveSheet
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            End With
            
            '-- If the current sheet wasn't first it finds the last of data on the master and pastes the new data two lines below
            mainworkBook.Sheets("Master").Range("A" & LastRow + 2).Select
            mainworkBook.Sheets("Master").Paste
            
            '-- Uses the address of the paste option to get the first and last cell of the new data
            CurrentRange = Selection.Address
            CurrentRangeRows = Selection.Rows.Count
            CurrentRangeSplit = Split(Replace(CurrentRange, ":", ""), "$")
            
        End If
        
        '----------------------------Copy Function For Data-------------------------------------------------------------
        If WorkSheetName = "Deleted Items" Then
            colLetter = "H"
        Else
            colLetter = "J"
        End If
        mainworkBook.Sheets(WorkSheetName).Range("A3:" & colLetter & LastRow + 1).Copy
        mainworkBook.Sheets("Master").Select
        
        '-- Uses the Boolean passed in to see if it's the "First" Worksheet, Pastes all copied data starting at Cell A14
        If First Then
            mainworkBook.Sheets("Master").Range("A14").Select
            mainworkBook.Sheets("Master").Paste
            
            '-- Uses the address of the paste option to get the first and last cell of the new data
            CurrentRange = Selection.Address
            CurrentRangeRows = Selection.Rows.Count
            CurrentRangeSplit = Split(Replace(CurrentRange, ":", ""), "$")
            
            '-- Uses the adress from the last function to add the drawing number to the master BOM
            For DrawingNumberLooper = 0 To CurrentRangeRows - 1
                thisPart = mainworkBook.Sheets("Master").Range("F" & CurrentRangeSplit(2) + DrawingNumberLooper).value
                thisMfg = mainworkBook.Sheets("Master").Range("E" & CurrentRangeSplit(2) + DrawingNumberLooper).value
                If thisPart <> "" Or thisMfg <> "" Then
                    mainworkBook.Sheets("Master").Range("K" & CurrentRangeSplit(2) + DrawingNumberLooper).value = WorkSheetName
                Else
                    'mainworkBook.Sheets("Master").Range("K" & CurrentRangeSplit(2) + DrawingNumberLooper).value = "Misfire?"
                End If
            Next

        Else
            '-- Gets the last column and row of the current work from the "Index" list
            With ActiveSheet
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            End With
            
            '-- If the current sheet wasn't first it finds the last of data on the master and pastes the new data two lines below
            mainworkBook.Sheets("Master").Range("A" & LastRow + 3).Select
            mainworkBook.Sheets("Master").Paste
            
            '-- Uses the address of the paste option to get the first and last cell of the new data
            CurrentRange = Selection.Address
            CurrentRangeRows = Selection.Rows.Count
            CurrentRangeSplit = Split(Replace(CurrentRange, ":", ""), "$")
            
            '-- Uses the adress from the last function to add the drawing number to the master BOM
            For DrawingNumberLooper = 0 To CurrentRangeRows - 1
                If mainworkBook.Sheets("Master").Range("F" & CurrentRangeSplit(2) + DrawingNumberLooper).value <> "" Then
                    mainworkBook.Sheets("Master").Range("K" & CurrentRangeSplit(2) + DrawingNumberLooper).value = WorkSheetName
                Else
                    'Do Nothing
                End If
            Next
            
        End If
        
    End If
    
    '-- Adjust qunatities for multiple parts
    If mainworkBook.Sheets(WorkSheetName).Range("F1").value = 1 Then
        'Do Nothing
    Else
        If WorkSheetName = "Deleted Items" Then
        'Do Nothing
        Else
            Call Multiplier(WorkSheetName)
        End If
        
    End If
    
End Sub '-- CopyAndPaste


Sub AddDrawing_click()
    
    Dim strDrawingNumber As String, strDescription As String, intMultiplier As Integer
    Dim wksLoop As Worksheet
    Dim wksTemplate As Worksheet
    
    '-- Prompt for Drawing Number
        strDrawingNumber = InputBox("Enter drawing number to be added", "Add Drawing")
        strDescription = InputBox("Enter Description of Drawing", "Add Drawing")
        intMultiplier = 1
        If Trim(strDrawingNumber) <> "" Then
            Call AddDrawing(strDrawingNumber, strDescription, intMultiplier)

            '-- Refresh Index
            Call RefreshIndex
        Else
            MsgBox "Drawing was not added"
        End If
        


End Sub '-- AddDrawing_click()


Sub AddDrawing(strDrawingNumber As String, strDescription As String, intMultiplier As Integer)
    
    
    'Dim strDrawingNumber As String
    Dim wksLoop As Worksheet
    Dim wksTemplate As Worksheet

    '-- Verify not already a tab
        pastMaster = False
        thisName = ""
        For Each wksLoop In Application.Worksheets
            thisName = wksLoop.name
            If LCase(thisName) = LCase(strDrawingNumber) Then
                MsgBox "Drawing already exists"
                Exit Sub
            End If
        Next
    
    '-- Locate Hidden Template Tab
        With Sheets("Template")
            .Visible = True
            .Copy After:=Sheets(Sheets.Count - 1)
            .Visible = False
        End With
        ActiveSheet.name = strDrawingNumber
        
        '-- Joe's additions
        MyFullName = ThisWorkbook.FullName
        ActiveSheet.Range("G1").value = "=HYPERLINK(""[" & MyFullName & "]'" & strDrawingNumber & "'!A2"",""" & strDescription & """)"
        
        '-- Travis Addition (4/18/2016)
        If intMultiplier > 1 Then
            ActiveSheet.Range("F1").value = intMultiplier
        End If

End Sub '-- AddDrawing()



Sub AddDrawingBatch_click()
    Dim strDrawingNumber As String, strDescription As String, intMultiplier As Integer
    
    Dim wksLoop As Worksheet
    Dim wksTemplate As Worksheet
    
    Dim rngSource As Range, rngRow As Range
    
    Dim strSourceWorkbook As String
    Dim i As Integer
    
    
    '-- Get Source Content
    strSourceWorkbook = "Admin"
    Set rngSource = ActiveWorkbook.Sheets(strSourceWorkbook).Range("I3:K40")
    
    i = 0
    For Each rngRow In rngSource.Rows
        strDrawingNumber = rngRow.Cells(1, 1)
        strDescription = rngRow.Cells(1, 2)
        intMultiplier = rngRow.Cells(1, 3)
        If strDrawingNumber <> "" Then
            Call AddDrawing(strDrawingNumber, strDescription, intMultiplier)
            i = i + 1
        Else
            Exit For
        End If
    Next
    MsgBox i & " drawings added!"

End Sub '-- AddDrawingBatch()


Sub RefreshIndex()
    
    Dim arrExcludeSheets
    arrExcludeSheets = GetExcludes()
    Set mainworkBook = ActiveWorkbook
    
    
    '-- Populate Index, link to it and grab description
        Dim thisName As String, intRowOffset As Integer, intRowCounter As Integer, boolAddtoIndex As Boolean
        
        '-- Get filename for hyperlink
        MyFullName = ThisWorkbook.FullName
        
        '-- Clear Index
        With Sheets("Index")
            .Unprotect
            .Range("A6:B44").Clear
            .Activate
        End With
        
        
        
        thisName = ""
        intRowOffset = 6 '-- Which row will we start populating data at
        intRowCounter = 0 '-- Track which row we are on
        
        
        For Each wksLoop In Application.Worksheets
            thisName = wksLoop.name
            boolAddtoIndex = True
            
            
            '-- Check for Excluded Names
            For x = 0 To UBound(arrExcludeSheets)
                If LCase(thisName) = LCase(arrExcludeSheets(x)) Then
                    boolAddtoIndex = False
                End If
            Next
                    
            If boolAddtoIndex Then
                Sheets("Index").Range("A" & intRowCounter + intRowOffset).value = "=HYPERLINK(""[" & MyFullName & "]'" & thisName & "'!A2"",""" & thisName & """)"
                Sheets("Index").Range("B" & intRowCounter + intRowOffset).value = "='" & thisName & "'!G1"
                intRowCounter = intRowCounter + 1
            End If
            
        Next
        
    mainworkBook.Sheets("Index").Select
    ActiveSheet.Protect "", True, True
    
End Sub '-- RefreshIndex()


Public Sub QBPartValidation()

    On Error Resume Next
    
    Dim LastRow As Long
    Dim LastCol As Long
    Dim NumberOfParts As Long
    Dim PartNumber As String
    Dim boolMultiplier As Boolean
    Dim boolDeletedItems As Boolean
    
    
    Dim PartNumberValidated As String
    Dim strXMLURL As String
    
    
    Dim xmlObject As MSXML2.XMLHTTP60
    Dim xmlDoc As MSXML2.DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument
    Set xmlObject = New MSXML2.XMLHTTP60
    Set mainworkBook = ActiveWorkbook

    xmlDoc.async = False
    
    Dim shtMaster As Worksheet
    
    
    Set shtMaster = ActiveWorkbook.Sheets("Master")
    

    
    '-- Selects the "Master" Worksheet
    'mainworkBook.Sheets("Master").Select
    shtMaster.Select
    
    
    '-- Date validation for Part Validation
    Call WhoToBlame("H4")
    
    '-- Counts all of the rows that Have part numbers
    With shtMaster
        LastRow = .Cells(.Rows.Count, "F").End(xlUp).Row

    End With
    
    '-- Loops through a for loop that skips over the general information at the top of the "Master" sheet
    For NumberOfParts = 1 To LastRow - 12
    
        '-- Using the current part number it pings the intranet
        With shtMaster
            boolMultiplier = (.Range("E" & NumberOfParts + 12).value = "Multiplier:")
            boolDeletedItems = (.Range("G" & NumberOfParts + 12).value = "Deleted Items")
            PartNumber = LTrim(RTrim(.Range("F" & NumberOfParts + 12).value))
            PartNumber = XMLRequestVariableEncode(PartNumber)
            
            strXMLURL = "http://intranet.americangovernor.com/lists/itemsXML.asp?q=" & PartNumber
            
            If boolDeletedItems Then
                '-- Stop loop
                Exit Sub
            End If
            If Not boolMultiplier Then
            
                '-- Get the XML response built into the intranet
                With xmlObject
                    Call .Open("GET", strXMLURL, False)
                    Call .send
                End With
                
                '-- Talk to Travis
                xmlDoc.LoadXML (xmlObject.responseXML.XML)
                
                
                '-- Clear all target cells first
                With shtMaster.Range("L" & NumberOfParts + 12)
                    .Clear
                    .Select
                End With
                   
                '-- Remove Hyperlinks
                ActiveCell.Hyperlinks.Delete
                
                
                '-- Skips over any blank cells
                If PartNumber <> "" Then
                    
                    intMatches = xmlDoc.SelectSingleNode("/items").Attributes.getNamedItem("results").Text
                    
                    If intMatches = "" Then intMatches = 0
                    
                    
                    '-- Checks to see if it is a part in quickbooks
                    If xmlDoc.Text = "Item not found " & PartNumber Then
                        PartNumberValidated = "Part Not in QuickBooks"
                        '-- Changes the cell's font and fill color to red and writes not in QB
                        With shtMaster.Range("L" & NumberOfParts + 12)
                            .value = PartNumberValidated
                            .Interior.Color = RGB(255, 199, 206)
                            .Font.Color = RGB(156, 0, 6)
                        End With
                        
                    '-- Checks to see if there is more than one part in quickbooks
                    Else
                        
                        If intMatches > 1 Then
                            '-- Changes the cell's font and fill to yellow and writes multiple parts in QB
                            PartNumberValidated = "Multiple Parts in QuickBooks (" & intMatches & ")"
                            With shtMaster.Range("L" & NumberOfParts + 12)
                                .value = PartNumberValidated
                                .Interior.Color = RGB(255, 235, 156)
                                .Font.Color = RGB(156, 101, 0)
                            End With
                        Else
                            '-- If there is only one part with the same number in qb
                            PartNumberValidated = xmlDoc.SelectSingleNode("/items/item/fullname").Text
                            thisColor = RGB(198, 239, 206)
                            If Err.Number > 0 Then
                                PartNumberValidated = "#Error#"
                                'PartNumberValidated = strXMLURL
                                thisColor = RGB(255, 199, 206)
                                Err.Clear
                            End If
                            
                            
                            '-- Changes the cell's font and fill to green and puts the qb part number in the cell
                            With shtMaster.Range("L" & NumberOfParts + 12)
                                .value = PartNumberValidated
                                .Interior.Color = thisColor
                                .Font.Color = RGB(0, 97, 0)
                            End With
                            
                        End If
                        
                    End If
                    
                    '-- Add Hyperlinks
                    ActiveCell.Hyperlinks.Add ActiveCell, "http://intranet.americangovernor.com/items.asp?searchitem=" & PartNumber, , , PartNumberValidated
                Else
                    ActiveCell.value = ""
                End If '-- boolMultiplier = False
            End If
        End With
        
        DoEvents
    Next
    
End Sub '-- PartValidation()



Function XMLRequestVariableEncode(value)
    value = Replace(value, "%", "%25")
    value = Replace(value, "&", "%26")
    value = Replace(value, "+", "%2B")
    XMLRequestVariableEncode = value
End Function '-- XMLRequestVariableEncode


Public Sub AGCPart()
    Dim strXMLURL As String
    Dim xmlObject As MSXML2.XMLHTTP60
    Dim xmlDoc As MSXML2.DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument
    Set xmlObject = New MSXML2.XMLHTTP60
    whatCase = 1
    
    Dim name As String
    name = "0825125"
    
    strXMLURL = "http://intranet.americangovernor.com/lists/itemsXML.asp?q=" & name
    
    With xmlObject
        Call .Open("GET", strXMLURL, False)
        Call .send
    End With
    
    MsgBox xmlObject.responseXML.XML
    
End Sub

Public Sub WhoToBlame(TargetCell)
    Range(TargetCell).value = Now
    Range(TargetCell).Offset(1, 0).value = Application.UserName
End Sub

Public Sub ClearWhoToBlame(TargetCell)
    Range(TargetCell).value = ""
    Range(TargetCell).Offset(1, 0).value = ""
End Sub

Public Sub SelectionTest()
    MsgBox Selection.Address
End Sub


Public Sub Multiplier(WKSName)
    '-- Setting variables and workbook
    Set mainworkBook = ActiveWorkbook
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim Multiplier, MultRows, x As Integer
    
    'Defining patterns for regex to extract numeric value
    With regex
        .Pattern = "\d*[0-9\.]+|\d*[^0-9\.]+"
        .Global = True
    End With
    
    '-- Reset counter for each group
    MultRows = 1
    Multiplier = mainworkBook.Sheets(WKSName).Range("F1").value
    
    '-- Getting current Location
    CurrentRange = Selection.Address
    
    '-- Counts how many rows in the current Selection
    CurrentRangeRows = Selection.Rows.Count
    
    '-- Breaks the current selction to useable values, and stores it
    CurrentRangeSplit = Split(Replace(CurrentRange, ":", ""), "$")
    CurrentCell = CurrentRangeSplit(2)
    
    '-- For loop to change values of each quantity
    For MultRows = 1 To CurrentRangeRows
        CurrentQTY = mainworkBook.Sheets("Master").Range("D" & CurrentCell).value
        
        'Skipping all cells that are blank
        If CurrentQTY = "" Then
        
        'Do Nothing
        
        Else
            
            'If the QTY is a just a number it is multipled
            If IsNumeric(CurrentQTY) Then
                NewQTY = CurrentQTY * Multiplier
                mainworkBook.Sheets("Master").Range("D" & CurrentCell).value = NewQTY
                
            'If the QTY is a number with a unit it splits the value, multiplies it, and concatenates the number and unit
            Else
                Set CurrentQTYSplit = regex.Execute(CurrentQTY)
                NewQTY = CurrentQTYSplit(0) * Multiplier
                mainworkBook.Sheets("Master").Range("D" & CurrentCell).value = NewQTY & CurrentQTYSplit(1)

            End If
            
        End If
        
        'Iterates to the next cell
        CurrentCell = CurrentCell + 1
        
    Next
    

End Sub







