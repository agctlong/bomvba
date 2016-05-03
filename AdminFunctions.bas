Attribute VB_Name = "AdminFunctions"


Public Sub SaveBackupFile()
    'On Error GoTo SaveBackupFile_ErrHandler
    
        Dim strBackupPath As String
        strBackupPath = ActiveWorkbook.Path & "\Backups\"
        
        strFileName = ActiveWorkbook.name
        
        strFileDate = Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2)
        strFileTime = Right("00" & Hour(Now), 2) & Right("00" & Minute(Now), 2)
        strFileUser = Application.UserName
        
        
        strFileNameBackup = Replace(strFileName, ".xlsm", "") & "_" & strFileDate & "_" & strFileTime & "_" & strFileUser & ".xlsm"
        
        'MsgBox strFileNameBackup
        answer = MsgBox("Would you like to backup this BOM?" & vbNewLine & vbNewLine & strBackupPath, vbYesNo + vbQuestion, "Backup BOM")
         
        If answer = vbYes Then
            'ActiveProject.SaveAs strFileNameBackup
            'MsgBox strFileNameBackup & vbCrLf & strBackupPath
            copyFile2 strFileName, ActiveWorkbook.FullName, strFileNameBackup, strBackupPath
        End If
    
SaveBackupFile_Exit:
    Exit Sub
    
SaveBackupFile_ErrHandler:
    MsgBox "There was an error with the backup! Operation failed."
    Err.Clear
    
End Sub

Sub copyFile2(ByVal strSourceName As String, ByVal strSourceFile As String, ByVal strNewName, ByVal strNewPath As String)
    Dim FSO
     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'If FSO.FileExists(strNewPath & strNewName) Then
        FSO.copyFile strSourceFile, strNewPath
        Set myFile = FSO.GetFile(strNewPath & strSourceName)
        FSO.MoveFile strNewPath & strSourceName, strNewPath & strNewName
    'End If

End Sub



Sub AddDataValidation()
    
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            With Ws.Range("A3", "A300").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="='Revision Log'!$A$8:$A$35"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Revision Not Found"
                .InputMessage = ""
                .ErrorMessage = "Revision must be created on Revision Log worksheet"
                .ShowInput = True
                .ShowError = True
            End With
        End If
    Next Ws
    
End Sub '-- AddDataValidation



Sub RemoveDataValidation()
    
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            Ws.Cells.Validation.Delete
        End If
    Next Ws
End Sub '-- RemoveDataValidation





Sub AddProcureFormulaFormat()
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            If thisName = "Master" Then
                With Ws.Range("J13", "J350")
                    .Formula = "=IF(AND(NOT(ISBLANK(I3)),NOT(ISerror(MATCH(I3,Index!G:G,0)))),INDEX(Index!I:I,MATCH(I3,Index!G:G,0)),"""")"
                    .NumberFormat = "m/d/yyyy"
                    .HorizontalAlignment = xlRight
                End With
            Else
                With Ws.Range("J3", "J300")
                    .Formula = "=IF(AND(NOT(ISBLANK(I3)),NOT(ISerror(MATCH(I3,Index!G:G,0)))),INDEX(Index!I:I,MATCH(I3,Index!G:G,0)),"""")"
                    .NumberFormat = "m/d/yyyy"
                    .HorizontalAlignment = xlRight
                End With
            End If
        End If
    Next Ws
End Sub '-- AddProcureFormulaFormat



Sub RemoveProcureFormulaFormat()
    
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            With Ws.Range("J3", "J300")
                .Clear
            End With
        End If
    Next Ws
    
End Sub '-- RemoveProcureFormulaFormat




Sub AddRevisionConditionalFormats()
    
    Dim rngTargetRange As Range
        
        
    '-- Master conditional formats (not required as they are copied from individual sheets)
    Set rngTargetRange = Worksheets("Master").Range("A13", "A400")
    ApplyWorksheetConditional rngTargetRange
    

    '--Get formats, pass it a range
    'MsgBoxCellFormat Range("A6")
    'Exit Sub


    '-- Worksheet exclude list
    arrExcludeList = GetValidationExcludes()
    Dim Ws As Worksheet
    
    
   '-- Loop through drawings
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = False
        For x = 0 To UBound(arrExcludeList)
            
            '-- We want to skip master for this one!
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If thisName = "Sample" Then doWork = True
        
        If doWork Then
            Set rngTargetRange = Ws.Range("A3", "A300")
            ApplyWorksheetConditional rngTargetRange
        End If
    Next Ws
    
End Sub '-- AddRevisionConditionalFormats

Sub MsgBoxCellFormat(rngTarget As Range)
    strOutput = "Nothing"
    '-- Capture Background Color and Font Color
    With rngTarget
        strOutput = "InteriorColor:" & .Interior.ColorIndex & vbCrLf
        strOutput = strOutput & "FontColor:" & .Font.ColorIndex
    End With
    MsgBox strOutput
End Sub



Sub rangeClearConditional(rngTarget As Range)
     
     With rngTarget
        .FormatConditions.Delete
    End With

End Sub

Sub ApplyWorksheetConditional(rngTarget As Range)
     
     With rngTarget
        .FormatConditions.Delete
        
        '-- Add equal to nothing
        .FormatConditions.Add xlCellValue, xlEqual, "="""""
        With .FormatConditions(1)
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = 1
            .StopIfTrue = True
        End With
        
        
        '-- Equal to Rev on Master form
        .FormatConditions.Add xlCellValue, xlEqual, "=Master!$C$10"
        With .FormatConditions(2)
            .Interior.ColorIndex = 35
            .Font.ColorIndex = 10
            .StopIfTrue = True
        End With
        
        '-- Less than Rev on Master form
        .FormatConditions.Add xlCellValue, xlLess, "=Master!$C$10"
        With .FormatConditions(3)
            .Interior.ColorIndex = 36
            .Font.ColorIndex = 53
            .StopIfTrue = True
        End With
        
        '-- Greater than Rev on Master form
        .FormatConditions.Add xlCellValue, xlGreater, "=Master!$C$10"
        With .FormatConditions(4)
            .Interior.ColorIndex = 38
            .Font.ColorIndex = 9
            .StopIfTrue = True
        End With
    End With

End Sub


Sub RemoveRevisionConditionalFormats()
    
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            With Ws.Range("A3", "A300")
                .FormatConditions.Delete
            End With
        End If
    Next Ws
    
End Sub '-- AddRevisionConditionalFormats


Sub AddDrawingFiltering()
    
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            With Ws
                .Activate
                LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                '.Range("A2:J" & LastRow + 1).Select
                .Range(.Cells(2, 1), .Cells(LastRow, 10)).Select
                .Range(.Cells(2, 1), .Cells(LastRow, 10)).AutoFilter
                .Range("A1").Select
            End With
            'MsgBox thisName & " - Row = " & LastRow & "; Col = " & LastCol
            'Exit Sub
        End If
    Next Ws
    
    '-- back to Admin page
    Worksheets("Admin").Activate
    

End Sub


Sub RemoteDrawingFiltering()
    '-- Exclude list
    arrExcludeList = GetValidationExcludes()
    
    '-- There were linking issues
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        thisName = Ws.name
        
        doWork = True
        For x = 0 To UBound(arrExcludeList)
            If thisName = arrExcludeList(x) Then
                doWork = False
                Exit For
            End If
        Next
        
        If doWork Then
            With Ws
                .AutoFilterMode = False
            End With
        End If
    Next Ws
End Sub



Sub ResetToTemplate()

    If MsgBox("This will delete all data from the document!", vbYesNoCancel, "Warning!") = vbYes Then
        If MsgBox("Make no mistake, this will delete everything from this BOM!!", vbYesNoCancel, "Warning!") = vbYes Then
            'MsgBox "Crazy Talk, feature not ready yet!"
            
            Call SaveBackupFile
            
            arrIgnoreSheets = GetDefaultSheetnames
            
            '-- Remove Drawing Tabs
            Application.DisplayAlerts = False
            Dim Ws As Worksheet
            
            For Each Ws In Worksheets
                thisName = Ws.name
                
                doWork = True
                For x = 0 To UBound(arrIgnoreSheets)
                    If thisName = arrIgnoreSheets(x) Then
                        doWork = False
                        Exit For
                    End If
                Next
                
                If doWork Then
                    Ws.Delete
                End If
            Next Ws
            Application.DisplayAlerts = True
            
            
            
            '-- QBBOM Cleanup
            Set Ws = ActiveWorkbook.Worksheets("QBBOM")
            With Ws
                .Range("C1:C7").value = ""
                .Range("A11:L400").value = ""
                .Range("M3:M4").value = ""
            End With
            
            
            '-- Index Logs Cleanup
            Set Ws = ActiveWorkbook.Worksheets("Index")
            With Ws
                .Unprotect
                .Range("H4:H100").value = ""
            End With
            
            
            '-- Revision Logs: Cleanup
            Set Ws = ActiveWorkbook.Worksheets("Revision Log")
            With Ws
                .Range("A9:D36").value = ""
                .Range("G9:J36").value = ""
                .Range("EngineerEmail").value = "Enter engineer email..."
                .Range("EngineerEmail").Hyperlinks.Delete
                .Range("AdminEmail").value = "Enter admin email..."
                .Range("AdminEmail").Hyperlinks.Delete
                .Range("CCEmail").value = "Enter carbon copy email..."
                .Range("CCEmail").Hyperlinks.Delete
            End With
            
            '-- Master: Reset key values
            Set Ws = ActiveWorkbook.Worksheets("Master")
            With Ws
                .Range("DocNum").value = "Enter document number..."
                .Range("CustomerName").value = "Enter customer name..."
                .Range("PONum").value = "Enter customer PO..."
                .Range("SalesOrderID").value = "QB ID..."
                .Range("H2:H7").value = ""
            End With
            
            '-- Deleted Items: Cleanup
            Set Ws = ActiveWorkbook.Worksheets("Deleted Items")
            With ActiveWorkbook.Worksheets("Deleted Items")
                .Range("A3:I300").value = ""
            End With
            
            
            '-- Computed Listings
            Call RefreshIndex
            Call CopyWorksheets
            
            
            'ActiveWorkbook.Worksheets("Instructions").
        End If
        
    End If
End Sub
