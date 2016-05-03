Attribute VB_Name = "FindLinks"

Sub Fetch_Links()
    Dim aLinks As Variant
    aLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        Sheets.Add
        For i = 1 To UBound(aLinks)
            Cells(i, 1).value = aLinks(i)
        Next i
    End If
End Sub


Sub ShowAllLinksInfo()
'Author:        JLLatham
'Purpose:       Identify which cells in which worksheets are using Linked Data
'Requirements:  requires a worksheet to be added to the workbook and named LinksList
'Modified From: http://answers.microsoft.com/en-us/office/forum/office_2007-excel/workbook-links-cannot-be-updated/b8242469-ec57-e011-8dfc-68b599b31bf5?page=1&tm=1301177444768
    Dim aLinks           As Variant
    Dim i                As Integer
    Dim Ws               As Worksheet
    Dim anyWS            As Worksheet
    Dim anyCell          As Range
    Dim reportWS         As Worksheet
    Dim nextReportRow    As Long
    Dim shtName          As String
 
    shtName = "LinksList"
 
    'Create the result sheet if one does not already exist
    For Each Ws In Application.Worksheets
        If Ws.name = shtName Then bWsExists = True
    Next Ws
    If bWsExists = False Then
        Application.DisplayAlerts = False
        Set Ws = ActiveWorkbook.Worksheets.Add(Type:=xlWorksheet)
        Ws.name = shtName
        Ws.Select
        Ws.Move After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Application.DisplayAlerts = True
    End If
 
    'Now start looking of linked data cells
    Set reportWS = ThisWorkbook.Worksheets(shtName)
    reportWS.Cells.Clear
    reportWS.Range("A1") = "Worksheet"
    reportWS.Range("B1") = "Cell"
    reportWS.Range("C1") = "Formula"
 
    aLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        'there are links somewhere in the workbook
        For Each anyWS In ThisWorkbook.Worksheets
            If anyWS.name <> reportWS.name Then
                For Each anyCell In anyWS.UsedRange
                    If anyCell.HasFormula Then
                        If InStr(anyCell.Formula, "[") > 0 Then
                            nextReportRow = reportWS.Range("A" & Rows.Count).End(xlUp).Row + 1
                            reportWS.Range("A" & nextReportRow) = anyWS.name
                            reportWS.Range("B" & nextReportRow) = anyCell.Address
                            reportWS.Range("C" & nextReportRow) = "'" & anyCell.Formula
                        End If
                    End If
                Next    ' end anyCell loop
            End If
        Next    ' end anyWS loop
    Else
        MsgBox "No links to Excel worksheets detected."
    End If
    'housekeeping
    Set reportWS = Nothing
    Set Ws = Nothing
End Sub

