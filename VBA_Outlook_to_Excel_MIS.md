## Excel VBA code to fetch Outlook email content to Excel spreadsheet

This file uses Outlook.Application object and if outlook is running, it uses outlook namespace "MAPI".
Seaches for specified template in outlook emails and loads all such email data in tabular format in excel

### Output of the exercise would look like following:

![alt text](https://github.com/SomeshUCLA/Excel_VBA_Projects/blob/master/VBA_Outlook_to_Excel_MIS_files.jpg "MIS Excel Published")

### VBA Code

```

Sub MIS()
    Application.ScreenUpdating = False
    Dim OLApp As Outlook.Application
    On Error Resume Next
    Set OLApp = GetObject(, "Outlook.Application")    
    If Err Then
       MsgBox "Please start Microsoft Outlook before importing MIS data."
    Else
        Set objApp = Application
        Dim olns As Outlook.Namespace
        Set olns = Outlook.GetNamespace("MAPI")
        Dim wkb As Excel.Workbook
        Dim wks As Excel.Worksheet
        Dim EachElement()
        Dim applicantPosition As Integer
        Dim rowWithData As Integer
        Dim lastRow As Integer
        Dim selectedMessageCount As Integer
        Dim rowsAdded As Boolean
        rowsAdded = True
        Dim chkOverwrite As Boolean
        chkOverwrite = False
        Set wkb = ThisWorkbook
        Set wks = wkb.Sheets(1)
        wks.Unprotect Password:="MIS123"
        rowWithData = wks.Range("A1").End(xlDown).Row
        If wks.Shapes("Option Button 2").OLEFormat.Object.Value = 1 Then
            chkOverwrite = True
        End If
        If chkOverwrite Or rowWithData = 0 Then
            wks.Range("A2:Q" & rowWithData + 2).ClearContents
            With wks
                StartCount = 1 'how many emails (start at 1 to leave row one for headers)
                selectedMessageCount = 0
                strEmailContents = ""
                For Each outlookmessage In Outlook.Application.ActiveExplorer.Selection
                    StartCount = StartCount + 1 'increment email count
                    selectedMessageCount = selectedMessageCount + 1
                    UseCol = 1
                    FullMsg = outlookmessage.Body
                    AllLines = Split(FullMsg, vbCrLf)
                    For FullLine = LBound(AllLines) To UBound(AllLines)
                        On Error Resume Next
                      'Here is where you could decide to process only certain lines
                        eachVal = Split(AllLines(FullLine), ",") 'for a comma delimited file
                        For EachDataPoint = LBound(eachVal) To UBound(eachVal) 'load each value to an array
                            UseCol = UseCol + 1
                            ReDim Preserve EachElement(UseCol)
                            EachElement(UseCol - 1) = eachVal(EachDataPoint)
                        Next
                    Next
                    applicantPosition = 0
                    testFlag = False
                    For Each x In EachElement
                        applicantPosition = applicantPosition + 1
                        If x = "Applicant name" Then
                            If EachElement(applicantPosition + 17) = "CAM received date" Then
                                testFlag = True
                            End If
                            Exit For
                        End If
                    Next
                    If testFlag Then
                        wks.Cells(StartCount, 1) = StartCount - 1
                        wks.Cells(StartCount, 2) = Trim(UCase(EachElement(applicantPosition)))
                        wks.Cells(StartCount, 3) = Trim(UCase(EachElement(applicantPosition + 2)))
                        wks.Cells(StartCount, 4) = Trim(UCase(EachElement(applicantPosition + 4)))
                        wks.Cells(StartCount, 5) = Trim(UCase(EachElement(applicantPosition + 6)))
                        wks.Cells(StartCount, 6) = Trim(UCase(EachElement(applicantPosition + 8)))
                        wks.Cells(StartCount, 7) = EachElement(applicantPosition + 10) / 100
                        wks.Cells(StartCount, 8) = EachElement(applicantPosition + 12) / 100
                        wks.Cells(StartCount, 9) = EachElement(applicantPosition + 14) / 100
                        wks.Cells(StartCount, 10) = EachElement(applicantPosition + 16) / 100
                        wks.Cells(StartCount, 12) = "Approved"
                        wks.Cells(StartCount, 14) = EachElement(applicantPosition + 18)
                        wks.Cells(StartCount, 15) = Format(outlookmessage.ReceivedTime, "DD-MMM-YYYY")
                        wks.Cells(StartCount, 16).Formula = "=networkdays(N" & StartCount & ",O" & StartCount & ",)"
                        If InStr(outlookmessage.Categories, "Green Category") Then
                            wks.Cells(StartCount, 17) = "A"
                        ElseIf InStr(outlookmessage.Categories, "Blue Category") Then
                            wks.Cells(StartCount, 17) = "B"
                        ElseIf InStr(outlookmessage.Categories, "Yellow Category") Then
                            wks.Cells(StartCount, 17) = "C"
                        ElseIf InStr(outlookmessage.Categories, "Red Category") Then
                            wks.Cells(StartCount, 17) = "D"
                        End If
                    Else
                        StartCount = StartCount - 1
                    End If
                Next
            End With
        Else
            With wks
                rowWithData = rowWithData - 1
                StartCount = 1 'how many emails (start at 1 to leave row one for headers)
                selectedMessageCount = 0
                strEmailContents = ""
                For Each outlookmessage In Outlook.Application.ActiveExplorer.Selection
                    StartCount = StartCount + 1 'increment email count
                    selectedMessageCount = selectedMessageCount + 1
                    UseCol = 1
                    FullMsg = outlookmessage.Body
                    AllLines = Split(FullMsg, vbCrLf)
                    For FullLine = LBound(AllLines) To UBound(AllLines)
                        On Error Resume Next
                      'Here is where you could decide to process only certain lines
                        eachVal = Split(AllLines(FullLine), ",") 'for a comma delimited file
                        For EachDataPoint = LBound(eachVal) To UBound(eachVal) 'load each value to an array
                            UseCol = UseCol + 1
                            ReDim Preserve EachElement(UseCol)
                            EachElement(UseCol - 1) = eachVal(EachDataPoint)
                        Next
                    Next
                    applicantPosition = 0
                    testFlag = False
                    For Each x In EachElement
                        applicantPosition = applicantPosition + 1
                        If x = "Applicant name" Then
                            If EachElement(applicantPosition + 17) = "CAM received date" Then
                                testFlag = True
                            End If
                            Exit For
                        End If
                    Next
                    If testFlag Then
                        wks.Cells(StartCount + rowWithData, 1) = StartCount + rowWithData - 1
                        wks.Cells(StartCount + rowWithData, 2) = Trim(UCase(EachElement(applicantPosition)))
                        wks.Cells(StartCount + rowWithData, 3) = Trim(UCase(EachElement(applicantPosition + 2)))
                        wks.Cells(StartCount + rowWithData, 4) = Trim(UCase(EachElement(applicantPosition + 4)))
                        wks.Cells(StartCount + rowWithData, 5) = Trim(UCase(EachElement(applicantPosition + 6)))
                        wks.Cells(StartCount + rowWithData, 6) = Trim(UCase(EachElement(applicantPosition + 8)))
                        wks.Cells(StartCount + rowWithData, 7) = EachElement(applicantPosition + 10) / 100
                        wks.Cells(StartCount + rowWithData, 8) = EachElement(applicantPosition + 12) / 100
                        wks.Cells(StartCount + rowWithData, 9) = EachElement(applicantPosition + 14) / 100
                        wks.Cells(StartCount + rowWithData, 10) = EachElement(applicantPosition + 16) / 100
                        wks.Cells(StartCount + rowWithData, 12) = "Approved"
                        wks.Cells(StartCount + rowWithData, 14) = EachElement(applicantPosition + 18)
                        wks.Cells(StartCount + rowWithData, 15) = Format(outlookmessage.ReceivedTime, "DD-MMM-YYYY")
                        wks.Cells(StartCount + rowWithData, 16).Formula = "=networkdays(N" & StartCount + rowWithData & ",O" & StartCount + rowWithData & ",)"
                        If InStr(outlookmessage.Categories, "Green Category") Then
                            wks.Cells(StartCount + rowWithData, 17) = "A"
                        ElseIf InStr(outlookmessage.Categories, "Blue Category") Then
                            wks.Cells(StartCount + rowWithData, 17) = "B"
                        ElseIf InStr(outlookmessage.Categories, "Yellow Category") Then
                            wks.Cells(StartCount + rowWithData, 17) = "C"
                        ElseIf InStr(outlookmessage.Categories, "Red Category") Then
                            wks.Cells(StartCount + rowWithData, 17) = "D"
                        End If
                    Else
                        StartCount = StartCount - 1
                    End If
                Next
            End With
        End If
        lastRow = wks.Range("A1").End(xlDown).Row
        If lastRow <> 0 Then
            If Sheets(1).ComboBox1.Value = "No Sort" Or lastRow = 0 Then
            ElseIf Sheets(1).ComboBox1.Value = "Sort all rows" Or chkOverwrite Or rowWithData = 0 Then
                Columns("B:Q").Sort key1:=Range("O2"), key2:=Range("N2"), order1:=xlAscending, order2:=xlDescending, Header:=xlYes
            ElseIf Sheets(1).ComboBox1.Value = "Sort only imported rows" And StartCount > 2 Then
                Range("B" & rowWithData + 1 & ":Q" & lastRow & "").Sort key1:=Range("O" & rowWithData + 1 & ":O" & lastRow & ""), key2:=Range("N" & rowWithData + 1 & ":N" & lastRow & ""), order1:=xlAscending, order2:=xlDescending, Header:=xlYes
            End If
        End If
        If StartCount = 1 Then
            rowsAdded = False
            If selectedMessageCount = 1 Then
                MsgBox "1 message was selected in Outlook." & vbCrLf & "No data was imported from Outlook."
            Else
                MsgBox selectedMessageCount & " messages were selected in Outlook." & vbCrLf & "No data was imported from Outlook."
            End If
        ElseIf StartCount = 2 Then
            If selectedMessageCount = 1 Then
                MsgBox "1 message was selected in Outlook." & vbCrLf & "1 row was imported from Outlook."
            Else
                MsgBox selectedMessageCount & " messages were selected in Outlook." & vbCrLf & "1 row was imported from Outlook."
            End If
        Else
            MsgBox selectedMessageCount & " messages were selected in Outlook." & vbCrLf & "" & StartCount - 1 & " rows were imported from Outlook."
        End If
        Set olns = Nothing
        Set myinbox = Nothing
        Set myItems = Nothing
    End If
    Set myOlApp = Nothing
    If rowsAdded Then
        RemoveDuplicates (lastRow)
    End If    
    wks.Protect Password:="MIS123"
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
    With Sheets(1).ComboBox1
        .Clear
        .AddItem "No Sort"
        .AddItem "Sort only imported rows"
        .AddItem "Sort all rows"
        .Value = "Sort only imported rows"
    End With
    ThisWorkbook.Sheets(1).Shapes("Option Button 3").OLEFormat.Object.Value = 1
End Sub

Sub RemoveDuplicates(currentlastRowNumber As Integer)
    Dim lastRowAfterRemoveDuplicate As Integer
    ThisWorkbook.Sheets(1).Activate
    ActiveSheet.Range("A2:Q" & currentlastRowNumber).RemoveDuplicates Columns:=Array(2, 5, 14), Header:=xlNo
    lastRowAfterRemoveDuplicate = ThisWorkbook.Sheets(1).Range("A1").End(xlDown).Row
    If lastRowAfterRemoveDuplicate < currentlastRowNumber Then
        ActiveSheet.Range("A2:Q" & currentlastRowNumber).Locked = False
        For i = 1 To lastRowAfterRemoveDuplicate - 1
            Cells(i + 1, 1) = i
        Next
        If (currentlastRowNumber - lastRowAfterRemoveDuplicate) = 1 Then
            MsgBox (currentlastRowNumber - lastRowAfterRemoveDuplicate) & " duplicate row removed."
        Else
            MsgBox (currentlastRowNumber - lastRowAfterRemoveDuplicate) & " duplicate rows removed."
        End If
    End If
End Sub

```

