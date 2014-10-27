Dim DB_PATH As String
Dim PPWK_PATH As String
Dim FILENAME As String

Public Sub scheduleInit()
    Dim sched As ListObject
    Dim i As Integer
    Dim addDate As String
    Dim addStart As Date
    Dim addEnd As Date
    Dim addName As String

    sched = Worksheets("Schedule").ListObjects("eventItems")
    scheduleList.Clear()
    If Not sched.Range.Cells(2, 1) = vbNullString Then

        For i = 2 To sched.DataBodyRange.Rows.Count + 1
            If i > 0 And sched.Range(i, 1) = sched.Range(i - 1, 1) Then
                addDate = ""
            Else
                addDate = sched.Range(i, 1)
            End If
            addStart = sched.Range(i, 2)
            addEnd = sched.Range(i, 3)
            addName = sched.Range(i, 4)
            With scheduleList
                .addItem()
                .List(i - 2, 0) = Format(addDate, "mm/dd/yyyy")
                .List(i - 2, 1) = Format(addStart, "hh:mm am/pm")
                .List(i - 2, 2) = Format(addEnd, "hh:mm am/pm")
                .List(i - 2, 3) = addName
                .List(i - 2, 4) = Format(addEnd - addStart, "hh:mm")
            End With

        Next
    End If


End Sub
Public Sub itemInit()
    Dim tbl As ListObject
    Dim adj As ListObject

    tbl = Worksheets(4).ListObjects("costItems")
    adj = Worksheets(4).ListObjects("adjTable")


    itemDateField.Value = ""
    itemSelect.Value = ""
    itemSelect.Text = ""
    qtyField.Value = ""
    hrsField.Value = ""
    qtyInc.Value = 0
    hrsInc.Value = 0

    Dim rngBlanks As Excel.Range

    With tbl
        On Error Resume Next
        If Not .DataBodyRange.Rows.Count = 1 Then
            rngBlanks = Intersect(.DataBodyRange, .ListColumns("Venue").Range).SpecialCells(xlCellTypeBlanks)
            On Error GoTo 0
            If Not rngBlanks Is Nothing Then
                rngBlanks.Delete()
            End If
        End If
    End With
    itemList.Clear()
    If Not tbl.DataBodyRange Is Nothing Then
        itemList.List = tbl.DataBodyRange.Value
    End If
    Dim lngIndex As Long
    With itemList
        For lngIndex = 0 To .ListCount - 1
            .List(lngIndex, 5) = (Format(Val(.List(lngIndex, 5)), "$#,##0.00"))
        Next
    End With
    subtotalList.List = tbl.TotalsRowRange.Value
    subtotalList.List(0, 5) = (Format(Val(subtotalList.List(0, 5)), "$#,##0.00"))
    adj.DataBodyRange(1, 6).Value = tbl.TotalsRowRange(1, 6) * adj.DataBodyRange(1, 2)
    adj.DataBodyRange(2, 6).Value = tbl.TotalsRowRange(1, 6) * adj.DataBodyRange(2, 2)
    adj.DataBodyRange(3, 6).Value = (tbl.TotalsRowRange(1, 6) + adj.DataBodyRange(1, 6) + adj.DataBodyRange(2, 6)) * adj.DataBodyRange(3, 2)
    adj.DataBodyRange(4, 6).Value = -((tbl.TotalsRowRange(1, 6) + adj.DataBodyRange(1, 6) + adj.DataBodyRange(2, 6) + adj.DataBodyRange(3, 6)) * adj.DataBodyRange(4, 2))
    With adjustmentList
        .Clear()
        If Select_DC.Value = True Then
            .addItem()
            .List(0, 0) = adj.DataBodyRange(1, 1) & " (" & Format(adj.DataBodyRange(1, 2), "0%") & ")"
            .List(0, 5) = adj.DataBodyRange(1, 6)
            .addItem()
            .List(1, 0) = "TOTAL"
            .List(1, 5) = (adj.DataBodyRange(1, 6).Value + tbl.TotalsRowRange(1, 6).Value)
        ElseIf Select_P.Value = True Then
            .addItem()
            .List(0, 0) = adj.DataBodyRange(1, 1) & " (" & Format(adj.DataBodyRange(1, 2), "0%") & ")"
            .List(0, 5) = adj.DataBodyRange(1, 6)
            .addItem()
            .List(1, 0) = adj.DataBodyRange(2, 1) & " (" & Format(adj.DataBodyRange(2, 2), "0%") & ")"
            .List(1, 5) = adj.DataBodyRange(2, 6)
            .addItem()
            .List(2, 0) = adj.DataBodyRange(3, 1) & " (" & Format(adj.DataBodyRange(3, 2), "0%") & ")"
            .List(2, 5) = adj.DataBodyRange(3, 6)
            .addItem()
            .List(3, 0) = "TOTAL"
            .List(3, 5) = (adj.DataBodyRange(1, 6).Value + adj.DataBodyRange(2, 6).Value + adj.DataBodyRange(3, 6) + tbl.TotalsRowRange(1, 6).Value)
        ElseIf Select_NFP.Value = True Then
            .addItem()
            .List(0, 0) = adj.DataBodyRange(1, 1) & " (" & Format(adj.DataBodyRange(1, 2), "0%") & ")"
            .List(0, 5) = adj.DataBodyRange(1, 6)
            .addItem()
            .List(1, 0) = adj.DataBodyRange(2, 1) & " (" & Format(adj.DataBodyRange(2, 2), "0%") & ")"
            .List(1, 5) = adj.DataBodyRange(2, 6)
            .addItem()
            .List(2, 0) = adj.DataBodyRange(3, 1) & " (" & Format(adj.DataBodyRange(3, 2), "0%") & ")"
            .List(2, 5) = adj.DataBodyRange(3, 6)
            .addItem()
            .List(3, 5) = adj.DataBodyRange(4, 6)
            .addItem()
            .List(4, 0) = "TOTAL"
            .List(4, 5) = (adj.DataBodyRange(1, 6).Value + adj.DataBodyRange(2, 6).Value + tbl.TotalsRowRange(1, 6).Value + adj.DataBodyRange(3, 6).Value)

        End If
        For lngIndex = 0 To .ListCount - 1
            .List(lngIndex, 5) = (Format(Val(.List(lngIndex, 5)), "$#,##0.00"))
        Next
    End With


End Sub



Private Sub boNotes_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Box Office Notes").DataBodyRange = boNotes.Value
End Sub

Private Sub delItem_Click()
    Dim tbl As ListObject
    Dim adj As ListObject
    tbl = Worksheets(4).ListObjects("costItems")
    adj = Worksheets(4).ListObjects("adjTable")

    Dim index As Integer


    For i = 0 To itemList.ListCount - 1
        If itemList.Selected(i) Then
            index = i
        End If
    Next
    tbl.ListRows(index + 1).Delete()
    Call itemInit()


End Sub

Private Sub editItem_Click()
    Dim tbl As ListObject
    Dim adj As ListObject
    Dim useTable As Range

    tbl = Worksheets(4).ListObjects("costItems")
    adj = Worksheets(4).ListObjects("adjTable")

    Dim index As Integer

    For i = 0 To itemList.ListCount - 1
        If itemList.Selected(i) Then
            index = i
        End If
    Next

    Dim editRow As ListRow

    With itemSelect
        .TextColumn = 2
        Select Case .Column(0, .ListIndex)
            Case "ERH"
                useTable = [erh_Items]
            Case "NT"
                useTable = [nt_Items]
            Case "MH"
                useTable = [mh_Items]
            Case "BW"
                useTable = [bw_Items]
            Case "RS"
                useTable = [rs_Items]
            Case "B2"
                useTable = [b2_Items]
            Case "B3"
                useTable = [b3_Items]
            Case "SR"
                useTable = [sr_Items]
        End Select

        For Each row In useTable.Rows
            If .Column(1, .ListIndex) = row.Columns(1) Then
                If row.Columns(3) = "Hourly" Then
                    itemCost = qtyField * hrsField * row.Columns(2)
                Else
                    itemCost = qtyField * row.Columns(2)
                End If
            End If
        Next


    End With
    editRow = tbl.ListRows(index + 1)
    editRow.Range = Array(itemDateField.Value, itemSelect.Value, itemSelect.Text, qtyField.Value, hrsField.Value, itemCost)
    Call itemInit()

End Sub
Private Sub hrsIncr_Change()
    hrsField.Value = hrsInc.Value

End Sub




Private Sub addItem_Click()
    Dim tbl As ListObject
    Dim adj As ListObject
    Dim newRow As ListRow
    Dim useTable As Range
    Dim itemCost As Double
    Dim itemPrice As Double
    Dim amtProfit As Double

    With itemSelect
        .TextColumn = 2
        Select Case .Column(0, .ListIndex)
            Case "ERH"
                useTable = [erh_Items]
            Case "NT"
                useTable = [nt_Items]
            Case "MH"
                useTable = [mh_Items]
            Case "BW"
                useTable = [bw_Items]
            Case "RS"
                useTable = [rs_Items]
            Case "B2"
                useTable = [b2_Items]
            Case "B3"
                useTable = [b3_Items]
            Case "SR"
                useTable = [sr_Items]
        End Select

        For Each row In useTable.Rows
            If .Column(1, .ListIndex) = row.Columns(1) Then
                If row.Columns(3) = "Hourly" Then
                    itemCost = qtyField * hrsField * row.Columns(2)
                Else
                    itemCost = qtyField * row.Columns(2)
                End If
                Dim itemCat
                itemCat = row.Columns(4)
                If hrsField = 0 Then
                    hrsField = Null
                End If
            End If
        Next


    End With

    tbl = Worksheets(4).ListObjects("costItems")
    newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    adj = Worksheets(4).ListObjects("adjTable")

    newRow.Range = Array(itemDateField.Value, itemSelect.Value, itemSelect.Text, qtyField.Value, hrsField.Value, itemCost, itemCat)
    Call itemInit()


End Sub





Private Sub hmInc_Change()
    numHM.Value = hmInc.Value
End Sub

Private Sub fohNotes_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("FOH Notes").DataBodyRange = fohNotes.Value
End Sub

Private Sub hrsField_Change()
    If Not IsNumeric(hrsField.Value) And Not hrsField.Value = "" Then
        MsgBox("Please enter the number of hours in a numeric format")
        hrsField.Value = 0
        hrsInc.Value = 0
    ElseIf Not hrsField.Value = "" Then
        hrsInc.Value = hrsField.Value
    End If

End Sub

Private Sub hrsInc_SpinUp()
    hrsField.Value = hrsField.Value + 1

End Sub
Private Sub hrsInc_SpinDown()
    hrsField.Value = hrsField.Value - 1
End Sub

Private Sub intermissionLength_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Intermission").DataBodyRange = intermissionLength.Value
End Sub

Private Sub itemList_Click()
    Dim tbl As ListObject
    Dim adj As ListObject

    With itemList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                itemDateField.Value = .List(i, 0)
                itemSelect.Value = .List(i, 1)
                itemSelect.Text = .List(i, 2)
                qtyField.Value = .List(i, 3)
                qtyInc.Value = .List(i, 3)
                hrsField.Value = .List(i, 4)
                hrsInc.Value = .List(i, 4)
            End If
        Next
    End With



End Sub

Private Sub itemSelect_Change()

    Dim useTable As Range

    With itemSelect
        If Not .Value = "" Then

            .TextColumn = 2
            Select Case .Column(0, .ListIndex)
                Case "ERH"
                    useTable = [erh_Items]
                Case "NT"
                    useTable = [nt_Items]
                Case "MH"
                    useTable = [mh_Items]
                Case "BW"
                    useTable = [bw_Items]
                Case "RS"
                    useTable = [rs_Items]
                Case "B2"
                    useTable = [b2_Items]
                Case "B3"
                    useTable = [b3_Items]
                Case "SR"
                    useTable = [sr_Items]
            End Select
            For Each row In useTable.Rows
                If .Column(1, .ListIndex) = row.Columns(1) Then
                    Select Case row.Columns(3)
                        Case "Hourly"
                            itemDateField.Enabled = True
                            dayLabel.Enabled = True
                            hrsLabel.Enabled = True
                            hrsField.Enabled = True
                            hrsInc.Enabled = True
                        Case "Daily"
                            itemDateField.Enabled = True
                            dayLabel.Enabled = True
                            hrsLabel.Enabled = False
                            hrsField.Enabled = False
                            hrsInc.Enabled = False
                        Case "Once"
                            itemDateField.Enabled = False
                            itemDateField.Value = "--/--/--"
                            dayLabel.Enabled = False
                            hrsLabel.Enabled = False
                            hrsField.Enabled = False
                            hrsInc.Enabled = False
                    End Select
                End If
            Next
        End If

    End With

End Sub



Private Sub marketingNotes_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Marketing Notes").DataBodyRange = marketingNotes.Value
End Sub

Private Sub MultiPage1_Change()
    scheduleStart.CustomFormat = "hh:mm tt"
    scheduleEnd.CustomFormat = "hh:mm: tt"
    scheduleStart.Format = dtpCustom
    scheduleEnd.Format = dtpCustom
End Sub

Private Sub numHM_Change()
    If Not IsNumeric(numHM.Value) And Not numVM.Value = "" Then
        MsgBox("Please enter the number of HMs in numeric format")
        numHM.Value = ""
        hmInc.Value = 0
    ElseIf Not numHM.Value = "" Then
        hmInc.Value = numHM.Value
    End If


End Sub

Private Sub numUsh_Change()
    If Not IsNumeric(numUsh.Value) And Not numUsh.Value = "" Then
        MsgBox("Please enter the number of ushers in numeric format")
        numUsh.Value = ""
        ushInc.Value = 0
    ElseIf Not numUsh.Value = "" Then
        ushInc.Value = numUsh.Value
    End If

End Sub

Private Sub numVM_Change()
    If Not IsNumeric(numVM.Value) And Not numVM.Value = "" Then
        MsgBox("Please enter the number of VMs in numeric format")
        numVM.Value = ""
        vmInc.Value = 0
    ElseIf Not numVM.Value = "" Then
        vmInc.Value = numVM.Value
    End If

End Sub

Private Sub overviewCreate_Click()
    Call saveAs_Click()
    Dim wrdApp
    Dim wrdDoc
    Dim wrdRange

    Dim ventbl As ListObject
    Dim data As Range
    Dim firstDate As Date
    Dim lastDate As Date


    ventbl = Worksheets("Venues").ListObjects("Venues")
    data = ventbl.DataBodyRange
    firstDate = Format("1/1/9999", "dd/mm/yyyy")
    lastDate = Format("1/1/1900", "dd/mm/yyyy")

    Dim itemTbl As Range
    itemTbl = [costItems]


    For i = 1 To data.Rows.Count
        If Not data.Cells(i, 4) = vbNullString And Not data.Cells(i, 4) > firstDate Then
            firstDate = data.Cells(i, 4)
        End If
        If Not data.Cells(i, 5) = vbNullString And Not data.Cells(i, 5) < lastDate Then
            lastDate = data.Cells(i, 5)
        End If
    Next

    wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True
    wrdDoc = wrdApp.Documents.Add
    wrdRange = wrdDoc.Range
    With wrdApp.Selection
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .Font.Size = 12
        .Font.Bold = True
        .TypeText "EVENT OVERVIEW" & vbCr & vbCr
        .Font.Bold = False
        .TypeText compNameField.Value & vbCr
        .TypeText eventNameField.Value & vbCr
        .TypeText firstDate & " - " & lastDate
        .Font.Size = 10
        .TypeParagraph()
        .TypeParagraph()
        wrdDoc.Tables.Add.Range, 1, 2
        Dim wrdTbl
        wrdTbl = wrdDoc.Tables(1)
        wrdTbl.Cell(1, 1).Range.Text = "Primary Contact:" & Chr(11) & contact1NameField.Value & Chr(11) & contact1PhoneField.Value & Chr(11) & contact1EmailField.Value
        wrdTbl.Cell(1, 2).Range.Text = "Secondary Contact:" & Chr(11) & contact2NameField.Value & Chr(11) & contact2PhoneField.Value & Chr(11) & contact2EmailField.Value

        For i = 1 To 10
            .MoveDown()
        Next
        .TypeParagraph()
        .TypeText "Run time: " & runTime.Value & vbCr
        .TypeText "Intermission: " & intermissionLength.Value & vbCr & vbCr
        .Font.Bold = True
        .TypeText "Summary:" & vbCr
        .Font.Bold = False
        .TypeText summaryField.Value & vbCr
        .TypeText "-------------------" & vbCr
        .Font.Bold = True
        .TypeText "Tech Notes:" & vbCr
        .Font.Bold = False
        .TypeText techNotes.Value & vbCr & vbCr
        wrdDoc.Tables.Add.Range, 2, 5
        Dim techTbl
        techTbl = wrdDoc.Tables(2)
        techTbl.Cell(1, 1).Range.Text = "Date"
        techTbl.Cell(1, 2).Range.Text = "Item"
        techTbl.Cell(1, 3).Range.Text = "Quantity"
        techTbl.Cell(1, 4).Range.Text = "Hours"
        techTbl.Cell(1, 5).Range.Text = "Venue"
        Dim rowIndex As Integer
        rowIndex = 2
        For Each row In itemTbl.Rows
            If row.Columns(7) = "TECH" Then
                techTbl.Cell(rowIndex, 1) = row.Columns(1)
                techTbl.Cell(rowIndex, 2) = row.Columns(3)
                techTbl.Cell(rowIndex, 3) = row.Columns(4)
                techTbl.Cell(rowIndex, 4) = row.Columns(5)
                techTbl.Cell(rowIndex, 5) = row.Columns(2)
                techTbl.Rows.Add()
                rowIndex = rowIndex + 1
            End If
        Next
        For i = 1 To 10
            .MoveDown()
        Next

        .TypeText "-------------------" & vbCr
        .Font.Bold = True
        .TypeText "FOH Notes:" & vbCr
        .Font.Bold = False
        .TypeText fohNotes.Value & vbCr & vbCr
        wrdDoc.Tables.Add.Range, 2, 5
        Dim fohTbl
        fohTbl = wrdDoc.Tables(3)
        fohTbl.Cell(1, 1).Range.Text = "Date"
        fohTbl.Cell(1, 2).Range.Text = "Item"
        fohTbl.Cell(1, 3).Range.Text = "Quantity"
        fohTbl.Cell(1, 4).Range.Text = "Hours"
        fohTbl.Cell(1, 5).Range.Text = "Venue"
        rowIndex = 2
        For Each row In itemTbl.Rows
            If row.Columns(7) = "FOH" Then
                fohTbl.Cell(rowIndex, 1) = row.Columns(1)
                fohTbl.Cell(rowIndex, 2) = row.Columns(3)
                fohTbl.Cell(rowIndex, 3) = row.Columns(4)
                fohTbl.Cell(rowIndex, 4) = row.Columns(5)
                fohTbl.Cell(rowIndex, 5) = row.Columns(2)
                fohTbl.Rows.Add()
                rowIndex = rowIndex + 1

            End If
        Next
        For i = 1 To 10
            .MoveDown()
        Next
        .TypeText "-------------------" & vbCr
        .Font.Bold = True
        .TypeText "Box Office Notes:" & vbCr
        .Font.Bold = False
        .TypeText boNotes.Value & vbCr & vbCr
        wrdDoc.Tables.Add.Range, 2, 5
        Dim boTbl
        boTbl = wrdDoc.Tables(4)
        boTbl.Cell(1, 1).Range.Text = "Date"
        boTbl.Cell(1, 2).Range.Text = "Item"
        boTbl.Cell(1, 3).Range.Text = "Quantity"
        boTbl.Cell(1, 4).Range.Text = "Hours"
        boTbl.Cell(1, 5).Range.Text = "Venue"
        rowIndex = 2
        For Each row In itemTbl.Rows
            If row.Columns(7) = "BOX OFFICE" Then
                boTbl.Cell(rowIndex, 1) = row.Columns(1)
                boTbl.Cell(rowIndex, 2) = row.Columns(3)
                boTbl.Cell(rowIndex, 3) = row.Columns(4)
                boTbl.Cell(rowIndex, 4) = row.Columns(5)
                boTbl.Cell(rowIndex, 5) = row.Columns(2)
                boTbl.Rows.Add()
                rowIndex = rowIndex + 1

            End If
        Next
        For i = 1 To 10
            .MoveDown()
        Next
        '.TypeText "{list of box office items here}"
        .TypeParagraph()
        .TypeText "-------------------" & vbCr
        .Font.Bold = True
        .TypeText "Marketing Notes:" & vbCr
        .Font.Bold = False
        .TypeText marketingNotes.Value & vbCr & vbCr
        wrdDoc.Tables.Add.Range, 2, 5
        Dim mktTbl
        mktTbl = wrdDoc.Tables(5)
        mktTbl.Cell(1, 1).Range.Text = "Date"
        mktTbl.Cell(1, 2).Range.Text = "Item"
        mktTbl.Cell(1, 3).Range.Text = "Quantity"
        mktTbl.Cell(1, 4).Range.Text = "Hours"
        mktTbl.Cell(1, 5).Range.Text = "Venue"
        rowIndex = 2
        For Each row In itemTbl.Rows
            If row.Columns(7) = "MARKETING" Then
                mktTbl.Cell(rowIndex, 1) = row.Columns(1)
                mktTbl.Cell(rowIndex, 2) = row.Columns(3)
                mktTbl.Cell(rowIndex, 3) = row.Columns(4)
                mktTbl.Cell(rowIndex, 4) = row.Columns(5)
                mktTbl.Cell(rowIndex, 5) = row.Columns(2)
                mktTbl.Rows.Add()
                rowIndex = rowIndex + 1


            End If
        Next
        '.TypeText "{list of marketing items here}"
        For i = 1 To 10
            .MoveDown()
        Next
        wrdDoc.Content.Select()
        With wrdApp.Selection.Find
            .Forward = True
            .Wrap = wdFindStop
            .Text = "Primary Contact:"
            .Execute()
        End With
        wrdApp.Selection.Font.Bold = True

        wrdDoc.Content.Select()
        With wrdApp.Selection.Find
            .Forward = True
            .Wrap = wdFindStop
            .Text = "Secondary Contact:"
            .Execute()
        End With
        wrdApp.Selection.Font.Bold = True
        techTbl.Rows(1).Range.Font.Bold = True
        fohTbl.Rows(1).Range.Font.Bold = True
        boTbl.Rows(1).Range.Font.Bold = True
        mktTbl.Rows(1).Range.Font.Bold = True

    End With



    With wrdDoc
        '    .saveAs ("E:\Projects\BPAC\Test.doc")
        .saveAs(PPWK_PATH & FILENAME & "_OVERVIEW.doc")
    End With
    wrdApp = Nothing
    wrdDoc = Nothing


End Sub

Private Sub postMortemText_Change()
    Dim saveRng As Range
    saveRng = Worksheets("Post-Mortem").Cells(1, 1)
    saveRng.Value = postMortemText.Value
End Sub

Private Sub qtyField_Change()
    If Not IsNumeric(qtyField.Value) And Not qtyField.Value = "" Then
        MsgBox("Please enter a quantity in numeric format")
        qtyField.Value = 0
        qtyInc.Value = 0
    ElseIf Not qtyField.Value = "" Then
        qtyInc.Value = qtyField.Value
    End If
End Sub

Private Sub runTime_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Run Time").DataBodyRange = runTime.Value
End Sub

Private Sub saveAs_Click()
    Dim saveDate As Date
    Dim ventbl As ListObject
    Dim data As Range

    ventbl = Worksheets("Venues").ListObjects("Venues")
    data = ventbl.DataBodyRange
    saveDate = Format("1/1/9999", "dd/mm/yyyy")

    For i = 1 To data.Rows.Count
        If Not data.Cells(i, 4) = vbNullString And Not data.Cells(i, 4) > saveDate Then
            saveDate = data.Cells(i, 4)
            '        MsgBox (saveDate)
        End If
    Next
    Dim BAD_CHAR
    BAD_CHAR = Array("!", "@", "#", "$", "%", "^", "&", "*", "{", "}", "=", "+", "/", "<", ">", ",", ".", "?", "\", "|", "`")
    FILENAME = Format(saveDate, "yy") & "-" & Format(saveDate, "mm") & "-" & Format(saveDate, "dd") & "_" & Replace(compNameField.Value, " ", "-") & "_" & Replace(eventNameField, " ", "-")
    For Each element In BAD_CHAR
        FILENAME = Replace(FILENAME, CStr(element), "")
    Next

    ' MsgBox (DB_PATH & FILENAME)
    ThisWorkbook.saveAs(DB_PATH & FILENAME, FileFormat:=52)
    'MsgBox ("E:\Projects\" & Format(saveDate, "yy") & "-" & Format(saveDate, "mm") & "-" & Format(saveDate, "dd") & "_" & Replace(compNameField.Value, " ", "-") & "_" & Replace(eventNameField, " ", "-"))
End Sub


Private Sub schedName_Change()

End Sub

Private Sub schedNotes_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Schedule Notes").DataBodyRange = schedNotes.Value
End Sub

Private Sub scheduleAdd_Click()
    Dim rngBlanks As Excel.Range
    Dim sched As ListObject
    Dim newRow As ListRow

    sched = Worksheets("Schedule").ListObjects("eventItems")

    newRow = sched.ListRows.Add(AlwaysInsert:=True)

    newRow.Range = Array(scheduleDate.Value, scheduleStart.Value, scheduleEnd.Value, schedName.Value)
    With sched
        On Error Resume Next
        If Not .DataBodyRange.Rows.Count = 1 Then
            rngBlanks = Intersect(.DataBodyRange, .ListColumns("Date").Range).SpecialCells(xlCellTypeBlanks)
            On Error GoTo 0
            If Not rngBlanks Is Nothing Then
                rngBlanks.Delete()
            End If
        End If
        .Sort.SortFields.Clear()
        .Sort.SortFields.Add( _
            Key:=sched.ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal)
        .Sort.SortFields.Add( _
            Key:=sched.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal)
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply()
        End With
    End With
    Call scheduleInit()
End Sub


Private Sub scheduleCreate_Click()


    Call saveAs_Click()

    Dim EXCL_ARRAY(1 To 3) As String
    EXCL_ARRAY(1) = "CALL"
    EXCL_ARRAY(2) = "OUT"
    EXCL_ARRAY(3) = "EXCLUDE"

    Dim wrdApp
    Dim wrdDoc
    Dim wrdRange

    Dim ventbl As ListObject
    Dim data As Range
    Dim firstDate As Date
    Dim lastDate As Date

    Dim schedAll As ListObject
    Dim schedRng As Range
    schedAll = Worksheets("Schedule").ListObjects("eventItems")
    schedRng = schedAll.DataBodyRange

    ventbl = Worksheets("Venues").ListObjects("Venues")
    data = ventbl.DataBodyRange
    firstDate = Format("1/1/9999", "dd/mm/yyyy")
    lastDate = Format("1/1/1900", "dd/mm/yyyy")

    For i = 1 To data.Rows.Count
        If Not data.Cells(i, 4) = vbNullString And Not data.Cells(i, 4) > firstDate Then
            firstDate = data.Cells(i, 4)
        End If
        If Not data.Cells(i, 5) = vbNullString And Not data.Cells(i, 5) < lastDate Then
            lastDate = data.Cells(i, 5)
        End If
    Next

    wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True
    wrdDoc = wrdApp.Documents.Add
    wrdRange = wrdDoc.Range
    With wrdApp.Selection
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .Font.Size = 12
        .Font.Bold = True
        .TypeText "CLIENT SCHEDULE" & vbCr & vbCr
        .Font.Bold = False
        .TypeText compNameField.Value & vbCr
        .TypeText eventNameField.Value & vbCr
        .TypeText firstDate & " - " & lastDate
        .Font.Size = 10
        .TypeParagraph()
        .TypeParagraph()
        wrdDoc.Tables.Add.Range, 1, 2
        Dim wrdTbl
        wrdTbl = wrdDoc.Tables(1)
        wrdTbl.Cell(1, 1).Range.Text = "Primary Contact:" & Chr(11) & contact1NameField.Value & Chr(11) & contact1PhoneField.Value & Chr(11) & contact1EmailField.Value
        wrdTbl.Cell(1, 2).Range.Text = "Secondary Contact:" & Chr(11) & contact2NameField.Value & Chr(11) & contact2PhoneField.Value & Chr(11) & contact2EmailField.Value

        For i = 1 To 10
            .MoveDown()
        Next
        .TypeParagraph()
        .TypeText "Run time: " & runTime.Value & vbCr
        .TypeText "Intermission: " & intermissionLength.Value & vbCr & vbCr
        .TypeText "-------------------" & vbCr

        wrdDoc.Tables.Add.Range, 2, 4
        Dim schedTbl
        schedTbl = wrdDoc.Tables(2)

        schedTbl.Cell(1, 1).Range.Text = "Date"
        schedTbl.Cell(1, 2).Range.Text = "Start Time"
        schedTbl.Cell(1, 3).Range.Text = "End Time"
        schedTbl.Cell(1, 4).Range.Text = "Scheduled Item"


        Dim printRow As Boolean
        Dim numRows As Integer
        numRows = 2
        For i = 1 To schedRng.Rows.Count
            printRow = True
            For Each phrase In EXCL_ARRAY
                If Not InStr(schedRng.Cells(i, 4), phrase) = 0 Then
                    printRow = False
                End If
            Next
            If printRow = True Then
                Dim checkDate As String
                Dim printDate As Boolean
                checkDate = schedRng.Cells(i, 1)
                printDate = True
                For j = 1 To numRows
                    If Not InStr(schedTbl.Cell(j, 1), checkDate) = 0 Then
                        printDate = False
                    End If
                Next
                If numRows = 2 Then
                    printDate = True
                End If

                If printDate = True Then
                    schedTbl.Cell(numRows, 1).Range.Text = schedRng.Cells(i, 1)

                Else
                    schedTbl.Cell(numRows, 1).Range.Text = vbNullString
                End If
                schedTbl.Cell(numRows, 2).Range.Text = Format(schedRng.Cells(i, 2), "hh:mm am/pm")
                schedTbl.Cell(numRows, 3).Range.Text = Format(schedRng.Cells(i, 3), "hh:mm am/pm")
                schedTbl.Cell(numRows, 4).Range.Text = schedRng.Cells(i, 4)
                schedTbl.Rows.Add()
                numRows = numRows + 1
            End If

        Next
        schedTbl.Rows(1).Range.Font.Bold = True
        schedTbl.Rows(1).Range.Font.Underline = True



    End With

    wrdDoc.saveAs PPWK_PATH & FILENAME & "_CLIENT_SCHEDULE"
    wrdDoc = Nothing
    wrdDoc = wrdApp.Documents.Add
    wrdRange = wrdDoc.Range
    With wrdApp.Selection
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .Font.Size = 12
        .Font.Bold = True
        .TypeText "INTERNAL SCHEDULE" & vbCr & vbCr
        .Font.Bold = False
        .TypeText compNameField.Value & vbCr
        .TypeText eventNameField.Value & vbCr
        .TypeText firstDate & " - " & lastDate
        .Font.Size = 10
        .TypeParagraph()
        .TypeParagraph()
        wrdDoc.Tables.Add.Range, 1, 2
        wrdTbl = wrdDoc.Tables(1)
        wrdTbl.Cell(1, 1).Range.Text = "Primary Contact:" & Chr(11) & contact1NameField.Value & Chr(11) & contact1PhoneField.Value & Chr(11) & contact1EmailField.Value
        wrdTbl.Cell(1, 2).Range.Text = "Secondary Contact:" & Chr(11) & contact2NameField.Value & Chr(11) & contact2PhoneField.Value & Chr(11) & contact2EmailField.Value

        For i = 1 To 10
            .MoveDown()
        Next
        .TypeParagraph()
        .TypeText "Run time: " & runTime.Value & vbCr
        .TypeText "Intermission: " & intermissionLength.Value & vbCr & vbCr
        .TypeText "-------------------" & vbCr

        wrdDoc.Tables.Add.Range, 2, 4
        schedTbl = wrdDoc.Tables(2)

        schedTbl.Cell(1, 1).Range.Text = "Date"
        schedTbl.Cell(1, 2).Range.Text = "Start Time"
        schedTbl.Cell(1, 3).Range.Text = "End Time"
        schedTbl.Cell(1, 4).Range.Text = "Scheduled Item"


        numRows = 2
        For i = 1 To schedRng.Rows.Count


            checkDate = schedRng.Cells(i, 1)
            printDate = True
            For j = 1 To numRows
                If Not InStr(schedTbl.Cell(j, 1), checkDate) = 0 Then
                    printDate = False
                End If
            Next
            If numRows = 2 Then
                printDate = True
            End If

            If printDate = True Then
                schedTbl.Cell(numRows, 1).Range.Text = schedRng.Cells(i, 1)

            Else
                schedTbl.Cell(numRows, 1).Range.Text = vbNullString
            End If
            schedTbl.Cell(numRows, 2).Range.Text = Format(schedRng.Cells(i, 2), "hh:mm am/pm")
            schedTbl.Cell(numRows, 3).Range.Text = Format(schedRng.Cells(i, 3), "hh:mm am/pm")
            schedTbl.Cell(numRows, 4).Range.Text = schedRng.Cells(i, 4)
            schedTbl.Rows.Add()
            numRows = numRows + 1


        Next
        schedTbl.Rows(1).Range.Font.Bold = True
        schedTbl.Rows(1).Range.Font.Underline = True



    End With

    wrdDoc.saveAs PPWK_PATH & FILENAME & "_INTERNAL_SCHEDULE.doc"
    wrdDoc = Nothing
    wrdApp = Nothing


End Sub


Private Sub scheduleDelete_Click()

    Dim index As Integer
    Dim sched As ListObject

    sched = Worksheets("Schedule").ListObjects("eventItems")


    With scheduleList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                index = i
            End If
        Next
        sched.ListRows(index + 1).Delete()
        Call scheduleInit()
    scheduleDate.Value = Date
        scheduleStart.Value = "12:00 am"
        scheduleEnd.Value = "12:00 am"
        schedName.Value = ""

    End With
End Sub
Private Sub scheduleEdit_Click()
    Dim schedTbl As ListObject
    Dim thisRow As ListRow
    schedTbl = Worksheets("Schedule").ListObjects("eventItems")

    rowNum = scheduleList.ListIndex + 1

    thisRow = schedTbl.ListRows(rowNum)
    thisRow.Range = Array(scheduleDate.Value, scheduleStart.Value, scheduleEnd.Value, schedName.Value)
    Call scheduleInit()
End Sub
Private Sub scheduleList_Click()
    Dim index As Integer
    With scheduleList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                index = i
            End If
        Next
        Dim dateToUse As String
        dateToUse = .List(index, 0)
        Dim lineNum As Integer
        lineNum = index
        Do While .List(lineNum, 0) = vbNullString
            lineNum = lineNum - 1
        Loop
        scheduleDate.Value = .List(lineNum, 0)
        scheduleStart.Value = .List(index, 1)
        scheduleEnd.Value = .List(index, 2)
        schedName.Value = .List(index, 3)
    End With

End Sub

Private Sub Select_DC_Click()
    Call itemInit()

    [OverviewData[Billing Option]].Value = "DC"
End Sub
Private Sub Select_NFP_Click()
    Call itemInit()

    [OverviewData[Billing Option]].Value = "NFP"
End Sub
Private Sub Select_P_Click()
    Call itemInit()

    [OverviewData[Billing Option]].Value = "P"
End Sub

Private Sub summaryField_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Summary").DataBodyRange = summaryField.Value
End Sub
Private Sub taxIDNum_Change()
    [OverviewData[Tax ID]].Value = taxIDNum.Value
End Sub


Private Sub techNotes_Change()
    Dim ovwData As ListObject
    ovwData = Worksheets("OverviewSheet").ListObjects("OverviewData")
    ovwData.ListColumns("Tech Notes").DataBodyRange = techNotes.Value
End Sub


Private Sub UserForm_Initialize()

    DB_PATH = Application.ActiveWorkbook.PATH & "\"
    PPWK_PATH = Application.ActiveWorkbook.PATH & "\"

    Dim tbl As ListObject
    Dim adj As ListObject
    tbl = Worksheets(4).ListObjects("costItems")
    adj = Worksheets(4).ListObjects("adjTable")

scheduleDate.Value = Date
    scheduleStart.Value = "12:00:00"
    scheduleEnd.Value = "12:00:00"
    scheduleStart.CustomFormat = "hh:mm tt"
    scheduleEnd.CustomFormat = "hh:mm tt"
    scheduleStart.UpDown = True
    scheduleEnd.UpDown = True

    With newEvent
        .MultiPage1.Value = 0
        For Each row In [Venues].Rows

            .Controls(row.Columns(6).Value).Value = row.Columns(3)
            .Controls(row.Columns(2) & "Start").Value = row.Columns(4)
            .Controls(row.Columns(2) & "End").Value = row.Columns(5)

        Next

    .Controls("compNameField").Value = [OverviewData[Company Name]].Value
    .Controls("eventNameField").Value = [OverviewData[Event Name]].Value
    .Controls("contact1NameField").Value = [OverviewData[Primary Contact Name]].Value
    .Controls("contact1PhoneField").Value = [OverviewData[Primary Contact Phone]].Value
    .Controls("contact1EmailField").Value = [OverviewData[Primary Contact Email]].Value
    .Controls("contact2NameField").Value = [OverviewData[Secondary Contact Name]].Value
    .Controls("contact2PhoneField").Value = [OverviewData[Secondary Contact Phone]].Value
    .Controls("contact2EmailField").Value = [OverviewData[Secondary Contact Email]].Value
    Website.Value = [OverviewData[URL]].Value
    taxIDNum.Value = [OverviewData[Tax ID]].Value
    runTime.Value = [OverviewData[Run Time]].Value
    intermissionLength.Value = [OverviewData[Intermission]].Value
    summaryField.Value = [OverviewData[Summary]].Value
    schedNotes.Value = [OverviewData[Schedule Notes]].Value
    techNotes.Value = [OverviewData[Tech Notes]].Value
    fohNotes.Value = [OverviewData[FOH Notes]].Value
    boNotes.Value = [OverviewData[Box Office Notes]].Value
    marketingNotes.Value = [OverviewData[Marketing Notes]].Value
        postMortemText.Value = Worksheets("Post-Mortem").Cells(1, 1).Value
    If [OverviewData[Billing Option]].Value = "" Then
            .Controls("Select_P").Value = True
       [OverviewData[Billing Option]].Value = "P"
        Else
        .Controls("Select_" & [OverviewData[Billing Option]].Value).Value = True
        End If

    End With
    If Not tbl.DataBodyRange Is Nothing Then
        itemList.List = tbl.DataBodyRange.Value
    End If
    subtotalList.List = tbl.TotalsRowRange.Value

    Call itemInit()
    Call scheduleInit()
    Dim sched As ListObject
    sched = Worksheets("Schedule").ListObjects("eventItems")

End Sub

Private Sub erhStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With erhStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            [Venues[StartDate]].Rows(1).Value = .Value
        End If
    End With
End Sub

Private Sub ntStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With ntStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(2).Value = .Value
        End If
    End With
End Sub
Private Sub mhStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With mhStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(3).Value = .Value
        End If
    End With
End Sub
Private Sub bwStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With bwStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(4).Value = .Value
        End If
    End With
End Sub
Private Sub rsStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With rsStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(5).Value = .Value
        End If
    End With
End Sub
Private Sub b2Start_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With b2Start
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(6).Value = .Value
        End If
    End With
End Sub
Private Sub b3Start_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With b3Start
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(7).Value = .Value
        End If
    End With
End Sub
Private Sub srStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With srStart
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
             [Venues[StartDate]].Rows(8).Value = .Value
        End If
    End With
End Sub

Private Sub erhEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With erhEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < erhStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(1).Value = .Value
            End If
        End If
    End With
End Sub
Private Sub ntEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With ntEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < ntStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(2).Value = .Value
            End If
        End If
    End With
End Sub
Private Sub mhEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With mhEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < mhStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(3).Value = .Value
            End If
        End If
    End With
End Sub
Private Sub bwEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With bwEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < bwStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(4).Value = .Value
            End If
        End If
    End With
End Sub
Private Sub rsEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With rsEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < rsStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(5).Value = .Value
            End If
        End If
    End With
End Sub

Private Sub b2End_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With b2End
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < b2Start.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(6).Value = .Value
            End If
        End If
    End With
End Sub

Private Sub b3End_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With b3End
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < b3Start.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(7).Value = .Value
            End If
        End If
    End With
End Sub

Private Sub srEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With srEnd
        If Not IsDate(.Value) And Not .Value = Blank Then
            MsgBox("Date must be in the format 'mm/dd/yyyy'", vbExclamation, "Date Error")
            Cancel = True
            .Value = ""
        Else
            .Value = Format(.Value, "mm/dd/yyyy")
            If .Value < srStart.Value And Not .Value = Blank Then
                MsgBox("Event ends before it begins... please check your dates!", vbExclamation, "Date Error")
                Cancel = True
                .Value = ""
            Else
                 [Venues[EndDate]].Rows(8).Value = .Value
            End If
        End If
    End With
End Sub


Private Sub useERH_Click()

    Dim i As Integer
    Dim pos As Integer


    With useERH
        If .Value = False Then
            erhStartLabel.Enabled = False
            erhEndLabel.Enabled = False
            erhStart.Enabled = False
            erhEnd.Enabled = False
            [Venues].Rows(1).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "ERH" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            erhStartLabel.Enabled = True
            erhEndLabel.Enabled = True
            erhStart.Enabled = True
            erhEnd.Enabled = True
            [Venues].Rows(1).Columns(3).Value = "TRUE"
        For Each row In [erh_Items[Item/Svc]].Rows
                itemSelect.addItem("ERH")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row

            Next

        End If
    End With

End Sub


Private Sub useNT_Click()

    Dim i As Integer

    With useNT
        If .Value = False Then
            ntStartLabel.Enabled = False
            ntEndLabel.Enabled = False
            ntStart.Enabled = False
            ntEnd.Enabled = False
            [Venues].Rows(2).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "NT" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            ntStartLabel.Enabled = True
            ntEndLabel.Enabled = True
            ntStart.Enabled = True
            ntEnd.Enabled = True
            [Venues].Rows(2).Columns(3).Value = "TRUE"
        For Each row In [nt_Items[Item/Svc]].Rows
                itemSelect.addItem("NT")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next

        End If
    End With

End Sub


Private Sub useMH_Click()

    Dim i As Integer

    With useMH
        If .Value = False Then
            mhStartLabel.Enabled = False
            mhEndLabel.Enabled = False
            mhStart.Enabled = False
            mhEnd.Enabled = False
            [Venues].Rows(3).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "MH" Then
                    itemSelect.RemoveItem(i)
                End If
            Next
        Else
            mhStartLabel.Enabled = True
            mhEndLabel.Enabled = True
            mhStart.Enabled = True
            mhEnd.Enabled = True
            [Venues].Rows(3).Columns(3).Value = "TRUE"
        For Each row In [mh_Items[Item/Svc]].Rows
                itemSelect.addItem("MH")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row

            Next
        End If
    End With

End Sub


Private Sub useBW_Click()

    Dim i As Integer

    With useBW
        If .Value = False Then
            bwStartLabel.Enabled = False
            bwEndLabel.Enabled = False
            bwStart.Enabled = False
            bwEnd.Enabled = False
            [Venues].Rows(4).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "BW" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            bwStartLabel.Enabled = True
            bwEndLabel.Enabled = True
            bwStart.Enabled = True
            bwEnd.Enabled = True
            [Venues].Rows(4).Columns(3).Value = "TRUE"
        For Each row In [bw_Items[Item/Svc]].Rows
                itemSelect.addItem("BW")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next

        End If
    End With

End Sub


Private Sub useRS_Click()

    Dim i As Integer

    With useRS
        If .Value = False Then
            rsStartLabel.Enabled = False
            rsEndLabel.Enabled = False
            rsStart.Enabled = False
            rsEnd.Enabled = False
            [Venues].Rows(5).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "RS" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            rsStartLabel.Enabled = True
            rsEndLabel.Enabled = True
            rsStart.Enabled = True
            rsEnd.Enabled = True
            [Venues].Rows(5).Columns(3).Value = "TRUE"
        For Each row In [rs_Items[Item/Svc]].Rows
                itemSelect.addItem("RS")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next
        End If
    End With

End Sub


Private Sub useB2_Click()

    Dim i As Integer

    With useB2
        If .Value = False Then
            b2StartLabel.Enabled = False
            b2EndLabel.Enabled = False
            b2Start.Enabled = False
            b2End.Enabled = False
            [Venues].Rows(6).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "B2" Then
                    itemSelect.RemoveItem(i)
                End If
            Next
        Else
            b2StartLabel.Enabled = True
            b2EndLabel.Enabled = True
            b2Start.Enabled = True
            b2End.Enabled = True
            [Venues].Rows(6).Columns(3).Value = "TRUE"
        For Each row In [b2_Items[Item/Svc]].Rows
                itemSelect.addItem("B2")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next
        End If
    End With

End Sub


Private Sub useB3_Click()

    Dim i As Integer

    With useB3
        If .Value = False Then
            b3StartLabel.Enabled = False
            b3EndLabel.Enabled = False
            b3Start.Enabled = False
            b3End.Enabled = False
            [Venues].Rows(7).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "B3" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            b3StartLabel.Enabled = True
            b3EndLabel.Enabled = True
            b3Start.Enabled = True
            b3End.Enabled = True
            [Venues].Rows(7).Columns(3).Value = "TRUE"
        For Each row In [b3_Items[Item/Svc]].Rows
                itemSelect.addItem("B3")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next

        End If
    End With

End Sub


Private Sub useSR_Click()

    Dim i As Integer

    With useSR
        If .Value = False Then
            srStartLabel.Enabled = False
            srEndLabel.Enabled = False
            srStart.Enabled = False
            srEnd.Enabled = False
            [Venues].Rows(8).Columns(3).Value = "FALSE"
            For i = itemSelect.ListCount - 1 To 0 Step -1
                If itemSelect.Column(0, i) = "SR" Then
                    itemSelect.RemoveItem(i)
                End If
            Next i
        Else
            srStartLabel.Enabled = True
            srEndLabel.Enabled = True
            srStart.Enabled = True
            srEnd.Enabled = True
            [Venues].Rows(8).Columns(3).Value = "TRUE"
        For Each row In [sr_Items[Item/Svc]].Rows
                itemSelect.addItem("SR")
                i = itemSelect.ListCount - 1
                itemSelect.Column(1, i) = row
            Next
        End If
    End With

End Sub

Private Sub compNameField_Change()
    With compNameField
        [OverviewData[Company Name]].Value = .Value
    End With

End Sub
Private Sub eventNameField_Change()
    With eventNameField
        [OverviewData[Event Name]].Value = .Value
    End With

End Sub
Private Sub contact1NameField_Change()
    With contact1NameField
        [OverviewData[Primary Contact Name]].Value = .Value
    End With

End Sub
Private Sub contact1PhoneField_Change()
    With contact1PhoneField
        [OverviewData[Primary Contact Phone]].Value = .Value
    End With

End Sub
Private Sub contact1EmailField_Change()
    With contact1EmailField
        [OverviewData[Primary Contact Email]].Value = .Value
    End With

End Sub
Private Sub contact2NameField_Change()
    With contact2NameField
        [OverviewData[Secondary Contact Name]].Value = .Value
    End With

End Sub
Private Sub contact2PhoneField_Change()
    With contact2PhoneField
        [OverviewData[Secondary Contact Phone]].Value = .Value
    End With

End Sub
Private Sub contact2EmailField_Change()
    With contact2EmailField
        [OverviewData[Secondary Contact Email]].Value = .Value
    End With

End Sub

Private Sub qtyInc_Change()
    qtyField.Value = qtyInc.Value
End Sub

Private Sub ushInc_Change()
    numUsh.Value = ushInc.Value
End Sub

Private Sub vmInc_Change()
    numVM.Value = vmInc.Value
End Sub
Private Sub Website_Change()
    [OverviewData[URL]].Value = Website.Value
End Sub