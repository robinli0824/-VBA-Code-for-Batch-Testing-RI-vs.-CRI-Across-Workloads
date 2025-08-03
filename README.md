# -VBA-Code-for-Batch-Testing-RI-vs.-CRI-Across-Workloads
' Full integration module: RI realistic accumulation strategy &amp; CRI (both insertion and removal use standard CRI formula)
Dim ReshuffleTotal As Long
Dim currentCount As Long

' --------------------------------------------------------------------
' Main entry point for batch testing RI vs. CRI
' --------------------------------------------------------------------
Sub BatchTestRIvsCRI()
    Dim workloads As Variant
    Dim trial As Integer
    Dim i As Integer
    Dim repeatN As Integer
    Dim resultRow As Integer, resultCol As Integer
    Dim ws As Worksheet
    
    ' Define the workload levels and number of repetitions
    workloads = Array(720, 960, 1200, 1440, 1680, 1920, 2160)
    repeatN = 8
    
    Set ws = ActiveSheet
    ws.Activate
    ' Clear previous results
    ws.Range("AP1:AW100").Clear
    
    ' Write header row in English
    ws.Range("AP1:AW1").Value = Array( _
        "NumContainers", "Utilization", "RI_Mean", "RI_Min", "RI_Max", _
        "CRI_Mean", "CRI_Min", "CRI_Max" _
    )
    
    resultRow = 2
    resultCol = 42  ' Column AP
    
    ' Loop over each workload
    For i = LBound(workloads) To UBound(workloads)
        Dim N As Integer
        N = workloads(i)
        
        Dim riArr() As Long, criArr() As Long
        ReDim riArr(1 To repeatN)
        ReDim criArr(1 To repeatN)
        
        ' Repeat experiments
        For trial = 1 To repeatN
            Call GenerateEvents_For_Batch(N)
            ws.Range("L4:AO13").ClearContents
            Call RunAllEvents_DynamicYard_RI_Realistic_10x6x5
            riArr(trial) = ReshuffleTotal
            
            Call GenerateEvents_For_Batch(N)
            ws.Range("L4:AO13").ClearContents
            Call RunAllEvents_DynamicYard_CRI_10x6x5
            criArr(trial) = ReshuffleTotal
        Next trial
        
        ' Compute statistics for RI and CRI
        Dim sumRI As Long, minRI As Long, maxRI As Long
        Dim sumCRI As Long, minCRI As Long, maxCRI As Long
        
        sumRI = 0: minRI = riArr(1): maxRI = riArr(1)
        sumCRI = 0: minCRI = criArr(1): maxCRI = criArr(1)
        
        Dim idx As Integer
        For idx = 1 To repeatN
            sumRI = sumRI + riArr(idx)
            If riArr(idx) < minRI Then minRI = riArr(idx)
            If riArr(idx) > maxRI Then maxRI = riArr(idx)
            
            sumCRI = sumCRI + criArr(idx)
            If criArr(idx) < minCRI Then minCRI = criArr(idx)
            If criArr(idx) > maxCRI Then maxCRI = criArr(idx)
        Next idx
        
        ' Output results
        ws.Cells(resultRow, resultCol).Value = N
        ws.Cells(resultRow, resultCol + 1).Value = N / 1200      ' Utilization relative to total slots
        ws.Cells(resultRow, resultCol + 2).Value = sumRI / repeatN
        ws.Cells(resultRow, resultCol + 3).Value = minRI
        ws.Cells(resultRow, resultCol + 4).Value = maxRI
        ws.Cells(resultRow, resultCol + 5).Value = sumCRI / repeatN
        ws.Cells(resultRow, resultCol + 6).Value = minCRI
        ws.Cells(resultRow, resultCol + 7).Value = maxCRI
        
        resultRow = resultRow + 1
    Next i
    
    MsgBox "Batch testing of RI vs. CRI completed. Results are in AP:AW."
End Sub

' --------------------------------------------------------------------
' Event queue generation for batch experiments
' --------------------------------------------------------------------
Sub GenerateEvents_For_Batch(N As Integer)
    Dim i As Integer
    Dim arrivalTime As Integer, departureTime As Integer
    Dim startRow As Integer: startRow = 4
    
    ' Clear previous event data
    Range("A4:F583").ClearContents
    
    ' Generate container arrival and departure times
    For i = 1 To N
        Cells(startRow + i - 1, 1).Value = i
        arrivalTime = WorksheetFunction.RandBetween(1, 20)
        Cells(startRow + i - 1, 2).Value = arrivalTime
        departureTime = arrivalTime + WorksheetFunction.RandBetween(30, 50)
        Cells(startRow + i - 1, 3).Value = departureTime
    Next i
    
    ' Build and sort event list
    Dim totalEvents As Integer: totalEvents = N * 2
    Dim events() As Variant
    ReDim events(1 To totalEvents, 1 To 3)
    
    Dim idx As Integer: idx = 1
    For i = 1 To N
        ' Arrival event
        events(idx, 1) = Cells(startRow + i - 1, 2).Value
        events(idx, 2) = "A"
        events(idx, 3) = Cells(startRow + i - 1, 1).Value
        idx = idx + 1
        ' Departure event
        events(idx, 1) = Cells(startRow + i - 1, 3).Value
        events(idx, 2) = "D"
        events(idx, 3) = Cells(startRow + i - 1, 1).Value
        idx = idx + 1
    Next i
    
    ' Simple bubble sort by event time
    Dim j As Integer, tmp1, tmp2, tmp3
    For i = 1 To totalEvents - 1
        For j = i + 1 To totalEvents
            If events(i, 1) > events(j, 1) Then
                tmp1 = events(i, 1): tmp2 = events(i, 2): tmp3 = events(i, 3)
                events(i, 1) = events(j, 1): events(i, 2) = events(j, 2): events(i, 3) = events(j, 3)
                events(j, 1) = tmp1: events(j, 2) = tmp2: events(j, 3) = tmp3
            End If
        Next j
    Next i
    
    ' Write sorted events back to sheet
    For i = 1 To totalEvents
        Cells(startRow + i - 1, 4).Value = events(i, 1)
        Cells(startRow + i - 1, 5).Value = events(i, 2)
        Cells(startRow + i - 1, 6).Value = events(i, 3)
    Next i
End Sub

' --------------------------------------------------------------------
' Execute events using RI strategy (realistic stacking)
' --------------------------------------------------------------------
Sub RunAllEvents_DynamicYard_RI_Realistic_10x6x5()
    Dim i As Long, typeAD As String, containerID As Variant
    ReshuffleTotal = 0: currentCount = 0
    For i = 4 To 583
        typeAD = Range("E" & i).Value
        containerID = Range("F" & i).Value
        If typeAD = "A" Then
            Call RI_Insert_Realistic(containerID)
        ElseIf typeAD = "D" Then
            Call RemoveContainerFromYard_FullAuto_Param_RI_Realistic(containerID)
        End If
    Next i
End Sub

' --------------------------------------------------------------------
' Execute events using CRI strategy (insertion/removal by CRI)
' --------------------------------------------------------------------
Sub RunAllEvents_DynamicYard_CRI_40x6x5()
    Dim i As Long, typeAD As String, containerID As Variant
    ReshuffleTotal = 0: currentCount = 0
    For i = 4 To 583
        typeAD = Range("E" & i).Value
        containerID = Range("F" & i).Value
        If typeAD = "A" Then
            Call CRI_Insert_EmptyPriority_WithSelf(containerID)
        ElseIf typeAD = "D" Then
            Call RemoveContainerFromYard_FullAuto_Param_CRI_EmptyPriority(containerID)
        End If
    Next i
End Sub

' --------------------------------------------------------------------
' RI insertion (realistic stacking)
' --------------------------------------------------------------------
Sub RI_Insert_Realistic(containerID As Variant)
    Dim yardStartCol As Integer: yardStartCol = Range("L4").Column
    Dim Lmax As Integer: Lmax = 5
    Dim bestRI As Integer: bestRI = 9999
    Dim bestRow As Integer, bestSlot As Integer, bestLayer As Integer
    Dim maxHeight As Integer: maxHeight = -1
    Dim row As Integer, slot As Integer, layer As Integer, col As Integer
    Dim RI As Integer, curDepTime As Variant, depTime As Variant

    On Error Resume Next
    curDepTime = Application.WorksheetFunction.VLookup(containerID, Range("A4:C583"), 3, False)
    On Error GoTo 0

    For row = 1 To 10
        For slot = 1 To 6
            Dim height As Integer: height = 0
            For layer = 1 To Lmax
                col = yardStartCol + (slot - 1) * 5 + (layer - 1)
                If Cells(row + 3, col).Value <> "" Then
                    height = height + 1
                End If
            Next layer
            If height < Lmax Then
                col = yardStartCol + (slot - 1) * 5 + height
                RI = 0
                For c = yardStartCol + (slot - 1) * 5 To yardStartCol + (slot - 1) * 5 + height - 1
                    If Cells(row + 3, c).Value <> "" Then
                        On Error Resume Next
                        depTime = Application.WorksheetFunction.VLookup(Cells(row + 3, c).Value, Range("A4:C583"), 3, False)
                        On Error GoTo 0
                        If depTime < curDepTime Then RI = RI + 1
                    End If
                Next c
                If (height > maxHeight) Or (height = maxHeight And RI < bestRI) Then
                    maxHeight = height
                    bestRI = RI
                    bestRow = row
                    bestSlot = slot
                    bestLayer = height + 1
                End If
            End If
        Next slot
    Next row

    If maxHeight >= 0 Then
        col = yardStartCol + (bestSlot - 1) * 5 + (bestLayer - 1)
        Cells(bestRow + 3, col).Value = containerID
        currentCount = currentCount + 1
    End If
End Sub

' --------------------------------------------------------------------
' RI removal (realistic stacking)
' --------------------------------------------------------------------
Sub RemoveContainerFromYard_FullAuto_Param_RI_Realistic(containerID As Variant)
    Dim row As Integer, col As Integer, slot As Integer, layer As Integer
    Dim yardStartCol As Integer: yardStartCol = Range("L4").Column
    Dim yardEndCol As Integer: yardEndCol = Range("AO4").Column
    Dim found As Boolean: found = False
    Dim i As Integer

    For row = 4 To 13
        For col = yardStartCol To yardEndCol
            If Cells(row, col).Value = containerID Then
                found = True
                slot = ((col - yardStartCol) \ 5) + 1
                layer = ((col - yardStartCol) Mod 5) + 1
                Exit For
            End If
        Next col
        If found Then Exit For
    Next row
    If Not found Then Exit Sub

    Dim topCol As Integer
    For i = 5 To layer + 1 Step -1
        topCol = yardStartCol + (slot - 1) * 5 + (i - 1)
        If i > layer And Cells(row, topCol).Value <> "" Then
            Dim moveID As Variant: moveID = Cells(row, topCol).Value
            ReshuffleTotal = ReshuffleTotal + 1
            Cells(row, topCol).Value = ""
            currentCount = currentCount - 1
            Call RI_Insert_Realistic(moveID)
        End If
    Next i
    Cells(row, yardStartCol + (slot - 1) * 5 + (layer - 1)).Value = ""
    currentCount = currentCount - 1
End Sub

' --------------------------------------------------------------------
' CRI insertion (standard CRI formula)
' --------------------------------------------------------------------
Sub CRI_Insert_EmptyPriority_WithSelf(containerID As Variant)
    Dim yardStartCol As Integer: yardStartCol = Range("L4").Column
    Dim Lmax As Integer: Lmax = 5
    Dim N As Integer: N = 10
    Dim S As Integer: S = 6

    Dim depDict As Object: Set depDict = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 4 To 583
        If Not depDict.Exists(Cells(i, 1).Value) Then
            depDict(Cells(i, 1).Value) = Cells(i, 3).Value
        End If
    Next i

    Dim curDepTime As Variant: curDepTime = depDict(containerID)

    Dim boxList() As Variant
    Dim boxCount As Long: boxCount = 0
    ReDim boxList(1 To 300, 1 To 4)
    Dim row As Integer, slot As Integer, layer As Integer, col As Integer
    For row = 1 To N
        For slot = 1 To S
            For layer = 1 To Lmax
                col = yardStartCol + (slot - 1) * 5 + (layer - 1)
                If Cells(row + 3, col).Value <> "" Then
                    boxCount = boxCount + 1
                    boxList(boxCount, 1) = row
                    boxList(boxCount, 2) = slot
                    boxList(boxCount, 3) = layer
                    boxList(boxCount, 4) = depDict(Cells(row + 3, col).Value)
                End If
            Next layer
        Next slot
    Next row

    Dim bestCRI As Double: bestCRI = 1E+30
    Dim bestRow As Integer, bestSlot As Integer, bestLayer As Integer

    For row = 1 To N
        For slot = 1 To S
            For layer = 1 To Lmax
                col = yardStartCol + (slot - 1) * 5 + (layer - 1)
                If Cells(row + 3, col).Value = "" Then
                    Dim cri As Double: cri = 0
                    For i = 1 To boxCount
                        Dim boxRow As Integer: boxRow = boxList(i, 1)
                        Dim boxSlot As Integer: boxSlot = boxList(i, 2)
                        Dim boxLayer As Integer: boxLayer = boxList(i, 3)
                        Dim depTime As Variant: depTime = boxList(i, 4)
                        If depTime < curDepTime Then
                            Dim stackHeight As Integer: stackHeight = 0
                            Dim h As Integer
                            For h = 1 To Lmax
                                If Cells(boxRow + 3, yardStartCol + (boxSlot - 1) * 5 + (h - 1)).Value <> "" Then
                                    stackHeight = stackHeight + 1
                                End If
                            Next h
                            Dim omega As Double: omega = 1 + stackHeight / Lmax
                            Dim dist As Integer: dist = Abs(boxRow - row) + Abs(boxSlot - slot)
                            cri = cri + omega / (1 + dist)
                        End If
                    Next i
                    If cri < bestCRI Then
                        bestCRI = cri
                        bestRow = row: bestSlot = slot: bestLayer = layer
                    End If
                End If
            Next layer
        Next slot
    Next row

    If bestCRI < 1E+30 Then
        col = yardStartCol + (bestSlot - 1) * 5 + (bestLayer - 1)
        Cells(bestRow + 3, col).Value = containerID
        currentCount = currentCount + 1
    End If
End Sub

' --------------------------------------------------------------------
' CRI removal (standard CRI formula)
' --------------------------------------------------------------------
Sub RemoveContainerFromYard_FullAuto_Param_CRI_EmptyPriority(containerID As Variant)
    Dim row As Integer, col As Integer, slot As Integer, layer As Integer
    Dim yardStartCol As Integer: yardStartCol = Range("L4").Column
    Dim yardEndCol As Integer: yardEndCol = Range("AO4").Column
    Dim found As Boolean: found = False
    Dim i As Integer

    For row = 4 To 13
        For col = yardStartCol To yardEndCol
            If Cells(row, col).Value = containerID Then
                found = True
                slot = ((col - yardStartCol) \ 5) + 1
                layer = ((col - yardStartCol) Mod 5) + 1
                Exit For
            End If
        Next col
        If found Then Exit For
    Next row
    If Not found Then Exit Sub

    Dim depDict As Object: Set depDict = CreateObject("Scripting.Dictionary")
    For i = 4 To 583
        If Not depDict.Exists(Cells(i, 1).Value) Then
            depDict(Cells(i, 1).Value) = Cells(i, 3).Value
        End If
    Next i

    Dim topCol As Integer
    For i = 5 To layer + 1 Step -1
        topCol = yardStartCol + (slot - 1) * 5 + (i - 1)
        If i > layer And Cells(row, topCol).Value <> "" Then
            Dim moveID As Variant: moveID = Cells(row, topCol).Value
            ReshuffleTotal = ReshuffleTotal + 1
            Cells(row, topCol).Value = ""
            currentCount = currentCount - 1
            Call CRI_Insert_EmptyPriority_WithSelf(moveID)
        End If
    Next i
    Cells(row, yardStartCol + (slot - 1) * 5 + (layer - 1)).Value = ""
    currentCount = currentCount - 1
End Sub

