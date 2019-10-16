VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExtractVBTParas()


    Dim tmpVBComponent As VBComponent
    Dim LineNumCnt As Long
    Dim LineIdx  As Long
    For Each tmpVBComponent In ThisWorkbook.VBProject.VBComponents
    
        
        If InStr(1, tmpVBComponent.Name, "VBT_basic") Then
            'Debug.Print tmpVBComponent.Name
            
            LineNumCnt = tmpVBComponent.CodeModule.CountOfLines
            For i = 1 To LineNumCnt - 1
                tmpStr = tmpVBComponent.CodeModule.Lines(i, 1)
                If InStr(1, tmpStr, "Public Function ") > 0 And InStr(1, tmpStr, "Basic_") > 0 Then
                    tmpStr = Replace(tmpStr, ",", " ")
                    tmpStr = Replace(tmpStr, " _", " ")
                    tmpStr = Replace(tmpStr, "(", " ")
                    tmpStr = Replace(tmpStr, ")", " ")
                    tmpStr = Replace(tmpStr, "  ", " ")
                    TmpStr_Split = Split(tmpStr, " ")
                    On Error Resume Next
                    
                    
                    Debug.Print TmpStr_Split(2)
                    
                    Set WS1 = ThisWorkbook.Worksheets(TmpStr_Split(2))
                    
                    If WS1 Is Nothing Or WS1.Worksheet.Name <> TmpStr_Split(2) Then
                        ThisWorkbook.Worksheets.Add.Name = TmpStr_Split(2)
                        Set WS1 = ThisWorkbook.Worksheets(TmpStr_Split(2))
                    Else
                        WS1.Range("a1:g60").Clear
                    End If
                    WS1.Cells(2, 2).Value = TmpStr_Split(2)
                    WS1.Cells(1, 1).Value = "测试项 "
                    WS1.Cells(1, 2) = "调用VBT对象"
                    WS1.Cells(1, 3) = "函数接口名称"
                    WS1.Cells(1, 4) = "类型"
                    WS1.Cells(1, 5) = "意义"
                    WS1.Cells(1, 6) = "配置建议"
                    WS1.Cells(1, 7) = "备注"

                    LineIdx = 2
                    For k = i To i + 20
                        tmpStr = tmpVBComponent.CodeModule.Lines(k, 1)
                        tmpStr = Replace(tmpStr, ") As", " ")
                        tmpStr = Replace(tmpStr, ",", " ")
                        tmpStr = Replace(tmpStr, " _", " ")
                        tmpStr = Replace(tmpStr, "(", " ")
                        tmpStr = Replace(tmpStr, ")", " ")
                        tmpStr = Replace(tmpStr, "  ", " ")

                        TmpStr_Split = Split(tmpStr, " ")
                        For j = LBound(TmpStr_Split) To UBound(TmpStr_Split)
                            If TmpStr_Split(j) = "As" Then
                                 WS1.Cells(LineIdx, 3) = TmpStr_Split(j - 1)
                                WS1.Cells(LineIdx, 4) = TmpStr_Split(j + 1)
                                If TmpStr_Split(j - 2) = "Optional" Then
                                    WS1.Cells(LineIdx, 6) = TmpStr_Split(j - 2)
                                    If TmpStr_Split(j + 2) = "=" Then
                                        WS1.Cells(LineIdx, 7) = TmpStr_Split(j + 3)
                                    End If
                                End If
                                LineIdx = LineIdx + 1
                            End If
                        Next j
                        tmpStr = tmpVBComponent.CodeModule.Lines(k, 1)
                        If InStr(1, tmpStr, ") As") > 0 Then
                            Exit For
                        End If
                    Next k
                    i = k


                    '''''change format
                    WS1.Select
                    With WS1.Range("A1: G1")
                        .Font.Bold = True '加粗
                        .Size = 11 '字号
                        .HorizontalAlignment = xlCenter '水平居中
                        .VerticalAlignment = xlCenter '垂直居中
                        .Interior.Color = RGB(253, 233, 217) '背景的颜色
                    End With

                    With WS1.Range(Cells(1, 1), Cells((LineIdx - 1), 7)).Borders
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With

                    With WS1.Range(Cells(1, 1), Cells((LineIdx - 1), 7))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                        .Borders(xlEdgeTop).Weight = xlMedium
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeRight).Weight = xlMedium
                        .WrapText = True
                    End With
                    With WS1.Range(Cells(1, 1), Cells(1, 7))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                        .Borders(xlEdgeTop).Weight = xlMedium
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeRight).Weight = xlMedium
                    End With

                    WS1.Range(Cells(2, 1), Cells((LineIdx - 1), 1)).Merge
                    WS1.Range(Cells(2, 2), Cells((LineIdx - 1), 2)).Merge
                    With WS1.Range("A2")
                        .HorizontalAlignment = xlCenter '水平居中
                        .VerticalAlignment = xlCenter '垂直居中
                    End With
                    With WS1.Range("B2")
                        .HorizontalAlignment = xlCenter '水平居中
                        .VerticalAlignment = xlCenter '垂直居中
                    End With

                    WS1.Range(Cells(1, 1), Cells((LineIdx - 1), 1)).ColumnWidth = 25
                    WS1.Range(Cells(1, 2), Cells((LineIdx - 1), 2)).ColumnWidth = 25
                    WS1.Range(Cells(1, 3), Cells((LineIdx - 1), 3)).ColumnWidth = 18
                    WS1.Range(Cells(1, 4), Cells((LineIdx - 1), 4)).ColumnWidth = 14
                    WS1.Range(Cells(1, 5), Cells((LineIdx - 1), 5)).ColumnWidth = 40
                    WS1.Range(Cells(1, 6), Cells((LineIdx - 1), 6)).ColumnWidth = 40
                    WS1.Range(Cells(1, 7), Cells((LineIdx - 1), 7)).ColumnWidth = 12

                End If
            Next i
            
        End If

      Next



End Sub




Sub Sortsheet()
Dim i As Integer, j As Integer, n As Integer
On Error Resume Next
n = Sheets.Count
If n = 1 Then End
For i = 1 To n - 1
  For j = i + 1 To n
    If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
      Sheets(j).Move before:=Sheets(i)
    End If
  Next j
Next i
End Sub


Sub RemoveVBT()


    Dim tmpVBComponent As VBComponent
    Dim LineNumCnt As Long
    Dim LineIdx  As Long
    For Each tmpVBComponent In ThisWorkbook.VBProject.VBComponents
    
        
        If InStr(1, tmpVBComponent.Name, "VBT_basic") Then
        ThisWorkbook.VBProject.VBComponents.Remove tmpVBComponent
        End If
    Next
End Sub

Sub makesuperlink()
    Application.DisplayAlerts = False   'do not display message window
    For i = 1 To ActiveWorkbook.Worksheets.Count
        sheetname = ActiveWorkbook.Worksheets(i).Name
        If sheetname = "Home" Then
            ActiveWorkbook.Worksheets(i).Delete
            Exit For
        End If
    Next i
    Application.DisplayAlerts = True
    
    Call addsheet("Home")
    
    'get the "Home" sheet index
    For i = 1 To ActiveWorkbook.Worksheets.Count
        sheetname = ActiveWorkbook.Worksheets(i).Name
        If sheetname = "Home" Then
            HomeSheetIndex = i
        End If
    Next i
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
    
        If i <> HomeSheetIndex Then
            SheetAddress = ActiveWorkbook.Worksheets(i).Name & "!A1"
        
            With ActiveWorkbook.Worksheets(HomeSheetIndex)
                .Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:= _
                SheetAddress, TextToDisplay:=ActiveWorkbook.Worksheets(i).Name
            End With
        End If
    Next i
End Sub

Sub addsheet(sheetname As String)
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets.Add(before:=ActiveWorkbook.Worksheets(1))
    sh.Name = sheetname
End Sub
