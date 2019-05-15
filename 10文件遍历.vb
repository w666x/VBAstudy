Sub narr_search()
'模糊查找的代码
fltr_file = ActiveCell.Value()
I = 0
If fltr_file = "" Then MsgBox "空值"
fltr_sheet = UCase(Split(fltr_file, ".")(0))
If InStr(fltr_sheet, "_") Then fltr_sheet = Split(fltr_sheet, "_")(1)
fltr_db = UCase(Split(fltr_file, ".")(1))
test1 = ThisWorkbook.Path & "\物理设计表\" & "*.xls?"
myfile = Dir(test1)
    Do While Len(myfile) > 0
       ' MsgBox myFile     '只输出文件名,而不显示路径,用处很大.
        I = I + 1
        If InStr(myfile, fltr_sheet) Then
            dest_file = ThisWorkbook.Path & "\物理设计表\" & myfile
            Workbooks.Open Filename:=dest_file
            'MsgBox "test1"
            For I = 1 To Workbooks(2).Worksheets.Count
               'Workbooks(2).Activate
               Set sh2 = ActiveWorkbook.Worksheets(I)
               sh2.Activate
               n = sh2.Name
               db_name1 = "表编码:  " & fltr_db
               Set Rng = Cells.Find(db_name1)
               If Not Rng Is Nothing Then
                    'On Error Resume Next
                    Cells.Find(what:=db_name1, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , MatchByte:=False, SearchFormat:=False).Activate
                   'Selection.End(xlToLeft).Select
                   Exit Do
                   Workbooks.Close
               End If
            Next
         End If
         'MsgBox "不在 " & myfile & " 表中"
         myfile = Dir           '只需对Dir进行循环操作即可,因为上面已经有一行查找命令了.
     Loop
     If I > 22 Then
     MsgBox "异常共找到" & I & "个文件，从msp中查看！"
     'myfile = Dir(test1)
     'Do While Len(myfile) > 0
       dest_file = ThisWorkbook.Path & "\物理设计表\" & "MSP_物理设计说明书V1.0.xls"
       'MsgBox dest_file
       Workbooks.Open Filename:=dest_file
       For I = 1 To Workbooks(2).Worksheets.Count
          'Workbooks(2).Activate
          Set sh2 = ActiveWorkbook.Worksheets(I)
          sh2.Activate
          n = sh2.Name
          db_name1 = "表编码:  " & fltr_db
          Set Rng = Cells.Find(db_name1)
          If Not Rng Is Nothing Then
               'On Error Resume Next
               Cells.Find(what:=db_name1, LookIn:=xlFormulas, LookAt:= _
               xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
               , MatchByte:=False, SearchFormat:=False).Activate
              'Selection.End(xlToLeft).Select
              Exit For
              Workbooks.Close
          End If
       Next
    'myfile = Dir
    'Loop
    'MsgBox "不在 " & myfile & " 表中"
     Else: MsgBox "第" & I & "个文件查到"
     End If
End Sub
