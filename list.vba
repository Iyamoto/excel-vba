

Public Sub buildlists()
Dim p0, p1, p2, s, r, add As Integer
Dim base, full, num As String
'Dim buff() As String

r = 109
p0 = 0
p1 = 0
p2 = 3
For s = 1 To r
ActiveWorkbook.Sheets("b").Copy _
           Before:=ActiveWorkbook.Sheets("b")
           ActiveSheet.Name = s
           ActiveSheet.Cells(4, 25).Value = p0
           ActiveSheet.Cells(4, 26).Value = p1
           ActiveSheet.Cells(4, 27).Value = p2
           p2 = p2 + 1
           If p2 > 9 Then
            p1 = p1 + 1
            p2 = 0
           End If
           If p1 = 10 Then
            p1 = 0
            p0 = p0 + 1
           End If
           

           'add = add + 1
           'num = CStr(add)
           full = ActiveWorkbook.Sheets("data").Cells(s, 1).Value
           For i = 1 To Len(full)
                ActiveSheet.Cells(10, 1 + i).Value = Mid$(full, i, 1)
           Next
                   
           full = ActiveWorkbook.Sheets("data").Cells(s, 2).Value
           'MsgBox full
           For i = 1 To Len(full)
                ActiveSheet.Cells(34, 21 + i).Value = Mid$(full, i, 1)
                ActiveSheet.Cells(66, 21 + i).Value = Mid$(full, i, 1)
           Next
           
           full = ActiveWorkbook.Sheets("data").Cells(s, 3).Value
           For i = 1 To Len(full)
                ActiveSheet.Cells(74, 21 + i).Value = Mid$(full, i, 1)
           Next
           
           full = ActiveWorkbook.Sheets("data").Cells(s, 4).Value
           For i = 1 To Len(full)
                ActiveSheet.Cells(82, 21 + i).Value = Mid$(full, i, 1)
                ActiveSheet.Cells(126, 21 + i).Value = Mid$(full, i, 1)
           Next
           
           full = ActiveWorkbook.Sheets("data").Cells(s, 5).Value * 10000
           If full = 10000 Then
                ActiveSheet.Cells(78, 22).Value = 1
            
           End If
           If full < 10000 Then
           For i = 1 To Len(full)
                ActiveSheet.Cells(78, 23 + i).Value = Mid$(full, i, 1)
           Next
           End If
Next
End Sub


