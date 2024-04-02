Function uniqueList() As Object

    '중복 제거된 학년-반을 리스트로 만들기
    
    Dim MyList As Object
    Set MyList = CreateObject("System.Collections.ArrayList")


    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("RR1").Select
    ActiveSheet.Paste
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    
    Range("RR2").Select
    
    Dim cell As Range
    For Each cell In Range(Selection, Selection.End(xlDown))
        MyList.Add cell.Value
    Next cell
    
    
    Range("RR1").Select
    Range(Selection, Selection.End(xlDown)).Delete
    
    Set uniqueList = MyList
    
End Function

Sub InsertRow()

    Cells.Find(What:="1-1", LookIn:=xlValues, LookAt:=xlWhole).Activate
        
    Dim MyList As Object
    Set MyList = uniqueList()
        
    Dim item As Variant
    For Each item In MyList
        Cells.Find(What:=item, LookIn:=xlValues, LookAt:=xlWhole).Select
        Rows(Selection.Row).Select
        Selection.Insert Shift:=xlDown
    Next item
    
End Sub



