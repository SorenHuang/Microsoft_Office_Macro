Attribute VB_Name = "Module1"
Sub 複製資料炫轉90度()

'
' 複製資料炫轉90度
'
'
    Sheets(1).Select
    Range("A1").Select
    Selection.End(xlToRight).Select
    Dim ColRang As Range
    Set ColRang = Selection
    Dim colNo As Integer
    colNo = ColRang.Column
    
    Range("A1").Select
    Selection.End(xlDown).Select
    Dim RowRang As Range
    Set RowRang = Selection
    Dim rowNo As Integer
    rowNo = RowRang.Row
    
  
    
    ReDim arr(rowNo, colNo)
    For CIndex = 0 To colNo - 1
       
       For RIndex = 0 To rowNo - 1

          Dim getrow, getcol As Integer
          
          getrow = rowNo - RIndex
          getcol = CIndex + 1
          
          
          
          arr(RIndex, CIndex) = Cells(getrow, getcol).Value
       
        
        
        Next
    
    Next
    
  
     
    
    
      Sheets.Add After:=ActiveSheet
      
    
    
    For RIndex = 0 To rowNo - 1
       For CIndex = 0 To colNo - 1
         
          getrow = RIndex + 1
          getcol = CIndex + 1
        
          Cells(getcol, getrow).Value = arr(RIndex, CIndex)

       Next
    
    Next
    
    Cells.Select
    Cells.EntireColumn.AutoFit
     
      MsgBox "End Col : " & colNo & "  End Row : " & rowNo
    ' Sheets(1).Select
    ' Sheets(Sheets.Count).Select
    ' Selection.Copy
    ' ActiveSheet.Paste

    
End Sub

