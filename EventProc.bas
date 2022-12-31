
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    Dim arrExcelFilePaths() As Variant
    
    'Excelファイルパスを取得
    arrExcelFilePaths() = GetFilePaths(ThisWorkbook.Path, "xlsm")
    
    'Excelシートとして展開
    If ExcelFilesToExcelSheet(arrExcelFilePaths) = False Then
        GoTo 正方形長方形1_Click_Exit
    End If
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
End Sub
