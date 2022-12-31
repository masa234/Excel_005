

'【概要】次の連番シート名取得
Public Function GetSheetNameWithSeqNumber(ByVal objWb As Excel.Workbook, _
                            ByVal strBaseSheetName As String) As String
On Error GoTo GetSheetNameWithSeqNumber_Err
    
    Dim lngCount As Long
    Dim strSheetName As String
    
    '100回繰り返す
    For lngCount = 1 To 100
        'シート名設定
        strSheetName = strBaseSheetName & "_" & CStr(lngCount)
        'シートが存在しない場合、終了
        If IsExistsSheet(objWb, strSheetName) = False Then
            GetSheetNameWithSeqNumber = strSheetName
            GoTo GetSheetNameWithSeqNumber_Exit
        End If
    Next lngCount
    
GetSheetNameWithSeqNumber_Err:

GetSheetNameWithSeqNumber_Exit:

End Function


