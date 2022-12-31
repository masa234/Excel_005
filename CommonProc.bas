Public Const DATA_SHEET_NAME = "データ"
Public Const EXCEL_FILE_OUTPUT_FAILED = "Excelファイルの出力に失敗しました。"


'【概要】シートが存在するか
Public Function IsExistsSheet(ByVal objWb As Excel.Workbook, _
                                ByVal strSheetName As String) As Boolean
On Error GoTo IsExistsSheet_Err

    IsExistsSheet = False
    
    Dim objWs As Excel.Worksheet
    
    '引数のシート名でシートオブジェクトを参照する
    'シートが存在しない場合、エラーが発生する
    Set objWs = objWb.Worksheets(strSheetName)
    
    IsExistsSheet = True
    
IsExistsSheet_Err:

IsExistsSheet_Exit:
    Set objWs = Nothing
End Function


'【概要】Excelシートからシートへの転記
Public Function ExcelSheetToExcelSheet(ByVal objPasteWb As Excel.Workbook, _
                            ByVal strPasteSheetName As String, _
                            ByVal objPastedWb As Excel.Workbook, _
                            ByVal strPastedSheetName As String) As Boolean
On Error GoTo ExcelSheetToExcelSheet_Err

    ExcelSheetToExcelSheet = False

    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim lngPasteRow As Long
    Dim lngPastedRow As Long
    
    With objPasteWb.Worksheets(strPasteSheetName)
        '貼り付け元最終行を取得
        lngLastRow = .Cells(1, 1).End(xlDown).Row
        '貼り付け先行初期化
        lngPastedRow = 1
        '貼り付け元最終行まで繰り返す
        For lngPasteRow = 1 To lngLastRow
            '貼り付け元→貼り付け先
            objPastedWb.Worksheets(strPastedSheetName).Cells(lngPastedRow, 1).Value = .Cells(lngPasteRow, 1).Value
            '貼り付け先行をカウントアップ
            lngPastedRow = lngPastedRow + 1
        Next lngPasteRow
    End With
    
    ExcelSheetToExcelSheet = True
    
ExcelSheetToExcelSheet_Err:

ExcelSheetToExcelSheet_Exit:

End Function


'【概要】Excelシートからシートへの転記
Public Function ExcelFilesToExcelSheet(ByVal arrExcelFilePaths As Variant) As Boolean
On Error GoTo ExcelFilesToExcelSheet_Err
    
    ExcelFilesToExcelSheet = False
    
    Dim lngArrIdx As Long
    Dim objPasteWb As Excel.Workbook
    Dim objPastedWb As Excel.Workbook
    Dim objWs As Excel.Worksheet
    
    'Excelファイルパスの数だけ繰り返す
    For lngArrIdx = 0 To UBound(arrExcelFilePaths)
        'ファイルを開く
        Workbooks.Open arrExcelFilePaths(lngArrIdx)
        '貼り付け元ブック
        Set objPasteWb = ActiveWorkbook
        'Excelファイル作成
        Set objPastedWb = Workbooks.Add
        'シート名をDATAにする
        ActiveSheet.Name = DATA_SHEET_NAME
        'Excelシートの数だけ繰り返す
        For Each objWs In objPasteWb.Worksheets
            'Excelシートの転記を開始する
            If ExcelSheetToExcelSheet(objPasteWb, objWs.Name, objPastedWb, DATA_SHEET_NAME) = False Then
                GoTo ExcelFilesToExcelSheet_Exit
            End If
        Next objWs
    Next lngArrIdx
    
    ExcelFilesToExcelSheet = True
    
ExcelFilesToExcelSheet_Err:

ExcelFilesToExcelSheet_Exit:
    Set objPasteWb = Nothing
    Set objPastedWb = Nothing
    Set objWs = Nothing
End Function


'【概要】特定のディレクトリ内の特定の拡張子のファイルパス（複数）を取得する
Public Function GetFilePaths(ByVal strDirectoryPath As String, _
                            ByVal strExtensionName As String) As Variant
On Error GoTo GetFilePaths_Err
    
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    '初期化
    lngArrIdx = 0
    
    With objFso
        'Fsoのフォルダ内のファイルだけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            'ファイル名が自身以外の場合
            If objFile.Name <> ThisWorkbook.Name Then
                'ファイルの拡張子が指定のものだった場合、
                If .GetExtensionName(objFile.Name) = strExtensionName Then
                    '配列再宣言
                    ReDim Preserve arrRet(lngArrIdx)
                    '配列格納
                    arrRet(lngArrIdx) = objFile.Path
                    '配列の要素番号を1つ進める
                    lngArrIdx = lngArrIdx + 1
                End If
            End If
        Next objFile
    End With
    
    GetFilePaths = arrRet
    
GetFilePaths_Err:

GetFilePaths_Exit:
    Set objFso = Nothing
End Function

