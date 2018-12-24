Attribute VB_Name = "paint"

Option Explicit
Const px = 1.5

' excel方眼紙を作る
Public Sub adjustCellSize()
    With Cells
        .ColumnWidth = px * 0.1
        .RowHeight = px
    End With
End Sub

' RGBそれぞれの値の二次元配列を読み込み、excel方眼紙に出力する
Public Sub importRGBTextToBGColor()
    '// add declarations
    On Error GoTo catchError
    Focus = True
    
    Dim first As Boolean: first = True ' Foreachの初回のみtrueになる
    Dim strLine As String ' .txtを1行ずつ読み込んだときの変数
    Dim arrLine() As String ' strLineを配列に変換する
    Dim map() As Long ' RGBの値を入れる2次元配列
    Dim row As Long ' 行数カウンタ
    Dim col As Long ' 列数カウンタ

    ' RGBそれぞれで繰り返す
    Dim rgb() As String: rgb = Split("Red,Green,Blue", ",")
    Dim color As Long
    For color = 0 To UBound(rgb)
        ' 行数リセット
        row = 0
        
        ' RGBの判断
        Dim colorPos As Long: colorPos = 256 ^ color

        ' 読み込むファイルを選択
        Dim openFileName As Variant
        openFileName = Application.GetOpenFilename("colormap, *.txt;*.csv", Title:=rgb(color))
        
        ' １行ずつ読み込む
        Dim n As Long
        n = FreeFile
        Open openFileName For Input As #n
        Do Until EOF(n)
            Line Input #n, strLine
            arrLine = Split(strLine, ",")

            ' 最初に読み込んだ行列数で2次元配列を作成
            If first Then
                ' 行数取得
                Dim lastRow As Long
                With CreateObject("Scripting.FileSystemObject")
                    lastRow = .OpenTextFile(openFileName, 8).Line - 1
                End With

                ReDim map(lastRow - 1, UBound(arrLine))
                ' 以降は配列の長さを操作しない
                first = False
            End If
            
            ' 代入
            For col = 0 To UBound(arrLine)
                map(row, col) = map(row, col) + CSng(arrLine(col)) * colorPos
            Next col

            ' 行数カウンタを加算
            row = row + 1
        Loop
        
        ' textファイルを閉じる
        Close #n
    Next color

    ' mapを元にセルの色を変える。Foreachだとvalueしか渡せないので原始的なfor文で回す
    For row = 0 To UBound(map, 1)
        For col = 0 To UBound(map, 2)
            Cells(row + 1, col + 1).Interior.color = "&H" & Hex$(map(row, col))
        Next col
    Next row

exitSub:
    Focus = False
    Exit Sub
    
catchError:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCrLf & "Detail: " & Err.Description
    End If
    Err.Clear
    GoTo exitSub
End Sub

' マクロが早くなるおまじない
Property Let Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

