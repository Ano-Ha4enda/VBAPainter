Attribute VB_Name = "paint"

Option Explicit
Const px = 1.5

' excel���ᎆ�����
Public Sub adjustCellSize()
    With Cells
        .ColumnWidth = px * 0.1
        .RowHeight = px
    End With
End Sub

' RGB���ꂼ��̒l�̓񎟌��z���ǂݍ��݁Aexcel���ᎆ�ɏo�͂���
Public Sub importRGBTextToBGColor()
    '// add declarations
    On Error GoTo catchError
    Focus = True
    
    Dim first As Boolean: first = True ' Foreach�̏���̂�true�ɂȂ�
    Dim strLine As String ' .txt��1�s���ǂݍ��񂾂Ƃ��̕ϐ�
    Dim arrLine() As String ' strLine��z��ɕϊ�����
    Dim map() As Long ' RGB�̒l������2�����z��
    Dim row As Long ' �s���J�E���^
    Dim col As Long ' �񐔃J�E���^

    ' RGB���ꂼ��ŌJ��Ԃ�
    Dim rgb() As String: rgb = Split("Red,Green,Blue", ",")
    Dim color As Long
    For color = 0 To UBound(rgb)
        ' �s�����Z�b�g
        row = 0
        
        ' RGB�̔��f
        Dim colorPos As Long: colorPos = 256 ^ color

        ' �ǂݍ��ރt�@�C����I��
        Dim openFileName As Variant
        openFileName = Application.GetOpenFilename("colormap, *.txt;*.csv", Title:=rgb(color))
        
        ' �P�s���ǂݍ���
        Dim n As Long
        n = FreeFile
        Open openFileName For Input As #n
        Do Until EOF(n)
            Line Input #n, strLine
            arrLine = Split(strLine, ",")

            ' �ŏ��ɓǂݍ��񂾍s�񐔂�2�����z����쐬
            If first Then
                ' �s���擾
                Dim lastRow As Long
                With CreateObject("Scripting.FileSystemObject")
                    lastRow = .OpenTextFile(openFileName, 8).Line - 1
                End With

                ReDim map(lastRow - 1, UBound(arrLine))
                ' �ȍ~�͔z��̒����𑀍삵�Ȃ�
                first = False
            End If
            
            ' ���
            For col = 0 To UBound(arrLine)
                map(row, col) = map(row, col) + CSng(arrLine(col)) * colorPos
            Next col

            ' �s���J�E���^�����Z
            row = row + 1
        Loop
        
        ' text�t�@�C�������
        Close #n
    Next color

    ' map�����ɃZ���̐F��ς���BForeach����value�����n���Ȃ��̂Ō��n�I��for���ŉ�
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

' �}�N���������Ȃ邨�܂��Ȃ�
Property Let Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

