Option Explicit
Option Private Module

'#VBA100本ノック 66本目
'ブック自身のあるフォルダ以下の全サブフォルダを検索し、
'自身と同一名称（拡張子含めて）のファイルを探してください。
'同一名称のファイルが見つかったら､シートに出力してください｡
'・A列：フルパス
'・B列：更新日時
'・C列：ファイルサイズ
'※シートは任意

Private FSO  As New Scripting.FileSystemObject   '参照設定:Microsoft Scripting Runtime

'シートの列
Private Enum E_Col_Find
    Data_S = 1
        フルパス = E_Col_Find.Data_S
        更新日時
        サイズ
    Data_E = E_Col_Find.サイズ
End Enum

'検索結果用
Private FileData(E_Col_Find.Data_S To E_Col_Find.Data_E)

'UNICODE版かANCI版か判定
Private Const Unicode   As Boolean = True

'ファイル検索用API
Private Const INVALID_HANDLE_VALUE = -1

Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long

'世界協定時刻 (UTC) に基づくファイル時刻を、ローカルのファイル時刻へ変換
Private Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

'64 ビット形式のファイル時刻を、システム日時形式へ変換
Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'構造体：ファイル日時
'1601年1月1日から 100 ナノ秒間隔の数を表す２つの 32 ビットメンバを持つ 64 ビットの値
Private Type FILETIME
    dwLowDateTime       As Long 'ファイル時刻(下位32ビット)
    dwHighDateTime      As Long 'ファイル時刻(上位32ビット)
End Type

Private Type SYSTEMTIME
    Year        As Integer
    Month       As Integer
    DayOfWeek   As Integer
    Day         As Integer
    Hour    As Integer
    Minute  As Integer
    Second  As Integer
    Milliseconds As Integer
End Type

Private Const MAX_PATH = 260
Private Const MAX_PATH_W = MAX_PATH * 2 - 1 'Unicodeの最大パス
Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long         'ファイル属性
    ftCreationTime      As FILETIME     '作成日
    ftLastAccessTime    As FILETIME     '最終アクセス日
    ftLastWriteTime     As FILETIME     '更新日
    nFileSizeHigh   As Long             'ファイルサイズ(上位32ビット)
    nFileSizeLow    As Long             'ファイルサイズ(下位32ビット)
    dwReserved0     As Long             'リパースタグ
    dwReserved1     As Long             '予約
    cFileName(MAX_PATH_W)   As Byte     'ファイル名
    cAlternate(14 * 2 - 1)  As Byte     'ファイル名(8.3Ver)
End Type
Private T_FileData  As WIN32_FIND_DATA

Private Type TP_File
    Path         As String   'ファイルパス
    FolderPath   As String   'フォルダパス
    Name         As String   'ファイル名
    Extension    As String   'ファイル拡張子
    
    Size_RawB    As Long     '実際のサイズ
    Size_RawKB   As Long     '実際のサイズ
    Size_DiscB   As Long     'ディスク上のサイズ
    Size_DiscKB  As Long     'ディスク上のサイズ
    
    Attributes       As VbFileAttribute  'ファイル属性
    CreationTime     As Variant  '作成日時
    LastWriteTime    As Variant  '更新日時
    LastAccessTime   As Variant  'アクセス日時
End Type

Private Function f100_064_by_WinAPI()
    
    Dim ActSht  As Worksheet
    Set ActSht = ActiveSheet
    With ActSht
        
        .Cells.ClearContents
        
        Dim Path_Folder As String
        Path_Folder = ThisWorkbook.Path
        
        Dim Dic_File    As Scripting.Dictionary
        Set Dic_File = New Scripting.Dictionary
        
        'ヘッダーを格納
        FileData(E_Col_Find.フルパス) = "フルパス"
        FileData(E_Col_Find.サイズ) = "サイズ"
        FileData(E_Col_Find.更新日時) = "更新日時"
        Dic_File.Item("") = FileData
        
        'ファイルを検索
        Call prSerch(Path_Folder, ThisWorkbook.Name, Dic_File)
        
        'データを2次元配列化
        Dim DataAry As Variant
        DataAry = fArray_Dim11_to_Dim2(Dic_File.Items)
        
        'シートに貼り付ける
        Call fArray_Paste(.Cells(1, 1), DataAry)
        
        '完了メッセージ
        If Dic_File.Count = 1 Then
            Call MsgBox("同名のファイルは見つかりませんでした") '自分自身を含めてるから,あり得ないけど一応...
        Else
            Call MsgBox("同名のファイルが見つかりました" & vbCrLf & Dic_File.Count - 1 & "個")
        End If
        
    End With
    
End Function

Private Function fArray_DimCount(DataAry As Variant) As Long
'配列の次元数を取得
    
    On Error GoTo Terminate
    Dim CntDim  As Long
    For CntDim = 1 To 100
        Dim AryLen       As Long
        AryLen = UBound(DataAry, CntDim) - LBound(DataAry, CntDim) + 1
        If AryLen = 0 Then Exit For
    Next
Terminate:
    fArray_DimCount = CntDim - 1
    
End Function

Private Function fArray_Paste(Cell As Range, DataAry As Variant)
'指定セルにデータ配列を貼り付ける(簡易版)
    
    Dim Row_Cnt As Long
    Dim Col_Cnt As Long
    
    '配列の次元数で分岐
    Dim CntDim          As Long
    CntDim = fArray_DimCount(DataAry)
    Select Case CntDim
    Case 1
        Row_Cnt = 1
        Col_Cnt = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
    
    Case 2
        Row_Cnt = UBound(DataAry, 1) - LBound(DataAry, 1) + 1
        Col_Cnt = UBound(DataAry, 2) - LBound(DataAry, 2) + 1
        
    Case Else
        Exit Function
        
    End Select
    
    '配列の貼付け
    Dim Area    As Range
    Set Area = Cell.Resize(Row_Cnt, Col_Cnt)
    Area.Value = DataAry
    
End Function

Private Function fArray_Dim11_to_Dim2(DataArys As Variant) As Variant
'1次元x1次元配列を2次元配列に変換する(簡易版)
'使用例：DictionaryのItemにデータ(1次元配列)を溜めておき、2次元配列に変換する

    '1次元配列のみ対象
    Dim CntDim          As Long
    CntDim = fArray_DimCount(DataArys)
    If CntDim <> 1 Then Exit Function
    
    '最初の1次元配列を取得
    Dim Dim1Ary As Variant
    Dim1Ary = DataArys(LBound(DataArys, 1))
    CntDim = fArray_DimCount(Dim1Ary)
    If CntDim <> 1 Then Exit Function
    
    '2次元配列を準備
    Dim Dim2Ary As Variant
    ReDim Dim2Ary(LBound(DataArys, 1) To UBound(DataArys, 1), LBound(Dim1Ary, 1) To UBound(Dim1Ary, 1))
    
    '各1次元配列のデータを2次元配列に格納
    Dim i   As Long
    For i = LBound(DataArys, 1) To UBound(DataArys, 1)
        
        Dim1Ary = DataArys(i)
        
        Dim j   As Long
        For j = LBound(Dim1Ary, 1) To UBound(Dim1Ary, 1)
            Dim2Ary(i, j) = Dim1Ary(j)
        Next
        
    Next
    
    fArray_Dim11_to_Dim2 = Dim2Ary
    
End Function

Private Function prSerch(ByVal Path_Folder As String, FindName As String, _
                         ByRef Dic_File As Scripting.Dictionary)
    
    If Dic_File Is Nothing Then Set Dic_File = New Scripting.Dictionary
    
    '末尾のセパレータを補完
    Path_Folder = fPath_RightSeparatorFix(Path_Folder)
    
    'サブフォルダ格納辞書を生成
    Dim Dic_Sub As Scripting.Dictionary
    Set Dic_Sub = New Scripting.Dictionary
    
    'ファイル検索
    Dim hFile   As LongPtr
    hFile = prFindFirstFile(Path_Folder, T_FileData)
    If hFile <> 0 Then
        
        Do
            'ファイル名を調整
            Dim Name    As String
            Name = T_FileData.cFileName
            Name = Left$(Name, InStr(Name, vbNullChar) - 1)
            
            If fPath_IsUsable(Name) = True Then
                
                Dim Path_FileOrFolder   As String
                Path_FileOrFolder = Path_Folder & Name
                
                'サブフォルダパスを格納
                If T_FileData.dwFileAttributes And vbDirectory Then
                    Dic_Sub.Item(Dic_Sub.Count) = Path_FileOrFolder
                    
                'ファイルパスを格納
                Else
                    
                    'ファイルの情報を取得
                    Dim T_File  As TP_File
                    T_File = prFileFindData_to_FileObject(Path_FileOrFolder, T_FileData)
                    
                    'ファイルの情報を取得
                    If LCase(T_File.Name) = LCase(FindName) Then
                        FileData(E_Col_Find.フルパス) = T_File.Path
                        FileData(E_Col_Find.サイズ) = T_File.Size_RawKB
                        FileData(E_Col_Find.更新日時) = T_File.LastWriteTime
                        Dic_File.Item(T_File.Path) = FileData
                    End If
                    
                End If
                
            End If
            
            '次のファイルを検索
            Dim hNext   As Long
            hNext = FindNextFile(hFile, T_FileData)
            If hNext = 0 Then Exit Do
            
        Loop
    End If
    Call FindClose(hFile)
    
    'サブフォルダに再帰処理
    Dim LoopFolder  As Variant
    For Each LoopFolder In Dic_Sub.Items
        Call prSerch(CStr(LoopFolder), FindName, Dic_File)
    Next
    Set Dic_Sub = Nothing
    
End Function

Private Function fPath_RightSeparatorFix(Path As String) As String
'パス末尾にセパレータを付ける
    
    If Path = "" Then Exit Function
    
    '末尾のセパレータを一旦除去
    Dim Ret As String
    Ret = fPath_RightSeparatorChop(Path)
    If Ret = "" Then Exit Function
    
    '末尾に１つだけセパレータを付加して返す
    fPath_RightSeparatorFix = Ret & Application.PathSeparator
    
End Function

Private Function fPath_RightSeparatorChop(Path As String) As String
'パス末尾のセパレータを除去する
    
    If Path = "" Then Exit Function
    
    Dim PS  As String
    PS = Application.PathSeparator
    
    Dim Ret As String
    Ret = Path
    Do
        If Right$(Ret, 1) <> PS Then Exit Do
        Ret = Left$(Ret, Len(Ret) - 1)
    Loop
    
    fPath_RightSeparatorChop = Ret
    
End Function

Private Function fPath_IsUsable(Name As String) As Boolean
'※Dir等で列挙された名前で、必要かどうかを判定
    
    Dim Flg_Use As Boolean
    Select Case Name
    Case "", ".", ".."
    Case "Thumbs.db"    'Explorer用のファイル
    Case "desktop.ini"  'Explorer用のファイル
    Case Else
        Flg_Use = True
    End Select
    If Flg_Use = False Then Exit Function
    
    'Officeの制御用ファイル
    If Left$(Name, 1) = "~" Then Exit Function
    
    fPath_IsUsable = True
    
End Function

Private Function prFindFirstFile(Path_Folder As String, ByRef T_FileData As WIN32_FIND_DATA) As LongPtr
    
    Dim Path_FD As String
    Dim hFile   As LongPtr
    
    'Unicode版用の調整
    If Unicode = False Then
        
        Path_FD = Path_Folder
        
        hFile = FindFirstFile(Path_Folder & "*", T_FileData)
        
    Else
        
        Path_FD = prPath_Convert(Path_Folder)
        
        'ファイル情報を取得
        hFile = FindFirstFile(StrPtr(Path_FD & "*"), T_FileData)
        
    End If
    
    If hFile = 0 Then Exit Function
    
    prFindFirstFile = hFile
    
End Function

Private Function prPath_Convert(Path As String) As String
    
    Dim Ret As String
    Ret = Path
    
    If Unicode = True Then
        If fString_Left_With(Path, "\\") = True Then
            Ret = "\\?\UNC" & Mid$(Path, 2)
        Else
            Ret = "\\?\" & Path
        End If
    End If
    
    prPath_Convert = Ret
    
End Function

Private Function fString_Left_With(Value As String, Find As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    
    If Value = "" Or Find = "" Then Exit Function
    
    If InStr(1, Value, Find, Compare) = 1 Then
        fString_Left_With = True
    End If
    
End Function

Private Function prFileFindData_to_FileObject(Path As String, T_Data As WIN32_FIND_DATA) As TP_File
    
    Dim T_File  As TP_File
    With T_File
        
        'ファイルパス,ファイル名
        .Path = Path
        .Name = FSO.GetFileName(Path)
        .Extension = FSO.GetExtensionName(.Name)
        
        'ファイル属性
        .Attributes = T_Data.dwFileAttributes
        
        'ファイルサイズ
        Dim Size1   As Long
        Size1 = T_Data.nFileSizeLow
        
        .Size_RawB = Size1
        .Size_RawKB = fMath_Round_Up(.Size_RawB / 1024, 0)
        .Size_DiscB = prFileSizeRound(.Size_RawB)
        .Size_DiscKB = .Size_DiscB / 1024
        
        '日時
        .CreationTime = prFileTime_to_SystemTime(T_Data.ftCreationTime)
        .LastAccessTime = prFileTime_to_SystemTime(T_Data.ftLastAccessTime)
        .LastWriteTime = prFileTime_to_SystemTime(T_Data.ftLastWriteTime)
        
    End With
    prFileFindData_to_FileObject = T_File
    
End Function

Private Function fMath_Round_Up(Num As Double, Optional Digits As Long = 0) As Double
    
    '簡易版
    fMath_Round_Up = Application.WorksheetFunction.RoundUp(Num, Digits)
    
End Function

Private Function prFileSizeRound(Size As Long) As Long
    
    'クラスターサイズ
    Const myClusterSize As Long = 1024 * 4
    
    'ディスク上のサイズに調整
    If Size = 0 Then Exit Function
    Do
        Dim i   As Long
        i = i + 1
        Dim Max As Long
        Max = myClusterSize * i
        If Size <= Max Then Exit Do
    Loop
    
    prFileSizeRound = Max
    
End Function

Private Function prFileTime_to_SystemTime(T_FileTime As FILETIME) As Variant
    
    'ローカル日時に変換
    Dim T_LocalTime As FILETIME
    If FileTimeToLocalFileTime(T_FileTime, T_LocalTime) = 0 Then Exit Function
    
    'システム日時に変換
    Dim T_SystemTime As SYSTEMTIME
    If FileTimeToSystemTime(T_LocalTime, T_SystemTime) = 0 Then Exit Function
    
    'Date型に変換
    Dim Date_System As Date
    With T_SystemTime
        Date_System = CDate(DateSerial(.Year, .Month, .Day) & " " & TimeSerial(.Hour, .Minute, .Second))
    End With
    
    '処理可能な日時でなければ終了
    Dim Date_Empty  As Date
    If Date_System < Date_Empty Then Exit Function
    
    prFileTime_to_SystemTime = Date_System
    
End Function
