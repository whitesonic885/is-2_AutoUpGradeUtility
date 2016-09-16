'
'   ファイル・コピーの為の汎用クラス
'   ファイルシェアやhttpでのコピー
Imports System.IO
Imports System.Net

Public Class FileCopy

    Private _ErrorMessage As String
    Private _PathName As String
    Private _FileName As String
    Private _TimeOut As Integer     'httpsは5分
    Private _TimeOut2 As Integer    'httpは30秒

    Private constDirSeparator As String = Path.DirectorySeparatorChar.ToString()
    Private constVolumeSeparator As String = Path.VolumeSeparatorChar.ToString()

#Region "ログ出力"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    '
    'コンストラクタ
    '
    Public Sub New()
        MyBase.New()
        '変数の初期化
        _ErrorMessage = ""
        _PathName = ""
        _FileName = ""
        'MOD-S 20050829 タイムアウト３０秒→５分
        '_TimeOut = 30000
        _TimeOut = 300000
        _TimeOut2 = 30000
        'MOD-E 20050829 タイムアウト３０秒→５分
        With _myLogRecord
            .Status = ""
            .SubStatus = "ファイルチェック"
        End With
    End Sub

    Public Sub New(ByVal TimeOut As Integer)
        Me.New()
        _TimeOut = TimeOut

    End Sub

    '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 START
    '//Public ReadOnly Property ErrorMesage()
    Public ReadOnly Property ErrorMesage() As String
        '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 END
        Get
            Return _ErrorMessage
        End Get
    End Property

    '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 START
    '//Public ReadOnly Property PathName()
    Public ReadOnly Property PathName() As String
        '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 END
        Get
            Return _PathName
        End Get
    End Property

    '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 START
    '//Public ReadOnly Property FileName()
    Public ReadOnly Property FileName() As String
        '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 END
        Get
            Return _FileName
        End Get
    End Property

    '   指定されたファイルを指定ディレクトリにコピーする
    '       指定ディレクトリが指定されていなければ、ファイル名のチェックのみ
    '       外部から呼出すエントリーポイントになります
    Protected Friend Function CopyFile(ByVal sourceFileName As String, ByVal destPathName As String) As Boolean
        Dim blnRet As Boolean
        _ErrorMessage = ""
        _PathName = ""
        _FileName = ""

        With _myLogRecord
            .Target = sourceFileName & "を" & destPathName & "へコピーを開始しました"
            .Result = "OK"
            .Remark = "コピー処理の開始"
        End With
        _myLog.WriteLog(_myLogRecord)
        'ファイル名の指定方法のチェック
        blnRet = CheckFileName(sourceFileName)
        If blnRet = False Then Return False

        If destPathName = String.Empty Then Return blnRet
        'ファイルのコピー
        blnRet = GetFileToCopy(sourceFileName, destPathName)
        With _myLogRecord
            If blnRet Then
                .Target = sourceFileName & "を" & destPathName & "へコピー完了しました"
                .Result = "OK"
                .Remark = "コピー処理の終了"
            Else
                .Target = sourceFileName & "を" & destPathName & "へのコピーに失敗しました"
                .Result = "NG"
                .Remark = "コピー処理の終了"
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return blnRet

    End Function

    '
    '   ファイル名が以下の規則に従っているかをチェックにします
    '       パスとファイル名をプロパティに設定します
    '       c:\のドライブレター方式、\\savernameの共有方式、http:のURL方式
    Private Function CheckFileName(ByVal fullPathFileName As String) As Boolean
        If fullPathFileName = String.Empty Then
            _ErrorMessage = "ファイル名が指定されていません"
            With _myLogRecord
                .Target = fullPathFileName & _ErrorMessage
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If
        'パスの開始文字とファイルの終了文字をチェック
        If CheckPathStart(fullPathFileName) = False Then Return False
        If CheckPathEnd(fullPathFileName) = False Then Return False
        'ファイル名を取得
        Dim FileName As String
        FileName = GetFileName(fullPathFileName)
        If FileName = String.Empty Then Return False
        '分離したパスとファイル名をプロパティに設定
        _FileName = FileName
        _PathName = fullPathFileName.Substring(0, fullPathFileName.Length - FileName.Length)
        With _myLogRecord
            .Target = "FileName=" & _FileName & " ,Dir=" & _PathName
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True

    End Function

    '
    '   パス形式が正しい文字列で始まっているかをチェック
    '       http://、https://、\\、x:\、x:、\、.
    '
    Private Function CheckPathStart(ByVal fullPathFileName As String) As Boolean
        If CheckPathIsUrl(fullPathFileName) Then Return True
        If fullPathFileName.StartsWith(String.Concat(constDirSeparator, constDirSeparator)) Then Return True
        If fullPathFileName.StartsWith(constDirSeparator) Then Return True
        If fullPathFileName.StartsWith(".") Then Return True
        If fullPathFileName.Substring(1, 2) = String.Concat(constVolumeSeparator, constDirSeparator) Then Return True
        If fullPathFileName.Substring(1, 1) = constVolumeSeparator Then Return True

        _ErrorMessage = "正しいパス文字列で始まっていません"
        With _myLogRecord
            .Target = fullPathFileName & "が" & _ErrorMessage
            .Result = "NG"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return False

    End Function

    '
    '   パス文字列が/や\や.で終了していないかをチェック
    '
    Private Function CheckPathEnd(ByVal fullPathFileName As String) As Boolean
        If fullPathFileName.EndsWith("/") = False Then
            If fullPathFileName.EndsWith(constDirSeparator) = False Then
                If fullPathFileName.EndsWith(".") = False Then Return True
            End If
        End If
        _ErrorMessage = "ファイル名が指定されていません"
        With _myLogRecord
            .Target = fullPathFileName & "が" & _ErrorMessage
            .Result = "NG"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        Return False

    End Function

    '
    '   パスがURL形式で始まるかどうかをチェック
    '       http://又はhttps://
    Protected Friend Function CheckPathIsUrl(ByVal fullPathFileName As String) As Boolean
        Dim FileName As String
        FileName = fullPathFileName.ToLower()
        If FileName.StartsWith("http://") Then Return True
        If FileName.StartsWith("https://") Then Return True

        Return False
    End Function

    '
    '   ファイルのコピーの作成
    '       ターゲットのファイル名とコピー先のディレクトリ
    Private Function GetFileToCopy(ByVal sourceFileName As String, ByVal destinationPathName As String) As Boolean
        Dim blnRet As Boolean
        Dim IsUrl As Boolean = CheckPathIsUrl(sourceFileName)
        Dim FileName As String = GetFileName(sourceFileName)
        If FileName = String.Empty Then Return False
        blnRet = CheckOrMakeDestinationDir(destinationPathName)
        If blnRet = False Then Return False

        Dim DestinationFile As String = Path.Combine(destinationPathName, FileName)
        If IsUrl Then
            'httpでのダウンロード
            With _myLogRecord
                .Target = sourceFileName & "を" & DestinationFile & "へhttpダウンロード開始"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            blnRet = GetHttpFileDownload(sourceFileName, DestinationFile)
        Else
            'ファイルのコピー
            With _myLogRecord
                .Target = sourceFileName & "を" & DestinationFile & "へファイルコピー開始"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            blnRet = GetFileDownload(sourceFileName, DestinationFile)
        End If

        Return blnRet

    End Function

    '
    '   httpでファイルをダウンロードして、ファイルを保存します。
    '
    Private Function GetHttpFileDownload(ByVal sourceUrl As String, ByVal destinationFile As String) As Boolean
        Dim myHttpRequest As HttpWebRequest
        Dim myHttpResponse As HttpWebResponse
        Dim myFileStream As FileStream
        Dim myBinaryWriter As BinaryWriter
        Dim myBuff() As Byte
        Dim myLength As Long
        Dim intRet As Integer
        Dim intReadByte As Integer
        Dim intReadToByte As Integer

        'httpリクエストの作成
        myHttpRequest = CType(WebRequest.Create(sourceUrl), HttpWebRequest)
        'mod-s 20050829 タイムアウト　httpsは5分、httpは30秒
        'myHttpRequest.Timeout = _TimeOut
        If sourceUrl.Substring(0, 5) = "https" Then
            myHttpRequest.Timeout = _TimeOut
        Else
            myHttpRequest.Timeout = _TimeOut2
        End If
        'mod-e 20050829 タイムアウト　httpsは5分、httpは30秒
        Try
            'httpレスポンスの取得
            myHttpResponse = myHttpRequest.GetResponse()
            If myHttpResponse.StatusCode <> HttpStatusCode.OK Then
                _ErrorMessage = "ファイルが見つかりません"
                With _myLogRecord
                    .Target = sourceUrl & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = myHttpResponse.StatusCode.ToString()
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End If

            'httpレスポンスの読み込み
            myLength = myHttpResponse.ContentLength()
            ReDim myBuff(myLength - 1)

            intReadByte = 0
            intReadToByte = myLength

            While intReadToByte > 0
                intRet = myHttpResponse.GetResponseStream.Read(myBuff, intReadByte, intReadToByte)
                If intRet = 0 Then
                    Exit While
                End If
                intReadByte += intRet
                intReadToByte -= intRet
            End While

            Try
                'ファイルへの保存
                '//MOD 2005.07.22 東都）伊賀 ファイルモードの変更 START
                'myFileStream = New FileStream(destinationFile, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
                myFileStream = New FileStream(destinationFile, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)
                '//MOD 2005.07.22 東都）伊賀 ファイルモードの変更 END
                myBinaryWriter = New BinaryWriter(myFileStream)
                myBinaryWriter.Write(myBuff)
                myBinaryWriter.Close()
                myFileStream.Close()
            Catch e As Exception
                _ErrorMessage = "ファイルコピー中のエラーです"
                With _myLogRecord
                    .Target = sourceUrl & "を" & _ErrorMessage
                    .Result = "NG"
                    .Remark = e.Message
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End Try

            myHttpResponse.Close()

        Catch e As WebException
            'タイムアウト
            If e.Status = WebExceptionStatus.ConnectFailure Then
                _ErrorMessage = WebExceptionStatus.ConnectFailure.ToString()
            Else
                _ErrorMessage = "タイムアウト"
            End If
            With _myLogRecord
                .Target = sourceUrl & "が" & _ErrorMessage
                .Result = "NG"
                .Remark = e.Message & "(status=" & e.Status.ToString() & ")"
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Catch e As Exception
            _ErrorMessage = e.Message
            With _myLogRecord
                .Target = sourceUrl & "が" & e.Message
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False

        End Try
        With _myLogRecord
            .Target = sourceUrl & "が正常にダウンロードできました"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True

    End Function

    '
    '   ファイルをコピーします
    '
    Private Function GetFileDownload(ByVal sourceFile As String, ByVal destinationFile As String) As Boolean

        Try
            'ファイルが存在するかをチェックして、コピー
            If File.Exists(sourceFile) = False Then
                _ErrorMessage = "ファイルが見つかりません"
                With _myLogRecord
                    .Target = sourceFile & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End If

            File.Copy(sourceFile, destinationFile, True)
        Catch e As Exception
            _ErrorMessage = "ファイルコピー中のエラーです"
            With _myLogRecord
                .Target = sourceFile & "が" & destinationFile & "への" & _ErrorMessage
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
        With _myLogRecord
            .Target = sourceFile & "が" & destinationFile & "へ正常にコピーできました"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        Return True
    End Function

    '
    '   ファイル名を取得します
    '
    Protected Friend Function GetFileName(ByVal fullPathFileName As String) As String
        If CheckPathIsUrl(fullPathFileName) Then
            'URLのファイル名を選択
            Dim FileIndex As Integer
            FileIndex = fullPathFileName.Replace("\", "/").LastIndexOf("/") + 1
            If FileIndex <= 0 Then
                _ErrorMessage = "ファイル名が指定されていません"
                With _myLogRecord
                    .Target = fullPathFileName & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If
            If FileIndex >= fullPathFileName.Length Then
                _ErrorMessage = "ファイル名が指定されていません"
                With _myLogRecord
                    .Target = fullPathFileName & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If

            Return fullPathFileName.Substring(FileIndex)
        Else
            'ファイルパスでのファイル名を選択
            Dim FileName As String
            FileName = Path.GetFileName(fullPathFileName)
            If FileName Is Nothing Then
                _ErrorMessage = "ファイル名が指定されていません"
                With _myLogRecord
                    .Target = fullPathFileName & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If
            If FileName = String.Empty Then
                _ErrorMessage = "ファイル名が指定されていません"
                With _myLogRecord
                    .Target = fullPathFileName & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If

            Return FileName
        End If

    End Function

    '
    '   ディレクトリーをチェックし、存在しなければ作成する
    '
    Protected Friend Function CheckOrMakeDestinationDir(ByVal dirName As String) As Boolean

        If Not (Directory.Exists(dirName)) Then
            Try
                Dim myDirInfo As DirectoryInfo
                myDirInfo = Directory.CreateDirectory(dirName)
                With _myLogRecord
                    .Target = dirName & "を作成しました"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            Catch e As Exception
                _ErrorMessage = "ディレクトリ作成エラー"
                With _myLogRecord
                    .Target = dirName & "が" & _ErrorMessage
                    .Result = "NG"
                    .Remark = e.Message
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End Try

        End If

        Return True
    End Function

End Class
