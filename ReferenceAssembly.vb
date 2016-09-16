Imports System.Reflection
Imports System.Threading
Imports System.IO

Public Class ReferenceAssembly
    Inherits MarshalByRefObject

    Private _SourcePath As String
    Private _DestinationAppPath As String
    Private _MyAppDomain As AppDomain
    Private _AssembliesAfterLoad As ArrayList
    Private _NotFindAssemblies As ArrayList

    Private Const constFileExt As String = ".dll"
    'チェックしないアセンブリのパターン
    Private _NoCheckAssemblies() As String = {"System.", "mscorlib"}
    Private Const constSystemName As String = "System"

#Region "ログ出力"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    Public Sub New()
        MyBase.New()

        _myLogRecord.Status = ""
        _myLogRecord.SubStatus = "参照アセンブリ"

    End Sub

    Private WriteOnly Property SourcePath() As String
        Set(ByVal Value As String)
            _SourcePath = Value
        End Set
    End Property

    Private WriteOnly Property DestinationAppPath() As String
        Set(ByVal Value As String)
            _DestinationAppPath = Value
        End Set
    End Property

    '
    '   アセンブリをロードして、参照アセンブリ配列を返す
    '       最初のメンバーにロードしたアセンブリ名を戻します
    '   
    Private Function GetReference(ByVal fileOrAssembly As String, _
                                  ByVal download As Boolean) As ArrayList
        Dim myArray As ArrayList
        Dim myName As String
        Dim myAssembly As [Assembly]
        Dim myAssemblies() As [AssemblyName]
        Dim myAssemblyName As AssemblyName
        Dim myMessages() As String
        Dim myRet As Boolean

        myArray = New ArrayList()

        myName = GetAssemblyName(fileOrAssembly, True)

        Try
            With _myLogRecord
                ReDim myMessages(1)
                myMessages(0) = fileOrAssembly
                myMessages(1) = "をロードします"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            If myName = String.Empty Then
                'ファイル名としてロード
                myAssembly = [Assembly].LoadFrom(fileOrAssembly)
            Else
                'アセンブリ・アイデンティティでロード
                myAssembly = [Assembly].Load(fileOrAssembly)
                'バージョンが違っているかどうか
                If myAssembly.FullName <> fileOrAssembly Then
                    'ダウンロードする
                    download = True
                    Throw New Exception("バージョン違い")
                End If
            End If
        Catch e As Exception
            If download Then
                'アセンブリをロード出来ない場合は、アセンブリをコピーしてロードします
                Dim myFileCopy As FileCopy = New FileCopy()
                Dim myFullPathFileName As String

                With _myLogRecord
                    ReDim myMessages(3)
                    myMessages(0) = _SourcePath
                    myMessages(1) = "から"
                    myMessages(2) = myName
                    myMessages(3) = "をダウンロードします"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    .Remark = e.Message
                    ReDim myMessages(1)
                End With
                _myLog.WriteLog(_myLogRecord)
                'URL形式かどうかによるダウンロード先を設定
                If myFileCopy.CheckPathIsUrl(_SourcePath) Then
                    myFullPathFileName = String.Concat(_SourcePath, myName)
                Else
                    myFullPathFileName = Path.Combine(_SourcePath, myName)
                End If
                myRet = myFileCopy.CopyFile(myFullPathFileName, _DestinationAppPath)
                If myRet = False Then
                    With _myLogRecord
                        myMessages(0) = myFullPathFileName
                        myMessages(1) = "をロードできませんでした"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "NG"
                        .Remark = e.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    Return Nothing
                End If
                Try
                    'ファイル名としてロード
                    myName = Path.Combine(_DestinationAppPath, myName)
                    With _myLogRecord
                        myMessages(0) = myName
                        myMessages(1) = "をロードします"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    myAssembly = [Assembly].LoadFrom(myName)
                Catch ee As Exception
                    With _myLogRecord
                        myMessages(0) = myName
                        myMessages(1) = "をロードできませんでした"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "NG"
                        .Remark = ee.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    Return Nothing
                End Try
            Else
                With _myLogRecord
                    '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 START
                    ReDim myMessages(1)
                    '// MOD 2016.09.16 Vivouac）菊池 Visual Studio 2013変換に伴う修正 END
                    myMessages(0) = fileOrAssembly
                    myMessages(1) = "をロードできませんでした"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "NG"
                    .Remark = e.Message
                End With
                _myLog.WriteLog(_myLogRecord)
                Return Nothing
            End If
        End Try

        '参照アセンブリ名を取得します
        myAssemblies = myAssembly.GetReferencedAssemblies()
        myArray.Add(myAssembly.FullName())  'アセンブリ名を追加

        For Each myAssemblyName In myAssemblies
            myArray.Add(myAssemblyName.FullName)
            With _myLogRecord
                ReDim myMessages(2)
                myMessages(0) = fileOrAssembly
                myMessages(1) = "の参照アセンブリは、"
                myMessages(2) = myAssemblyName.FullName
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = "List"
            End With
            _myLog.WriteLog(_myLogRecord)
        Next

        With _myLogRecord
            ReDim myMessages(4)
            myMessages(0) = "ロードした"
            myMessages(1) = fileOrAssembly
            myMessages(2) = "の参照アセンブリは"
            myMessages(3) = myAssemblies.Length.ToString()
            myMessages(4) = "件です"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myArray


    End Function

    '
    '   アセンブリ・アイデンティティからアセンブリ・フレンドリ名を取得
    '       拡張子DLLを付加
    Private Function GetAssemblyName(ByVal myAssemblyName As String, _
                                     ByVal addExpand As Boolean) As String
        Dim myPos As Integer
        Dim myName As String
        Dim myMessages() As String
        ReDim myMessages(1)

        'アイデンティティの区切り文字カンマの位置を取得
        myPos = myAssemblyName.IndexOf(",")
        If myPos <= 0 Then
            If addExpand Then
                With _myLogRecord
                    myMessages(0) = myAssemblyName
                    myMessages(1) = "はファイル名が指定されました"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            Return String.Empty
        End If
        'アセンブリ名を取得
        myName = myAssemblyName.Substring(0, myPos)
        If addExpand Then
            myName = String.Concat(myName, constFileExt)
            With _myLogRecord
                myMessages(0) = myName
                myMessages(1) = "アセンブリ名が選択されました"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
        End If

        Return myName
    End Function

    '
    '   現在のスレッドに読み込まれているアセンブリを取得する
    '       処理を開始する初めに、読み込み済みのアセンブリをリストし
    '       ロードするかどうかの判断に利用する
    Private Function GetAssemblies() As [Assembly]()
        Dim myAssemblies() As [Assembly]
        Try
            With _myLogRecord
                .Target = "読み込み済みのアセンブリを取得します"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            myAssemblies = Thread.GetDomain.GetAssemblies()
        Catch e As Exception
            With _myLogRecord
                .Target = "読み込み済みのアセンブリを取得できませんでした"
                .Result = "OK"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return Nothing
        End Try
        With _myLogRecord
            .Target = "読み込み済みのアセンブリを取得しました"
            .Result = "OK"
            Dim myMessages() As String = {myAssemblies.Length.ToString(), "件"}
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myAssemblies
    End Function


    '
    '   参照アセンブリを検索し、無ければダウンロードを試みる
    '       外部からの呼出しポイント
    '   最終はProtected Friend
    Protected Friend Function SearchAssemblies(ByVal fileName As String, _
                                               ByVal SourcePath As String, _
                                               ByVal DestinationAppPath As String) As Boolean
        Dim myRet As Boolean
        Dim myMBR As ReferenceAssembly
        Dim myAssemblies() As [Assembly]
        Dim myAssembly As [Assembly]
        Dim myArray As ArrayList
        Dim myFullPathFileName As String

        'アプリケーション・ドメインの作成
        With _myLogRecord
            .Target = "アプリケーション・ドメインを作成します"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        myRet = CreateDomain()
        If myRet = False Then Return False

        '参照アセンブリチェック用のインスタンスの作成
        Try
            myMBR = _MyAppDomain.CreateInstanceAndUnwrap([Assembly].GetExecutingAssembly.FullName, _
                                                    "Microsoft.jp.DevMktg.TE.ReferenceAssembly")
            With _myLogRecord
                .Target = "チェック用のインスタンスを作成しました"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
        Catch e As Exception
            With _myLogRecord
                .Target = "チェック用のインスタンスの作成に失敗しました"
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try

        _SourcePath = SourcePath
        _DestinationAppPath = DestinationAppPath
        myMBR.SourcePath = SourcePath
        myMBR.DestinationAppPath = DestinationAppPath

        '基準になるロードされたアセンブリの一覧を取得します
        myAssemblies = myMBR.GetAssemblies()
        If IsNothing(myAssemblies) Then
            AppDomain.Unload(_MyAppDomain)
            With _myLogRecord
                .Target = "終了：基準アセンブリの読込み失敗"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If
        With _myLogRecord
            .Target = "基準アセンブリを読込みました"
            .Result = "OK"
            .Remark = myAssemblies.Length.ToString() & "件"
        End With
        _myLog.WriteLog(_myLogRecord)

        _AssembliesAfterLoad = New ArrayList()
        _NotFindAssemblies = New ArrayList()
        For Each myAssembly In myAssemblies
            _AssembliesAfterLoad.Add(myAssembly.FullName())
        Next
        _AssembliesAfterLoad.Sort()

        '最初の参照アセンブリを取得します
        myFullPathFileName = Path.Combine(_DestinationAppPath, fileName)
        With _myLogRecord
            .Target = "最初のアセンブリをロードします"
            .Result = "OK"
            .Remark = myFullPathFileName
        End With
        _myLog.WriteLog(_myLogRecord)
        myArray = New ArrayList()
        myArray = myMBR.GetReference(myFullPathFileName, False)
        If IsNothing(myArray) Then
            _NotFindAssemblies.Add(myFullPathFileName)
            With _myLogRecord
                .Target = "終了：参照アセンブリの読込み失敗"
                .Result = "NG"
                .Remark = myFullPathFileName
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If
        With _myLogRecord
            .Target = "参照アセンブリの読込みました"
            .Result = "OK"
            .Remark = myFullPathFileName
        End With
        _myLog.WriteLog(_myLogRecord)

        '参照アセンブリの存在チェックを行います
        myRet = CheckAssemblies(myArray, myMBR, 1)
        With _myLogRecord
            If myRet Then
                .Target = "終了：チェックが正常に終了しました"
                .Result = "OK"
            Else
                .Target = "終了：チェックが異常終了しました"
                .Result = "NG"
                .Remark = ""
            End If
            Dim myMessages() As String = {"読込んだアセンブリ=", _AssembliesAfterLoad.Count.ToString(), _
                                          "件、読込めないアセンブリ=", _NotFindAssemblies.Count.ToString(), "件"}
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        AppDomain.Unload(_MyAppDomain)

        Return myRet
    End Function

    '
    '   チェック用のAppDomainを作成する
    Private Function CreateDomain() As Boolean
        Dim myAppSetup As AppDomainSetup
        Dim myName As String
        Dim myMessages() As String
        ReDim myMessages(1)

        myAppSetup = New AppDomainSetup()
        myAppSetup.LoaderOptimization = LoaderOptimization.MultiDomain

        myName = GetAssemblyName([Assembly].GetExecutingAssembly.FullName, True)
        Try
            _MyAppDomain = AppDomain.CreateDomain(myName, _
                                                  AppDomain.CurrentDomain.Evidence, _
                                                  myAppSetup)
        Catch e As Exception
            With _myLogRecord
                myMessages(0) = myName
                myMessages(1) = "アプリケーション・ドメインの作成に失敗しました"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
        With _myLogRecord
            myMessages(0) = myName
            myMessages(1) = "アプリケーション・ドメインの作成に成功しました"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True
    End Function

    '
    '   参照アセンブリをチェックします
    '       この処理は、再帰コールされます
    Private Function CheckAssemblies(ByVal myArray As ArrayList, _
                                     ByRef myMBR As ReferenceAssembly, _
                                     ByVal TireLevel As Long) As Boolean
        Dim myRet As Boolean
        Dim myObject As Object
        Dim myName As String
        Dim myRefArray As ArrayList
        Dim i, j As Long
        Dim myMessages() As String

        ReDim myMessages(1)

        'ロード済みのアセンブリをチェック済みとする
        If _AssembliesAfterLoad.BinarySearch(myArray.Item(0)) < 0 Then
            _AssembliesAfterLoad.Add(myArray.Item(0))
            myArray.RemoveAt(0)
        End If
        _AssembliesAfterLoad.Sort()
        With _myLogRecord
            ReDim myMessages(1)
            myMessages(0) = TireLevel.ToString()
            myMessages(1) = ":読込み済みのアセンブリ件数"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = _AssembliesAfterLoad.Count.ToString() : myMessages(1) = "件"
            .Remark = _myLog.AppendStrings(myMessages)
            myMessages(0) = TireLevel.ToString()
        End With
        _myLog.WriteLog(_myLogRecord)

        'ロード済みのアセンブリを参照アセンブリリストから削除する
        For i = myArray.Count - 1 To 0 Step -1
            If _AssembliesAfterLoad.BinarySearch(myArray.Item(i)) >= 0 Then
                myArray.RemoveAt(i)
            Else
                Dim myAssemblyName As String
                myAssemblyName = GetAssemblyName(myArray.Item(i), False)
                If constSystemName = myAssemblyName Then
                    myArray.RemoveAt(i)
                Else
                    For j = 0 To _NoCheckAssemblies.Length - 1
                        If myAssemblyName.StartsWith(_NoCheckAssemblies(j)) Then
                            myArray.RemoveAt(i)
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
        With _myLogRecord
            myMessages(1) = ":参照アセンブリのチェック対象件数"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = myArray.Count.ToString() : myMessages(1) = "件"
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        '残りのアセンブリ参照メンバーをロードしてチェックする
        For Each myObject In myArray
            myName = CType(myObject, String)
            With _myLogRecord
                ReDim myMessages(1)
                myMessages(0) = TireLevel.ToString()
                myMessages(1) = ":個別のアセンブリをチェックします"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = myName
            End With
            _myLog.WriteLog(_myLogRecord)
            myRefArray = myMBR.GetReference(myName, True)
            If myRefArray Is Nothing Then
                'ロードできないアセンブリリストに追加
                _NotFindAssemblies.Add(myName)
                With _myLogRecord
                    myMessages(1) = ":ロード出来ませんでした"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "NG"
                    .Remark = myName
                End With
                _myLog.WriteLog(_myLogRecord)
            Else
                With _myLogRecord
                    myMessages(1) = ":再帰チェックを開始します"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    myMessages(0) = myRefArray.Count.ToString()
                    myMessages(1) = "件"
                    .Remark = _myLog.AppendStrings(myMessages)
                    myMessages(0) = TireLevel.ToString()
                End With
                _myLog.WriteLog(_myLogRecord)

                myRet = CheckAssemblies(myRefArray, myMBR, TireLevel + 1)
                With _myLogRecord
                    If myRet = False Then
                        _NotFindAssemblies.Add(myName)
                        myMessages(1) = ":再帰チェックに失敗しました"
                        .Result = "NG"
                    Else
                        myMessages(1) = ":再帰チェックに成功しました"
                        .Result = "OK"
                    End If
                    .Target = _myLog.AppendStrings(myMessages)
                    ReDim myMessages(3)
                    myMessages(0) = myName
                    myMessages(1) = "の参照アセンブリが、"
                    myMessages(2) = myRefArray.Count.ToString()
                    myMessages(3) = "件でした"
                    .Remark = _myLog.AppendStrings(myMessages)
                End With
                _myLog.WriteLog(_myLogRecord)

                myRefArray.Clear()
            End If
        Next
        With _myLogRecord
            ReDim myMessages(1)
            myMessages(0) = TireLevel.ToString()
            myMessages(1) = ":個別アセンブリのチェックが終了しました"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = myArray.Count.ToString()
            myMessages(1) = "件"
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True
    End Function
End Class

