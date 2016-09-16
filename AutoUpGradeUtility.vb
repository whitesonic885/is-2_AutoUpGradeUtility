Imports System.IO
Imports System.Reflection
Imports System.Net
Imports System.Threading

Imports System.Text
Imports System.Security.Cryptography

	'/// <summary>
	'/// 自動更新アプリ部品[AutoUpGradeUtility.dll]
	'/// </summary>
	'//--------------------------------------------------------------------------
	'// 修正履歴
	'//--------------------------------------------------------------------------
	'// ADD 2008.06.11 東都）高木 [conime.exe]の起動 
	'// ADD 2008.06.11 東都）高木 ログイン画面を常にノーマル表示に 
	'// ADD 2008.06.11 東都）高木 [CopyAutoUpGrade]を非表示に 
	'// ADD 2008.06.11 東都）高木 [EXPAND]を非表示に 
	'//--------------------------------------------------------------------------
	'// ADD 2009.07.29 東都）高木 プロキシ対応 
	'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 
	'//                           （次期リリース時にはコメントにする）
	'// ADD 2009.07.30 東都）高木 exeのdll化対応 
	'// ADD 2009.10.02 東都）高木 ２重起動のチェック 
	'// ADD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 
	'// ADD 2009.10.06 東都）高木 パソコンの日付設定簡易チェック 
	'//--------------------------------------------------------------------------
	'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 
	'// MOD 2010.06.21 東都）高木 ログの暗号化解除 
	'//                           （暗号化を多用すると解読されやすくなる為）
	'//--------------------------------------------------------------------------
Public Class AutoUpGradeUtility

    Private _FileName As String
    Private _SourcePath As String
    Private _DestinationPath As String
    Private _DestinationAppPath As String
    Private _DestinationConfigPath As String
    Private Const constAppDir As String = "App"
    Private Const constConfigDir As String = "Config"
    Private Const constConfigExt As String = ".config"
    Private _AppDir As String
    Private _ConfigDir As String
    Public waitDlg As WaitDialog

'// ADD 2005.05.31 東都）伊賀 CopyAutoUpGrade対応 START
    Public _ClientMutex As String
'// ADD 2005.05.31 東都）伊賀 CopyAutoUpGrade対応 END
'// ADD 2009.07.29 東都）高木 プロキシ対応 START
    Dim gsInitProxy As String = AppDomain.CurrentDomain.BaseDirectory _
             + "\proxy.ini"
    Dim gbInitProxyExists As Boolean = False
    Dim gsアプリフォルダ As String = ""

    Dim gsProxyAdr As String = ""
    Dim gsProxyAdrUserSet As String = ""
    Dim gsProxyAdrSecure As String = ""
    Dim gsProxyAdrHttp As String = ""
    Dim gsProxyAdrAll As String = ""
    Dim giProxyNo As Integer = 0
    Dim giProxyNoUserSet As Integer = 0
    Dim giProxyNoSecure As Integer = 0
    Dim giProxyNoHttp As Integer = 0
    Dim giProxyNoAll As Integer = 0
    Dim giConnectTimeOut As Integer = 0
    Dim gProxy As System.Net.WebProxy
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
	Dim gbProxyOnUserSet   As Boolean = False
	Dim gbProxyIdOnUserSet As Boolean = False
    Dim gsProxyIdUserSet   As String  = ""
    Dim gsProxyPaUserSet   As String  = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END

    Dim sv_init As is2init.Service1 = Nothing
'// ADD 2009.07.29 東都）高木 プロキシ対応 END

#Region "ログ出力"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    '
    '   コンストラクタ
    '
    Public Sub New()
        '初期化
        _DestinationPath = AppDomain.CurrentDomain.BaseDirectory
        _AppDir = Path.Combine(_DestinationPath, constAppDir)
        _ConfigDir = Path.Combine(_DestinationPath, constConfigDir)
        _DestinationAppPath = Path.Combine(_DestinationPath, _AppDir)
        _DestinationConfigPath = Path.Combine(_DestinationPath, _ConfigDir)

        With _myLogRecord
            .Status = "開始"
            .SubStatus = "AutoUpGrade"
        End With

    End Sub

'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
    '
    '   ファイルスタンプの変更
    '
    Public Function ChangeFileStamp(ByVal sFileName As String, _
                                 ByVal sTimeStamp As String) As Boolean

        Try
            Dim sFullFileName As String = Path.Combine(_DestinationAppPath, sFileName)
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 START
'//         Dim sFileStamp As String = _
'//             sTimeStamp.Substring(0, 4) + "/" + _
'//             sTimeStamp.Substring(4, 2) + "/" + _
'//             sTimeStamp.Substring(6, 2) + " " + _
'//             sTimeStamp.Substring(8, 2) + ":" + _
'//             sTimeStamp.Substring(10, 2) + ":" + _
'//             sTimeStamp.Substring(12, 2)
            Dim sFileStamp As String = _
                sTimeStamp.Substring(0, 4) + "/" + _
                sTimeStamp.Substring(4, 2) + "/" + _
                sTimeStamp.Substring(6, 2) + " " + _
                sTimeStamp.Substring(8, 2) + ":" + _
                sTimeStamp.Substring(10, 2) + ":" + _
                "00"
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 END

            Dim dtFileStamp As Date = Date.Parse(sFileStamp)
            'With _myLogRecord
            '    .Target = "ファイルスタンプの変更：" + dtFileStamp.ToString()
            '    .Result = "OK"
            '    .Remark = ""
            'End With
            '_myLog.WriteLog(_myLogRecord)

            'ファイルスタンプの変更
            File.SetLastWriteTime(sFullFileName, dtFileStamp)

        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
        End Try

        Return True
    End Function
    '
    '   ＣＡＢファイルのダウンロード
    '
    Public Function DownloadCabFile(ByVal srcFullName As String, _
                                        ByVal desFullName As String, _
                                        ByVal sourceDirName As String, _
                                        ByVal sFileName As String) As Boolean

        Dim myRet As Boolean
        Try
            Dim myFileCopy As FileCopy = New FileCopy

            'ＣＡＢファイルのダウンロード
            myRet = myFileCopy.CopyFile(sourceDirName & sFileName, _DestinationPath)
            If myRet Then
                With _myLogRecord
                    .Target = sourceDirName & sFileName & "を" & _DestinationPath & "へダウンロードに成功しました"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)

                'ＣＡＢバージョンファイルを上位階層にコピーする
                File.Copy(srcFullName, desFullName, True)
                'ＣＡＢバージョンファイルのファイルスタンプの変更
                File.SetLastWriteTime(desFullName, File.GetLastWriteTime(srcFullName))

                Dim flgFullName As String

                sFileName = sFileName.Substring(0, sFileName.Length - 3) + "cab"
                flgFullName = Path.Combine(_DestinationPath, sFileName) + "_flg"

                '上位階層のフラグファイルの削除
                If System.IO.File.Exists(flgFullName) Then
                    System.IO.File.Delete(flgFullName)
                End If
            Else
                With _myLogRecord
                    .Target = sourceDirName & sFileName & "のダウンロードに失敗しました"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
        End Try

        Return True
    End Function
    '
    '   ＣＡＢフラグファイルの削除
    '
    Public Function DeleteCabFlagFile(ByVal sFileName As String) As Boolean

        Try
            'ＣＡＢバージョンファイルがダウンロードされた時
            If sFileName.Substring(sFileName.LastIndexOf(".") + 1).Equals("txt") Then
                Dim flgFullName As String

                sFileName = sFileName.Substring(0, sFileName.Length - 3) + "cab"
                flgFullName = Path.Combine(_DestinationPath, sFileName) + "_flg"

                '上位階層のフラグファイルの削除
                If System.IO.File.Exists(flgFullName) Then
                    System.IO.File.Delete(flgFullName)
                End If
            End If

        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
        End Try

        Return True
    End Function
    '
    '   ＣＡＢバージョンファイルのチェック
    '
    Public Function CheckCabVerFile(ByVal sourceDirName As String, _
                                        ByVal sFileName As String) As Boolean

        Try
            Dim srcFullName, desFullName, cabFullName As String
            srcFullName = Path.Combine(_DestinationAppPath, sFileName)
            desFullName = Path.Combine(_DestinationPath, sFileName)
            cabFullName = desFullName.Substring(0, desFullName.Length - 3) + "cab"

            sFileName = sFileName.Substring(0, sFileName.Length - 3) + "cab"

            If sourceDirName.EndsWith("/ReleaseAdmin/") Then
                sourceDirName = sourceDirName.Replace("/ReleaseAdmin/", "/ReleaseCab/")
            Else
                sourceDirName = sourceDirName.Replace("/Release/", "/ReleaseCab/")
            End If

            'ＣＡＢバージョンファイルが存在しない時、ＣＡＢファイルをダウンロード
            If File.Exists(desFullName) = False Then
                'ファイルをダウンロードする
                DownloadCabFile(srcFullName, desFullName, _
                                sourceDirName, sFileName)

                'ＣＡＢバージョンファイルの更新日時が新しい時、ＣＡＢファイルをダウンロード
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 START
'//         ElseIf File.GetLastWriteTime(srcFullName) > File.GetLastWriteTime(desFullName) Then
                'タイムスタンプは、年月日時分まで比較する（秒は比較しない）
            ElseIf File.GetLastWriteTime(srcFullName).ToString("yyyyMMddHHmm") > File.GetLastWriteTime(desFullName).ToString("yyyyMMddHHmm") Then
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 END
                'ファイルをダウンロードする
                DownloadCabFile(srcFullName, desFullName, _
                                sourceDirName, sFileName)

                'ＣＡＢファイルが存在しない時、ＣＡＢファイルをダウンロード
            ElseIf File.Exists(cabFullName) = False Then
                'ファイルをダウンロードする
                DownloadCabFile(srcFullName, desFullName, _
                                sourceDirName, sFileName)
            End If

        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
        End Try

        Return True
    End Function
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END

    '
    '   バージョン情報の取得
    '
    Public Function GetVersion(ByVal sourceDirName As String, _
                                 ByVal sourceFileName As String, _
                                 ByVal download As Boolean, _
                                 ByVal downloadOnly As Boolean, _
                                 ByVal localExecute As Boolean, _
                                 ByVal logClear As Boolean) As Boolean

'// ADD 2008.06.11 東都）高木 [conime.exe]の起動 START
        Dim pConIme As New System.Diagnostics.Process
        Try
            With pConIme.StartInfo
                .FileName = "cmd.exe"
                .Arguments = "/C conime.exe"
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            End With
            pConIme.Start()
        Catch ex As System.Exception
            'その他
        Finally
            If Not (pConIme Is Nothing) Then pConIme.Close()
        End Try
'// ADD 2008.06.11 東都）高木 [conime.exe]の起動 END

        If logClear Then _myLog.Clear()

'// ADD 2009.10.02 東都）高木 ２重起動のチェック START
        Dim sCheckExeName As String = ""
        If sourceDirName.EndsWith("/ReleaseAdmin/") Then
            sCheckExeName = "is2AdminClient"
        Else
            sCheckExeName = "is2Client"
        End If
        Dim iExeCnt As Integer = 0
        Try
            Dim hMutex As New System.Threading.Mutex(False, sCheckExeName)
            Do While hMutex.WaitOne(0, False) = False
                iExeCnt += 1
                If iExeCnt > 2 Then
                    With _myLogRecord
                        .Target = "2秒待ちましたが、終了しませんでした" _
                          + "[" + sCheckExeName + "]"
                        .Result = "NG"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    Windows.Forms.MessageBox.Show( _
                     "アプリケーションは既に起動されています　　　" & vbCrLf _
                     & "タスクバーやタスクマネージャ等をご確認ください　　　", _
                     sCheckExeName, _
                     Windows.Forms.MessageBoxButtons.OK, _
                     Windows.Forms.MessageBoxIcon.Error)

                    Return True
                End If
                System.Threading.Thread.Sleep(1000)
            Loop

            ' Mutex を開放する
            hMutex.ReleaseMutex()
            hMutex.Close()
        Catch ex As ApplicationException
            With _myLogRecord
                .Target = "Mutexを開放できませんでした"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
'// ADD 2009.10.02 東都）高木 ２重起動のチェック END
'// ADD 2009.10.06 東都）高木 パソコンの日付設定簡易チェック START
        Dim dtNow As New Date
        Dim sNow = dtNow.Now.ToString("yyyy/MM/dd")
        If dtNow.Now.Year < 2009 Then
            With _myLogRecord
                .Target = "パソコンの日付設定に誤りがあります" _
                  & "[" & sNow & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Windows.Forms.MessageBox.Show( _
             "パソコンの日付設定に誤りがあります　　　" & vbCrLf _
             & "（" & sNow & "）" & "　　　", _
             sCheckExeName, _
             Windows.Forms.MessageBoxButtons.OK, _
             Windows.Forms.MessageBoxIcon.Error)

            Return False
        End If
'// ADD 2009.10.06 東都）高木 パソコンの日付設定簡易チェック END

        Dim myRet As Boolean
        waitDlg = New WaitDialog

        Try
            Dim myFileCopy As FileCopy = New FileCopy
'//            Dim myCheckFileName As String = Path.Combine(_DestinationConfigPath, myFileCopy.GetFileName(sourceFileName))
'//            'パス形式がおかしい
'//            If Not (localExecute = True And File.Exists(myCheckFileName)) Then
'//                'コンフィグのダウンロード、DL出来なくても続行する
'//                myRet = myFileCopy.CopyFile(sourceDirName & sourceFileName & constConfigExt, _DestinationConfigPath)
'//                With _myLogRecord
'//                    If myRet Then
'//                        .Target = sourceDirName & sourceFileName & constConfigExt & "を" & _DestinationConfigPath & "へダウンロードできました"
'//                        .Result = "OK"
'//                        .Remark = ""
'//                    Else
'//現状、ネットワークとつながっていないと起動できないため
'//ローカル起動は行わない
'//                        If myFileCopy.ErrorMesage = WebExceptionStatus.ConnectFailure.ToString() Then
'//ネットワークエラーはローカル起動にする
'//                            localExecute = True
'//                            download = False
'//                            .Target = sourceDirName & sourceFileName & "をネットワークエラーによりローカル起動に切り替えます"
'//                            .Result = "OK"
'//                            .Remark = myFileCopy.ErrorMesage
'//                        Else
'//                            .Target = sourceDirName & sourceFileName & constConfigExt & "を" & _DestinationConfigPath & "へダウンロードできません"
'//                            .Result = "NG"
'//                            .Remark = myFileCopy.ErrorMesage
'//                        End If
'//                    End If
'//                End With
'//                _myLog.WriteLog(_myLogRecord)
'//            End If
'//            If Not download Then Return True

            Dim verCheck As New VersionCheck
            Dim ServerSerializer As New System.Xml.Serialization.XmlSerializer(GetType(VersionCheck.FileVersion()))
            Dim ServerVersion() As VersionCheck.FileVersion
            Dim ServerFile As VersionCheck.FileVersion
            Dim ClientSerializer As New System.Xml.Serialization.XmlSerializer(GetType(VersionCheck.FileVersion()))
            Dim ClientVersion() As VersionCheck.FileVersion
            Dim ClientFile As VersionCheck.FileVersion

'// ADD 2005.07.22 東都）小童谷  START
            Dim appFile As String
            Dim appSize As Integer = 0
'// ADD 2005.07.22 東都）小童谷  END

            ' 進行状況ダイアログを表示する
            waitDlg.Show()
            ' 進行状況ダイアログの初期化処理
            'waitDlg.Owner = Me ' ダイアログのオーナーを設定
'// MOD 2005.05.10 東都）高木 メッセージの変更 START
            '            waitDlg.MainMsg = "バージョンファイルを取得しています……" ' 処理の概要
            '            waitDlg.SubMsg = ""
            waitDlg.MainMsg = "起動準備中です．．．"
'// MOD 2005.05.10 東都）高木 メッセージの変更 END
            waitDlg.ProgressMsg = ""
            waitDlg.ProgressMin = 0 ' 処理件数の最小値（0件から開始）
            waitDlg.ProgressStep = 1 ' 何件ごとにメーターを進めるか
            waitDlg.ProgressValue = 0 ' 最初の件数
            waitDlg.Update()
            Dim iCount As Integer = 0   ' 何件目の処理かを示すカウンタ

'// ADD 2009.07.29 東都）高木 プロキシ対応 START
            '// カレントディレクトリの取得
            gsアプリフォルダ = AppDomain.CurrentDomain.BaseDirectory
            If Not (Directory.Exists(gsアプリフォルダ)) Then
                With _myLogRecord
                    .Target = "カレントディレクトリがみつかりませんでした\n" _
                      + "[" + gsアプリフォルダ + "]"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If

            gbInitProxyExists = False
            gsProxyAdrUserSet = ""
            giProxyNoUserSet = 0
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         giConnectTimeOut = 100  '// 初期値：１００秒
            giConnectTimeOut   = 50 '// 初期値：５０秒
	        gbProxyOnUserSet   = False
	        gbProxyIdOnUserSet = False
            gsProxyIdUserSet   = ""
            gsProxyPaUserSet   = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
            Dim sConnectTimeOut As String = ""
            Dim sProxyAdrUserSet As String = ""
            Dim sProxyNoUserSet As String = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
			Dim sProxyOnUserSet   As String = ""
			Dim sProxyIdOnUserSet As String = ""
			Dim sProxyIdUserSet   As String = ""
			Dim sProxyPaUserSet   As String = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
            Dim sr As StreamReader = Nothing
            Try
                sr = File.OpenText(gsInitProxy)
                gbInitProxyExists = True
                sConnectTimeOut = sr.ReadLine()
                sProxyAdrUserSet = sr.ReadLine()
                sProxyNoUserSet = sr.ReadLine()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//				if not(sr is Nothing) then sr.Close()
                sProxyOnUserSet   = sr.ReadLine()
                sProxyIdOnUserSet = sr.ReadLine()
                sProxyIdUserSet   = sr.ReadLine()
                sProxyPaUserSet   = sr.ReadLine()
                If sConnectTimeOut   Is Nothing Then sConnectTimeOut   = ""
                If sProxyAdrUserSet  Is Nothing Then sProxyAdrUserSet  = ""
                If sProxyNoUserSet   Is Nothing Then sProxyNoUserSet   = ""
                If sProxyOnUserSet   Is Nothing Then sProxyOnUserSet   = ""
                If sProxyIdOnUserSet Is Nothing Then sProxyIdOnUserSet = ""
                If sProxyIdUserSet   Is Nothing Then sProxyIdUserSet   = ""
                If sProxyPaUserSet   Is Nothing Then sProxyPaUserSet   = ""

                If sProxyIdUserSet.Length > 0 Then
                    sProxyIdUserSet  = 復号化２(sProxyIdUserSet)
                    sProxyPaUserSet  = 復号化２(sProxyPaUserSet)
                End If
'//				With _myLogRecord
'//					.Target = "" _
'//						& "[" & sProxyIdUserSet & "]" _
'//						& "[" & sProxyPaUserSet & "]"
'//					.Result = ""
'//					.Remark = ""
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
            Catch ex As System.IO.FileNotFoundException
                With _myLogRecord
                    .Target = "プロキシ設定ファイルがみつかりませんでした"
                    .Result = "NG"
                    .Remark = ex.Message
                End With
                _myLog.WriteLog(_myLogRecord)
            Catch ex As Exception
                With _myLogRecord
                    .Target = "プロキシ設定ファイル Exception:"
                    .Result = "NG"
                    .Remark = ex.Message
                End With
                _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
            Finally
                If Not (sr Is Nothing) Then sr.Close()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
            End Try

'// ADD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 START
            If gbInitProxyExists Then
'// ADD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 END
                Try
                    If sConnectTimeOut.Length > 0 Then
                        If 半角チェック(sConnectTimeOut, "タイムアウト") Then
                            If 数値チェック(sConnectTimeOut, "タイムアウト") Then
                                giConnectTimeOut = Integer.Parse(sConnectTimeOut)
                            End If
                        End If
                    End If
                    If Not (sProxyAdrUserSet Is Nothing) Then
                        If 半角チェック(sProxyAdrUserSet, "プロキシアドレス") Then
                            gsProxyAdrUserSet = sProxyAdrUserSet
                        End If
                    End If
                    If sProxyNoUserSet.Length > 0 Then
                        If 半角チェック(sProxyNoUserSet, "プロキシポート番号") Then
                            If 数値チェック(sProxyNoUserSet, "プロキシポート番号") Then
                                giProxyNoUserSet = Integer.Parse(sProxyNoUserSet)
                            End If
                        End If
                    End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
                    If 半角チェック(sProxyOnUserSet, "プロキシ設定") Then
	                    If sProxyOnUserSet.Equals("1") Then
	                        gbProxyOnUserSet = True
	                    End If
                    End If
                    If sProxyOnUserSet.Length = 0 Then
                        If gsProxyAdrUserSet.Length > 0 Then gbProxyOnUserSet = True
                    End If
                    If 半角チェック(sProxyIdOnUserSet, "プロキシＩＤ設定") Then
	                    If sProxyIdOnUserSet.Equals("1") Then
	                        gbProxyIdOnUserSet = True
	                    End If
                    End If
                    If 半角チェック(sProxyIdUserSet, "ユーザＩＤ") Then
                        gsProxyIdUserSet = sProxyIdUserSet
                    End If
                    If 半角チェック(sProxyPaUserSet, "パスワード") Then
                        gsProxyPaUserSet = sProxyPaUserSet
                    End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                Catch ex As Exception
'//保留　エラー処理
                    With _myLogRecord
                        .Target = "プロキシ設定値ワーニング Exception:"
                        .Result = "NG"
                        .Remark = ex.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                End Try

                '// プロキシの設定を取得
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             Dim iRet As Integer = -1
                Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                ＩＥプロキシ設定情報取得()

                sv_init = New is2init.Service1
                If giConnectTimeOut > 0 Then
                    sv_init.Timeout = giConnectTimeOut * 1000
                End If
                sv_init.Url = sv_init.Url.Replace("http://", "https://")

                '//--------------------------------------------------------
                '// プロキシ設定(ユーザ設定)
                '// （[proxy.ini]設定値）
                '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             If iRet <> 1 And gsProxyAdrUserSet.Length > 0 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                     .Target = "P設定(U)[" & 暗号化２(gsProxyAdrUserSet) _
'//                         & "][" & 暗号化２(giProxyNoUserSet.ToString("0000")) & "]"
                        .Target = "P設定(U)[" & gsProxyAdrUserSet _
                            & "][" & giProxyNoUserSet.ToString("0000") & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 iRet = プロキシ設定(gsProxyAdrUserSet, giProxyNoUserSet)
'//                 If iRet = 1 Then myRet = True
                    With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                     .Target = "P_1[" & 暗号化２(gbProxyOnUserSet.ToString()) _
'//                         & "] P_2[" & 暗号化２(gbProxyIdOnUserSet.ToString()) & "]"
                        .Target = "P_1[" & gbProxyOnUserSet.ToString() _
                            & "] P_2[" & gbProxyIdOnUserSet.ToString() & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    If gbProxyOnUserSet Then
                        If gbProxyIdOnUserSet Then
                            wRet = プロキシ設定２(gsProxyAdrUserSet, giProxyNoUserSet _
                                , gsProxyIdUserSet, gsProxyPaUserSet)
                        Else
                            wRet = プロキシ設定(gsProxyAdrUserSet, giProxyNoUserSet)
                        End If
                    Else
                        wRet = プロキシ設定("", 0) '// プロキシなし接続
                    End If
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                End If

                '//--------------------------------------------------------
                '// デフォルトプロキシ設定（以前と同じになるはず）
                '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    With _myLogRecord
                        .Target = "P設定(D)"
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 iRet = デフォルトプロキシ設定()
'//                 If iRet = 1 Then myRet = True
                    wRet = デフォルトプロキシ設定()
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                End If

                '//--------------------------------------------------------
                '// プロキシ設定(ＳＳＬ、排除リスト有効)←ベストだと思われる
                '// （ＩＥのレジストリ設定値）
                '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             If iRet <> 1 And gsProxyAdr.Length > 0 Then
                If wRet <> WebExceptionStatus.Success  And gsProxyAdr.Length > 0Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                     .Target = "P設定(S)[" & 暗号化２(gsProxyAdr) _
'//                         & "][" & 暗号化２(giProxyNo.ToString("0000")) & "]"
                        .Target = "P設定(S)[" & gsProxyAdr _
                            & "][" & giProxyNo.ToString("0000") & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 iRet = プロキシ設定(gsProxyAdr, giProxyNo)
'//                 If iRet = 1 Then myRet = True
                    wRet = プロキシ設定(gsProxyAdr, giProxyNo)
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                End If

                '//--------------------------------------------------------
                '// プロキシ未設定時
                '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 If gsProxyAdrUserSet.Length = 0 And gsProxyAdr.Length = 0 Then
                    If gbProxyOnUserSet = False And gsProxyAdrUserSet.Length = 0 And gsProxyAdr.Length = 0 Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                        Try
                            Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                         iRet = 1
                            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                            myRet = True
                        Catch ex As System.Net.WebException
                            With _myLogRecord
                                .Target = "P設定(N) s WebException:" & ex.Status.ToString()
                                .Result = "NG"
                                .Remark = ex.Message
                            End With
                            _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                         iRet = -1
'//                         Select Case ex.Status
'//                             Case WebExceptionStatus.NameResolutionFailure
'//                                 iRet = -11
'//                             Case WebExceptionStatus.Timeout
'//                                 iRet = -12
'//                             Case WebExceptionStatus.TrustFailure
'//                                 iRet = -13
'//                             Case WebExceptionStatus.ConnectFailure
'//                                 iRet = -14
'//                             Case Else
'//                                 iRet = -19
'//                         End Select
                            wRet = ex.Status
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                        Catch ex As Exception
                            With _myLogRecord
                                .Target = "P設定(N) s Exception:"
                                .Result = "NG"
                                .Remark = ex.Message
                            End With
                            _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                         iRet = -2
                            wRet = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                        End Try
                        '//--------------------------------------------------------
                        '// 開発機用
                        '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                     If iRet <> 1 Then
                        If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                            Try
                                sv_init = New is2init.Service1
                                sv_init.Timeout = 5000 '// 5秒
                                sv_init.Url = sv_init.Url.Replace("https://", "http://")
                                Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                             iRet = 1
                                wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                                myRet = True
                            Catch ex As System.Net.WebException
                                With _myLogRecord
                                    .Target = "P設定(N) _ WebException:" & ex.Status.ToString()
                                    .Result = "NG"
                                    .Remark = ex.Message
                                End With
                                _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                             iRet = -1
'//                             Select Case ex.Status
'//                                 Case WebExceptionStatus.NameResolutionFailure
'//                                     iRet = -11
'//                                 Case WebExceptionStatus.Timeout
'//                                     iRet = -12
'//                                 Case WebExceptionStatus.TrustFailure
'//                                     iRet = -13
'//                                 Case WebExceptionStatus.ConnectFailure
'//                                     iRet = -14
'//                                 Case Else
'//                                     iRet = -19
'//                             End Select
                                wRet = ex.Status
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                            Catch ex As Exception
                                With _myLogRecord
                                    .Target = "P設定(N) _ Exception:"
                                    .Result = "NG"
                                    .Remark = ex.Message
                                End With
                                _myLog.WriteLog(_myLogRecord)
                                myRet = False
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                             iRet = -2
                                wRet = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                            Finally
                                If giConnectTimeOut > 0 Then
                                    sv_init.Timeout = giConnectTimeOut * 1000
                                End If
                            End Try
                        End If
                    End If
                End If

                sv_init.Url = sv_init.Url.Replace("http://", "https://")

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
                Dim iDlgCnt = 0
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             Do While iRet <> 1
                Do While wRet <> WebExceptionStatus.Success
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
                    '//--------------------------------------------------------
                    '// 接続エラーメッセージの設定
                    '//--------------------------------------------------------
                    Dim sErrMsg As String = 接続エラーメッセージ編集(wRet)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    '//--------------------------------------------------------
                    '// WebExceptionStatus.TrustFailure ＳＳＬ通信の失敗
                    '// （→プロキシの設定ダイアログを表示しない）
                    '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 If iRet = -13 Then
'//                     Windows.Forms.MessageBox.Show( _
'//                      "is2サーバとの通信に失敗しました　　　" & vbCrLf _
'//                      & "パソコンの日付設定やＳＳＬ通信の設定などを確認してください　　　", _
'//                      "is2", _
'//                      Windows.Forms.MessageBoxButtons.OK, _
'//                      Windows.Forms.MessageBoxIcon.Error)
'//                     Exit Do
'//                 End If
                    If wRet = WebExceptionStatus.TrustFailure Then
                        Windows.Forms.MessageBox.Show( _
                        "サーバー接続エラー（"& sErrMsg &"）　　　　　" & vbCrLf _
                         & "パソコンの日付設定やＳＳＬ通信の設定などを確認してください　　　　　", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                        Exit Do
                    End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 '//--------------------------------------------------------
'//                 '// WebExceptionStatus.ConnectFailure 接続の失敗
'//                 '// （→プロキシの設定ダイアログを表示しない）
'//                 '//--------------------------------------------------------
'//                 If iRet = -14 Then
'//                     Exit Do
'//                 End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    '//--------------------------------------------------------
                    '// プロキシ設定が３回以上失敗したら終了
                    '//--------------------------------------------------------
                    iDlgCnt += 1
                    If iDlgCnt > 3 Then
                        Windows.Forms.MessageBox.Show( _
                        "サーバー接続エラー（"& sErrMsg &"）　　　　　" & vbCrLf _
                         & "プロキシ設定に３回失敗しましたので終了します　　　　　", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                        Exit Do
                    End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    '//--------------------------------------------------------
                    '// プロキシ設定確認メッセージの表示
                    '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 Dim dRet As Windows.Forms.DialogResult = Windows.Forms.MessageBox.Show( _
'//                     "is2サーバとの通信に失敗しました　　　" & vbCrLf _
'//                     & "プロキシの設定を行いますか？　　　", _
'//                     "is2", _
'//                     Windows.Forms.MessageBoxButtons.YesNo, _
'//                     Windows.Forms.MessageBoxIcon.Question)
                    Dim dRet As Windows.Forms.DialogResult = Windows.Forms.MessageBox.Show( _
                        "サーバー接続エラー（"& sErrMsg &"）　　　　　" & vbCrLf _
                        & "プロキシの設定を行いますか？　　　　　", _
                        "is2", _
                        Windows.Forms.MessageBoxButtons.YesNo, _
                        Windows.Forms.MessageBoxIcon.Error)
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    '[いいえ]ならそのまま終了
                    If dRet = Windows.Forms.DialogResult.No Then Exit Do
                    '//--------------------------------------------------------
                    '// プロキシ設定ダイアログを表示
                    '//--------------------------------------------------------
                    Dim dProxy As New SetProxy
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 dProxy.sProxyAdr = gsProxyAdr
'//                 dProxy.sProxyAdrUserSet = gsProxyAdrUserSet
'//                 dProxy.iProxyNo = giProxyNo
'//                 dProxy.iProxyNoUserSet = giProxyNoUserSet
                    dProxy.sProxyAdrUserSet  = gsProxyAdrUserSet
                    dProxy.iProxyNoUserSet   = giProxyNoUserSet
                    dProxy.bProxyOnUserSet   = gbProxyOnUserSet
                    dProxy.bProxyIdOnUserSet = gbProxyIdOnUserSet
                    dProxy.sProxyIdUserSet   = gsProxyIdUserSet
                    dProxy.sProxyPaUserSet   = gsProxyPaUserSet
                    dProxy.iConnectTimeOut   = giConnectTimeOut
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    dRet = dProxy.ShowDialog()
                    '[いいえ]ならそのまま終了
                    If dRet = Windows.Forms.DialogResult.No Then Exit Do
                    'キャンセルならそのまま終了
                    If dRet = Windows.Forms.DialogResult.Cancel Then Exit Do

                    gsProxyAdrUserSet = dProxy.sProxyAdrUserSet
                    giProxyNoUserSet = dProxy.iProxyNoUserSet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
                    gbProxyOnUserSet   = dProxy.bProxyOnUserSet
                    gbProxyIdOnUserSet = dProxy.bProxyIdOnUserSet
                    gsProxyIdUserSet   = dProxy.sProxyIdUserSet
                    gsProxyPaUserSet   = dProxy.sProxyPaUserSet
                    giConnectTimeOut   = dProxy.iConnectTimeOut
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                     .Target = "P設定(U)[" & 暗号化２(gsProxyAdrUserSet) _
'//                         & "][" & 暗号化２(giProxyNoUserSet.ToString("0000")) & "]"
                        .Target = "P設定(U)[" & gsProxyAdrUserSet _
                            & "][" & giProxyNoUserSet.ToString("0000") & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 '//--------------------------------------------------------
'//                 '// プロキシ設定ファイル[proxy.ini]の初期化
'//                 '//--------------------------------------------------------
'//                 If gsProxyAdrUserSet = "" And giProxyNoUserSet = 0 Then
'//                     'ファイルの書き込み
'//                     Dim sw As StreamWriter = Nothing
'//                     Try
'//                         'ファイルの書き込み
'//                         sw = File.CreateText(gsInitProxy)
'//                         sw.WriteLine("")
'//                         sw.WriteLine("")
'//                         sw.WriteLine("")
'//                         sw.Close()
'//                     Catch ex As Exception
'//                         With _myLogRecord
'//                             .Target = "ファイルの書き込み Exception:"
'//                             .Result = "NG"
'//                             .Remark = ex.Message
'//                         End With
'//                         _myLog.WriteLog(_myLogRecord)
'//                         If Not (sw Is Nothing) Then sw.Close()
'//                         Windows.Forms.MessageBox.Show( _
'//                          "ファイル[" & gsInitProxy & "]の書き込みに失敗しました　　　" & vbCrLf _
'//                          & "フォルダに書き込み権限を追加してください　　　", _
'//                          "is2", _
'//                          Windows.Forms.MessageBoxButtons.OK, _
'//                          Windows.Forms.MessageBoxIcon.Error)
'//                     End Try
'//                 End If
                    '//--------------------------------------------------------
                    '// プロキシ設定ファイル[proxy.ini]の書き込み
                    '//--------------------------------------------------------
                    Dim sw As StreamWriter = Nothing
                    Try
                        sw = File.CreateText(gsInitProxy)
                        If giConnectTimeOut = 0 Then
                            sw.WriteLine("")
                        Else
                            sw.WriteLine(giConnectTimeOut.ToString())
                        End If
                        sw.WriteLine(gsProxyAdrUserSet)
                        If giProxyNoUserSet = 0 Then
                            sw.WriteLine("")
                        Else
                            sw.WriteLine(giProxyNoUserSet.ToString())
                        End If
	                    If gbProxyOnUserSet Then
                            sw.WriteLine("1")
                        Else
                            sw.WriteLine("0")
                        End If
                        If gbProxyIdOnUserSet Then
                            sw.WriteLine("1")
                        Else
                            sw.WriteLine("0")
                        End If
                        sw.WriteLine(暗号化２(gsProxyIdUserSet))
                        sw.WriteLine(暗号化２(gsProxyPaUserSet))
                        sw.Close()
                    Catch ex As Exception
                        With _myLogRecord
                            .Target = "ファイルの書き込み Exception:"
                            .Result = "NG"
                            .Remark = ex.Message
                        End With
                        _myLog.WriteLog(_myLogRecord)
                        If Not (sw Is Nothing) Then sw.Close()
                        Windows.Forms.MessageBox.Show( _
                         "ファイル[" & gsInitProxy & "]の書き込みに失敗しました　　　" & vbCrLf _
                         & "フォルダに書き込み権限を追加してください　　　", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                    End Try
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    '//--------------------------------------------------------
                    '// プロキシ設定(ユーザ設定)
                    '//--------------------------------------------------------
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 iRet = プロキシ設定(gsProxyAdrUserSet, giProxyNoUserSet)
                    With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                     .Target = "P_1[" & 暗号化２(gbProxyOnUserSet.ToString()) _
'//                         & "] P_2[" & 暗号化２(gbProxyIdOnUserSet.ToString()) & "]"
                        .Target = "P_1[" & gbProxyOnUserSet.ToString() _
                            & "] P_2[" & gbProxyIdOnUserSet.ToString() & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    If gbProxyOnUserSet Then
                        If gbProxyIdOnUserSet Then
                            wRet = プロキシ設定２(gsProxyAdrUserSet, giProxyNoUserSet _
                                , gsProxyIdUserSet, gsProxyPaUserSet)
                        Else
                            wRet = プロキシ設定(gsProxyAdrUserSet, giProxyNoUserSet)
                        End If
                    Else
                        wRet = プロキシ設定("", 0)
                    End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                 If iRet = 1 Then myRet = True
'//                 If iRet = 1 Then
                    If wRet = WebExceptionStatus.Success Then
                        myRet = True
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                        gsProxyAdr = gsProxyAdrUserSet
                        giProxyNo = giProxyNoUserSet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                     'ファイルの書き込み
'//                     Dim sw As StreamWriter = Nothing
'//                     Try
'//                         'ファイルの書き込み
'//                         sw = File.CreateText(gsInitProxy)
'//タイムアウト
'//                         If giConnectTimeOut = 0 Then
'//                             sw.WriteLine("")
'//                         Else
'//                             sw.WriteLine(giConnectTimeOut.ToString())
'//                         End If
'//                         sw.WriteLine(gsProxyAdrUserSet)
'//                         If giProxyNoUserSet = 0 Then
'//                             sw.WriteLine("")
'//                         Else
'//                            sw.WriteLine(giProxyNoUserSet.ToString())
'//                         End If
'//                         sw.Close()
'//                     Catch ex As Exception
'//                         With _myLogRecord
'//                             .Target = "ファイルの書き込み Exception:"
'//                             .Result = "NG"
'//                             .Remark = ex.Message
'//                         End With
'//                         _myLog.WriteLog(_myLogRecord)
'//                         If Not (sw Is Nothing) Then sw.Close()
'//                         Windows.Forms.MessageBox.Show( _
'//                          "ファイル[" & gsInitProxy & "]の書き込みに失敗しました　　　" & vbCrLf _
'//                          & "フォルダに書き込み権限を追加してください　　　", _
'//                          "is2", _
'//                          Windows.Forms.MessageBoxButtons.OK, _
'//                          Windows.Forms.MessageBoxIcon.Error)
'//                     End Try
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                        Exit Do
                    End If
                Loop
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                    Return False
                End If
                sv_init.Timeout = 100 * 1000 '// デフォルト値の１００秒
'// ADD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 START
            End If
'// ADD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 END
'// ADD 2009.07.29 東都）高木 プロキシ対応 END

            'サーバーのバージョン情報の取得
            myRet = verCheck.XmlReadServer(sourceDirName, _DestinationConfigPath, ServerSerializer, ServerVersion)
            If Not myRet Then Return False

            'クライアントのバージョン情報の取得
            Dim cntRet As Boolean
            cntRet = verCheck.XmlReadClient(Path.Combine(_DestinationPath, "VersionFile.xml"), ClientSerializer, ClientVersion)

'// MOD 2005.05.10 東都）高木 メッセージの変更 START
'//         waitDlg.MainMsg = "ファイルを転送しています……" ' 処理の概要
            waitDlg.MainMsg = "しばらくお待ちください。．．．"
'// MOD 2005.05.10 東都）高木 メッセージの変更 END
            waitDlg.ProgressMax = ServerVersion.Length ' 全体の処理件数
            Dim verRet As Boolean
            For Each ServerFile In ServerVersion
                ' 処理中止かどうかをチェック
                If waitDlg.IsAborting = True Then
                    Return False
                End If

'// DEL 2005.05.10 東都）高木 メッセージの変更 START
'//              waitDlg.SubMsg = ServerFile.FileName
'// DEL 2005.05.10 東都）高木 メッセージの変更 END
                ' 進行状況ダイアログのメーターを設定
                waitDlg.ProgressMsg = _
                (CType((iCount * 100 / ServerVersion.Length), Integer)).ToString() & "%　" _
                & "（" + iCount.ToString() + " / " & ServerVersion.Length & "）"
                With _myLogRecord
                    .Status = "開始"
                    .Target = ServerFile.FileName & "のバージョンチェック"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                With _myLogRecord
                    .Status = ""
                    .SubStatus = "ファイルチェック"
                End With

                'ダウンロードディレクトリの設定
                Dim FileDirLength = ServerFile.FileName.Replace("\", "/").LastIndexOf("/")
                Dim FileDir As String = ""
                If FileDirLength <> -1 Then
                    FileDir = "/" & ServerFile.FileName.Substring(0, ServerFile.FileName.Replace("\", "/").LastIndexOf("/"))
                    myRet = myFileCopy.CheckOrMakeDestinationDir(constAppDir & FileDir)
                    If Not myRet Then Return False
                End If

                If cntRet Then
                    verRet = False
                    'クライアントにバージョンファイルがある場合は最新のみダウンロード
                    For Each ClientFile In ClientVersion
                        If ServerFile.FileName.Equals(ClientFile.FileName) Then
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 START
'//                         If ServerFile.TimeStamp > ClientFile.TimeStamp Or ServerFile.Size <> ClientFile.Size Then
                            'タイムスタンプは、年月日時分まで比較する（秒は比較しない）
                            If ServerFile.TimeStamp.Substring(0, 12) > ClientFile.TimeStamp.Substring(0, 12) _
                            Or ServerFile.Size <> ClientFile.Size Then
'// MOD 2007.11.28 東都）高木 タイムスタンプの秒比較廃止 END
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 START
'// MOD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 START
'//                             If Not( ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                And  ServerFile.TimeStamp.Substring(0,8).Equals("20081014")) Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                             If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                And gbInitProxyExists _
'//                                And ServerFile.TimeStamp.Substring(0, 8).Equals("20081014")) Then
                                If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
                                   And gbInitProxyExists _
                                   And (ServerFile.TimeStamp.Substring(0, 8).Equals("20081014") _
                                     Or ServerFile.TimeStamp.Substring(0, 8).Equals("20091019"))) Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 END
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 END
                                    'バージョンが違う場合ダウンロードを行う
                                    With _myLogRecord
                                        .Target = ServerFile.FileName & "の最新バージョンをダウンロードします"
                                        .Result = "OK"
                                        .Remark = ""
                                    End With
                                    _myLog.WriteLog(_myLogRecord)
                                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                    If Not myRet Then
                                        'ダウンロードに失敗した場合は起動を行わない
                                        Return False
                                    End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 START
                                End If
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 END
'// ADD 2005.07.22 東都）小童谷 START
                            Else
                                appFile = System.IO.Path.Combine(_DestinationAppPath, ClientFile.FileName)
                                If System.IO.File.Exists(appFile) Then
'// ADD 2007.10.26 東都）高木 ＣＡＢ対象のファイルはサイズ比較は行わない START
                                    If ClientFile.FileName <> "IS2Client.exe" _
                                    And ClientFile.FileName <> "is2AdminClient.exe" Then
'// ADD 2007.10.26 東都）高木 ＣＡＢ対象のファイルはサイズ比較は行わない END
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 START
'// MOD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 START
'//                                     If Not( ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                        And  ServerFile.TimeStamp.Substring(0,8).Equals("20081014")) Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//                                     If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                        And gbInitProxyExists _
'//                                        And ServerFile.TimeStamp.Substring(0, 8).Equals("20081014")) Then
                                        If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
                                           And gbInitProxyExists _
                                           And (ServerFile.TimeStamp.Substring(0, 8).Equals("20081014") _
                                             Or ServerFile.TimeStamp.Substring(0, 8).Equals("20091019"))) Then
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2009.10.03 東都）高木 [proxy.ini]が存在しない時、プロキシβ版機能停止 END
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 END
                                            'ダウンロードに失敗した場合
                                            Dim appfs As System.IO.FileStream = New System.IO.FileStream(appFile, System.IO.FileMode.Open)
                                            appSize = appfs.Length
                                            appfs.Close()
                                            If appSize <> ClientFile.Size Then
                                                With _myLogRecord
                                                    .Target = ServerFile.FileName & "の最新バージョンをダウンロードします"
                                                    .Result = "OK"
                                                    .Remark = ""
                                                End With
                                                _myLog.WriteLog(_myLogRecord)
                                                myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                                If Not myRet Then
                                                    'ダウンロードに失敗した場合は起動を行わない
                                                    Return False
                                                End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                                                DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END
                                            End If
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 START
                                        End If
'// MOD 2009.08.05 東都）高木 ブリヂストン様暫定対応 END
'// ADD 2007.10.26 東都）高木 ＣＡＢ対象のファイルはサイズ比較は行わない START
                                    End If
'// ADD 2007.10.26 東都）高木 ＣＡＢ対象のファイルはサイズ比較は行わない END
                                Else
                                    With _myLogRecord
                                        .Target = ServerFile.FileName & "の最新バージョンをダウンロードします"
                                        .Result = "OK"
                                        .Remark = ""
                                    End With
                                    _myLog.WriteLog(_myLogRecord)
                                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                    If Not myRet Then
                                        'ダウンロードに失敗した場合は起動を行わない
                                        Return False
                                    End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END
                                End If
'// ADD 2005.07.22 東都）小童谷 END
                            End If
                            verRet = True
                        End If
                    Next
                    If Not verRet Then
                        With _myLogRecord
                            .Target = "新規ファイル" & ServerFile.FileName & "をダウンロードします"
                            .Result = "OK"
                            .Remark = ""
                        End With
                        _myLog.WriteLog(_myLogRecord)
                        myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                        If Not myRet Then
                            'ダウンロードに失敗した場合は起動を行わない
                            Return False
                        End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                        DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END
                        verRet = True
                    Else
                        With _myLogRecord
                            .Target = ServerFile.FileName & "は最新バージョンです"
                            .Result = "OK"
                            .Remark = ""
                        End With
                        _myLog.WriteLog(_myLogRecord)
                    End If
                Else
                    With _myLogRecord
                        .Target = ServerFile.FileName & "をダウンロードします"
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    'クライアントにバージョンファイルがない場合は全てダウンロード
                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                    If Not myRet Then
                        'ダウンロードに失敗した場合は起動を行わない
                        Return False
                    End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END
                End If

'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 START
                If ServerFile.FileName.Substring(ServerFile.FileName.LastIndexOf(".") + 1).Equals("txt") Then
'// ADD 2007.10.12 東都）高木 ＣＡＢファイルの解凍 START
                    ChangeFileStamp(ServerFile.FileName, ServerFile.TimeStamp)
                    CheckCabVerFile(sourceDirName, ServerFile.FileName)
'// ADD 2007.10.12 東都）高木 ＣＡＢファイルの解凍 END

                    ' ＣＡＢファイルが存在し、CABフラグファイルがなければ、解凍をする
                    Dim srcFullName, flgFullName As String
                    srcFullName = Path.Combine(_DestinationPath, ServerFile.FileName.Substring(0, ServerFile.FileName.LastIndexOf(".") + 1) + "cab")
                    flgFullName = srcFullName + "_flg"
                    If System.IO.File.Exists(srcFullName) And System.IO.File.Exists(flgFullName) = False Then
                        '保留　アプリが起動していた場合には終了させる
                        'ＣＡＢファイルの解凍
                        myRet = UnCab(srcFullName, _DestinationAppPath)
                        'フラグファイルの作成
                        Dim fsflg As FileStream = System.IO.File.Create(flgFullName)
                        fsflg.Close()
                    End If
                End If
'// ADD 2007.10.05 東都）高木 ＣＡＢファイルの解凍 END

                ' 処理カウントを1ステップ進める
                iCount = iCount + 1
                waitDlg.PerformStep()
                waitDlg.Update()
                System.Windows.Forms.Application.DoEvents()
            Next

            If download Then
                'サーバーのバージョン情報をクライアントのバージョン情報にコピー
                Dim fs As New IO.FileStream(Path.Combine(_DestinationPath, "VersionFile.xml"), FileMode.Create)
                ServerSerializer.Serialize(fs, ServerVersion)
                fs.Close()
            End If

            'ローカルフォルダをチェックし不要ファイルを削除
            Dim localFile As New LocalFileCheck
            Dim localFileList As ArrayList
            Dim localFileName As String
            Dim localCheck As Boolean
            localFile.LocalFileList(_DestinationAppPath)
            localFileList = localFile.GetFileList

            For Each localFileName In localFileList
                localCheck = False
                For Each ServerFile In ServerVersion
'// MOD 2005.05.10 東都）高木 ファイル名比較の変更 START
'//                 If ServerFile.FileName.Replace("\", "/").Equals(localFileName.Substring(_DestinationAppPath.Length + 1).Replace("\", "/")) Then
                    If ServerFile.FileName.Replace("\", "/").ToUpper().Equals( _
                            localFileName.Substring(_DestinationAppPath.Length + 1).Replace("\", "/").ToUpper()) Then
'// MOD 2005.05.10 東都）高木 ファイル名比較の変更 END
                        localCheck = True
                    End If
                Next
                If Not localCheck Then
                    With _myLogRecord
                        .Target = "不要ファイル " & localFileName & "を削除します"
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    'ローカルファイルの削除
                    System.IO.File.Delete(localFileName)
                End If
            Next

        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
            ' 進行状況ダイアログを閉じる
            waitDlg.Close()
        End Try


        Return True
    End Function

    '
    '   アプリケーションの実行
    '
    Public Function ExecApp(ByVal sourceFileName As String, ByVal cmdPara() As String, ByVal localExecute As Boolean) As Boolean
        Dim myRet As Boolean
        Dim myAppDom As AppDomain

'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない START
'//     'アプリケーションドメインの設定
'//     myAppDom = GetAppDomain(sourceFileName, localExecute)
'//     If myAppDom Is Nothing Then Return False
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない END
'// ADD 2007.11.29 東都）高木 明示的ガベージコレクトを行う START
        GC.Collect()    'メモリークリーンアップ
'// ADD 2007.11.29 東都）高木 明示的ガベージコレクトを行う END

'// ADD 2009.07.30 東都）高木 exeのdll化対応 START
'//     'アプリケーションの実行
'//     myRet = Execute(myAppDom, sourceFileName, cmdPara, localExecute)
        Dim sDllFileName As String = ""
        Dim sDllFlagName As String = ""
        sDllFileName = sourceFileName.Substring(0, sourceFileName.Length - 4) + ".dll"
        sDllFlagName = sDllFileName & constConfigExt

        If System.IO.File.Exists(Path.Combine(_DestinationAppPath, sDllFlagName)) _
        And System.IO.File.Exists(Path.Combine(_DestinationAppPath, sDllFileName)) Then
            '//ＤＬＬの呼び出し
            myRet = ExecuteDll(sDllFileName, cmdPara)
        Else
            '//ＥＸＥの呼び出し
            myRet = Execute(myAppDom, sourceFileName, cmdPara, localExecute)
        End If
'// ADD 2009.07.30 東都）高木 exeのdll化対応 END
        Return myRet
    End Function

    '
    '   スタート・アップ処理
    '
    '   起動ファイルパス、ダウンロードするか、ダウンロードのみか、ローカル起動か、ログをクリアーするか
    '       
    Public Function Startup(ByVal sourceFileName As String, _
                            ByVal fileDir As String, _
                            ByVal download As Boolean, _
                            ByVal downloadOnly As Boolean, _
                            ByVal localExecute As Boolean, _
                            ByVal logClear As Boolean) As Boolean

        Dim myRet As Boolean

        Try
            Dim myFileCopy As FileCopy = New FileCopy

            With _myLogRecord
                .Status = "開始"
                .SubStatus = "AutoUpGrade"
                .Target = "処理を開始します"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            _myLogRecord.Status = ""

            'ダウンロードのみの場合のパラメータを設定
            '   ダウンロードは行う、ローカル起動は行わない
            If downloadOnly Then
                download = True
                localExecute = False
            End If
            'ファイル名チェック
            myRet = CheckFile(sourceFileName, myFileCopy)
            If myRet = False Then Return False

            'ファイルのダウンロード
            myRet = GetFile(sourceFileName, fileDir, download, localExecute, myFileCopy)
            If myRet = False Then Return False

'// DEL 2007.11.29 東都）高木 アセンブリダウンロードは負荷がかかるので行わない START
'//         'アセンブリの参照アセンブリをダウンロード
'//         If download And sourceFileName.Substring(sourceFileName.LastIndexOf(".") + 1).Equals("exe") Then
'//             myRet = CheckAssemblies()
'//             GC.Collect()    'メモリークリーンアップ
'//         End If
'// DEL 2007.11.29 東都）高木 アセンブリダウンロードは負荷がかかるので行わない END

            'ダウンロードのみ
            If downloadOnly Then
                With _myLogRecord
                    .Status = "終了"
                    .SubStatus = "AutoUpGrade"
                    If localExecute Then
                        .Target = sourceFileName & "のダウンロード処理に失敗しました"
                        .Result = "NG"
                        .Remark = ""
                        myRet = False
                    Else
                        .Target = sourceFileName & "のダウンロード処理が終了しました"
                        .Result = "OK"
                        .Remark = ""
                        myRet = True
                    End If
                End With
                _myLog.WriteLog(_myLogRecord)
            End If

        Catch ex As Exception
            With _myLogRecord
                .Target = ex.Message
                .Result = "EX"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            myRet = False
        End Try

        Return myRet

    End Function

    '
    '   ファイル名が正しいかどうかのチェック
    '
    Private Function CheckFile(ByVal sourceFileName As String, _
                               ByRef myFileCopy As FileCopy) As Boolean
        If sourceFileName = String.Empty Then
            With _myLogRecord
                .Target = "ファイル名が指定されていません"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If

        _FileName = myFileCopy.GetFileName(sourceFileName)
        If _FileName = String.Empty Then
            With _myLogRecord
                .Target = "ファイル名が指定されていません"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If

        Return True
    End Function


    '
    '   ファイルのダウンロードを行う
    '
    Private Function GetFile(ByVal sourceFileName As String, _
                        ByVal fileDir As String, _
                        ByRef download As Boolean, _
                        ByRef localExecute As Boolean, _
                        ByRef myFileCopy As FileCopy) As Boolean
        Dim myRet As Boolean = True

        If download Then
            'ファイルのダウンロード
            myRet = myFileCopy.CopyFile(sourceFileName, _DestinationAppPath & fileDir)
            _FileName = myFileCopy.FileName
            _SourcePath = myFileCopy.PathName
            If myRet Then
                With _myLogRecord
                    .Target = sourceFileName & "を" & _DestinationAppPath & fileDir & "へダウンロードに成功しました"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            Else
                With _myLogRecord
                    .Target = sourceFileName & "のダウンロードに失敗しました"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
        End If

        Return myRet

    End Function

    '
    '   アプリケーション・ドメインを作成します
    '
    Private Function GetAppDomain(ByVal sourceFileName As String, _
                                 ByVal localExecute As Boolean) As AppDomain
        Dim myAppSetup As AppDomainSetup
        Dim myAppDom As AppDomain
        Dim myName As String

        'アプリケーションドメインの設定
        myAppSetup = New AppDomainSetup
        myAppSetup.ConfigurationFile = Path.Combine(_DestinationAppPath, sourceFileName) & constConfigExt
'//     myAppSetup.ConfigurationFile = Path.Combine(_DestinationConfigPath, sourceFileName) & constConfigExt
        myAppSetup.ShadowCopyFiles = "true"
        If localExecute Then
            If Not File.Exists(Path.Combine(_DestinationAppPath, sourceFileName)) Then
                With _myLogRecord
                    .Status = "終了"
                    .Target = Path.Combine(_DestinationAppPath, sourceFileName) & "が見つかりません"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                _myLogRecord.Status = ""
                Return Nothing
            End If
            myAppSetup.ApplicationBase = _DestinationAppPath
        End If
        'アプリケーション・ドメイン名を実行中のアセンブリ名で取得
        myName = [Assembly].GetExecutingAssembly.FullName
        myName = myName.Substring(0, myName.IndexOf(","))
        myAppDom = AppDomain.CreateDomain(String.Concat(myName, " : ", sourceFileName), _
                                          AppDomain.CurrentDomain.Evidence, _
                                          myAppSetup)
        Return myAppDom
    End Function

    '
    '   アプリケーションを実行します
    '
    Private Function Execute(ByRef myAppDom As AppDomain, _
                            ByVal sourceFileName As String, _
                            ByVal cmdPara() As String, _
                            ByVal localExecute As Boolean) As Boolean
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない START
'//     Dim myName As String = myAppDom.FriendlyName
        Dim myName As String = " "
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない END

        'カレントディレクトリの変更
        'System.IO.Directory.SetCurrentDirectory(_DestinationAppPath)

        'アプリケーションの実行
        Try

            If localExecute Then
                With _myLogRecord
                    .Target = myName & "で" & Path.Combine(_DestinationAppPath, sourceFileName) & "を実行を行います"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                'myAppDom.ExecuteAssembly(Path.Combine(_DestinationAppPath, sourceFileName), myAppDom.Evidence, cmdPara)
'// MOD 2005.05.31 東都）伊賀 スレッドを廃止 START
                'スレッドを作成し、開始する
                Dim process As New System.Diagnostics.Process

                With process.StartInfo

                    .Arguments = cmdPara(0) ' コマンドライン引数
                    .WorkingDirectory = _DestinationAppPath ' 作業ディレクトリ
                    .FileName = Path.Combine(_DestinationAppPath, sourceFileName) ' 実行するファイル (*.exeでなくても良い)
'// ADD 2008.06.11 東都）高木 ログイン画面を常にノーマル表示に START
                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
'// ADD 2008.06.11 東都）高木 ログイン画面を常にノーマル表示に END
                End With
                Try
                    process.Start()

                Catch ex As System.ComponentModel.Win32Exception
                    ' ファイルが見つからなかった場合、
                    ' 関連付けられたアプリケーションが見つからなかった場合等

                Catch ex As System.Exception
                    'その他

                End Try
'// MOD 2005.05.31 東都）伊賀 スレッドを廃止 END
            Else
                With _myLogRecord
                    .Target = myName & "で" & sourceFileName & "を実行を行います"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない START
'//             myAppDom.ExecuteAssembly(sourceFileName)
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない END
            End If
            With _myLogRecord
                .Status = "終了"
                .Target = myName & "で実行が終了しました"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

        Catch e As Exception
            With _myLogRecord
                .Status = "終了"
                .Target = myName & "でエラーが発生しました"
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない START
'//         AppDomain.Unload(myAppDom)
'// DEL 2007.11.29 東都）高木 アプリケーションドメインの設定は負荷がかかるので行わない START
'// MOD 2005.05.31 東都）伊賀 スレッドを廃止 START
            Dim process As New System.Diagnostics.Process

            With process.StartInfo
                .Arguments = _ClientMutex ' コマンドライン引数
                .WorkingDirectory = _DestinationPath ' 作業ディレクトリ
'// MOD 2005.06.03 東都）伊賀 実行ファイルパスの変更 START
'//             '.FileName = Path.Combine(_DestinationPath, "CopyAutoUpGrade.exe") ' 実行するファイル (*.exeでなくても良い)
                .FileName = Path.Combine(_DestinationAppPath, "CopyAutoUpGrade.exe") ' 実行するファイル (*.exeでなくても良い)
'// MOD 2005.06.03 東都）伊賀 実行ファイルパスの変更 END
'// ADD 2008.06.11 東都）高木 [CopyAutoUpGrade]を非表示に START
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
'// ADD 2008.06.11 東都）高木 [CopyAutoUpGrade]を非表示に END
            End With
            Try
                process.Start()

            Catch ex As System.ComponentModel.Win32Exception
                ' ファイルが見つからなかった場合、
                ' 関連付けられたアプリケーションが見つからなかった場合等

            Catch ex As System.Exception
                'その他

            End Try
'// MOD 2005.05.31 東都）伊賀 スレッドを廃止 END

            'カレントディレクトリの変更
            'System.IO.Directory.SetCurrentDirectory(_DestinationPath)
        End Try

        Return True
    End Function

    '   アセンブリの参照情報を元に解決する
    '       
    Private Function CheckAssemblies() As Boolean
        Dim myRet As Boolean
        Dim myReference As ReferenceAssembly = New ReferenceAssembly
        With _myLogRecord
            .Target = "アセンブリ参照のチェックを開始します"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        myRet = myReference.SearchAssemblies(_FileName, _SourcePath, _DestinationAppPath)

        With _myLogRecord
            If myRet Then
                .Target = "アセンブリ参照のチェックを終了しました"
                .Result = "OK"
                .Remark = _FileName
            Else
                .Target = "アセンブリ参照のチェックでチェックできないものがありました"
                .Result = "NG"
                .Remark = _FileName
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myRet
    End Function

'// ADD 2007.06.15 東都）高木 ＣＡＢファイルの解凍 START
    '       
    '   ＣＡＢファイルの解凍する
    '       
    Private Function UnCab(ByVal cabFile As String, ByVal extractDir As String) As Boolean
        Dim myRet As Boolean

        myRet = False
        With _myLogRecord
            .Target = "ＣＡＢ解凍を開始します"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        'スレッドを作成し、開始する
        Dim process As New System.Diagnostics.Process

        With process.StartInfo
            .Arguments = " -r """ + cabFile + """" + " """ + extractDir + """" ' コマンドライン引数
            .WorkingDirectory = extractDir ' 作業ディレクトリ
            .FileName = "EXPAND.EXE"  ' 実行するファイル (*.exeでなくても良い)
'// ADD 2008.06.11 東都）高木 [EXPAND]を非表示に START
            '            .WindowStyle = ProcessWindowStyle.Minimized
            .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
'// ADD 2008.06.11 東都）高木 [EXPAND]を非表示に END
        End With

        Try
            process.Start()
'// ADD 2007.06.21 東都）高木 プロセスの終了を待つ START
            process.WaitForExit(25000) 'プロセス終了するか25秒経過するまで待つ
'// ADD 2007.06.21 東都）高木 プロセスの終了を待つ END
            myRet = True

        Catch ex As System.ComponentModel.Win32Exception
            ' ファイルが見つからなかった場合、
            ' 関連付けられたアプリケーションが見つからなかった場合等

        Catch ex As System.Exception
            'その他

        End Try

        With _myLogRecord
            If myRet Then
                .Target = "ＣＡＢ解凍が終了しました"
                .Result = "OK"
                .Remark = cabFile
            Else
                .Target = "ＣＡＢ解凍ができませんでした"
                .Result = "NG"
                .Remark = cabFile
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myRet
    End Function
'// ADD 2007.06.15 東都）高木 ＣＡＢファイルの解凍 END
'// ADD 2009.07.29 東都）高木 プロキシ対応 START
    '
    '   ＩＥのプロキシの設定内容をレジストリから取得
    '
    Private Sub ＩＥプロキシ設定情報取得()
        Try
            Dim regkey As Microsoft.Win32.RegistryKey = _
             Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Internet Settings", False)
            '//キーが存在しないときは null が返される
            If regkey Is Nothing Then Return

            '// 
            '// [自動構成スクリプトを使用する場合]のチェック有無
            '// 
            Dim sAutoConfigURL As String = ""
            If Not (regkey.GetValue("AutoConfigURL") Is Nothing) Then
                sAutoConfigURL = CType(regkey.GetValue("AutoConfigURL"), String)
                With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                 .Target = "R AC[" & 暗号化２(CType(regkey.GetValue("AutoConfigURL"), String)) & "]"
                    .Target = "R AC[" & CType(regkey.GetValue("AutoConfigURL"), String) & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            '// 
            '// [ＬＡＮにプロキシサーバを使用する場合]のチェック有無
            '// 
            Dim iProxyEnable As Integer = 0
            If Not (regkey.GetValue("ProxyEnable") Is Nothing) Then
                iProxyEnable = CType(regkey.GetValue("ProxyEnable"), Integer)
                With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                 .Target = "R PE[" & 暗号化２(iProxyEnable.ToString()) & "]"
                    .Target = "R PE[" & iProxyEnable.ToString() & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            Dim sProxyServer As String() = {""}
            If Not (regkey.GetValue("ProxyServer") Is Nothing) Then
                sProxyServer = CType(regkey.GetValue("ProxyServer"), String).Split(";")
                With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                 .Target = "R PS[" & 暗号化２(CType(regkey.GetValue("ProxyServer"), String)) & "]"
                    .Target = "R PS[" & CType(regkey.GetValue("ProxyServer"), String) & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If

            Dim sProxyOverride As String() = {""}
            If Not (regkey.GetValue("ProxyOverride") Is Nothing) Then
                sProxyOverride = CType(regkey.GetValue("ProxyOverride"), String).Split(";")
                With _myLogRecord
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 START
'//                 .Target = "R PO[" & 暗号化２(CType(regkey.GetValue("ProxyOverride"), String)) & "]"
                    .Target = "R PO[" & CType(regkey.GetValue("ProxyOverride"), String) & "]"
'// MOD 2010.06.21 東都）高木 ログの暗号化解除 END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            regkey.Close()
            If iProxyEnable = 1 Then
                For iCnt As Integer = 0 To sProxyServer.Length - 1
                    Dim sParam As String() = sProxyServer(iCnt).Split(":")
                    If sParam(0).StartsWith("https=") Then
                        gsProxyAdrSecure = sParam(0).Substring(6)
                        If sParam.Length >= 2 Then giProxyNoSecure = Integer.Parse(sParam(1))
                    ElseIf sParam(0).StartsWith("http=") Then
                        gsProxyAdrHttp = sParam(0).Substring(5)
                        If sParam.Length >= 2 Then giProxyNoHttp = Integer.Parse(sParam(1))
                    ElseIf sParam(0).StartsWith("ftp=") Then
                        ';
                    ElseIf sParam(0).StartsWith("socks=") Then
                        ';
                    Else
                        gsProxyAdrAll = sParam(0)
                        If sParam.Length >= 2 Then giProxyNoAll = Integer.Parse(sParam(1))
                    End If
                Next
                For iCnt As Integer = 0 To sProxyOverride.Length - 1
                    If sProxyOverride(iCnt).Equals("wwwis2.fukutsu.co.jp") Then Return
                    If sProxyOverride(iCnt).Equals("fukutsu.co.jp") Then Return
                Next
                If gsProxyAdrAll.Length > 0 Then
                    gsProxyAdr = gsProxyAdrAll
                    giProxyNo = giProxyNoAll
                ElseIf gsProxyAdrSecure.Length > 0 Then
                    gsProxyAdr = gsProxyAdrSecure
                    giProxyNo = giProxyNoSecure
                ElseIf gsProxyAdrHttp.Length > 0 Then
                    gsProxyAdr = gsProxyAdrHttp
                    giProxyNo = giProxyNoHttp
                End If
            End If
        Catch ex As Exception
            With _myLogRecord
                .Target = "IEP設定情報取得 Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            Return
        End Try
    End Sub
    '
    '   デフォルトプロキシの設定を行う
    '
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'// Private Function デフォルトプロキシ設定() As Integer
    Private Function デフォルトプロキシ設定() As WebExceptionStatus
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        '
        '   デフォルトプロキシの設定を行う
        '
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//     Dim iRet As Integer = 0
        Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Try
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()

            Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = 1
            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Catch ex As System.Net.WebException
            With _myLogRecord
                .Target = "P設定(D) WebException:" & ex.Status.ToString()
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = -1
'//         Select Case ex.Status
'//             Case WebExceptionStatus.NameResolutionFailure
'//                 iRet = -11
'//             Case WebExceptionStatus.Timeout
'//                 iRet = -12
'//             Case WebExceptionStatus.TrustFailure
'//                 iRet = -13
'//             Case WebExceptionStatus.ConnectFailure
'//                 iRet = -14
'//             Case Else
'//                 iRet = -19
'//         End Select
'//         Return iRet
            wRet = ex.Status
            Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Catch ex As Exception
            With _myLogRecord
                .Target = "P設定(D) Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = -2
'//         Return iRet
            wRet = WebExceptionStatus.UnknownError
            Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        End Try

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//     Return iRet
        Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
    End Function

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'// Private Function プロキシ設定(ByVal sProxyAdr As String, ByVal iProxyNo As Integer) As Integer
    Private Function プロキシ設定(ByVal sProxyAdr As String, ByVal iProxyNo As Integer) As WebExceptionStatus
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
        Return プロキシ設定２(sProxyAdr, iProxyNo, "", "")
    End Function

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'// Private Function プロキシ設定２(ByVal sProxyAdr As String, ByVal iProxyNo As Integer _
'//         , ByVal sProxyId As String, ByVal sProxyPa As String) As Integer
    Private Function プロキシ設定２(ByVal sProxyAdr As String, ByVal iProxyNo As Integer _
            , ByVal sProxyId As String, ByVal sProxyPa As String) As WebExceptionStatus
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//     Dim iRet As Integer = 0
        Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Try
            If sProxyAdr.Length > 0 Then
                If iProxyNo > 0 Then
                    gProxy = New System.Net.WebProxy(sProxyAdr, iProxyNo)
                Else
                    gProxy = New System.Net.WebProxy(sProxyAdr)
                End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
                If sProxyId.Length > 0 Then
                    gProxy.Credentials = New NetworkCredential(sProxyId, sProxyPa)
                End If
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
                System.Net.GlobalProxySelection.Select = gProxy
                sv_init.Proxy = gProxy
            Else
                System.Net.GlobalProxySelection.Select = System.Net.GlobalProxySelection.GetEmptyWebProxy()
                sv_init.Proxy = System.Net.GlobalProxySelection.GetEmptyWebProxy()
            End If

'//			try
            Dim sRet As String = sv_init.wakeupDB()
'//			catch ex as Exception
'//				With _myLogRecord
'//					.Target = "プロキシ設定 Exception:"
'//					.Result = "NG"
'//					.Remark = ex.Message
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'//				return false
'//			end try

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = 1
            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Catch ex As System.Net.WebException
            With _myLogRecord
                .Target = "P設定( ) WebException:" & ex.Status.ToString()
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            '//元の状態に戻す
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = -1
'//         Select Case ex.Status
'//             Case WebExceptionStatus.NameResolutionFailure
'//                 iRet = -11
'//             Case WebExceptionStatus.Timeout
'//                 iRet = -12
'//             Case WebExceptionStatus.TrustFailure
'//                 iRet = -13
'//             Case WebExceptionStatus.ConnectFailure
'//                 iRet = -14
'//             Case Else
'//                 iRet = -19
'//         End Select
'//         Return iRet
            wRet = ex.Status
            Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        Catch ex As Exception
            With _myLogRecord
                .Target = "P設定( ) Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            '//元の状態に戻す
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//         iRet = -2
'//         Return iRet
            wRet = WebExceptionStatus.UnknownError
            Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
        End Try

'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//     Return iRet
        Return wRet
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
    End Function
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
    '
    '   接続エラーメッセージの編集
    '
    Function 接続エラーメッセージ編集(wStatus As WebExceptionStatus) As String
        Dim sRet As String = ""
        Select Case wStatus
           Case WebExceptionStatus.ConnectFailure
				'// プロキシサーバ設定エラー
                sRet = "原因：プロキシ設定など"
           Case WebExceptionStatus.ConnectionClosed 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.KeepAliveFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.MessageLengthLimitExceeded
                sRet = wStatus.ToString()
           Case WebExceptionStatus.NameResolutionFailure 
				'// ケーブル接続エラー
				'// プロキシポート番号設定エラー
				'// ＤＮＳ設定エラー
                sRet = "原因：ケーブル接続、プロキシ設定、ＤＮＳ設定など"
           Case WebExceptionStatus.Pending 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.PipelineFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.ProtocolError 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.ProxyNameResolutionFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.ReceiveFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.RequestCanceled 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.SecureChannelFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.SendFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.ServerProtocolViolation 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.Success 
                sRet = "接続成功"
           Case WebExceptionStatus.Timeout 
                sRet = "タイムアウトエラー、原因：プロキシ設定など"
           Case WebExceptionStatus.TrustFailure 
				'// 日付設定エラー
                sRet = "原因：日付設定、ＳＳＬ設定など"
           Case WebExceptionStatus.UnknownError
                sRet = "未定義エラー"
        End Select
        Return sRet        
    End Function
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
    '
    '   ＳＪＩＳチェックを行う
    '
    Private Function ＳＪＩＳチェック(ByVal tex As String, ByVal name As String, ByRef sUnicode As String, ByRef bSjis As Byte()) As Boolean
        '//逆変換してＳＪＩＳ文字をチェックする
        Dim sRevUnicode As String = System.Text.Encoding.GetEncoding("shift-jis").GetString(bSjis)
        Dim sErrChars As String = ""
        For iPos As Integer = 0 To sUnicode.Length - 1
            If iPos >= sRevUnicode.Length Then Exit For
            If sUnicode.Chars(iPos) <> sRevUnicode.Chars(iPos) Then
                sErrChars += sUnicode.Chars(iPos)
            End If
        Next
        If sErrChars.Length > 0 Then
            '//				MessageBox.Show(name + "に使用できない文字があります\n" _
            '//					+ "『" + sErrChars + "』", _
            '//					"入力チェック",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "ＳＪＩＳチェック[" & name & "][" & sErrChars & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If
        Return True
    End Function
    '
    '   半角チェックを行う
    '
    Protected Function 半角チェック(ByVal tex As String, ByVal name As String) As Boolean
        Dim sUnicode As String = tex
        Dim bSjis As Byte() = System.Text.Encoding.GetEncoding("shift-jis").GetBytes(sUnicode)
        If Not ＳＪＩＳチェック(tex, name, sUnicode, bSjis) Then Return False
        If bSjis.Length <> sUnicode.Length Then
            '//			MessageBox.Show(name + "は半角文字で入力してください", _
            '//				"入力チェック",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "半角チェックエラー[" & name & "][" & tex & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If

'// MOD 2011.09.22 東都）高木 記号チェック廃止 START
'//		For i As Integer = 0 To tex.Length - 1
'//			'// [!"#$%&'()*+,]
'//			'// [:;<=>?]
'//			'// [[]^]
'//			'// [{|}]
'//			'// [\]
'//			If (tex.Chars(i) >= "!" And tex.Chars(i) <= ",") _
'//			 Or (tex.Chars(i) >= ":" And tex.Chars(i) <= "?") _
'//			 Or (tex.Chars(i) >= "[" And tex.Chars(i) <= "^") _
'//			 Or (tex.Chars(i) >= "{" And tex.Chars(i) <= "}") _
'//			 Or (tex.Chars(i) = "\") Then
'//				'//				MessageBox.Show(name + "に記号が入力されています","入力チェック",MessageBoxButtons.OK)
'//				With _myLogRecord
'//					.Target = "半角記号チェックエラー[" & name & "][" & tex & "]"
'//					.Result = "NG"
'//					.Remark = ""
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'//
'//				Return False
'//			End If
'//		Next
'// MOD 2011.09.22 東都）高木 記号チェック廃止 END
        Return True
    End Function
    '
    '   数値チェックを行う
    '
    Protected Function 数値チェック(ByVal tex As String, ByVal name As String) As Boolean
        Try
            Dim lChk As Long = Long.Parse(tex.Replace(",", ""))
            Return True
        Catch ex As Exception
            '//			MessageBox.Show(name + "に数値が入力されていません","入力チェック",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "数値チェックエラー[" & name & "][" & tex & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End Try
    End Function
'// ADD 2009.07.29 東都）高木 プロキシ対応 END

'// ADD 2009.07.30 東都）高木 exeのdll化対応 
    '
    '   dllの実行
    '
    Public Function ExecuteDll(ByVal sDllFileName As String, ByVal cmdPara() As String) As Boolean
        Dim myRet As Boolean
        Try
            'カレントディレクトリの変更
            System.IO.Directory.SetCurrentDirectory(_DestinationAppPath)

            With _myLogRecord
                .Target = "[" & sDllFileName & "]読込"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Dim asm As System.Reflection.Assembly
            asm = System.Reflection.Assembly.LoadFrom(sDllFileName)

            'ライブラリ読込み用変数
            Dim oLibrary As New Object
            '呼出す関数への引数
            Dim methodParams() As Object = {cmdPara}

            With _myLogRecord
                .Target = "[" & sDllFileName & "]呼出"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            'クラスをインスタンス化
            oLibrary = asm.CreateInstance("IS2Client.メニュー")

            'メソッド情報を取得する
            Dim mi As System.Reflection.MethodInfo = oLibrary.GetType.GetMethod("Main")

            'メソッドを実行する
            mi.Invoke(oLibrary, methodParams)

        Catch ex As System.IO.FileNotFoundException
            With _myLogRecord
                .Target = "ExecApp " & ex.FileName & "がみつかりませんでした"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)
        Catch ex As Exception
            With _myLogRecord
                .Target = "ExecApp Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)
        Finally
            'カレントディレクトリの変更
            System.IO.Directory.SetCurrentDirectory(_DestinationPath)

            Dim process As New System.Diagnostics.Process

            With process.StartInfo
                .Arguments = _ClientMutex ' コマンドライン引数
                .WorkingDirectory = _DestinationPath ' 作業ディレクトリ
                .FileName = Path.Combine(_DestinationAppPath, "CopyAutoUpGrade.exe")
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            End With
            Try
                process.Start()
            Catch ex As System.ComponentModel.Win32Exception
                ' ファイルが見つからなかった場合、
                ' 関連付けられたアプリケーションが見つからなかった場合等
            Catch ex As System.Exception
                'その他
            End Try
        End Try
        Return myRet
    End Function
'// ADD 2009.07.30 東都）高木 exeのdll化対応 END
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
    Dim bKey As Byte() = {51, 168, 96, 2, 36, 12, 74, 143, 11, 85, 61, 233, 202, 170, 114, 59}
    Dim bIv As Byte() = {100, 223, 207, 80, 29, 100, 53, 152}
    Dim sダミー As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz01234567890!#$%&()=~|`{+*}<>?_"
    Dim rndObj As New Random
    Function 暗号化２(ByVal sText As String) As String
        '//先頭および末尾にダミー文字を追加
        sText = sダミー.Substring(rndObj.Next(0, sダミー.Length - 1), 1) + sText
        sText = sText + sダミー.Substring(rndObj.Next(0, sダミー.Length - 1), 1)

        Return 暗号化(sText)
    End Function
    Function 暗号化(ByVal sText As String) As String
        Dim bText As Byte()
        Dim bRet As Byte()
        Dim sRet As String = ""

        Try
            Dim msEncrypt As MemoryStream = New MemoryStream
            Dim rc2 As RC2CryptoServiceProvider = New RC2CryptoServiceProvider

            '// CryptoStreamオブジェクトを作成する
            Dim transform As ICryptoTransform = rc2.CreateEncryptor(bKey, bIv) '// Encryptorを作成する
            Dim cryptoStream As cryptoStream = New cryptoStream(msEncrypt, transform, CryptoStreamMode.Write)

            '// 暗号化する対象をバイト配列として読み込む
            bText = Encoding.GetEncoding("shift-jis").GetBytes(sText)

            '// CryptoStreamによって暗号化して書き込む
            cryptoStream.Write(bText, 0, bText.Length)
            cryptoStream.FlushFinalBlock()

            bRet = msEncrypt.ToArray()
            '//sRet = Encoding.GetEncoding("shift-jis").GetString(bRet)
            For iCnt As Integer = 0 To bRet.Length - 1
                sRet = sRet + bRet(iCnt).ToString("X2")
            Next iCnt

            '// CryptoStreamを閉じる
            cryptoStream.Close()

            '// FileStreamを閉じる
            msEncrypt.Close()

            rc2 = Nothing
            cryptoStream = Nothing
            msEncrypt = Nothing
        Catch ex As Exception
            sRet = ex.Message
        End Try

        Return sRet
    End Function
    Function 復号化(ByVal sText As String) As String
	    Return 復号化Ａ(sText, False)
    End Function
    Function 復号化２(ByVal sText As String) As String
	    Return 復号化Ａ(sText, True)
    End Function
    Function 復号化Ａ(ByVal sText As String, ByVal bDummy As Boolean) As String
        Dim bText As Byte()
        Dim bRet As Byte()
        Dim sRet As String = ""

        Try
            '//bText = Encoding.GetEncoding("shift-jis").GetBytes(sText)
            Dim [sByte] As String = ""
            ReDim bText(sText.Length / 2 - 1)
            For iCnt As Integer = 0 To sText.Length - 1 Step 2
                [sByte] = sText.Substring(iCnt, 2)
                bText(iCnt / 2) = Convert.ToByte([sByte], 16)
            Next

            Dim inputStream As MemoryStream = New MemoryStream(bText)
            Dim rc2 As RC2CryptoServiceProvider = New RC2CryptoServiceProvider

            '// CryptoStreamオブジェクトを作成する
            Dim transform As ICryptoTransform = rc2.CreateDecryptor(bKey, bIv) '// Decryptorを作成する
            Dim csDecrypt As CryptoStream = New CryptoStream(inputStream, transform, CryptoStreamMode.Read)

            ReDim bRet(bText.Length - 1)
            '//Read the data out of the crypto stream.
            Dim iLen As Integer = csDecrypt.Read(bRet, 0, bRet.Length)

            '//Convert the byte array back into a string.
            sRet = Encoding.GetEncoding("shift-jis").GetString(bRet, 0, iLen)

            '// ストリームを閉じる
            csDecrypt.Close()
            inputStream.Close()

            rc2 = Nothing
            csDecrypt = Nothing
            inputStream = Nothing

			If bDummy And sRet.Length >= 2 Then
				'//先頭および末尾のダミー文字を削除
				sRet = sRet.Substring(1,sRet.Length-2)
			End If

        Catch ex As Exception
            sRet = ex.Message
        End Try

        Return sRet
    End Function
''// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
End Class
