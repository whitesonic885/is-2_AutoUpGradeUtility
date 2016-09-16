Public Class VersionCheck

#Region "ログ出力"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    Public Class FileVersion
        Public FileName As String
        Public TimeStamp As String
        Public Size As Integer
    End Class


    Public Sub New()
    End Sub

    Protected Friend Function XmlReadServer(ByRef SourceDirName As String, ByVal DestinationConfigPath As String _
                                , ByRef Serializer As System.Xml.Serialization.XmlSerializer _
                                , ByRef fileVer() As VersionCheck.FileVersion) As Boolean
        Dim myRet As Boolean

        Dim myFileCopy As FileCopy = New FileCopy
        Dim FileName As String = "VersionFile.xml"
        Dim myCheckFileName As String = System.IO.Path.Combine(DestinationConfigPath, myFileCopy.GetFileName(FileName))

        'サーバーのバージョン情報のダウンロード
        myRet = myFileCopy.CopyFile(SourceDirName & FileName, DestinationConfigPath)
        With _myLogRecord
            If myRet Then
                .Target = "サーバー:" & FileName & "を読込ました"
                .Result = "OK"
                .Remark = ""
                _myLog.WriteLog(_myLogRecord)
            Else
                SourceDirName = "http://" + SourceDirName.Substring(8)
                .Target = "httpに切り替えました"
                .Result = "OK"
                .Remark = ""
                _myLog.WriteLog(_myLogRecord)
                '// ADD 2005.06.24 東都）伊賀 httpsとhttpの切り替え START
                myRet = myFileCopy.CopyFile(SourceDirName & FileName, DestinationConfigPath)
                If myRet Then
                    .Target = "サーバー:" & FileName & "を読込ました"
                    .Result = "OK"
                    .Remark = ""
                    _myLog.WriteLog(_myLogRecord)
                Else
                    Return False
                End If
                '// ADD 2005.06.24 東都）伊賀 httpsとhttpの切り替え END
            End If
        End With

        'XmlSerializerオブジェクトの作成
        Serializer = New System.Xml.Serialization.XmlSerializer(GetType(FileVersion()))
        'ファイルを開く
        Dim fs As System.IO.FileStream
        Try
            fs = New System.IO.FileStream(System.IO.Path.Combine(DestinationConfigPath, FileName), System.IO.FileMode.Open)
        Catch e As Exception
            With _myLogRecord
                .Target = "サーバー:" & FileName & "読込エラー"
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try

        fileVer = CType(Serializer.Deserialize(fs), FileVersion())
        fs.Close()

        'サーバーのバージョン情報の削除
        System.IO.File.Delete(System.IO.Path.Combine(DestinationConfigPath, FileName))

        Return True
    End Function

    Protected Friend Function XmlReadClient(ByVal fileName As String _
                                , ByRef Serializer As System.Xml.Serialization.XmlSerializer _
                                , ByRef fileVer() As VersionCheck.FileVersion) As Boolean
        'ファイルを開く
        Dim fs As System.IO.FileStream

        If Not System.IO.File.Exists(fileName) Then Return False

        Try
            fs = New System.IO.FileStream(fileName, System.IO.FileMode.Open)
            With _myLogRecord
                .Target = "クライアント:" & fileName & "を読込ました"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
        Catch e As Exception
            With _myLogRecord
                .Target = "クライアント:" & fileName & "読込エラー"
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try

        fileVer = CType(Serializer.Deserialize(fs), FileVersion())
        fs.Close()

        Return True
    End Function

End Class
