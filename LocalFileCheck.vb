Public Class LocalFileCheck

    Private _FileList As ArrayList

    '
    'コンストラクタ
    '
    Public Sub New()
        MyBase.New()
        '変数の初期化
        _FileList = New ArrayList
    End Sub

    Public ReadOnly Property GetFileList()
        Get
            Return _FileList
        End Get
    End Property

    Protected Friend Function LocalFileList(ByVal path As String) As Boolean
        Dim directories As String()
        Dim directory As String
        Dim files As String()
        Dim fileName As String
        Dim i As Integer

        files = System.IO.Directory.GetFiles(path)
        For Each fileName In files
            If Not fileName.EndsWith(".dll") Then
                _FileList.Add(fileName)
            End If
        Next
        directories = System.IO.Directory.GetDirectories(path)
        For Each directory In directories
            LocalFileList(directory)
        Next
    End Function

End Class
