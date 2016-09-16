Public Class LocalFileCheck

    Private _FileList As ArrayList

    '
    'ƒRƒ“ƒXƒgƒ‰ƒNƒ^
    '
    Public Sub New()
        MyBase.New()
        '•Ï”‚Ì‰Šú‰»
        _FileList = New ArrayList
    End Sub

    '// MOD 2016.09.16 Vivouacj‹e’r Visual Studio 2013•ÏŠ·‚É”º‚¤C³ START
    '//Public ReadOnly Property GetFileList()
    Public ReadOnly Property GetFileList() As ArrayList
        '// MOD 2016.09.16 Vivouacj‹e’r Visual Studio 2013•ÏŠ·‚É”º‚¤C³ END
        Get
            Return _FileList
        End Get
    End Property

    Protected Friend Function LocalFileList(ByVal path As String) As Boolean
        Dim directories As String()
        Dim directory As String
        Dim files As String()
        Dim fileName As String
        '//Dim i As Integer

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
