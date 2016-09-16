'
'   �t�@�C���E�R�s�[�ׂ̈̔ėp�N���X
'   �t�@�C���V�F�A��http�ł̃R�s�[
Imports System.IO
Imports System.Net

Public Class FileCopy

    Private _ErrorMessage As String
    Private _PathName As String
    Private _FileName As String
    Private _TimeOut As Integer     'https��5��
    Private _TimeOut2 As Integer    'http��30�b

    Private constDirSeparator As String = Path.DirectorySeparatorChar.ToString()
    Private constVolumeSeparator As String = Path.VolumeSeparatorChar.ToString()

#Region "���O�o��"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    '
    '�R���X�g���N�^
    '
    Public Sub New()
        MyBase.New()
        '�ϐ��̏�����
        _ErrorMessage = ""
        _PathName = ""
        _FileName = ""
        'MOD-S 20050829 �^�C���A�E�g�R�O�b���T��
        '_TimeOut = 30000
        _TimeOut = 300000
        _TimeOut2 = 30000
        'MOD-E 20050829 �^�C���A�E�g�R�O�b���T��
        With _myLogRecord
            .Status = ""
            .SubStatus = "�t�@�C���`�F�b�N"
        End With
    End Sub

    Public Sub New(ByVal TimeOut As Integer)
        Me.New()
        _TimeOut = TimeOut

    End Sub

    '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� START
    '//Public ReadOnly Property ErrorMesage()
    Public ReadOnly Property ErrorMesage() As String
        '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� END
        Get
            Return _ErrorMessage
        End Get
    End Property

    '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� START
    '//Public ReadOnly Property PathName()
    Public ReadOnly Property PathName() As String
        '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� END
        Get
            Return _PathName
        End Get
    End Property

    '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� START
    '//Public ReadOnly Property FileName()
    Public ReadOnly Property FileName() As String
        '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� END
        Get
            Return _FileName
        End Get
    End Property

    '   �w�肳�ꂽ�t�@�C�����w��f�B���N�g���ɃR�s�[����
    '       �w��f�B���N�g�����w�肳��Ă��Ȃ���΁A�t�@�C�����̃`�F�b�N�̂�
    '       �O������ďo���G���g���[�|�C���g�ɂȂ�܂�
    Protected Friend Function CopyFile(ByVal sourceFileName As String, ByVal destPathName As String) As Boolean
        Dim blnRet As Boolean
        _ErrorMessage = ""
        _PathName = ""
        _FileName = ""

        With _myLogRecord
            .Target = sourceFileName & "��" & destPathName & "�փR�s�[���J�n���܂���"
            .Result = "OK"
            .Remark = "�R�s�[�����̊J�n"
        End With
        _myLog.WriteLog(_myLogRecord)
        '�t�@�C�����̎w����@�̃`�F�b�N
        blnRet = CheckFileName(sourceFileName)
        If blnRet = False Then Return False

        If destPathName = String.Empty Then Return blnRet
        '�t�@�C���̃R�s�[
        blnRet = GetFileToCopy(sourceFileName, destPathName)
        With _myLogRecord
            If blnRet Then
                .Target = sourceFileName & "��" & destPathName & "�փR�s�[�������܂���"
                .Result = "OK"
                .Remark = "�R�s�[�����̏I��"
            Else
                .Target = sourceFileName & "��" & destPathName & "�ւ̃R�s�[�Ɏ��s���܂���"
                .Result = "NG"
                .Remark = "�R�s�[�����̏I��"
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return blnRet

    End Function

    '
    '   �t�@�C�������ȉ��̋K���ɏ]���Ă��邩���`�F�b�N�ɂ��܂�
    '       �p�X�ƃt�@�C�������v���p�e�B�ɐݒ肵�܂�
    '       c:\�̃h���C�u���^�[�����A\\savername�̋��L�����Ahttp:��URL����
    Private Function CheckFileName(ByVal fullPathFileName As String) As Boolean
        If fullPathFileName = String.Empty Then
            _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
            With _myLogRecord
                .Target = fullPathFileName & _ErrorMessage
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If
        '�p�X�̊J�n�����ƃt�@�C���̏I���������`�F�b�N
        If CheckPathStart(fullPathFileName) = False Then Return False
        If CheckPathEnd(fullPathFileName) = False Then Return False
        '�t�@�C�������擾
        Dim FileName As String
        FileName = GetFileName(fullPathFileName)
        If FileName = String.Empty Then Return False
        '���������p�X�ƃt�@�C�������v���p�e�B�ɐݒ�
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
    '   �p�X�`����������������Ŏn�܂��Ă��邩���`�F�b�N
    '       http://�Ahttps://�A\\�Ax:\�Ax:�A\�A.
    '
    Private Function CheckPathStart(ByVal fullPathFileName As String) As Boolean
        If CheckPathIsUrl(fullPathFileName) Then Return True
        If fullPathFileName.StartsWith(String.Concat(constDirSeparator, constDirSeparator)) Then Return True
        If fullPathFileName.StartsWith(constDirSeparator) Then Return True
        If fullPathFileName.StartsWith(".") Then Return True
        If fullPathFileName.Substring(1, 2) = String.Concat(constVolumeSeparator, constDirSeparator) Then Return True
        If fullPathFileName.Substring(1, 1) = constVolumeSeparator Then Return True

        _ErrorMessage = "�������p�X������Ŏn�܂��Ă��܂���"
        With _myLogRecord
            .Target = fullPathFileName & "��" & _ErrorMessage
            .Result = "NG"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return False

    End Function

    '
    '   �p�X������/��\��.�ŏI�����Ă��Ȃ������`�F�b�N
    '
    Private Function CheckPathEnd(ByVal fullPathFileName As String) As Boolean
        If fullPathFileName.EndsWith("/") = False Then
            If fullPathFileName.EndsWith(constDirSeparator) = False Then
                If fullPathFileName.EndsWith(".") = False Then Return True
            End If
        End If
        _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
        With _myLogRecord
            .Target = fullPathFileName & "��" & _ErrorMessage
            .Result = "NG"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        Return False

    End Function

    '
    '   �p�X��URL�`���Ŏn�܂邩�ǂ������`�F�b�N
    '       http://����https://
    Protected Friend Function CheckPathIsUrl(ByVal fullPathFileName As String) As Boolean
        Dim FileName As String
        FileName = fullPathFileName.ToLower()
        If FileName.StartsWith("http://") Then Return True
        If FileName.StartsWith("https://") Then Return True

        Return False
    End Function

    '
    '   �t�@�C���̃R�s�[�̍쐬
    '       �^�[�Q�b�g�̃t�@�C�����ƃR�s�[��̃f�B���N�g��
    Private Function GetFileToCopy(ByVal sourceFileName As String, ByVal destinationPathName As String) As Boolean
        Dim blnRet As Boolean
        Dim IsUrl As Boolean = CheckPathIsUrl(sourceFileName)
        Dim FileName As String = GetFileName(sourceFileName)
        If FileName = String.Empty Then Return False
        blnRet = CheckOrMakeDestinationDir(destinationPathName)
        If blnRet = False Then Return False

        Dim DestinationFile As String = Path.Combine(destinationPathName, FileName)
        If IsUrl Then
            'http�ł̃_�E�����[�h
            With _myLogRecord
                .Target = sourceFileName & "��" & DestinationFile & "��http�_�E�����[�h�J�n"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            blnRet = GetHttpFileDownload(sourceFileName, DestinationFile)
        Else
            '�t�@�C���̃R�s�[
            With _myLogRecord
                .Target = sourceFileName & "��" & DestinationFile & "�փt�@�C���R�s�[�J�n"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            blnRet = GetFileDownload(sourceFileName, DestinationFile)
        End If

        Return blnRet

    End Function

    '
    '   http�Ńt�@�C�����_�E�����[�h���āA�t�@�C����ۑ����܂��B
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

        'http���N�G�X�g�̍쐬
        myHttpRequest = CType(WebRequest.Create(sourceUrl), HttpWebRequest)
        'mod-s 20050829 �^�C���A�E�g�@https��5���Ahttp��30�b
        'myHttpRequest.Timeout = _TimeOut
        If sourceUrl.Substring(0, 5) = "https" Then
            myHttpRequest.Timeout = _TimeOut
        Else
            myHttpRequest.Timeout = _TimeOut2
        End If
        'mod-e 20050829 �^�C���A�E�g�@https��5���Ahttp��30�b
        Try
            'http���X�|���X�̎擾
            myHttpResponse = myHttpRequest.GetResponse()
            If myHttpResponse.StatusCode <> HttpStatusCode.OK Then
                _ErrorMessage = "�t�@�C����������܂���"
                With _myLogRecord
                    .Target = sourceUrl & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = myHttpResponse.StatusCode.ToString()
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End If

            'http���X�|���X�̓ǂݍ���
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
                '�t�@�C���ւ̕ۑ�
                '//MOD 2005.07.22 ���s�j�ɉ� �t�@�C�����[�h�̕ύX START
                'myFileStream = New FileStream(destinationFile, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
                myFileStream = New FileStream(destinationFile, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)
                '//MOD 2005.07.22 ���s�j�ɉ� �t�@�C�����[�h�̕ύX END
                myBinaryWriter = New BinaryWriter(myFileStream)
                myBinaryWriter.Write(myBuff)
                myBinaryWriter.Close()
                myFileStream.Close()
            Catch e As Exception
                _ErrorMessage = "�t�@�C���R�s�[���̃G���[�ł�"
                With _myLogRecord
                    .Target = sourceUrl & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = e.Message
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End Try

            myHttpResponse.Close()

        Catch e As WebException
            '�^�C���A�E�g
            If e.Status = WebExceptionStatus.ConnectFailure Then
                _ErrorMessage = WebExceptionStatus.ConnectFailure.ToString()
            Else
                _ErrorMessage = "�^�C���A�E�g"
            End If
            With _myLogRecord
                .Target = sourceUrl & "��" & _ErrorMessage
                .Result = "NG"
                .Remark = e.Message & "(status=" & e.Status.ToString() & ")"
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Catch e As Exception
            _ErrorMessage = e.Message
            With _myLogRecord
                .Target = sourceUrl & "��" & e.Message
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False

        End Try
        With _myLogRecord
            .Target = sourceUrl & "������Ƀ_�E�����[�h�ł��܂���"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True

    End Function

    '
    '   �t�@�C�����R�s�[���܂�
    '
    Private Function GetFileDownload(ByVal sourceFile As String, ByVal destinationFile As String) As Boolean

        Try
            '�t�@�C�������݂��邩���`�F�b�N���āA�R�s�[
            If File.Exists(sourceFile) = False Then
                _ErrorMessage = "�t�@�C����������܂���"
                With _myLogRecord
                    .Target = sourceFile & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return False
            End If

            File.Copy(sourceFile, destinationFile, True)
        Catch e As Exception
            _ErrorMessage = "�t�@�C���R�s�[���̃G���[�ł�"
            With _myLogRecord
                .Target = sourceFile & "��" & destinationFile & "�ւ�" & _ErrorMessage
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
        With _myLogRecord
            .Target = sourceFile & "��" & destinationFile & "�֐���ɃR�s�[�ł��܂���"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        Return True
    End Function

    '
    '   �t�@�C�������擾���܂�
    '
    Protected Friend Function GetFileName(ByVal fullPathFileName As String) As String
        If CheckPathIsUrl(fullPathFileName) Then
            'URL�̃t�@�C������I��
            Dim FileIndex As Integer
            FileIndex = fullPathFileName.Replace("\", "/").LastIndexOf("/") + 1
            If FileIndex <= 0 Then
                _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
                With _myLogRecord
                    .Target = fullPathFileName & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If
            If FileIndex >= fullPathFileName.Length Then
                _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
                With _myLogRecord
                    .Target = fullPathFileName & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If

            Return fullPathFileName.Substring(FileIndex)
        Else
            '�t�@�C���p�X�ł̃t�@�C������I��
            Dim FileName As String
            FileName = Path.GetFileName(fullPathFileName)
            If FileName Is Nothing Then
                _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
                With _myLogRecord
                    .Target = fullPathFileName & "��" & _ErrorMessage
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                Return ""
            End If
            If FileName = String.Empty Then
                _ErrorMessage = "�t�@�C�������w�肳��Ă��܂���"
                With _myLogRecord
                    .Target = fullPathFileName & "��" & _ErrorMessage
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
    '   �f�B���N�g���[���`�F�b�N���A���݂��Ȃ���΍쐬����
    '
    Protected Friend Function CheckOrMakeDestinationDir(ByVal dirName As String) As Boolean

        If Not (Directory.Exists(dirName)) Then
            Try
                Dim myDirInfo As DirectoryInfo
                myDirInfo = Directory.CreateDirectory(dirName)
                With _myLogRecord
                    .Target = dirName & "���쐬���܂���"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            Catch e As Exception
                _ErrorMessage = "�f�B���N�g���쐬�G���["
                With _myLogRecord
                    .Target = dirName & "��" & _ErrorMessage
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
