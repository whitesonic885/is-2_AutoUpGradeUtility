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
    '�`�F�b�N���Ȃ��A�Z���u���̃p�^�[��
    Private _NoCheckAssemblies() As String = {"System.", "mscorlib"}
    Private Const constSystemName As String = "System"

#Region "���O�o��"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    Public Sub New()
        MyBase.New()

        _myLogRecord.Status = ""
        _myLogRecord.SubStatus = "�Q�ƃA�Z���u��"

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
    '   �A�Z���u�������[�h���āA�Q�ƃA�Z���u���z���Ԃ�
    '       �ŏ��̃����o�[�Ƀ��[�h�����A�Z���u������߂��܂�
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
                myMessages(1) = "�����[�h���܂�"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            If myName = String.Empty Then
                '�t�@�C�����Ƃ��ă��[�h
                myAssembly = [Assembly].LoadFrom(fileOrAssembly)
            Else
                '�A�Z���u���E�A�C�f���e�B�e�B�Ń��[�h
                myAssembly = [Assembly].Load(fileOrAssembly)
                '�o�[�W����������Ă��邩�ǂ���
                If myAssembly.FullName <> fileOrAssembly Then
                    '�_�E�����[�h����
                    download = True
                    Throw New Exception("�o�[�W�����Ⴂ")
                End If
            End If
        Catch e As Exception
            If download Then
                '�A�Z���u�������[�h�o���Ȃ��ꍇ�́A�A�Z���u�����R�s�[���ă��[�h���܂�
                Dim myFileCopy As FileCopy = New FileCopy()
                Dim myFullPathFileName As String

                With _myLogRecord
                    ReDim myMessages(3)
                    myMessages(0) = _SourcePath
                    myMessages(1) = "����"
                    myMessages(2) = myName
                    myMessages(3) = "���_�E�����[�h���܂�"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    .Remark = e.Message
                    ReDim myMessages(1)
                End With
                _myLog.WriteLog(_myLogRecord)
                'URL�`�����ǂ����ɂ��_�E�����[�h���ݒ�
                If myFileCopy.CheckPathIsUrl(_SourcePath) Then
                    myFullPathFileName = String.Concat(_SourcePath, myName)
                Else
                    myFullPathFileName = Path.Combine(_SourcePath, myName)
                End If
                myRet = myFileCopy.CopyFile(myFullPathFileName, _DestinationAppPath)
                If myRet = False Then
                    With _myLogRecord
                        myMessages(0) = myFullPathFileName
                        myMessages(1) = "�����[�h�ł��܂���ł���"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "NG"
                        .Remark = e.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    Return Nothing
                End If
                Try
                    '�t�@�C�����Ƃ��ă��[�h
                    myName = Path.Combine(_DestinationAppPath, myName)
                    With _myLogRecord
                        myMessages(0) = myName
                        myMessages(1) = "�����[�h���܂�"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    myAssembly = [Assembly].LoadFrom(myName)
                Catch ee As Exception
                    With _myLogRecord
                        myMessages(0) = myName
                        myMessages(1) = "�����[�h�ł��܂���ł���"
                        .Target = _myLog.AppendStrings(myMessages)
                        .Result = "NG"
                        .Remark = ee.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    Return Nothing
                End Try
            Else
                With _myLogRecord
                    '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� START
                    ReDim myMessages(1)
                    '// MOD 2016.09.16 Vivouac�j�e�r Visual Studio 2013�ϊ��ɔ����C�� END
                    myMessages(0) = fileOrAssembly
                    myMessages(1) = "�����[�h�ł��܂���ł���"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "NG"
                    .Remark = e.Message
                End With
                _myLog.WriteLog(_myLogRecord)
                Return Nothing
            End If
        End Try

        '�Q�ƃA�Z���u�������擾���܂�
        myAssemblies = myAssembly.GetReferencedAssemblies()
        myArray.Add(myAssembly.FullName())  '�A�Z���u������ǉ�

        For Each myAssemblyName In myAssemblies
            myArray.Add(myAssemblyName.FullName)
            With _myLogRecord
                ReDim myMessages(2)
                myMessages(0) = fileOrAssembly
                myMessages(1) = "�̎Q�ƃA�Z���u���́A"
                myMessages(2) = myAssemblyName.FullName
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = "List"
            End With
            _myLog.WriteLog(_myLogRecord)
        Next

        With _myLogRecord
            ReDim myMessages(4)
            myMessages(0) = "���[�h����"
            myMessages(1) = fileOrAssembly
            myMessages(2) = "�̎Q�ƃA�Z���u����"
            myMessages(3) = myAssemblies.Length.ToString()
            myMessages(4) = "���ł�"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myArray


    End Function

    '
    '   �A�Z���u���E�A�C�f���e�B�e�B����A�Z���u���E�t�����h�������擾
    '       �g���qDLL��t��
    Private Function GetAssemblyName(ByVal myAssemblyName As String, _
                                     ByVal addExpand As Boolean) As String
        Dim myPos As Integer
        Dim myName As String
        Dim myMessages() As String
        ReDim myMessages(1)

        '�A�C�f���e�B�e�B�̋�؂蕶���J���}�̈ʒu���擾
        myPos = myAssemblyName.IndexOf(",")
        If myPos <= 0 Then
            If addExpand Then
                With _myLogRecord
                    myMessages(0) = myAssemblyName
                    myMessages(1) = "�̓t�@�C�������w�肳��܂���"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            Return String.Empty
        End If
        '�A�Z���u�������擾
        myName = myAssemblyName.Substring(0, myPos)
        If addExpand Then
            myName = String.Concat(myName, constFileExt)
            With _myLogRecord
                myMessages(0) = myName
                myMessages(1) = "�A�Z���u�������I������܂���"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
        End If

        Return myName
    End Function

    '
    '   ���݂̃X���b�h�ɓǂݍ��܂�Ă���A�Z���u�����擾����
    '       �������J�n���鏉�߂ɁA�ǂݍ��ݍς݂̃A�Z���u�������X�g��
    '       ���[�h���邩�ǂ����̔��f�ɗ��p����
    Private Function GetAssemblies() As [Assembly]()
        Dim myAssemblies() As [Assembly]
        Try
            With _myLogRecord
                .Target = "�ǂݍ��ݍς݂̃A�Z���u�����擾���܂�"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            myAssemblies = Thread.GetDomain.GetAssemblies()
        Catch e As Exception
            With _myLogRecord
                .Target = "�ǂݍ��ݍς݂̃A�Z���u�����擾�ł��܂���ł���"
                .Result = "OK"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return Nothing
        End Try
        With _myLogRecord
            .Target = "�ǂݍ��ݍς݂̃A�Z���u�����擾���܂���"
            .Result = "OK"
            Dim myMessages() As String = {myAssemblies.Length.ToString(), "��"}
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myAssemblies
    End Function


    '
    '   �Q�ƃA�Z���u�����������A������΃_�E�����[�h�����݂�
    '       �O������̌ďo���|�C���g
    '   �ŏI��Protected Friend
    Protected Friend Function SearchAssemblies(ByVal fileName As String, _
                                               ByVal SourcePath As String, _
                                               ByVal DestinationAppPath As String) As Boolean
        Dim myRet As Boolean
        Dim myMBR As ReferenceAssembly
        Dim myAssemblies() As [Assembly]
        Dim myAssembly As [Assembly]
        Dim myArray As ArrayList
        Dim myFullPathFileName As String

        '�A�v���P�[�V�����E�h���C���̍쐬
        With _myLogRecord
            .Target = "�A�v���P�[�V�����E�h���C�����쐬���܂�"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)
        myRet = CreateDomain()
        If myRet = False Then Return False

        '�Q�ƃA�Z���u���`�F�b�N�p�̃C���X�^���X�̍쐬
        Try
            myMBR = _MyAppDomain.CreateInstanceAndUnwrap([Assembly].GetExecutingAssembly.FullName, _
                                                    "Microsoft.jp.DevMktg.TE.ReferenceAssembly")
            With _myLogRecord
                .Target = "�`�F�b�N�p�̃C���X�^���X���쐬���܂���"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
        Catch e As Exception
            With _myLogRecord
                .Target = "�`�F�b�N�p�̃C���X�^���X�̍쐬�Ɏ��s���܂���"
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

        '��ɂȂ郍�[�h���ꂽ�A�Z���u���̈ꗗ���擾���܂�
        myAssemblies = myMBR.GetAssemblies()
        If IsNothing(myAssemblies) Then
            AppDomain.Unload(_MyAppDomain)
            With _myLogRecord
                .Target = "�I���F��A�Z���u���̓Ǎ��ݎ��s"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If
        With _myLogRecord
            .Target = "��A�Z���u����Ǎ��݂܂���"
            .Result = "OK"
            .Remark = myAssemblies.Length.ToString() & "��"
        End With
        _myLog.WriteLog(_myLogRecord)

        _AssembliesAfterLoad = New ArrayList()
        _NotFindAssemblies = New ArrayList()
        For Each myAssembly In myAssemblies
            _AssembliesAfterLoad.Add(myAssembly.FullName())
        Next
        _AssembliesAfterLoad.Sort()

        '�ŏ��̎Q�ƃA�Z���u�����擾���܂�
        myFullPathFileName = Path.Combine(_DestinationAppPath, fileName)
        With _myLogRecord
            .Target = "�ŏ��̃A�Z���u�������[�h���܂�"
            .Result = "OK"
            .Remark = myFullPathFileName
        End With
        _myLog.WriteLog(_myLogRecord)
        myArray = New ArrayList()
        myArray = myMBR.GetReference(myFullPathFileName, False)
        If IsNothing(myArray) Then
            _NotFindAssemblies.Add(myFullPathFileName)
            With _myLogRecord
                .Target = "�I���F�Q�ƃA�Z���u���̓Ǎ��ݎ��s"
                .Result = "NG"
                .Remark = myFullPathFileName
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If
        With _myLogRecord
            .Target = "�Q�ƃA�Z���u���̓Ǎ��݂܂���"
            .Result = "OK"
            .Remark = myFullPathFileName
        End With
        _myLog.WriteLog(_myLogRecord)

        '�Q�ƃA�Z���u���̑��݃`�F�b�N���s���܂�
        myRet = CheckAssemblies(myArray, myMBR, 1)
        With _myLogRecord
            If myRet Then
                .Target = "�I���F�`�F�b�N������ɏI�����܂���"
                .Result = "OK"
            Else
                .Target = "�I���F�`�F�b�N���ُ�I�����܂���"
                .Result = "NG"
                .Remark = ""
            End If
            Dim myMessages() As String = {"�Ǎ��񂾃A�Z���u��=", _AssembliesAfterLoad.Count.ToString(), _
                                          "���A�Ǎ��߂Ȃ��A�Z���u��=", _NotFindAssemblies.Count.ToString(), "��"}
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        AppDomain.Unload(_MyAppDomain)

        Return myRet
    End Function

    '
    '   �`�F�b�N�p��AppDomain���쐬����
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
                myMessages(1) = "�A�v���P�[�V�����E�h���C���̍쐬�Ɏ��s���܂���"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
        With _myLogRecord
            myMessages(0) = myName
            myMessages(1) = "�A�v���P�[�V�����E�h���C���̍쐬�ɐ������܂���"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True
    End Function

    '
    '   �Q�ƃA�Z���u�����`�F�b�N���܂�
    '       ���̏����́A�ċA�R�[������܂�
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

        '���[�h�ς݂̃A�Z���u�����`�F�b�N�ς݂Ƃ���
        If _AssembliesAfterLoad.BinarySearch(myArray.Item(0)) < 0 Then
            _AssembliesAfterLoad.Add(myArray.Item(0))
            myArray.RemoveAt(0)
        End If
        _AssembliesAfterLoad.Sort()
        With _myLogRecord
            ReDim myMessages(1)
            myMessages(0) = TireLevel.ToString()
            myMessages(1) = ":�Ǎ��ݍς݂̃A�Z���u������"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = _AssembliesAfterLoad.Count.ToString() : myMessages(1) = "��"
            .Remark = _myLog.AppendStrings(myMessages)
            myMessages(0) = TireLevel.ToString()
        End With
        _myLog.WriteLog(_myLogRecord)

        '���[�h�ς݂̃A�Z���u�����Q�ƃA�Z���u�����X�g����폜����
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
            myMessages(1) = ":�Q�ƃA�Z���u���̃`�F�b�N�Ώی���"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = myArray.Count.ToString() : myMessages(1) = "��"
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        '�c��̃A�Z���u���Q�ƃ����o�[�����[�h���ă`�F�b�N����
        For Each myObject In myArray
            myName = CType(myObject, String)
            With _myLogRecord
                ReDim myMessages(1)
                myMessages(0) = TireLevel.ToString()
                myMessages(1) = ":�ʂ̃A�Z���u�����`�F�b�N���܂�"
                .Target = _myLog.AppendStrings(myMessages)
                .Result = "OK"
                .Remark = myName
            End With
            _myLog.WriteLog(_myLogRecord)
            myRefArray = myMBR.GetReference(myName, True)
            If myRefArray Is Nothing Then
                '���[�h�ł��Ȃ��A�Z���u�����X�g�ɒǉ�
                _NotFindAssemblies.Add(myName)
                With _myLogRecord
                    myMessages(1) = ":���[�h�o���܂���ł���"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "NG"
                    .Remark = myName
                End With
                _myLog.WriteLog(_myLogRecord)
            Else
                With _myLogRecord
                    myMessages(1) = ":�ċA�`�F�b�N���J�n���܂�"
                    .Target = _myLog.AppendStrings(myMessages)
                    .Result = "OK"
                    myMessages(0) = myRefArray.Count.ToString()
                    myMessages(1) = "��"
                    .Remark = _myLog.AppendStrings(myMessages)
                    myMessages(0) = TireLevel.ToString()
                End With
                _myLog.WriteLog(_myLogRecord)

                myRet = CheckAssemblies(myRefArray, myMBR, TireLevel + 1)
                With _myLogRecord
                    If myRet = False Then
                        _NotFindAssemblies.Add(myName)
                        myMessages(1) = ":�ċA�`�F�b�N�Ɏ��s���܂���"
                        .Result = "NG"
                    Else
                        myMessages(1) = ":�ċA�`�F�b�N�ɐ������܂���"
                        .Result = "OK"
                    End If
                    .Target = _myLog.AppendStrings(myMessages)
                    ReDim myMessages(3)
                    myMessages(0) = myName
                    myMessages(1) = "�̎Q�ƃA�Z���u�����A"
                    myMessages(2) = myRefArray.Count.ToString()
                    myMessages(3) = "���ł���"
                    .Remark = _myLog.AppendStrings(myMessages)
                End With
                _myLog.WriteLog(_myLogRecord)

                myRefArray.Clear()
            End If
        Next
        With _myLogRecord
            ReDim myMessages(1)
            myMessages(0) = TireLevel.ToString()
            myMessages(1) = ":�ʃA�Z���u���̃`�F�b�N���I�����܂���"
            .Target = _myLog.AppendStrings(myMessages)
            .Result = "OK"
            myMessages(0) = myArray.Count.ToString()
            myMessages(1) = "��"
            .Remark = _myLog.AppendStrings(myMessages)
        End With
        _myLog.WriteLog(_myLogRecord)

        Return True
    End Function
End Class

