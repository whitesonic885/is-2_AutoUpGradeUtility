Imports System.IO
Imports System.Reflection
Imports System.Net
Imports System.Threading

Imports System.Text
Imports System.Security.Cryptography

	'/// <summary>
	'/// �����X�V�A�v�����i[AutoUpGradeUtility.dll]
	'/// </summary>
	'//--------------------------------------------------------------------------
	'// �C������
	'//--------------------------------------------------------------------------
	'// ADD 2008.06.11 ���s�j���� [conime.exe]�̋N�� 
	'// ADD 2008.06.11 ���s�j���� ���O�C����ʂ���Ƀm�[�}���\���� 
	'// ADD 2008.06.11 ���s�j���� [CopyAutoUpGrade]���\���� 
	'// ADD 2008.06.11 ���s�j���� [EXPAND]���\���� 
	'//--------------------------------------------------------------------------
	'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� 
	'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� 
	'//                           �i���������[�X���ɂ̓R�����g�ɂ���j
	'// ADD 2009.07.30 ���s�j���� exe��dll���Ή� 
	'// ADD 2009.10.02 ���s�j���� �Q�d�N���̃`�F�b�N 
	'// ADD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ 
	'// ADD 2009.10.06 ���s�j���� �p�\�R���̓��t�ݒ�ȈՃ`�F�b�N 
	'//--------------------------------------------------------------------------
	'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� 
	'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� 
	'//                           �i�Í����𑽗p����Ɖ�ǂ���₷���Ȃ�ׁj
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

'// ADD 2005.05.31 ���s�j�ɉ� CopyAutoUpGrade�Ή� START
    Public _ClientMutex As String
'// ADD 2005.05.31 ���s�j�ɉ� CopyAutoUpGrade�Ή� END
'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� START
    Dim gsInitProxy As String = AppDomain.CurrentDomain.BaseDirectory _
             + "\proxy.ini"
    Dim gbInitProxyExists As Boolean = False
    Dim gs�A�v���t�H���_ As String = ""

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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
	Dim gbProxyOnUserSet   As Boolean = False
	Dim gbProxyIdOnUserSet As Boolean = False
    Dim gsProxyIdUserSet   As String  = ""
    Dim gsProxyPaUserSet   As String  = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END

    Dim sv_init As is2init.Service1 = Nothing
'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� END

#Region "���O�o��"
    Private _myLog As LogWriter = LogWriter.CreateInstance()
    Private _myLogRecord As LogWriter.LogRecord
#End Region

    '
    '   �R���X�g���N�^
    '
    Public Sub New()
        '������
        _DestinationPath = AppDomain.CurrentDomain.BaseDirectory
        _AppDir = Path.Combine(_DestinationPath, constAppDir)
        _ConfigDir = Path.Combine(_DestinationPath, constConfigDir)
        _DestinationAppPath = Path.Combine(_DestinationPath, _AppDir)
        _DestinationConfigPath = Path.Combine(_DestinationPath, _ConfigDir)

        With _myLogRecord
            .Status = "�J�n"
            .SubStatus = "AutoUpGrade"
        End With

    End Sub

'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
    '
    '   �t�@�C���X�^���v�̕ύX
    '
    Public Function ChangeFileStamp(ByVal sFileName As String, _
                                 ByVal sTimeStamp As String) As Boolean

        Try
            Dim sFullFileName As String = Path.Combine(_DestinationAppPath, sFileName)
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ START
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
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ END

            Dim dtFileStamp As Date = Date.Parse(sFileStamp)
            'With _myLogRecord
            '    .Target = "�t�@�C���X�^���v�̕ύX�F" + dtFileStamp.ToString()
            '    .Result = "OK"
            '    .Remark = ""
            'End With
            '_myLog.WriteLog(_myLogRecord)

            '�t�@�C���X�^���v�̕ύX
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
    '   �b�`�a�t�@�C���̃_�E�����[�h
    '
    Public Function DownloadCabFile(ByVal srcFullName As String, _
                                        ByVal desFullName As String, _
                                        ByVal sourceDirName As String, _
                                        ByVal sFileName As String) As Boolean

        Dim myRet As Boolean
        Try
            Dim myFileCopy As FileCopy = New FileCopy

            '�b�`�a�t�@�C���̃_�E�����[�h
            myRet = myFileCopy.CopyFile(sourceDirName & sFileName, _DestinationPath)
            If myRet Then
                With _myLogRecord
                    .Target = sourceDirName & sFileName & "��" & _DestinationPath & "�փ_�E�����[�h�ɐ������܂���"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)

                '�b�`�a�o�[�W�����t�@�C������ʊK�w�ɃR�s�[����
                File.Copy(srcFullName, desFullName, True)
                '�b�`�a�o�[�W�����t�@�C���̃t�@�C���X�^���v�̕ύX
                File.SetLastWriteTime(desFullName, File.GetLastWriteTime(srcFullName))

                Dim flgFullName As String

                sFileName = sFileName.Substring(0, sFileName.Length - 3) + "cab"
                flgFullName = Path.Combine(_DestinationPath, sFileName) + "_flg"

                '��ʊK�w�̃t���O�t�@�C���̍폜
                If System.IO.File.Exists(flgFullName) Then
                    System.IO.File.Delete(flgFullName)
                End If
            Else
                With _myLogRecord
                    .Target = sourceDirName & sFileName & "�̃_�E�����[�h�Ɏ��s���܂���"
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
    '   �b�`�a�t���O�t�@�C���̍폜
    '
    Public Function DeleteCabFlagFile(ByVal sFileName As String) As Boolean

        Try
            '�b�`�a�o�[�W�����t�@�C�����_�E�����[�h���ꂽ��
            If sFileName.Substring(sFileName.LastIndexOf(".") + 1).Equals("txt") Then
                Dim flgFullName As String

                sFileName = sFileName.Substring(0, sFileName.Length - 3) + "cab"
                flgFullName = Path.Combine(_DestinationPath, sFileName) + "_flg"

                '��ʊK�w�̃t���O�t�@�C���̍폜
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
    '   �b�`�a�o�[�W�����t�@�C���̃`�F�b�N
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

            '�b�`�a�o�[�W�����t�@�C�������݂��Ȃ����A�b�`�a�t�@�C�����_�E�����[�h
            If File.Exists(desFullName) = False Then
                '�t�@�C�����_�E�����[�h����
                DownloadCabFile(srcFullName, desFullName, _
                                sourceDirName, sFileName)

                '�b�`�a�o�[�W�����t�@�C���̍X�V�������V�������A�b�`�a�t�@�C�����_�E�����[�h
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ START
'//         ElseIf File.GetLastWriteTime(srcFullName) > File.GetLastWriteTime(desFullName) Then
                '�^�C���X�^���v�́A�N���������܂Ŕ�r����i�b�͔�r���Ȃ��j
            ElseIf File.GetLastWriteTime(srcFullName).ToString("yyyyMMddHHmm") > File.GetLastWriteTime(desFullName).ToString("yyyyMMddHHmm") Then
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ END
                '�t�@�C�����_�E�����[�h����
                DownloadCabFile(srcFullName, desFullName, _
                                sourceDirName, sFileName)

                '�b�`�a�t�@�C�������݂��Ȃ����A�b�`�a�t�@�C�����_�E�����[�h
            ElseIf File.Exists(cabFullName) = False Then
                '�t�@�C�����_�E�����[�h����
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
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END

    '
    '   �o�[�W�������̎擾
    '
    Public Function GetVersion(ByVal sourceDirName As String, _
                                 ByVal sourceFileName As String, _
                                 ByVal download As Boolean, _
                                 ByVal downloadOnly As Boolean, _
                                 ByVal localExecute As Boolean, _
                                 ByVal logClear As Boolean) As Boolean

'// ADD 2008.06.11 ���s�j���� [conime.exe]�̋N�� START
        Dim pConIme As New System.Diagnostics.Process
        Try
            With pConIme.StartInfo
                .FileName = "cmd.exe"
                .Arguments = "/C conime.exe"
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            End With
            pConIme.Start()
        Catch ex As System.Exception
            '���̑�
        Finally
            If Not (pConIme Is Nothing) Then pConIme.Close()
        End Try
'// ADD 2008.06.11 ���s�j���� [conime.exe]�̋N�� END

        If logClear Then _myLog.Clear()

'// ADD 2009.10.02 ���s�j���� �Q�d�N���̃`�F�b�N START
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
                        .Target = "2�b�҂��܂������A�I�����܂���ł���" _
                          + "[" + sCheckExeName + "]"
                        .Result = "NG"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    Windows.Forms.MessageBox.Show( _
                     "�A�v���P�[�V�����͊��ɋN������Ă��܂��@�@�@" & vbCrLf _
                     & "�^�X�N�o�[��^�X�N�}�l�[�W���������m�F���������@�@�@", _
                     sCheckExeName, _
                     Windows.Forms.MessageBoxButtons.OK, _
                     Windows.Forms.MessageBoxIcon.Error)

                    Return True
                End If
                System.Threading.Thread.Sleep(1000)
            Loop

            ' Mutex ���J������
            hMutex.ReleaseMutex()
            hMutex.Close()
        Catch ex As ApplicationException
            With _myLogRecord
                .Target = "Mutex���J���ł��܂���ł���"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End Try
'// ADD 2009.10.02 ���s�j���� �Q�d�N���̃`�F�b�N END
'// ADD 2009.10.06 ���s�j���� �p�\�R���̓��t�ݒ�ȈՃ`�F�b�N START
        Dim dtNow As New Date
        Dim sNow = dtNow.Now.ToString("yyyy/MM/dd")
        If dtNow.Now.Year < 2009 Then
            With _myLogRecord
                .Target = "�p�\�R���̓��t�ݒ�Ɍ�肪����܂�" _
                  & "[" & sNow & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Windows.Forms.MessageBox.Show( _
             "�p�\�R���̓��t�ݒ�Ɍ�肪����܂��@�@�@" & vbCrLf _
             & "�i" & sNow & "�j" & "�@�@�@", _
             sCheckExeName, _
             Windows.Forms.MessageBoxButtons.OK, _
             Windows.Forms.MessageBoxIcon.Error)

            Return False
        End If
'// ADD 2009.10.06 ���s�j���� �p�\�R���̓��t�ݒ�ȈՃ`�F�b�N END

        Dim myRet As Boolean
        waitDlg = New WaitDialog

        Try
            Dim myFileCopy As FileCopy = New FileCopy
'//            Dim myCheckFileName As String = Path.Combine(_DestinationConfigPath, myFileCopy.GetFileName(sourceFileName))
'//            '�p�X�`������������
'//            If Not (localExecute = True And File.Exists(myCheckFileName)) Then
'//                '�R���t�B�O�̃_�E�����[�h�ADL�o���Ȃ��Ă����s����
'//                myRet = myFileCopy.CopyFile(sourceDirName & sourceFileName & constConfigExt, _DestinationConfigPath)
'//                With _myLogRecord
'//                    If myRet Then
'//                        .Target = sourceDirName & sourceFileName & constConfigExt & "��" & _DestinationConfigPath & "�փ_�E�����[�h�ł��܂���"
'//                        .Result = "OK"
'//                        .Remark = ""
'//                    Else
'//����A�l�b�g���[�N�ƂȂ����Ă��Ȃ��ƋN���ł��Ȃ�����
'//���[�J���N���͍s��Ȃ�
'//                        If myFileCopy.ErrorMesage = WebExceptionStatus.ConnectFailure.ToString() Then
'//�l�b�g���[�N�G���[�̓��[�J���N���ɂ���
'//                            localExecute = True
'//                            download = False
'//                            .Target = sourceDirName & sourceFileName & "���l�b�g���[�N�G���[�ɂ�胍�[�J���N���ɐ؂�ւ��܂�"
'//                            .Result = "OK"
'//                            .Remark = myFileCopy.ErrorMesage
'//                        Else
'//                            .Target = sourceDirName & sourceFileName & constConfigExt & "��" & _DestinationConfigPath & "�փ_�E�����[�h�ł��܂���"
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

'// ADD 2005.07.22 ���s�j�����J  START
            Dim appFile As String
            Dim appSize As Integer = 0
'// ADD 2005.07.22 ���s�j�����J  END

            ' �i�s�󋵃_�C�A���O��\������
            waitDlg.Show()
            ' �i�s�󋵃_�C�A���O�̏���������
            'waitDlg.Owner = Me ' �_�C�A���O�̃I�[�i�[��ݒ�
'// MOD 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX START
            '            waitDlg.MainMsg = "�o�[�W�����t�@�C�����擾���Ă��܂��c�c" ' �����̊T�v
            '            waitDlg.SubMsg = ""
            waitDlg.MainMsg = "�N���������ł��D�D�D"
'// MOD 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX END
            waitDlg.ProgressMsg = ""
            waitDlg.ProgressMin = 0 ' ���������̍ŏ��l�i0������J�n�j
            waitDlg.ProgressStep = 1 ' �������ƂɃ��[�^�[��i�߂邩
            waitDlg.ProgressValue = 0 ' �ŏ��̌���
            waitDlg.Update()
            Dim iCount As Integer = 0   ' �����ڂ̏������������J�E���^

'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� START
            '// �J�����g�f�B���N�g���̎擾
            gs�A�v���t�H���_ = AppDomain.CurrentDomain.BaseDirectory
            If Not (Directory.Exists(gs�A�v���t�H���_)) Then
                With _myLogRecord
                    .Target = "�J�����g�f�B���N�g�����݂���܂���ł���\n" _
                      + "[" + gs�A�v���t�H���_ + "]"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If

            gbInitProxyExists = False
            gsProxyAdrUserSet = ""
            giProxyNoUserSet = 0
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//         giConnectTimeOut = 100  '// �����l�F�P�O�O�b
            giConnectTimeOut   = 50 '// �����l�F�T�O�b
	        gbProxyOnUserSet   = False
	        gbProxyIdOnUserSet = False
            gsProxyIdUserSet   = ""
            gsProxyPaUserSet   = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
            Dim sConnectTimeOut As String = ""
            Dim sProxyAdrUserSet As String = ""
            Dim sProxyNoUserSet As String = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
			Dim sProxyOnUserSet   As String = ""
			Dim sProxyIdOnUserSet As String = ""
			Dim sProxyIdUserSet   As String = ""
			Dim sProxyPaUserSet   As String = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
            Dim sr As StreamReader = Nothing
            Try
                sr = File.OpenText(gsInitProxy)
                gbInitProxyExists = True
                sConnectTimeOut = sr.ReadLine()
                sProxyAdrUserSet = sr.ReadLine()
                sProxyNoUserSet = sr.ReadLine()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
                    sProxyIdUserSet  = �������Q(sProxyIdUserSet)
                    sProxyPaUserSet  = �������Q(sProxyPaUserSet)
                End If
'//				With _myLogRecord
'//					.Target = "" _
'//						& "[" & sProxyIdUserSet & "]" _
'//						& "[" & sProxyPaUserSet & "]"
'//					.Result = ""
'//					.Remark = ""
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
            Catch ex As System.IO.FileNotFoundException
                With _myLogRecord
                    .Target = "�v���L�V�ݒ�t�@�C�����݂���܂���ł���"
                    .Result = "NG"
                    .Remark = ex.Message
                End With
                _myLog.WriteLog(_myLogRecord)
            Catch ex As Exception
                With _myLogRecord
                    .Target = "�v���L�V�ݒ�t�@�C�� Exception:"
                    .Result = "NG"
                    .Remark = ex.Message
                End With
                _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
            Finally
                If Not (sr Is Nothing) Then sr.Close()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
            End Try

'// ADD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ START
            If gbInitProxyExists Then
'// ADD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ END
                Try
                    If sConnectTimeOut.Length > 0 Then
                        If ���p�`�F�b�N(sConnectTimeOut, "�^�C���A�E�g") Then
                            If ���l�`�F�b�N(sConnectTimeOut, "�^�C���A�E�g") Then
                                giConnectTimeOut = Integer.Parse(sConnectTimeOut)
                            End If
                        End If
                    End If
                    If Not (sProxyAdrUserSet Is Nothing) Then
                        If ���p�`�F�b�N(sProxyAdrUserSet, "�v���L�V�A�h���X") Then
                            gsProxyAdrUserSet = sProxyAdrUserSet
                        End If
                    End If
                    If sProxyNoUserSet.Length > 0 Then
                        If ���p�`�F�b�N(sProxyNoUserSet, "�v���L�V�|�[�g�ԍ�") Then
                            If ���l�`�F�b�N(sProxyNoUserSet, "�v���L�V�|�[�g�ԍ�") Then
                                giProxyNoUserSet = Integer.Parse(sProxyNoUserSet)
                            End If
                        End If
                    End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
                    If ���p�`�F�b�N(sProxyOnUserSet, "�v���L�V�ݒ�") Then
	                    If sProxyOnUserSet.Equals("1") Then
	                        gbProxyOnUserSet = True
	                    End If
                    End If
                    If sProxyOnUserSet.Length = 0 Then
                        If gsProxyAdrUserSet.Length > 0 Then gbProxyOnUserSet = True
                    End If
                    If ���p�`�F�b�N(sProxyIdOnUserSet, "�v���L�V�h�c�ݒ�") Then
	                    If sProxyIdOnUserSet.Equals("1") Then
	                        gbProxyIdOnUserSet = True
	                    End If
                    End If
                    If ���p�`�F�b�N(sProxyIdUserSet, "���[�U�h�c") Then
                        gsProxyIdUserSet = sProxyIdUserSet
                    End If
                    If ���p�`�F�b�N(sProxyPaUserSet, "�p�X���[�h") Then
                        gsProxyPaUserSet = sProxyPaUserSet
                    End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                Catch ex As Exception
'//�ۗ��@�G���[����
                    With _myLogRecord
                        .Target = "�v���L�V�ݒ�l���[�j���O Exception:"
                        .Result = "NG"
                        .Remark = ex.Message
                    End With
                    _myLog.WriteLog(_myLogRecord)
                End Try

                '// �v���L�V�̐ݒ���擾
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             Dim iRet As Integer = -1
                Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                �h�d�v���L�V�ݒ���擾()

                sv_init = New is2init.Service1
                If giConnectTimeOut > 0 Then
                    sv_init.Timeout = giConnectTimeOut * 1000
                End If
                sv_init.Url = sv_init.Url.Replace("http://", "https://")

                '//--------------------------------------------------------
                '// �v���L�V�ݒ�(���[�U�ݒ�)
                '// �i[proxy.ini]�ݒ�l�j
                '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             If iRet <> 1 And gsProxyAdrUserSet.Length > 0 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                     .Target = "P�ݒ�(U)[" & �Í����Q(gsProxyAdrUserSet) _
'//                         & "][" & �Í����Q(giProxyNoUserSet.ToString("0000")) & "]"
                        .Target = "P�ݒ�(U)[" & gsProxyAdrUserSet _
                            & "][" & giProxyNoUserSet.ToString("0000") & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 iRet = �v���L�V�ݒ�(gsProxyAdrUserSet, giProxyNoUserSet)
'//                 If iRet = 1 Then myRet = True
                    With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                     .Target = "P_1[" & �Í����Q(gbProxyOnUserSet.ToString()) _
'//                         & "] P_2[" & �Í����Q(gbProxyIdOnUserSet.ToString()) & "]"
                        .Target = "P_1[" & gbProxyOnUserSet.ToString() _
                            & "] P_2[" & gbProxyIdOnUserSet.ToString() & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    If gbProxyOnUserSet Then
                        If gbProxyIdOnUserSet Then
                            wRet = �v���L�V�ݒ�Q(gsProxyAdrUserSet, giProxyNoUserSet _
                                , gsProxyIdUserSet, gsProxyPaUserSet)
                        Else
                            wRet = �v���L�V�ݒ�(gsProxyAdrUserSet, giProxyNoUserSet)
                        End If
                    Else
                        wRet = �v���L�V�ݒ�("", 0) '// �v���L�V�Ȃ��ڑ�
                    End If
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                End If

                '//--------------------------------------------------------
                '// �f�t�H���g�v���L�V�ݒ�i�ȑO�Ɠ����ɂȂ�͂��j
                '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    With _myLogRecord
                        .Target = "P�ݒ�(D)"
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 iRet = �f�t�H���g�v���L�V�ݒ�()
'//                 If iRet = 1 Then myRet = True
                    wRet = �f�t�H���g�v���L�V�ݒ�()
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                End If

                '//--------------------------------------------------------
                '// �v���L�V�ݒ�(�r�r�k�A�r�����X�g�L��)���x�X�g���Ǝv����
                '// �i�h�d�̃��W�X�g���ݒ�l�j
                '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             If iRet <> 1 And gsProxyAdr.Length > 0 Then
                If wRet <> WebExceptionStatus.Success  And gsProxyAdr.Length > 0Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                     .Target = "P�ݒ�(S)[" & �Í����Q(gsProxyAdr) _
'//                         & "][" & �Í����Q(giProxyNo.ToString("0000")) & "]"
                        .Target = "P�ݒ�(S)[" & gsProxyAdr _
                            & "][" & giProxyNo.ToString("0000") & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 iRet = �v���L�V�ݒ�(gsProxyAdr, giProxyNo)
'//                 If iRet = 1 Then myRet = True
                    wRet = �v���L�V�ݒ�(gsProxyAdr, giProxyNo)
                    If wRet = WebExceptionStatus.Success Then myRet = True
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                End If

                '//--------------------------------------------------------
                '// �v���L�V���ݒ莞
                '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 If gsProxyAdrUserSet.Length = 0 And gsProxyAdr.Length = 0 Then
                    If gbProxyOnUserSet = False And gsProxyAdrUserSet.Length = 0 And gsProxyAdr.Length = 0 Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                        Try
                            Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                         iRet = 1
                            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                            myRet = True
                        Catch ex As System.Net.WebException
                            With _myLogRecord
                                .Target = "P�ݒ�(N) s WebException:" & ex.Status.ToString()
                                .Result = "NG"
                                .Remark = ex.Message
                            End With
                            _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                        Catch ex As Exception
                            With _myLogRecord
                                .Target = "P�ݒ�(N) s Exception:"
                                .Result = "NG"
                                .Remark = ex.Message
                            End With
                            _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                         iRet = -2
                            wRet = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                        End Try
                        '//--------------------------------------------------------
                        '// �J���@�p
                        '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                     If iRet <> 1 Then
                        If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                            Try
                                sv_init = New is2init.Service1
                                sv_init.Timeout = 5000 '// 5�b
                                sv_init.Url = sv_init.Url.Replace("https://", "http://")
                                Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                             iRet = 1
                                wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                                myRet = True
                            Catch ex As System.Net.WebException
                                With _myLogRecord
                                    .Target = "P�ݒ�(N) _ WebException:" & ex.Status.ToString()
                                    .Result = "NG"
                                    .Remark = ex.Message
                                End With
                                _myLog.WriteLog(_myLogRecord)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                            Catch ex As Exception
                                With _myLogRecord
                                    .Target = "P�ݒ�(N) _ Exception:"
                                    .Result = "NG"
                                    .Remark = ex.Message
                                End With
                                _myLog.WriteLog(_myLogRecord)
                                myRet = False
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                             iRet = -2
                                wRet = WebExceptionStatus.UnknownError 
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                            Finally
                                If giConnectTimeOut > 0 Then
                                    sv_init.Timeout = giConnectTimeOut * 1000
                                End If
                            End Try
                        End If
                    End If
                End If

                sv_init.Url = sv_init.Url.Replace("http://", "https://")

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
                Dim iDlgCnt = 0
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             Do While iRet <> 1
                Do While wRet <> WebExceptionStatus.Success
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
                    '//--------------------------------------------------------
                    '// �ڑ��G���[���b�Z�[�W�̐ݒ�
                    '//--------------------------------------------------------
                    Dim sErrMsg As String = �ڑ��G���[���b�Z�[�W�ҏW(wRet)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    '//--------------------------------------------------------
                    '// WebExceptionStatus.TrustFailure �r�r�k�ʐM�̎��s
                    '// �i���v���L�V�̐ݒ�_�C�A���O��\�����Ȃ��j
                    '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 If iRet = -13 Then
'//                     Windows.Forms.MessageBox.Show( _
'//                      "is2�T�[�o�Ƃ̒ʐM�Ɏ��s���܂����@�@�@" & vbCrLf _
'//                      & "�p�\�R���̓��t�ݒ��r�r�k�ʐM�̐ݒ�Ȃǂ��m�F���Ă��������@�@�@", _
'//                      "is2", _
'//                      Windows.Forms.MessageBoxButtons.OK, _
'//                      Windows.Forms.MessageBoxIcon.Error)
'//                     Exit Do
'//                 End If
                    If wRet = WebExceptionStatus.TrustFailure Then
                        Windows.Forms.MessageBox.Show( _
                        "�T�[�o�[�ڑ��G���[�i"& sErrMsg &"�j�@�@�@�@�@" & vbCrLf _
                         & "�p�\�R���̓��t�ݒ��r�r�k�ʐM�̐ݒ�Ȃǂ��m�F���Ă��������@�@�@�@�@", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                        Exit Do
                    End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 '//--------------------------------------------------------
'//                 '// WebExceptionStatus.ConnectFailure �ڑ��̎��s
'//                 '// �i���v���L�V�̐ݒ�_�C�A���O��\�����Ȃ��j
'//                 '//--------------------------------------------------------
'//                 If iRet = -14 Then
'//                     Exit Do
'//                 End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    '//--------------------------------------------------------
                    '// �v���L�V�ݒ肪�R��ȏ㎸�s������I��
                    '//--------------------------------------------------------
                    iDlgCnt += 1
                    If iDlgCnt > 3 Then
                        Windows.Forms.MessageBox.Show( _
                        "�T�[�o�[�ڑ��G���[�i"& sErrMsg &"�j�@�@�@�@�@" & vbCrLf _
                         & "�v���L�V�ݒ�ɂR�񎸔s���܂����̂ŏI�����܂��@�@�@�@�@", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                        Exit Do
                    End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    '//--------------------------------------------------------
                    '// �v���L�V�ݒ�m�F���b�Z�[�W�̕\��
                    '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 Dim dRet As Windows.Forms.DialogResult = Windows.Forms.MessageBox.Show( _
'//                     "is2�T�[�o�Ƃ̒ʐM�Ɏ��s���܂����@�@�@" & vbCrLf _
'//                     & "�v���L�V�̐ݒ���s���܂����H�@�@�@", _
'//                     "is2", _
'//                     Windows.Forms.MessageBoxButtons.YesNo, _
'//                     Windows.Forms.MessageBoxIcon.Question)
                    Dim dRet As Windows.Forms.DialogResult = Windows.Forms.MessageBox.Show( _
                        "�T�[�o�[�ڑ��G���[�i"& sErrMsg &"�j�@�@�@�@�@" & vbCrLf _
                        & "�v���L�V�̐ݒ���s���܂����H�@�@�@�@�@", _
                        "is2", _
                        Windows.Forms.MessageBoxButtons.YesNo, _
                        Windows.Forms.MessageBoxIcon.Error)
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    '[������]�Ȃ炻�̂܂܏I��
                    If dRet = Windows.Forms.DialogResult.No Then Exit Do
                    '//--------------------------------------------------------
                    '// �v���L�V�ݒ�_�C�A���O��\��
                    '//--------------------------------------------------------
                    Dim dProxy As New SetProxy
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    dRet = dProxy.ShowDialog()
                    '[������]�Ȃ炻�̂܂܏I��
                    If dRet = Windows.Forms.DialogResult.No Then Exit Do
                    '�L�����Z���Ȃ炻�̂܂܏I��
                    If dRet = Windows.Forms.DialogResult.Cancel Then Exit Do

                    gsProxyAdrUserSet = dProxy.sProxyAdrUserSet
                    giProxyNoUserSet = dProxy.iProxyNoUserSet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
                    gbProxyOnUserSet   = dProxy.bProxyOnUserSet
                    gbProxyIdOnUserSet = dProxy.bProxyIdOnUserSet
                    gsProxyIdUserSet   = dProxy.sProxyIdUserSet
                    gsProxyPaUserSet   = dProxy.sProxyPaUserSet
                    giConnectTimeOut   = dProxy.iConnectTimeOut
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                     .Target = "P�ݒ�(U)[" & �Í����Q(gsProxyAdrUserSet) _
'//                         & "][" & �Í����Q(giProxyNoUserSet.ToString("0000")) & "]"
                        .Target = "P�ݒ�(U)[" & gsProxyAdrUserSet _
                            & "][" & giProxyNoUserSet.ToString("0000") & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 '//--------------------------------------------------------
'//                 '// �v���L�V�ݒ�t�@�C��[proxy.ini]�̏�����
'//                 '//--------------------------------------------------------
'//                 If gsProxyAdrUserSet = "" And giProxyNoUserSet = 0 Then
'//                     '�t�@�C���̏�������
'//                     Dim sw As StreamWriter = Nothing
'//                     Try
'//                         '�t�@�C���̏�������
'//                         sw = File.CreateText(gsInitProxy)
'//                         sw.WriteLine("")
'//                         sw.WriteLine("")
'//                         sw.WriteLine("")
'//                         sw.Close()
'//                     Catch ex As Exception
'//                         With _myLogRecord
'//                             .Target = "�t�@�C���̏������� Exception:"
'//                             .Result = "NG"
'//                             .Remark = ex.Message
'//                         End With
'//                         _myLog.WriteLog(_myLogRecord)
'//                         If Not (sw Is Nothing) Then sw.Close()
'//                         Windows.Forms.MessageBox.Show( _
'//                          "�t�@�C��[" & gsInitProxy & "]�̏������݂Ɏ��s���܂����@�@�@" & vbCrLf _
'//                          & "�t�H���_�ɏ������݌�����ǉ����Ă��������@�@�@", _
'//                          "is2", _
'//                          Windows.Forms.MessageBoxButtons.OK, _
'//                          Windows.Forms.MessageBoxIcon.Error)
'//                     End Try
'//                 End If
                    '//--------------------------------------------------------
                    '// �v���L�V�ݒ�t�@�C��[proxy.ini]�̏�������
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
                        sw.WriteLine(�Í����Q(gsProxyIdUserSet))
                        sw.WriteLine(�Í����Q(gsProxyPaUserSet))
                        sw.Close()
                    Catch ex As Exception
                        With _myLogRecord
                            .Target = "�t�@�C���̏������� Exception:"
                            .Result = "NG"
                            .Remark = ex.Message
                        End With
                        _myLog.WriteLog(_myLogRecord)
                        If Not (sw Is Nothing) Then sw.Close()
                        Windows.Forms.MessageBox.Show( _
                         "�t�@�C��[" & gsInitProxy & "]�̏������݂Ɏ��s���܂����@�@�@" & vbCrLf _
                         & "�t�H���_�ɏ������݌�����ǉ����Ă��������@�@�@", _
                         "is2", _
                         Windows.Forms.MessageBoxButtons.OK, _
                         Windows.Forms.MessageBoxIcon.Error)
                    End Try
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    '//--------------------------------------------------------
                    '// �v���L�V�ݒ�(���[�U�ݒ�)
                    '//--------------------------------------------------------
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 iRet = �v���L�V�ݒ�(gsProxyAdrUserSet, giProxyNoUserSet)
                    With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                     .Target = "P_1[" & �Í����Q(gbProxyOnUserSet.ToString()) _
'//                         & "] P_2[" & �Í����Q(gbProxyIdOnUserSet.ToString()) & "]"
                        .Target = "P_1[" & gbProxyOnUserSet.ToString() _
                            & "] P_2[" & gbProxyIdOnUserSet.ToString() & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                        .Result = ""
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)

                    If gbProxyOnUserSet Then
                        If gbProxyIdOnUserSet Then
                            wRet = �v���L�V�ݒ�Q(gsProxyAdrUserSet, giProxyNoUserSet _
                                , gsProxyIdUserSet, gsProxyPaUserSet)
                        Else
                            wRet = �v���L�V�ݒ�(gsProxyAdrUserSet, giProxyNoUserSet)
                        End If
                    Else
                        wRet = �v���L�V�ݒ�("", 0)
                    End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                 If iRet = 1 Then myRet = True
'//                 If iRet = 1 Then
                    If wRet = WebExceptionStatus.Success Then
                        myRet = True
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                        gsProxyAdr = gsProxyAdrUserSet
                        giProxyNo = giProxyNoUserSet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                     '�t�@�C���̏�������
'//                     Dim sw As StreamWriter = Nothing
'//                     Try
'//                         '�t�@�C���̏�������
'//                         sw = File.CreateText(gsInitProxy)
'//�^�C���A�E�g
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
'//                             .Target = "�t�@�C���̏������� Exception:"
'//                             .Result = "NG"
'//                             .Remark = ex.Message
'//                         End With
'//                         _myLog.WriteLog(_myLogRecord)
'//                         If Not (sw Is Nothing) Then sw.Close()
'//                         Windows.Forms.MessageBox.Show( _
'//                          "�t�@�C��[" & gsInitProxy & "]�̏������݂Ɏ��s���܂����@�@�@" & vbCrLf _
'//                          & "�t�H���_�ɏ������݌�����ǉ����Ă��������@�@�@", _
'//                          "is2", _
'//                          Windows.Forms.MessageBoxButtons.OK, _
'//                          Windows.Forms.MessageBoxIcon.Error)
'//                     End Try
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                        Exit Do
                    End If
                Loop
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//             If iRet <> 1 Then
                If wRet <> WebExceptionStatus.Success Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
                    Return False
                End If
                sv_init.Timeout = 100 * 1000 '// �f�t�H���g�l�̂P�O�O�b
'// ADD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ START
            End If
'// ADD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ END
'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� END

            '�T�[�o�[�̃o�[�W�������̎擾
            myRet = verCheck.XmlReadServer(sourceDirName, _DestinationConfigPath, ServerSerializer, ServerVersion)
            If Not myRet Then Return False

            '�N���C�A���g�̃o�[�W�������̎擾
            Dim cntRet As Boolean
            cntRet = verCheck.XmlReadClient(Path.Combine(_DestinationPath, "VersionFile.xml"), ClientSerializer, ClientVersion)

'// MOD 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX START
'//         waitDlg.MainMsg = "�t�@�C����]�����Ă��܂��c�c" ' �����̊T�v
            waitDlg.MainMsg = "���΂炭���҂����������B�D�D�D"
'// MOD 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX END
            waitDlg.ProgressMax = ServerVersion.Length ' �S�̂̏�������
            Dim verRet As Boolean
            For Each ServerFile In ServerVersion
                ' �������~���ǂ������`�F�b�N
                If waitDlg.IsAborting = True Then
                    Return False
                End If

'// DEL 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX START
'//              waitDlg.SubMsg = ServerFile.FileName
'// DEL 2005.05.10 ���s�j���� ���b�Z�[�W�̕ύX END
                ' �i�s�󋵃_�C�A���O�̃��[�^�[��ݒ�
                waitDlg.ProgressMsg = _
                (CType((iCount * 100 / ServerVersion.Length), Integer)).ToString() & "%�@" _
                & "�i" + iCount.ToString() + " / " & ServerVersion.Length & "�j"
                With _myLogRecord
                    .Status = "�J�n"
                    .Target = ServerFile.FileName & "�̃o�[�W�����`�F�b�N"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                With _myLogRecord
                    .Status = ""
                    .SubStatus = "�t�@�C���`�F�b�N"
                End With

                '�_�E�����[�h�f�B���N�g���̐ݒ�
                Dim FileDirLength = ServerFile.FileName.Replace("\", "/").LastIndexOf("/")
                Dim FileDir As String = ""
                If FileDirLength <> -1 Then
                    FileDir = "/" & ServerFile.FileName.Substring(0, ServerFile.FileName.Replace("\", "/").LastIndexOf("/"))
                    myRet = myFileCopy.CheckOrMakeDestinationDir(constAppDir & FileDir)
                    If Not myRet Then Return False
                End If

                If cntRet Then
                    verRet = False
                    '�N���C�A���g�Ƀo�[�W�����t�@�C��������ꍇ�͍ŐV�̂݃_�E�����[�h
                    For Each ClientFile In ClientVersion
                        If ServerFile.FileName.Equals(ClientFile.FileName) Then
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ START
'//                         If ServerFile.TimeStamp > ClientFile.TimeStamp Or ServerFile.Size <> ClientFile.Size Then
                            '�^�C���X�^���v�́A�N���������܂Ŕ�r����i�b�͔�r���Ȃ��j
                            If ServerFile.TimeStamp.Substring(0, 12) > ClientFile.TimeStamp.Substring(0, 12) _
                            Or ServerFile.Size <> ClientFile.Size Then
'// MOD 2007.11.28 ���s�j���� �^�C���X�^���v�̕b��r�p�~ END
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� START
'// MOD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ START
'//                             If Not( ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                And  ServerFile.TimeStamp.Substring(0,8).Equals("20081014")) Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                             If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                And gbInitProxyExists _
'//                                And ServerFile.TimeStamp.Substring(0, 8).Equals("20081014")) Then
                                If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
                                   And gbInitProxyExists _
                                   And (ServerFile.TimeStamp.Substring(0, 8).Equals("20081014") _
                                     Or ServerFile.TimeStamp.Substring(0, 8).Equals("20091019"))) Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ END
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� END
                                    '�o�[�W�������Ⴄ�ꍇ�_�E�����[�h���s��
                                    With _myLogRecord
                                        .Target = ServerFile.FileName & "�̍ŐV�o�[�W�������_�E�����[�h���܂�"
                                        .Result = "OK"
                                        .Remark = ""
                                    End With
                                    _myLog.WriteLog(_myLogRecord)
                                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                    If Not myRet Then
                                        '�_�E�����[�h�Ɏ��s�����ꍇ�͋N�����s��Ȃ�
                                        Return False
                                    End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� START
                                End If
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� END
'// ADD 2005.07.22 ���s�j�����J START
                            Else
                                appFile = System.IO.Path.Combine(_DestinationAppPath, ClientFile.FileName)
                                If System.IO.File.Exists(appFile) Then
'// ADD 2007.10.26 ���s�j���� �b�`�a�Ώۂ̃t�@�C���̓T�C�Y��r�͍s��Ȃ� START
                                    If ClientFile.FileName <> "IS2Client.exe" _
                                    And ClientFile.FileName <> "is2AdminClient.exe" Then
'// ADD 2007.10.26 ���s�j���� �b�`�a�Ώۂ̃t�@�C���̓T�C�Y��r�͍s��Ȃ� END
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� START
'// MOD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ START
'//                                     If Not( ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                        And  ServerFile.TimeStamp.Substring(0,8).Equals("20081014")) Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//                                     If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
'//                                        And gbInitProxyExists _
'//                                        And ServerFile.TimeStamp.Substring(0, 8).Equals("20081014")) Then
                                        If Not (ServerFile.FileName.Equals("AutoUpGradeUtility.dll") _
                                           And gbInitProxyExists _
                                           And (ServerFile.TimeStamp.Substring(0, 8).Equals("20081014") _
                                             Or ServerFile.TimeStamp.Substring(0, 8).Equals("20091019"))) Then
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2009.10.03 ���s�j���� [proxy.ini]�����݂��Ȃ����A�v���L�V���ŋ@�\��~ END
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� END
                                            '�_�E�����[�h�Ɏ��s�����ꍇ
                                            Dim appfs As System.IO.FileStream = New System.IO.FileStream(appFile, System.IO.FileMode.Open)
                                            appSize = appfs.Length
                                            appfs.Close()
                                            If appSize <> ClientFile.Size Then
                                                With _myLogRecord
                                                    .Target = ServerFile.FileName & "�̍ŐV�o�[�W�������_�E�����[�h���܂�"
                                                    .Result = "OK"
                                                    .Remark = ""
                                                End With
                                                _myLog.WriteLog(_myLogRecord)
                                                myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                                If Not myRet Then
                                                    '�_�E�����[�h�Ɏ��s�����ꍇ�͋N�����s��Ȃ�
                                                    Return False
                                                End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                                                DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END
                                            End If
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� START
                                        End If
'// MOD 2009.08.05 ���s�j���� �u���a�X�g���l�b��Ή� END
'// ADD 2007.10.26 ���s�j���� �b�`�a�Ώۂ̃t�@�C���̓T�C�Y��r�͍s��Ȃ� START
                                    End If
'// ADD 2007.10.26 ���s�j���� �b�`�a�Ώۂ̃t�@�C���̓T�C�Y��r�͍s��Ȃ� END
                                Else
                                    With _myLogRecord
                                        .Target = ServerFile.FileName & "�̍ŐV�o�[�W�������_�E�����[�h���܂�"
                                        .Result = "OK"
                                        .Remark = ""
                                    End With
                                    _myLog.WriteLog(_myLogRecord)
                                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                                    If Not myRet Then
                                        '�_�E�����[�h�Ɏ��s�����ꍇ�͋N�����s��Ȃ�
                                        Return False
                                    End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END
                                End If
'// ADD 2005.07.22 ���s�j�����J END
                            End If
                            verRet = True
                        End If
                    Next
                    If Not verRet Then
                        With _myLogRecord
                            .Target = "�V�K�t�@�C��" & ServerFile.FileName & "���_�E�����[�h���܂�"
                            .Result = "OK"
                            .Remark = ""
                        End With
                        _myLog.WriteLog(_myLogRecord)
                        myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                        If Not myRet Then
                            '�_�E�����[�h�Ɏ��s�����ꍇ�͋N�����s��Ȃ�
                            Return False
                        End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                        DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END
                        verRet = True
                    Else
                        With _myLogRecord
                            .Target = ServerFile.FileName & "�͍ŐV�o�[�W�����ł�"
                            .Result = "OK"
                            .Remark = ""
                        End With
                        _myLog.WriteLog(_myLogRecord)
                    End If
                Else
                    With _myLogRecord
                        .Target = ServerFile.FileName & "���_�E�����[�h���܂�"
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    '�N���C�A���g�Ƀo�[�W�����t�@�C�����Ȃ��ꍇ�͑S�ă_�E�����[�h
                    myRet = Startup(sourceDirName & ServerFile.FileName, FileDir, download, downloadOnly, localExecute, logClear)
                    If Not myRet Then
                        '�_�E�����[�h�Ɏ��s�����ꍇ�͋N�����s��Ȃ�
                        Return False
                    End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                    DeleteCabFlagFile(ServerFile.FileName)
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END
                End If

'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� START
                If ServerFile.FileName.Substring(ServerFile.FileName.LastIndexOf(".") + 1).Equals("txt") Then
'// ADD 2007.10.12 ���s�j���� �b�`�a�t�@�C���̉� START
                    ChangeFileStamp(ServerFile.FileName, ServerFile.TimeStamp)
                    CheckCabVerFile(sourceDirName, ServerFile.FileName)
'// ADD 2007.10.12 ���s�j���� �b�`�a�t�@�C���̉� END

                    ' �b�`�a�t�@�C�������݂��ACAB�t���O�t�@�C�����Ȃ���΁A�𓀂�����
                    Dim srcFullName, flgFullName As String
                    srcFullName = Path.Combine(_DestinationPath, ServerFile.FileName.Substring(0, ServerFile.FileName.LastIndexOf(".") + 1) + "cab")
                    flgFullName = srcFullName + "_flg"
                    If System.IO.File.Exists(srcFullName) And System.IO.File.Exists(flgFullName) = False Then
                        '�ۗ��@�A�v�����N�����Ă����ꍇ�ɂ͏I��������
                        '�b�`�a�t�@�C���̉�
                        myRet = UnCab(srcFullName, _DestinationAppPath)
                        '�t���O�t�@�C���̍쐬
                        Dim fsflg As FileStream = System.IO.File.Create(flgFullName)
                        fsflg.Close()
                    End If
                End If
'// ADD 2007.10.05 ���s�j���� �b�`�a�t�@�C���̉� END

                ' �����J�E���g��1�X�e�b�v�i�߂�
                iCount = iCount + 1
                waitDlg.PerformStep()
                waitDlg.Update()
                System.Windows.Forms.Application.DoEvents()
            Next

            If download Then
                '�T�[�o�[�̃o�[�W���������N���C�A���g�̃o�[�W�������ɃR�s�[
                Dim fs As New IO.FileStream(Path.Combine(_DestinationPath, "VersionFile.xml"), FileMode.Create)
                ServerSerializer.Serialize(fs, ServerVersion)
                fs.Close()
            End If

            '���[�J���t�H���_���`�F�b�N���s�v�t�@�C�����폜
            Dim localFile As New LocalFileCheck
            Dim localFileList As ArrayList
            Dim localFileName As String
            Dim localCheck As Boolean
            localFile.LocalFileList(_DestinationAppPath)
            localFileList = localFile.GetFileList

            For Each localFileName In localFileList
                localCheck = False
                For Each ServerFile In ServerVersion
'// MOD 2005.05.10 ���s�j���� �t�@�C������r�̕ύX START
'//                 If ServerFile.FileName.Replace("\", "/").Equals(localFileName.Substring(_DestinationAppPath.Length + 1).Replace("\", "/")) Then
                    If ServerFile.FileName.Replace("\", "/").ToUpper().Equals( _
                            localFileName.Substring(_DestinationAppPath.Length + 1).Replace("\", "/").ToUpper()) Then
'// MOD 2005.05.10 ���s�j���� �t�@�C������r�̕ύX END
                        localCheck = True
                    End If
                Next
                If Not localCheck Then
                    With _myLogRecord
                        .Target = "�s�v�t�@�C�� " & localFileName & "���폜���܂�"
                        .Result = "OK"
                        .Remark = ""
                    End With
                    _myLog.WriteLog(_myLogRecord)
                    '���[�J���t�@�C���̍폜
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
            ' �i�s�󋵃_�C�A���O�����
            waitDlg.Close()
        End Try


        Return True
    End Function

    '
    '   �A�v���P�[�V�����̎��s
    '
    Public Function ExecApp(ByVal sourceFileName As String, ByVal cmdPara() As String, ByVal localExecute As Boolean) As Boolean
        Dim myRet As Boolean
        Dim myAppDom As AppDomain

'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� START
'//     '�A�v���P�[�V�����h���C���̐ݒ�
'//     myAppDom = GetAppDomain(sourceFileName, localExecute)
'//     If myAppDom Is Nothing Then Return False
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� END
'// ADD 2007.11.29 ���s�j���� �����I�K�x�[�W�R���N�g���s�� START
        GC.Collect()    '�������[�N���[���A�b�v
'// ADD 2007.11.29 ���s�j���� �����I�K�x�[�W�R���N�g���s�� END

'// ADD 2009.07.30 ���s�j���� exe��dll���Ή� START
'//     '�A�v���P�[�V�����̎��s
'//     myRet = Execute(myAppDom, sourceFileName, cmdPara, localExecute)
        Dim sDllFileName As String = ""
        Dim sDllFlagName As String = ""
        sDllFileName = sourceFileName.Substring(0, sourceFileName.Length - 4) + ".dll"
        sDllFlagName = sDllFileName & constConfigExt

        If System.IO.File.Exists(Path.Combine(_DestinationAppPath, sDllFlagName)) _
        And System.IO.File.Exists(Path.Combine(_DestinationAppPath, sDllFileName)) Then
            '//�c�k�k�̌Ăяo��
            myRet = ExecuteDll(sDllFileName, cmdPara)
        Else
            '//�d�w�d�̌Ăяo��
            myRet = Execute(myAppDom, sourceFileName, cmdPara, localExecute)
        End If
'// ADD 2009.07.30 ���s�j���� exe��dll���Ή� END
        Return myRet
    End Function

    '
    '   �X�^�[�g�E�A�b�v����
    '
    '   �N���t�@�C���p�X�A�_�E�����[�h���邩�A�_�E�����[�h�݂̂��A���[�J���N�����A���O���N���A�[���邩
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
                .Status = "�J�n"
                .SubStatus = "AutoUpGrade"
                .Target = "�������J�n���܂�"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            _myLogRecord.Status = ""

            '�_�E�����[�h�݂̂̏ꍇ�̃p�����[�^��ݒ�
            '   �_�E�����[�h�͍s���A���[�J���N���͍s��Ȃ�
            If downloadOnly Then
                download = True
                localExecute = False
            End If
            '�t�@�C�����`�F�b�N
            myRet = CheckFile(sourceFileName, myFileCopy)
            If myRet = False Then Return False

            '�t�@�C���̃_�E�����[�h
            myRet = GetFile(sourceFileName, fileDir, download, localExecute, myFileCopy)
            If myRet = False Then Return False

'// DEL 2007.11.29 ���s�j���� �A�Z���u���_�E�����[�h�͕��ׂ�������̂ōs��Ȃ� START
'//         '�A�Z���u���̎Q�ƃA�Z���u�����_�E�����[�h
'//         If download And sourceFileName.Substring(sourceFileName.LastIndexOf(".") + 1).Equals("exe") Then
'//             myRet = CheckAssemblies()
'//             GC.Collect()    '�������[�N���[���A�b�v
'//         End If
'// DEL 2007.11.29 ���s�j���� �A�Z���u���_�E�����[�h�͕��ׂ�������̂ōs��Ȃ� END

            '�_�E�����[�h�̂�
            If downloadOnly Then
                With _myLogRecord
                    .Status = "�I��"
                    .SubStatus = "AutoUpGrade"
                    If localExecute Then
                        .Target = sourceFileName & "�̃_�E�����[�h�����Ɏ��s���܂���"
                        .Result = "NG"
                        .Remark = ""
                        myRet = False
                    Else
                        .Target = sourceFileName & "�̃_�E�����[�h�������I�����܂���"
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
    '   �t�@�C���������������ǂ����̃`�F�b�N
    '
    Private Function CheckFile(ByVal sourceFileName As String, _
                               ByRef myFileCopy As FileCopy) As Boolean
        If sourceFileName = String.Empty Then
            With _myLogRecord
                .Target = "�t�@�C�������w�肳��Ă��܂���"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If

        _FileName = myFileCopy.GetFileName(sourceFileName)
        If _FileName = String.Empty Then
            With _myLogRecord
                .Target = "�t�@�C�������w�肳��Ă��܂���"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        End If

        Return True
    End Function


    '
    '   �t�@�C���̃_�E�����[�h���s��
    '
    Private Function GetFile(ByVal sourceFileName As String, _
                        ByVal fileDir As String, _
                        ByRef download As Boolean, _
                        ByRef localExecute As Boolean, _
                        ByRef myFileCopy As FileCopy) As Boolean
        Dim myRet As Boolean = True

        If download Then
            '�t�@�C���̃_�E�����[�h
            myRet = myFileCopy.CopyFile(sourceFileName, _DestinationAppPath & fileDir)
            _FileName = myFileCopy.FileName
            _SourcePath = myFileCopy.PathName
            If myRet Then
                With _myLogRecord
                    .Target = sourceFileName & "��" & _DestinationAppPath & fileDir & "�փ_�E�����[�h�ɐ������܂���"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            Else
                With _myLogRecord
                    .Target = sourceFileName & "�̃_�E�����[�h�Ɏ��s���܂���"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
        End If

        Return myRet

    End Function

    '
    '   �A�v���P�[�V�����E�h���C�����쐬���܂�
    '
    Private Function GetAppDomain(ByVal sourceFileName As String, _
                                 ByVal localExecute As Boolean) As AppDomain
        Dim myAppSetup As AppDomainSetup
        Dim myAppDom As AppDomain
        Dim myName As String

        '�A�v���P�[�V�����h���C���̐ݒ�
        myAppSetup = New AppDomainSetup
        myAppSetup.ConfigurationFile = Path.Combine(_DestinationAppPath, sourceFileName) & constConfigExt
'//     myAppSetup.ConfigurationFile = Path.Combine(_DestinationConfigPath, sourceFileName) & constConfigExt
        myAppSetup.ShadowCopyFiles = "true"
        If localExecute Then
            If Not File.Exists(Path.Combine(_DestinationAppPath, sourceFileName)) Then
                With _myLogRecord
                    .Status = "�I��"
                    .Target = Path.Combine(_DestinationAppPath, sourceFileName) & "��������܂���"
                    .Result = "NG"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                _myLogRecord.Status = ""
                Return Nothing
            End If
            myAppSetup.ApplicationBase = _DestinationAppPath
        End If
        '�A�v���P�[�V�����E�h���C���������s���̃A�Z���u�����Ŏ擾
        myName = [Assembly].GetExecutingAssembly.FullName
        myName = myName.Substring(0, myName.IndexOf(","))
        myAppDom = AppDomain.CreateDomain(String.Concat(myName, " : ", sourceFileName), _
                                          AppDomain.CurrentDomain.Evidence, _
                                          myAppSetup)
        Return myAppDom
    End Function

    '
    '   �A�v���P�[�V���������s���܂�
    '
    Private Function Execute(ByRef myAppDom As AppDomain, _
                            ByVal sourceFileName As String, _
                            ByVal cmdPara() As String, _
                            ByVal localExecute As Boolean) As Boolean
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� START
'//     Dim myName As String = myAppDom.FriendlyName
        Dim myName As String = " "
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� END

        '�J�����g�f�B���N�g���̕ύX
        'System.IO.Directory.SetCurrentDirectory(_DestinationAppPath)

        '�A�v���P�[�V�����̎��s
        Try

            If localExecute Then
                With _myLogRecord
                    .Target = myName & "��" & Path.Combine(_DestinationAppPath, sourceFileName) & "�����s���s���܂�"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
                'myAppDom.ExecuteAssembly(Path.Combine(_DestinationAppPath, sourceFileName), myAppDom.Evidence, cmdPara)
'// MOD 2005.05.31 ���s�j�ɉ� �X���b�h��p�~ START
                '�X���b�h���쐬���A�J�n����
                Dim process As New System.Diagnostics.Process

                With process.StartInfo

                    .Arguments = cmdPara(0) ' �R�}���h���C������
                    .WorkingDirectory = _DestinationAppPath ' ��ƃf�B���N�g��
                    .FileName = Path.Combine(_DestinationAppPath, sourceFileName) ' ���s����t�@�C�� (*.exe�łȂ��Ă��ǂ�)
'// ADD 2008.06.11 ���s�j���� ���O�C����ʂ���Ƀm�[�}���\���� START
                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
'// ADD 2008.06.11 ���s�j���� ���O�C����ʂ���Ƀm�[�}���\���� END
                End With
                Try
                    process.Start()

                Catch ex As System.ComponentModel.Win32Exception
                    ' �t�@�C����������Ȃ������ꍇ�A
                    ' �֘A�t����ꂽ�A�v���P�[�V������������Ȃ������ꍇ��

                Catch ex As System.Exception
                    '���̑�

                End Try
'// MOD 2005.05.31 ���s�j�ɉ� �X���b�h��p�~ END
            Else
                With _myLogRecord
                    .Target = myName & "��" & sourceFileName & "�����s���s���܂�"
                    .Result = "OK"
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� START
'//             myAppDom.ExecuteAssembly(sourceFileName)
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� END
            End If
            With _myLogRecord
                .Status = "�I��"
                .Target = myName & "�Ŏ��s���I�����܂���"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

        Catch e As Exception
            With _myLogRecord
                .Status = "�I��"
                .Target = myName & "�ŃG���[���������܂���"
                .Result = "NG"
                .Remark = e.Message
            End With
            _myLog.WriteLog(_myLogRecord)
            Return False
        Finally
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� START
'//         AppDomain.Unload(myAppDom)
'// DEL 2007.11.29 ���s�j���� �A�v���P�[�V�����h���C���̐ݒ�͕��ׂ�������̂ōs��Ȃ� START
'// MOD 2005.05.31 ���s�j�ɉ� �X���b�h��p�~ START
            Dim process As New System.Diagnostics.Process

            With process.StartInfo
                .Arguments = _ClientMutex ' �R�}���h���C������
                .WorkingDirectory = _DestinationPath ' ��ƃf�B���N�g��
'// MOD 2005.06.03 ���s�j�ɉ� ���s�t�@�C���p�X�̕ύX START
'//             '.FileName = Path.Combine(_DestinationPath, "CopyAutoUpGrade.exe") ' ���s����t�@�C�� (*.exe�łȂ��Ă��ǂ�)
                .FileName = Path.Combine(_DestinationAppPath, "CopyAutoUpGrade.exe") ' ���s����t�@�C�� (*.exe�łȂ��Ă��ǂ�)
'// MOD 2005.06.03 ���s�j�ɉ� ���s�t�@�C���p�X�̕ύX END
'// ADD 2008.06.11 ���s�j���� [CopyAutoUpGrade]���\���� START
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
'// ADD 2008.06.11 ���s�j���� [CopyAutoUpGrade]���\���� END
            End With
            Try
                process.Start()

            Catch ex As System.ComponentModel.Win32Exception
                ' �t�@�C����������Ȃ������ꍇ�A
                ' �֘A�t����ꂽ�A�v���P�[�V������������Ȃ������ꍇ��

            Catch ex As System.Exception
                '���̑�

            End Try
'// MOD 2005.05.31 ���s�j�ɉ� �X���b�h��p�~ END

            '�J�����g�f�B���N�g���̕ύX
            'System.IO.Directory.SetCurrentDirectory(_DestinationPath)
        End Try

        Return True
    End Function

    '   �A�Z���u���̎Q�Ə������ɉ�������
    '       
    Private Function CheckAssemblies() As Boolean
        Dim myRet As Boolean
        Dim myReference As ReferenceAssembly = New ReferenceAssembly
        With _myLogRecord
            .Target = "�A�Z���u���Q�Ƃ̃`�F�b�N���J�n���܂�"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        myRet = myReference.SearchAssemblies(_FileName, _SourcePath, _DestinationAppPath)

        With _myLogRecord
            If myRet Then
                .Target = "�A�Z���u���Q�Ƃ̃`�F�b�N���I�����܂���"
                .Result = "OK"
                .Remark = _FileName
            Else
                .Target = "�A�Z���u���Q�Ƃ̃`�F�b�N�Ń`�F�b�N�ł��Ȃ����̂�����܂���"
                .Result = "NG"
                .Remark = _FileName
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myRet
    End Function

'// ADD 2007.06.15 ���s�j���� �b�`�a�t�@�C���̉� START
    '       
    '   �b�`�a�t�@�C���̉𓀂���
    '       
    Private Function UnCab(ByVal cabFile As String, ByVal extractDir As String) As Boolean
        Dim myRet As Boolean

        myRet = False
        With _myLogRecord
            .Target = "�b�`�a�𓀂��J�n���܂�"
            .Result = "OK"
            .Remark = ""
        End With
        _myLog.WriteLog(_myLogRecord)

        '�X���b�h���쐬���A�J�n����
        Dim process As New System.Diagnostics.Process

        With process.StartInfo
            .Arguments = " -r """ + cabFile + """" + " """ + extractDir + """" ' �R�}���h���C������
            .WorkingDirectory = extractDir ' ��ƃf�B���N�g��
            .FileName = "EXPAND.EXE"  ' ���s����t�@�C�� (*.exe�łȂ��Ă��ǂ�)
'// ADD 2008.06.11 ���s�j���� [EXPAND]���\���� START
            '            .WindowStyle = ProcessWindowStyle.Minimized
            .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
'// ADD 2008.06.11 ���s�j���� [EXPAND]���\���� END
        End With

        Try
            process.Start()
'// ADD 2007.06.21 ���s�j���� �v���Z�X�̏I����҂� START
            process.WaitForExit(25000) '�v���Z�X�I�����邩25�b�o�߂���܂ő҂�
'// ADD 2007.06.21 ���s�j���� �v���Z�X�̏I����҂� END
            myRet = True

        Catch ex As System.ComponentModel.Win32Exception
            ' �t�@�C����������Ȃ������ꍇ�A
            ' �֘A�t����ꂽ�A�v���P�[�V������������Ȃ������ꍇ��

        Catch ex As System.Exception
            '���̑�

        End Try

        With _myLogRecord
            If myRet Then
                .Target = "�b�`�a�𓀂��I�����܂���"
                .Result = "OK"
                .Remark = cabFile
            Else
                .Target = "�b�`�a�𓀂��ł��܂���ł���"
                .Result = "NG"
                .Remark = cabFile
            End If
        End With
        _myLog.WriteLog(_myLogRecord)

        Return myRet
    End Function
'// ADD 2007.06.15 ���s�j���� �b�`�a�t�@�C���̉� END
'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� START
    '
    '   �h�d�̃v���L�V�̐ݒ���e�����W�X�g������擾
    '
    Private Sub �h�d�v���L�V�ݒ���擾()
        Try
            Dim regkey As Microsoft.Win32.RegistryKey = _
             Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Internet Settings", False)
            '//�L�[�����݂��Ȃ��Ƃ��� null ���Ԃ����
            If regkey Is Nothing Then Return

            '// 
            '// [�����\���X�N���v�g���g�p����ꍇ]�̃`�F�b�N�L��
            '// 
            Dim sAutoConfigURL As String = ""
            If Not (regkey.GetValue("AutoConfigURL") Is Nothing) Then
                sAutoConfigURL = CType(regkey.GetValue("AutoConfigURL"), String)
                With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                 .Target = "R AC[" & �Í����Q(CType(regkey.GetValue("AutoConfigURL"), String)) & "]"
                    .Target = "R AC[" & CType(regkey.GetValue("AutoConfigURL"), String) & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            '// 
            '// [�k�`�m�Ƀv���L�V�T�[�o���g�p����ꍇ]�̃`�F�b�N�L��
            '// 
            Dim iProxyEnable As Integer = 0
            If Not (regkey.GetValue("ProxyEnable") Is Nothing) Then
                iProxyEnable = CType(regkey.GetValue("ProxyEnable"), Integer)
                With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                 .Target = "R PE[" & �Í����Q(iProxyEnable.ToString()) & "]"
                    .Target = "R PE[" & iProxyEnable.ToString() & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If
            Dim sProxyServer As String() = {""}
            If Not (regkey.GetValue("ProxyServer") Is Nothing) Then
                sProxyServer = CType(regkey.GetValue("ProxyServer"), String).Split(";")
                With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                 .Target = "R PS[" & �Í����Q(CType(regkey.GetValue("ProxyServer"), String)) & "]"
                    .Target = "R PS[" & CType(regkey.GetValue("ProxyServer"), String) & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
                    .Result = ""
                    .Remark = ""
                End With
                _myLog.WriteLog(_myLogRecord)
            End If

            Dim sProxyOverride As String() = {""}
            If Not (regkey.GetValue("ProxyOverride") Is Nothing) Then
                sProxyOverride = CType(regkey.GetValue("ProxyOverride"), String).Split(";")
                With _myLogRecord
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� START
'//                 .Target = "R PO[" & �Í����Q(CType(regkey.GetValue("ProxyOverride"), String)) & "]"
                    .Target = "R PO[" & CType(regkey.GetValue("ProxyOverride"), String) & "]"
'// MOD 2010.06.21 ���s�j���� ���O�̈Í������� END
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
                .Target = "IEP�ݒ���擾 Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            Return
        End Try
    End Sub
    '
    '   �f�t�H���g�v���L�V�̐ݒ���s��
    '
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'// Private Function �f�t�H���g�v���L�V�ݒ�() As Integer
    Private Function �f�t�H���g�v���L�V�ݒ�() As WebExceptionStatus
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        '
        '   �f�t�H���g�v���L�V�̐ݒ���s��
        '
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//     Dim iRet As Integer = 0
        Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Try
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()

            Dim sRet As String = sv_init.wakeupDB()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//         iRet = 1
            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Catch ex As System.Net.WebException
            With _myLogRecord
                .Target = "P�ݒ�(D) WebException:" & ex.Status.ToString()
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Catch ex As Exception
            With _myLogRecord
                .Target = "P�ݒ�(D) Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//         iRet = -2
'//         Return iRet
            wRet = WebExceptionStatus.UnknownError
            Return wRet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        End Try

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//     Return iRet
        Return wRet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
    End Function

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'// Private Function �v���L�V�ݒ�(ByVal sProxyAdr As String, ByVal iProxyNo As Integer) As Integer
    Private Function �v���L�V�ݒ�(ByVal sProxyAdr As String, ByVal iProxyNo As Integer) As WebExceptionStatus
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
        Return �v���L�V�ݒ�Q(sProxyAdr, iProxyNo, "", "")
    End Function

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'// Private Function �v���L�V�ݒ�Q(ByVal sProxyAdr As String, ByVal iProxyNo As Integer _
'//         , ByVal sProxyId As String, ByVal sProxyPa As String) As Integer
    Private Function �v���L�V�ݒ�Q(ByVal sProxyAdr As String, ByVal iProxyNo As Integer _
            , ByVal sProxyId As String, ByVal sProxyPa As String) As WebExceptionStatus
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//     Dim iRet As Integer = 0
        Dim wRet As WebExceptionStatus = WebExceptionStatus.UnknownError
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Try
            If sProxyAdr.Length > 0 Then
                If iProxyNo > 0 Then
                    gProxy = New System.Net.WebProxy(sProxyAdr, iProxyNo)
                Else
                    gProxy = New System.Net.WebProxy(sProxyAdr)
                End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
                If sProxyId.Length > 0 Then
                    gProxy.Credentials = New NetworkCredential(sProxyId, sProxyPa)
                End If
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
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
'//					.Target = "�v���L�V�ݒ� Exception:"
'//					.Result = "NG"
'//					.Remark = ex.Message
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'//				return false
'//			end try

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//         iRet = 1
            wRet = WebExceptionStatus.Success
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Catch ex As System.Net.WebException
            With _myLogRecord
                .Target = "P�ݒ�( ) WebException:" & ex.Status.ToString()
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            '//���̏�Ԃɖ߂�
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
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
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        Catch ex As Exception
            With _myLogRecord
                .Target = "P�ݒ�( ) Exception:"
                .Result = "NG"
                .Remark = ex.Message
            End With
            _myLog.WriteLog(_myLogRecord)

            '//���̏�Ԃɖ߂�
            System.Net.GlobalProxySelection.Select = System.Net.WebProxy.GetDefaultProxy()
            sv_init.Proxy = System.Net.WebProxy.GetDefaultProxy()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//         iRet = -2
'//         Return iRet
            wRet = WebExceptionStatus.UnknownError
            Return wRet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
        End Try

'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//     Return iRet
        Return wRet
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
    End Function
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
    '
    '   �ڑ��G���[���b�Z�[�W�̕ҏW
    '
    Function �ڑ��G���[���b�Z�[�W�ҏW(wStatus As WebExceptionStatus) As String
        Dim sRet As String = ""
        Select Case wStatus
           Case WebExceptionStatus.ConnectFailure
				'// �v���L�V�T�[�o�ݒ�G���[
                sRet = "�����F�v���L�V�ݒ�Ȃ�"
           Case WebExceptionStatus.ConnectionClosed 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.KeepAliveFailure 
                sRet = wStatus.ToString()
           Case WebExceptionStatus.MessageLengthLimitExceeded
                sRet = wStatus.ToString()
           Case WebExceptionStatus.NameResolutionFailure 
				'// �P�[�u���ڑ��G���[
				'// �v���L�V�|�[�g�ԍ��ݒ�G���[
				'// �c�m�r�ݒ�G���[
                sRet = "�����F�P�[�u���ڑ��A�v���L�V�ݒ�A�c�m�r�ݒ�Ȃ�"
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
                sRet = "�ڑ�����"
           Case WebExceptionStatus.Timeout 
                sRet = "�^�C���A�E�g�G���[�A�����F�v���L�V�ݒ�Ȃ�"
           Case WebExceptionStatus.TrustFailure 
				'// ���t�ݒ�G���[
                sRet = "�����F���t�ݒ�A�r�r�k�ݒ�Ȃ�"
           Case WebExceptionStatus.UnknownError
                sRet = "����`�G���["
        End Select
        Return sRet        
    End Function
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
    '
    '   �r�i�h�r�`�F�b�N���s��
    '
    Private Function �r�i�h�r�`�F�b�N(ByVal tex As String, ByVal name As String, ByRef sUnicode As String, ByRef bSjis As Byte()) As Boolean
        '//�t�ϊ����Ăr�i�h�r�������`�F�b�N����
        Dim sRevUnicode As String = System.Text.Encoding.GetEncoding("shift-jis").GetString(bSjis)
        Dim sErrChars As String = ""
        For iPos As Integer = 0 To sUnicode.Length - 1
            If iPos >= sRevUnicode.Length Then Exit For
            If sUnicode.Chars(iPos) <> sRevUnicode.Chars(iPos) Then
                sErrChars += sUnicode.Chars(iPos)
            End If
        Next
        If sErrChars.Length > 0 Then
            '//				MessageBox.Show(name + "�Ɏg�p�ł��Ȃ�����������܂�\n" _
            '//					+ "�w" + sErrChars + "�x", _
            '//					"���̓`�F�b�N",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "�r�i�h�r�`�F�b�N[" & name & "][" & sErrChars & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If
        Return True
    End Function
    '
    '   ���p�`�F�b�N���s��
    '
    Protected Function ���p�`�F�b�N(ByVal tex As String, ByVal name As String) As Boolean
        Dim sUnicode As String = tex
        Dim bSjis As Byte() = System.Text.Encoding.GetEncoding("shift-jis").GetBytes(sUnicode)
        If Not �r�i�h�r�`�F�b�N(tex, name, sUnicode, bSjis) Then Return False
        If bSjis.Length <> sUnicode.Length Then
            '//			MessageBox.Show(name + "�͔��p�����œ��͂��Ă�������", _
            '//				"���̓`�F�b�N",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "���p�`�F�b�N�G���[[" & name & "][" & tex & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End If

'// MOD 2011.09.22 ���s�j���� �L���`�F�b�N�p�~ START
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
'//				'//				MessageBox.Show(name + "�ɋL�������͂���Ă��܂�","���̓`�F�b�N",MessageBoxButtons.OK)
'//				With _myLogRecord
'//					.Target = "���p�L���`�F�b�N�G���[[" & name & "][" & tex & "]"
'//					.Result = "NG"
'//					.Remark = ""
'//				End With
'//				_myLog.WriteLog(_myLogRecord)
'//
'//				Return False
'//			End If
'//		Next
'// MOD 2011.09.22 ���s�j���� �L���`�F�b�N�p�~ END
        Return True
    End Function
    '
    '   ���l�`�F�b�N���s��
    '
    Protected Function ���l�`�F�b�N(ByVal tex As String, ByVal name As String) As Boolean
        Try
            Dim lChk As Long = Long.Parse(tex.Replace(",", ""))
            Return True
        Catch ex As Exception
            '//			MessageBox.Show(name + "�ɐ��l�����͂���Ă��܂���","���̓`�F�b�N",MessageBoxButtons.OK)
            With _myLogRecord
                .Target = "���l�`�F�b�N�G���[[" & name & "][" & tex & "]"
                .Result = "NG"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Return False
        End Try
    End Function
'// ADD 2009.07.29 ���s�j���� �v���L�V�Ή� END

'// ADD 2009.07.30 ���s�j���� exe��dll���Ή� 
    '
    '   dll�̎��s
    '
    Public Function ExecuteDll(ByVal sDllFileName As String, ByVal cmdPara() As String) As Boolean
        Dim myRet As Boolean
        Try
            '�J�����g�f�B���N�g���̕ύX
            System.IO.Directory.SetCurrentDirectory(_DestinationAppPath)

            With _myLogRecord
                .Target = "[" & sDllFileName & "]�Ǎ�"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            Dim asm As System.Reflection.Assembly
            asm = System.Reflection.Assembly.LoadFrom(sDllFileName)

            '���C�u�����Ǎ��ݗp�ϐ�
            Dim oLibrary As New Object
            '�ďo���֐��ւ̈���
            Dim methodParams() As Object = {cmdPara}

            With _myLogRecord
                .Target = "[" & sDllFileName & "]�ďo"
                .Result = "OK"
                .Remark = ""
            End With
            _myLog.WriteLog(_myLogRecord)

            '�N���X���C���X�^���X��
            oLibrary = asm.CreateInstance("IS2Client.���j���[")

            '���\�b�h�����擾����
            Dim mi As System.Reflection.MethodInfo = oLibrary.GetType.GetMethod("Main")

            '���\�b�h�����s����
            mi.Invoke(oLibrary, methodParams)

        Catch ex As System.IO.FileNotFoundException
            With _myLogRecord
                .Target = "ExecApp " & ex.FileName & "���݂���܂���ł���"
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
            '�J�����g�f�B���N�g���̕ύX
            System.IO.Directory.SetCurrentDirectory(_DestinationPath)

            Dim process As New System.Diagnostics.Process

            With process.StartInfo
                .Arguments = _ClientMutex ' �R�}���h���C������
                .WorkingDirectory = _DestinationPath ' ��ƃf�B���N�g��
                .FileName = Path.Combine(_DestinationAppPath, "CopyAutoUpGrade.exe")
                .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            End With
            Try
                process.Start()
            Catch ex As System.ComponentModel.Win32Exception
                ' �t�@�C����������Ȃ������ꍇ�A
                ' �֘A�t����ꂽ�A�v���P�[�V������������Ȃ������ꍇ��
            Catch ex As System.Exception
                '���̑�
            End Try
        End Try
        Return myRet
    End Function
'// ADD 2009.07.30 ���s�j���� exe��dll���Ή� END
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
    Dim bKey As Byte() = {51, 168, 96, 2, 36, 12, 74, 143, 11, 85, 61, 233, 202, 170, 114, 59}
    Dim bIv As Byte() = {100, 223, 207, 80, 29, 100, 53, 152}
    Dim s�_�~�[ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz01234567890!#$%&()=~|`{+*}<>?_"
    Dim rndObj As New Random
    Function �Í����Q(ByVal sText As String) As String
        '//�擪����і����Ƀ_�~�[������ǉ�
        sText = s�_�~�[.Substring(rndObj.Next(0, s�_�~�[.Length - 1), 1) + sText
        sText = sText + s�_�~�[.Substring(rndObj.Next(0, s�_�~�[.Length - 1), 1)

        Return �Í���(sText)
    End Function
    Function �Í���(ByVal sText As String) As String
        Dim bText As Byte()
        Dim bRet As Byte()
        Dim sRet As String = ""

        Try
            Dim msEncrypt As MemoryStream = New MemoryStream
            Dim rc2 As RC2CryptoServiceProvider = New RC2CryptoServiceProvider

            '// CryptoStream�I�u�W�F�N�g���쐬����
            Dim transform As ICryptoTransform = rc2.CreateEncryptor(bKey, bIv) '// Encryptor���쐬����
            Dim cryptoStream As cryptoStream = New cryptoStream(msEncrypt, transform, CryptoStreamMode.Write)

            '// �Í�������Ώۂ��o�C�g�z��Ƃ��ēǂݍ���
            bText = Encoding.GetEncoding("shift-jis").GetBytes(sText)

            '// CryptoStream�ɂ���ĈÍ������ď�������
            cryptoStream.Write(bText, 0, bText.Length)
            cryptoStream.FlushFinalBlock()

            bRet = msEncrypt.ToArray()
            '//sRet = Encoding.GetEncoding("shift-jis").GetString(bRet)
            For iCnt As Integer = 0 To bRet.Length - 1
                sRet = sRet + bRet(iCnt).ToString("X2")
            Next iCnt

            '// CryptoStream�����
            cryptoStream.Close()

            '// FileStream�����
            msEncrypt.Close()

            rc2 = Nothing
            cryptoStream = Nothing
            msEncrypt = Nothing
        Catch ex As Exception
            sRet = ex.Message
        End Try

        Return sRet
    End Function
    Function ������(ByVal sText As String) As String
	    Return �������`(sText, False)
    End Function
    Function �������Q(ByVal sText As String) As String
	    Return �������`(sText, True)
    End Function
    Function �������`(ByVal sText As String, ByVal bDummy As Boolean) As String
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

            '// CryptoStream�I�u�W�F�N�g���쐬����
            Dim transform As ICryptoTransform = rc2.CreateDecryptor(bKey, bIv) '// Decryptor���쐬����
            Dim csDecrypt As CryptoStream = New CryptoStream(inputStream, transform, CryptoStreamMode.Read)

            ReDim bRet(bText.Length - 1)
            '//Read the data out of the crypto stream.
            Dim iLen As Integer = csDecrypt.Read(bRet, 0, bRet.Length)

            '//Convert the byte array back into a string.
            sRet = Encoding.GetEncoding("shift-jis").GetString(bRet, 0, iLen)

            '// �X�g���[�������
            csDecrypt.Close()
            inputStream.Close()

            rc2 = Nothing
            csDecrypt = Nothing
            inputStream = Nothing

			If bDummy And sRet.Length >= 2 Then
				'//�擪����і����̃_�~�[�������폜
				sRet = sRet.Substring(1,sRet.Length-2)
			End If

        Catch ex As Exception
            sRet = ex.Message
        End Try

        Return sRet
    End Function
''// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
End Class
