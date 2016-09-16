	'/// <summary>
	'/// ���O�������ݗp�N���X
	'/// </summary>
	'//--------------------------------------------------------------------------
	'// �C������
	'//--------------------------------------------------------------------------
	'// MOD 2009.10.03 ���s�j���� ���O�o�͕����̏ȗ��i�t�@�C���`�F�b�N�j 
	'//--------------------------------------------------------------------------
'
'
Imports System.Reflection
Imports System.IO
Imports System.Text

Public Class LogWriter

    Protected Friend Structure LogRecord
        Public Status As String
        Public SubStatus As String
        Public Target As String
        Public Result As String
        Public Remark As String
    End Structure

    Private Const constLogDir As String = "Log"
    Private _LogFileName As String
    Private _LogFullDirName As String
    Private _LogFullPathFileName As String

    '
    '   �R���X�g���N�^
    '
    Private Sub New()
        MyBase.New()

        Dim myAssemblyName As String

        myAssemblyName = [Assembly].GetExecutingAssembly.FullName()
        _LogFileName = myAssemblyName.Substring(0, myAssemblyName.IndexOf(",")) & ".log"
        _LogFullDirName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, constLogDir)

        _LogFullPathFileName = Path.Combine(_LogFullDirName, _LogFileName)
    End Sub

    '
    '   �C���X�^���X�쐬���\�b�h
    '       ���������̂��߂ɃV���O���g�����̗p���Ă��܂�
    Protected Friend Shared Function CreateInstance() As LogWriter
        Static Dim myLogWriter As LogWriter

        If myLogWriter Is Nothing Then
            myLogWriter = New LogWriter()
        End If

        Return myLogWriter

    End Function

    '
    '   ���O�������݃��\�b�h
    '
    Protected Friend Function WriteLog(ByVal logRecord As LogRecord) As Boolean
        Dim blnRet As Boolean
        Dim myFileWriter As StreamWriter
        Dim myLogString As Text.StringBuilder = New Text.StringBuilder()
        blnRet = CheckDir()
        If blnRet = False Then Return False

        If File.Exists(_LogFullPathFileName) Then
            myFileWriter = File.AppendText(_LogFullPathFileName)
        Else
            myFileWriter = File.CreateText(_LogFullPathFileName)
        End If

        myLogString.Append(DateTime.Now.ToString())
        myLogString.Append(vbTab)
'// MOD 2009.10.03 ���s�j���� ���O�o�͕����̏ȗ��i�t�@�C���`�F�b�N�j START
'//     myLogString.Append(logRecord.Status)
'//     myLogString.Append(vbTab)
'//     myLogString.Append(logRecord.SubStatus)
'//     myLogString.Append(vbTab)
'// MOD 2009.10.03 ���s�j���� ���O�o�͕����̏ȗ��i�t�@�C���`�F�b�N�j END
        myLogString.Append(logRecord.Target)
        myLogString.Append(vbTab)
        myLogString.Append(logRecord.Result)
        myLogString.Append(vbTab)
        myLogString.Append(logRecord.Remark)
'// MOD 2009.10.03 ���s�j���� ���O�o�͕����̏ȗ��i�t�@�C���`�F�b�N�j START
'//     myLogString.Append(vbTab)
'//     myLogString.Append(AppDomain.GetCurrentThreadId().ToString())
'// MOD 2009.10.03 ���s�j���� ���O�o�͕����̏ȗ��i�t�@�C���`�F�b�N�j END

        myFileWriter.WriteLine(myLogString.ToString())

        myFileWriter.Flush()
        myFileWriter.Close()

	Return True

    End Function

    '
    '   ���O�t�@�C�����N���A�[
    '
    Protected Friend Function Clear() As Boolean
        Dim blnRet As Boolean
        Dim myFileStream As FileStream
        blnRet = CheckDir()
        If blnRet = False Then Return False

        myFileStream = New FileStream(_LogFullPathFileName, FileMode.Create)
        myFileStream.Close()

        Return True
    End Function

    '   ���O�f�B���N�g���̃`�F�b�N
    Private Function CheckDir() As Boolean
        Dim myFileCopy As New FileCopy()

        Return myFileCopy.CheckOrMakeDestinationDir(_LogFullDirName)
    End Function

    '
    '   ���b�Z�[�W�z������������ɂ��Ė߂�
    '
    Protected Friend Function AppendStrings(ByVal messages() As String) As String
        Dim myStringBuilder As StringBuilder
        Dim i As Long

        myStringBuilder = New StringBuilder(messages(0))
        For i = 1 To messages.Length - 1
            myStringBuilder.Append(messages(i))
        Next

        Return myStringBuilder.ToString()

    End Function
End Class
