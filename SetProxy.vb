Imports System.Windows.Forms

Public Class SetProxy
    Inherits System.Windows.Forms.Form

'// ADD 2008.10.29 ���s�j���� �v���L�V�Ή� START
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//	public sProxyAdr        as string  = ""
'//	public sProxyAdrUserSet as string  = ""
'//	public iProxyNo         as integer = 0
'//	public iProxyNoUserSet  as integer = 0
	public sProxyAdrUserSet  As String  = ""
	public iProxyNoUserSet   As Integer = 0
	public bProxyOnUserSet   As Boolean = True
	public bProxyIdOnUserSet As Boolean = False
	public sProxyIdUserSet   As String  = ""
	public sProxyPaUserSet   As String  = ""
	public iConnectTimeOut   As Integer = 0
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
'// ADD 2008.10.29 ���s�j���� �v���L�V�Ή� END

#Region " Windows �t�H�[�� �f�U�C�i�Ő������ꂽ�R�[�h "

    Public Sub New()
        MyBase.New()

        ' ���̌Ăяo���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        InitializeComponent()

        ' InitializeComponent() �Ăяo���̌�ɏ�������ǉ����܂��B

    End Sub

    ' Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    Private components As System.ComponentModel.IContainer

    ' ���� : �ȉ��̃v���V�[�W���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    'Windows �t�H�[�� �f�U�C�i���g���ĕύX���Ă��������B  
    ' �R�[�h �G�f�B�^���g���ĕύX���Ȃ��ł��������B
    Friend WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lab�ݒ�v���L�V�A�h���X As System.Windows.Forms.Label
    Friend WithEvents lab�ݒ�v���L�V�|�[�g�ԍ� As System.Windows.Forms.Label
    Friend WithEvents btn�ݒ� As System.Windows.Forms.Button
    Friend WithEvents btn���� As System.Windows.Forms.Button
    Friend WithEvents panel7 As System.Windows.Forms.Panel
    Friend WithEvents lab�v���L�V�ݒ�^�C�g�� As System.Windows.Forms.Label
    Private WithEvents tex�ݒ�v���L�V�A�h���X As System.Windows.Forms.TextBox
    Private WithEvents tex�ݒ�v���L�V�|�[�g�ԍ� As System.Windows.Forms.TextBox
    Private WithEvents tex�p�X���[ As System.Windows.Forms.TextBox
    Private WithEvents tex���[�U�[ As System.Windows.Forms.TextBox
    Friend WithEvents lab�p�X���[ As System.Windows.Forms.Label
    Friend WithEvents lab���[�U�[�h�c As System.Windows.Forms.Label
    Friend WithEvents cb�v���L�V�g�p As System.Windows.Forms.CheckBox
    Friend WithEvents cb���[�U�[�g�p As System.Windows.Forms.CheckBox
    Friend WithEvents lab�^�C���A�E�g As System.Windows.Forms.Label
    Friend WithEvents tex�^�C���A As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.lab�^�C���A�E�g = New System.Windows.Forms.Label
        Me.tex�^�C���A = New System.Windows.Forms.TextBox
        Me.cb���[�U�[�g�p = New System.Windows.Forms.CheckBox
        Me.cb�v���L�V�g�p = New System.Windows.Forms.CheckBox
        Me.tex�p�X���[ = New System.Windows.Forms.TextBox
        Me.tex���[�U�[ = New System.Windows.Forms.TextBox
        Me.tex�ݒ�v���L�V�|�[�g�ԍ� = New System.Windows.Forms.TextBox
        Me.tex�ݒ�v���L�V�A�h���X = New System.Windows.Forms.TextBox
        Me.lab���[�U�[�h�c = New System.Windows.Forms.Label
        Me.lab�p�X���[ = New System.Windows.Forms.Label
        Me.lab�ݒ�v���L�V�A�h���X = New System.Windows.Forms.Label
        Me.lab�ݒ�v���L�V�|�[�g�ԍ� = New System.Windows.Forms.Label
        Me.btn�ݒ� = New System.Windows.Forms.Button
        Me.btn���� = New System.Windows.Forms.Button
        Me.panel7 = New System.Windows.Forms.Panel
        Me.lab�v���L�V�ݒ�^�C�g�� = New System.Windows.Forms.Label
        Me.groupBox1.SuspendLayout()
        Me.panel7.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.lab�^�C���A�E�g)
        Me.groupBox1.Controls.Add(Me.tex�^�C���A)
        Me.groupBox1.Controls.Add(Me.cb���[�U�[�g�p)
        Me.groupBox1.Controls.Add(Me.cb�v���L�V�g�p)
        Me.groupBox1.Controls.Add(Me.tex�p�X���[)
        Me.groupBox1.Controls.Add(Me.tex���[�U�[)
        Me.groupBox1.Controls.Add(Me.tex�ݒ�v���L�V�|�[�g�ԍ�)
        Me.groupBox1.Controls.Add(Me.tex�ݒ�v���L�V�A�h���X)
        Me.groupBox1.Controls.Add(Me.lab���[�U�[�h�c)
        Me.groupBox1.Controls.Add(Me.lab�p�X���[)
        Me.groupBox1.Controls.Add(Me.lab�ݒ�v���L�V�A�h���X)
        Me.groupBox1.Controls.Add(Me.lab�ݒ�v���L�V�|�[�g�ԍ�)
        Me.groupBox1.Controls.Add(Me.btn�ݒ�)
        Me.groupBox1.Controls.Add(Me.btn����)
        Me.groupBox1.Location = New System.Drawing.Point(2, 28)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(384, 236)
        Me.groupBox1.TabIndex = 16
        Me.groupBox1.TabStop = False
        '
        'lab�^�C���A�E�g
        '
        Me.lab�^�C���A�E�g.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab�^�C���A�E�g.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab�^�C���A�E�g.Location = New System.Drawing.Point(236, 78)
        Me.lab�^�C���A�E�g.Name = "lab�^�C���A�E�g"
        Me.lab�^�C���A�E�g.Size = New System.Drawing.Size(84, 14)
        Me.lab�^�C���A�E�g.TabIndex = 58
        Me.lab�^�C���A�E�g.Text = "�^�C���A�E�g�i�b�j"
        '
        'tex�^�C���A
        '
        Me.tex�^�C���A.BackColor = System.Drawing.Color.Honeydew
        Me.tex�^�C���A.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tex�^�C���A.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex�^�C���A.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex�^�C���A.Location = New System.Drawing.Point(320, 78)
        Me.tex�^�C���A.MaxLength = 5
        Me.tex�^�C���A.Name = "tex�^�C���A"
        Me.tex�^�C���A.Size = New System.Drawing.Size(50, 16)
        Me.tex�^�C���A.TabIndex = 57
        Me.tex�^�C���A.TabStop = False
        Me.tex�^�C���A.Text = ""
        '
        'cb���[�U�[�g�p
        '
        Me.cb���[�U�[�g�p.ForeColor = System.Drawing.Color.LimeGreen
        Me.cb���[�U�[�g�p.Location = New System.Drawing.Point(14, 100)
        Me.cb���[�U�[�g�p.Name = "cb���[�U�[�g�p"
        Me.cb���[�U�[�g�p.Size = New System.Drawing.Size(238, 24)
        Me.cb���[�U�[�g�p.TabIndex = 3
        Me.cb���[�U�[�g�p.Text = "���[�U�[�h�c����уp�X���[�h���g�p����"
        '
        'cb�v���L�V�g�p
        '
        Me.cb�v���L�V�g�p.ForeColor = System.Drawing.Color.LimeGreen
        Me.cb�v���L�V�g�p.Location = New System.Drawing.Point(14, 18)
        Me.cb�v���L�V�g�p.Name = "cb�v���L�V�g�p"
        Me.cb�v���L�V�g�p.Size = New System.Drawing.Size(238, 24)
        Me.cb�v���L�V�g�p.TabIndex = 0
        Me.cb�v���L�V�g�p.Text = "�v���L�V�T�[�o���g�p����"
        '
        'tex�p�X���[
        '
        Me.tex�p�X���[.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex�p�X���[.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex�p�X���[.Location = New System.Drawing.Point(104, 152)
        Me.tex�p�X���[.MaxLength = 20
        Me.tex�p�X���[.Name = "tex�p�X���["
        Me.tex�p�X���[.Size = New System.Drawing.Size(150, 23)
        Me.tex�p�X���[.TabIndex = 5
        Me.tex�p�X���[.TabStop = False
        Me.tex�p�X���[.Text = ""
        '
        'tex���[�U�[
        '
        Me.tex���[�U�[.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex���[�U�[.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex���[�U�[.Location = New System.Drawing.Point(104, 124)
        Me.tex���[�U�[.MaxLength = 20
        Me.tex���[�U�[.Name = "tex���[�U�["
        Me.tex���[�U�[.Size = New System.Drawing.Size(150, 23)
        Me.tex���[�U�[.TabIndex = 4
        Me.tex���[�U�[.TabStop = False
        Me.tex���[�U�[.Text = ""
        '
        'tex�ݒ�v���L�V�|�[�g�ԍ�
        '
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.Location = New System.Drawing.Point(104, 74)
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.MaxLength = 5
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.Name = "tex�ݒ�v���L�V�|�[�g�ԍ�"
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.Size = New System.Drawing.Size(50, 23)
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.TabIndex = 2
        Me.tex�ݒ�v���L�V�|�[�g�ԍ�.Text = ""
        '
        'tex�ݒ�v���L�V�A�h���X
        '
        Me.tex�ݒ�v���L�V�A�h���X.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex�ݒ�v���L�V�A�h���X.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex�ݒ�v���L�V�A�h���X.Location = New System.Drawing.Point(104, 46)
        Me.tex�ݒ�v���L�V�A�h���X.MaxLength = 45
        Me.tex�ݒ�v���L�V�A�h���X.Name = "tex�ݒ�v���L�V�A�h���X"
        Me.tex�ݒ�v���L�V�A�h���X.Size = New System.Drawing.Size(268, 23)
        Me.tex�ݒ�v���L�V�A�h���X.TabIndex = 1
        Me.tex�ݒ�v���L�V�A�h���X.Text = ""
        '
        'lab���[�U�[�h�c
        '
        Me.lab���[�U�[�h�c.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab���[�U�[�h�c.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab���[�U�[�h�c.Location = New System.Drawing.Point(30, 128)
        Me.lab���[�U�[�h�c.Name = "lab���[�U�[�h�c"
        Me.lab���[�U�[�h�c.Size = New System.Drawing.Size(74, 14)
        Me.lab���[�U�[�h�c.TabIndex = 56
        Me.lab���[�U�[�h�c.Text = "���[�U�[�h�c"
        '
        'lab�p�X���[
        '
        Me.lab�p�X���[.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab�p�X���[.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab�p�X���[.Location = New System.Drawing.Point(30, 156)
        Me.lab�p�X���[.Name = "lab�p�X���["
        Me.lab�p�X���[.Size = New System.Drawing.Size(74, 14)
        Me.lab�p�X���[.TabIndex = 55
        Me.lab�p�X���[.Text = "�p�X���[�h"
        '
        'lab�ݒ�v���L�V�A�h���X
        '
        Me.lab�ݒ�v���L�V�A�h���X.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab�ݒ�v���L�V�A�h���X.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab�ݒ�v���L�V�A�h���X.Location = New System.Drawing.Point(20, 50)
        Me.lab�ݒ�v���L�V�A�h���X.Name = "lab�ݒ�v���L�V�A�h���X"
        Me.lab�ݒ�v���L�V�A�h���X.Size = New System.Drawing.Size(84, 14)
        Me.lab�ݒ�v���L�V�A�h���X.TabIndex = 52
        Me.lab�ݒ�v���L�V�A�h���X.Text = "�v���L�V�A�h���X"
        '
        'lab�ݒ�v���L�V�|�[�g�ԍ�
        '
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.Location = New System.Drawing.Point(20, 78)
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.Name = "lab�ݒ�v���L�V�|�[�g�ԍ�"
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.Size = New System.Drawing.Size(84, 14)
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.TabIndex = 51
        Me.lab�ݒ�v���L�V�|�[�g�ԍ�.Text = "�|�[�g�ԍ�"
        '
        'btn�ݒ�
        '
        Me.btn�ݒ�.BackColor = System.Drawing.Color.PaleGreen
        Me.btn�ݒ�.ForeColor = System.Drawing.Color.Blue
        Me.btn�ݒ�.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btn�ݒ�.Location = New System.Drawing.Point(168, 186)
        Me.btn�ݒ�.Name = "btn�ݒ�"
        Me.btn�ݒ�.Size = New System.Drawing.Size(96, 36)
        Me.btn�ݒ�.TabIndex = 6
        Me.btn�ݒ�.Text = "�ݒ�"
        '
        'btn����
        '
        Me.btn����.BackColor = System.Drawing.Color.PaleGreen
        Me.btn����.ForeColor = System.Drawing.Color.Red
        Me.btn����.Location = New System.Drawing.Point(274, 186)
        Me.btn����.Name = "btn����"
        Me.btn����.Size = New System.Drawing.Size(96, 36)
        Me.btn����.TabIndex = 7
        Me.btn����.TabStop = False
        Me.btn����.Text = "�L�����Z��"
        '
        'panel7
        '
        Me.panel7.BackColor = System.Drawing.Color.FromArgb(CType(44, Byte), CType(241, Byte), CType(83, Byte))
        Me.panel7.Controls.Add(Me.lab�v���L�V�ݒ�^�C�g��)
        Me.panel7.Location = New System.Drawing.Point(0, 0)
        Me.panel7.Name = "panel7"
        Me.panel7.Size = New System.Drawing.Size(394, 26)
        Me.panel7.TabIndex = 15
        '
        'lab�v���L�V�ݒ�^�C�g��
        '
        Me.lab�v���L�V�ݒ�^�C�g��.BackColor = System.Drawing.Color.FromArgb(CType(44, Byte), CType(241, Byte), CType(83, Byte))
        Me.lab�v���L�V�ݒ�^�C�g��.Font = New System.Drawing.Font("MS UI Gothic", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab�v���L�V�ݒ�^�C�g��.ForeColor = System.Drawing.Color.White
        Me.lab�v���L�V�ݒ�^�C�g��.Location = New System.Drawing.Point(12, 2)
        Me.lab�v���L�V�ݒ�^�C�g��.Name = "lab�v���L�V�ݒ�^�C�g��"
        Me.lab�v���L�V�ݒ�^�C�g��.Size = New System.Drawing.Size(264, 24)
        Me.lab�v���L�V�ݒ�^�C�g��.TabIndex = 0
        Me.lab�v���L�V�ݒ�^�C�g��.Text = "�v���L�V�ݒ�"
        '
        'SetProxy
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.BackColor = System.Drawing.Color.Honeydew
        Me.ClientSize = New System.Drawing.Size(388, 266)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.panel7)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(394, 298)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(394, 298)
        Me.Name = "SetProxy"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "is-2 �v���L�V�ݒ�"
        Me.groupBox1.ResumeLayout(False)
        Me.panel7.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	Private Sub �G���^�[�ړ�(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		select case(e.KeyCode)
			'// [Enter]�L�[�������ꂽ���A���̃R���g���[���փt�H�[�J�X�ړ�
			case Keys.Enter:
				Me.SelectNextControl(Me.ActiveControl, true, true, true, true)
			'// [Esc]�L�[�������ꂽ���A�t�H�[�������
			case Keys.Escape:
				Close()
		end select
    End Sub
	Private Sub �G���^�[�L�����Z��(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)  Handles MyBase.KeyPress
    	if e.KeyChar = Chr(13) then
			e.Handled = true
		end if
    End Sub

    Private Sub btn����_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn����.Click
        Me.DialogResult = Windows.Forms.DialogResult.No 
		Me.Close()
    End Sub

    Private Sub btn�ݒ�_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn�ݒ�.Click
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
        �R���g���[������()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
		'//�g����
		tex�ݒ�v���L�V�A�h���X.Text   = tex�ݒ�v���L�V�A�h���X.Text.Trim()
		tex�ݒ�v���L�V�|�[�g�ԍ�.Text = tex�ݒ�v���L�V�|�[�g�ԍ�.Text.Trim()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
		tex���[�U�[.Text               = tex���[�U�[.Text.Trim()
		tex�p�X���[.Text               = tex�p�X���[.Text.Trim()
		tex�^�C���A.Text               = tex�^�C���A.Text.Trim()
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END

'//		if Not �K�{�`�F�b�N(tex�ݒ�v���L�V�A�h���X,"�v���L�V�A�h���X") then return

		if Not ���p�`�F�b�N(tex�ݒ�v���L�V�A�h���X,"�v���L�V�A�h���X") then return
		if Not ���p�`�F�b�N(tex�ݒ�v���L�V�|�[�g�ԍ�,"�|�[�g�ԍ�") then return
		if tex�ݒ�v���L�V�|�[�g�ԍ�.Text.Length > 0 then
			if Not ���l�`�F�b�N(tex�ݒ�v���L�V�|�[�g�ԍ�,"�|�[�g�ԍ�") then return
        end if
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
		if Not ���p�`�F�b�N(tex���[�U�[,"���[�U�[") then return
		if Not ���p�`�F�b�N(tex�p�X���[,"�p�X���[�h") then return
		if Not ���p�`�F�b�N(tex�^�C���A,"�^�C���A�E�g") then return
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END

		sProxyAdrUserSet = tex�ݒ�v���L�V�A�h���X.Text
        if tex�ݒ�v���L�V�|�[�g�ԍ�.Text.Length = 0 then
            iProxyNoUserSet = 0
        else
            iProxyNoUserSet = integer.Parse(tex�ݒ�v���L�V�|�[�g�ԍ�.Text)
        end if
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
        bProxyOnUserSet   = cb�v���L�V�g�p.Checked
        bProxyIdOnUserSet = cb���[�U�[�g�p.Checked
		sProxyIdUserSet   = tex���[�U�[.Text
		sProxyPaUserSet   = tex�p�X���[.Text
        if tex�^�C���A.Text.Length = 0 then
            iConnectTimeOut = 100
        else
            iConnectTimeOut = integer.Parse(tex�^�C���A.Text)
        end if
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END

        Me.DialogResult = Windows.Forms.DialogResult.Yes
		Me.Close()
    End Sub

    Private Sub SetProxy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tex�ݒ�v���L�V�A�h���X.Text = sProxyAdrUserSet
        if iProxyNoUserSet = 0 then
		    tex�ݒ�v���L�V�|�[�g�ԍ�.Text = ""
        else 
    		tex�ݒ�v���L�V�|�[�g�ԍ�.Text = iProxyNoUserSet.ToString()
        end if
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//		tex���s�v���L�V�A�h���X.Text   = sProxyAdr
'//     if iProxyNo = 0 then
'//		    tex���s�v���L�V�|�[�g�ԍ�.Text = ""
'//     else 
'//		    tex���s�v���L�V�|�[�g�ԍ�.Text = iProxyNo.ToString()
'//     end if
        cb�v���L�V�g�p.Checked = bProxyOnUserSet
        cb���[�U�[�g�p.Checked = bProxyIdOnUserSet
        tex���[�U�[.Text = sProxyIdUserSet
        tex�p�X���[.Text = sProxyPaUserSet
        if iConnectTimeOut = 0 then
		    tex�^�C���A.Text = ""
        else 
    		tex�^�C���A.Text = iConnectTimeOut.ToString()
        end if
        �R���g���[������()
    End Sub
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
    Private Sub �R���g���[������()
        If cb�v���L�V�g�p.Checked Then
            tex�ݒ�v���L�V�A�h���X.Enabled   = True
            tex�ݒ�v���L�V�|�[�g�ԍ�.Enabled = True
            cb���[�U�[�g�p.Enabled = True
            If cb���[�U�[�g�p.Checked Then
                tex���[�U�[.Enabled = True
                tex�p�X���[.Enabled = True
            Else
                tex���[�U�[.Enabled = False
                tex�p�X���[.Enabled = False
            End If
        Else
            tex�ݒ�v���L�V�A�h���X.Enabled   = False
            tex�ݒ�v���L�V�|�[�g�ԍ�.Enabled = False
            cb���[�U�[�g�p.Enabled = False
            tex���[�U�[.Enabled = False
            tex�p�X���[.Enabled = False
        End If
    End Sub
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END

    Private Sub SetProxy_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		tex�ݒ�v���L�V�A�h���X.Text   = ""
		tex�ݒ�v���L�V�|�[�g�ԍ�.Text = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� START
'//		tex���s�v���L�V�A�h���X.Text   = ""
'//		tex���s�v���L�V�|�[�g�ԍ�.Text = ""
        cb�v���L�V�g�p.Checked         = False
        cb���[�U�[�g�p.Checked         = False
		tex���[�U�[.Text               = ""
		tex�p�X���[.Text               = ""
		tex�^�C���A.Text               = ""
'// MOD 2010.06.21 ���s�j���� �v���L�V�F�؂̒ǉ� END
		tex�ݒ�v���L�V�A�h���X.Focus()
    End Sub

	private function �K�{�`�F�b�N(tex as TextBox, name as string) as boolean
		if tex.Text.Length > 0 Then return true
		MessageBox.Show("�K�{����(" + name + ")�����͂���Ă��܂���", _
			"���̓`�F�b�N",MessageBoxButtons.OK)
		tex.Focus()
		return false
	end function

	private function �r�i�h�r�`�F�b�N(tex as TextBox, name as string, byref sUnicode as string, byref bSjis as byte()) as boolean
		'//�t�ϊ����Ăr�i�h�r�������`�F�b�N����
		dim sRevUnicode as string = System.Text.Encoding.GetEncoding("shift-jis").GetString(bSjis)
		dim sErrChars as string = ""
		for iPos as integer = 0 To sUnicode.Length - 1
            if iPos >= sRevUnicode.Length then exit for
			if sUnicode.Chars(iPos) <> sRevUnicode.Chars(iPos) then
					sErrChars += sUnicode.Chars(iPos)
			end if
		next
		if sErrChars.Length > 0 then
			MessageBox.Show(name + "�Ɏg�p�ł��Ȃ�����������܂�\n" _
				+ "�w" + sErrChars + "�x", _
				"���̓`�F�b�N",MessageBoxButtons.OK)
			tex.Focus()
			return false
		end if
		return true
	end function

	protected function ���p�`�F�b�N(tex as TextBox, name as string) as boolean 
		dim sUnicode as string = tex.Text
		dim bSjis as byte() = System.Text.Encoding.GetEncoding("shift-jis").GetBytes(sUnicode)
		if Not �r�i�h�r�`�F�b�N(tex, name, sUnicode, bSjis) Then return false
		if bSjis.Length <> sUnicode.Length Then
			MessageBox.Show(name + "�͔��p�����œ��͂��Ă�������", _
				"���̓`�F�b�N",MessageBoxButtons.OK)
			tex.Focus()
			return false
		End If

'// MOD 2011.09.22 ���s�j���� �L���`�F�b�N�p�~ START
'//		for i as integer = 0 To tex.Text.Length - 1
'//			'// [!"#$%&'()*+,]
'//			'// [:;<=>?]
'//			'// [[]^]
'//			'// [{|}]
'//			'// [\]
'//			if (tex.Text.Chars(i) >= "!" and tex.Text.Chars(i) <= ",") _
'//				Or (tex.Text.Chars(i) >= ":" and tex.Text.Chars(i) <= "?") _
'//				Or (tex.Text.Chars(i) >= "[" and tex.Text.Chars(i) <= "^") _
'//				Or (tex.Text.Chars(i) >= "{" and tex.Text.Chars(i) <= "}") _
'//				Or (tex.Text.Chars(i) = "\") Then
'//				MessageBox.Show(name + "�ɋL�������͂���Ă��܂�","���̓`�F�b�N",MessageBoxButtons.OK)
'//				tex.Focus()
'//				return false
'//			end if
'//		next
'// MOD 2011.09.22 ���s�j���� �L���`�F�b�N�p�~ END
		return true
	end function
	protected function ���l�`�F�b�N(tex as TextBox, name as string) as boolean
		try
			dim lChk as long = long.Parse(tex.Text.Replace(",",""))
			return true
		catch ex as Exception
			MessageBox.Show(name + "�ɐ��l�����͂���Ă��܂���","���̓`�F�b�N",MessageBoxButtons.OK)
			tex.Focus()
			
			return false
		end try
	end function

    Private Sub cb�v���L�V�g�p_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb�v���L�V�g�p.CheckedChanged
        �R���g���[������()
    End Sub

    Private Sub cb���[�U�[�g�p_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb���[�U�[�g�p.CheckedChanged
        �R���g���[������()
    End Sub
End Class
