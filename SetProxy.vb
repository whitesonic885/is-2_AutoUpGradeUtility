Imports System.Windows.Forms

Public Class SetProxy
    Inherits System.Windows.Forms.Form

'// ADD 2008.10.29 東都）高木 プロキシ対応 START
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
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
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
'// ADD 2008.10.29 東都）高木 プロキシ対応 END

#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更してください。  
    ' コード エディタを使って変更しないでください。
    Friend WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lab設定プロキシアドレス As System.Windows.Forms.Label
    Friend WithEvents lab設定プロキシポート番号 As System.Windows.Forms.Label
    Friend WithEvents btn設定 As System.Windows.Forms.Button
    Friend WithEvents btn閉じる As System.Windows.Forms.Button
    Friend WithEvents panel7 As System.Windows.Forms.Panel
    Friend WithEvents labプロキシ設定タイトル As System.Windows.Forms.Label
    Private WithEvents tex設定プロキシアドレス As System.Windows.Forms.TextBox
    Private WithEvents tex設定プロキシポート番号 As System.Windows.Forms.TextBox
    Private WithEvents texパスワー As System.Windows.Forms.TextBox
    Private WithEvents texユーザー As System.Windows.Forms.TextBox
    Friend WithEvents labパスワー As System.Windows.Forms.Label
    Friend WithEvents labユーザーＩＤ As System.Windows.Forms.Label
    Friend WithEvents cbプロキシ使用 As System.Windows.Forms.CheckBox
    Friend WithEvents cbユーザー使用 As System.Windows.Forms.CheckBox
    Friend WithEvents labタイムアウト As System.Windows.Forms.Label
    Friend WithEvents texタイムア As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.labタイムアウト = New System.Windows.Forms.Label
        Me.texタイムア = New System.Windows.Forms.TextBox
        Me.cbユーザー使用 = New System.Windows.Forms.CheckBox
        Me.cbプロキシ使用 = New System.Windows.Forms.CheckBox
        Me.texパスワー = New System.Windows.Forms.TextBox
        Me.texユーザー = New System.Windows.Forms.TextBox
        Me.tex設定プロキシポート番号 = New System.Windows.Forms.TextBox
        Me.tex設定プロキシアドレス = New System.Windows.Forms.TextBox
        Me.labユーザーＩＤ = New System.Windows.Forms.Label
        Me.labパスワー = New System.Windows.Forms.Label
        Me.lab設定プロキシアドレス = New System.Windows.Forms.Label
        Me.lab設定プロキシポート番号 = New System.Windows.Forms.Label
        Me.btn設定 = New System.Windows.Forms.Button
        Me.btn閉じる = New System.Windows.Forms.Button
        Me.panel7 = New System.Windows.Forms.Panel
        Me.labプロキシ設定タイトル = New System.Windows.Forms.Label
        Me.groupBox1.SuspendLayout()
        Me.panel7.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.labタイムアウト)
        Me.groupBox1.Controls.Add(Me.texタイムア)
        Me.groupBox1.Controls.Add(Me.cbユーザー使用)
        Me.groupBox1.Controls.Add(Me.cbプロキシ使用)
        Me.groupBox1.Controls.Add(Me.texパスワー)
        Me.groupBox1.Controls.Add(Me.texユーザー)
        Me.groupBox1.Controls.Add(Me.tex設定プロキシポート番号)
        Me.groupBox1.Controls.Add(Me.tex設定プロキシアドレス)
        Me.groupBox1.Controls.Add(Me.labユーザーＩＤ)
        Me.groupBox1.Controls.Add(Me.labパスワー)
        Me.groupBox1.Controls.Add(Me.lab設定プロキシアドレス)
        Me.groupBox1.Controls.Add(Me.lab設定プロキシポート番号)
        Me.groupBox1.Controls.Add(Me.btn設定)
        Me.groupBox1.Controls.Add(Me.btn閉じる)
        Me.groupBox1.Location = New System.Drawing.Point(2, 28)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(384, 236)
        Me.groupBox1.TabIndex = 16
        Me.groupBox1.TabStop = False
        '
        'labタイムアウト
        '
        Me.labタイムアウト.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labタイムアウト.ForeColor = System.Drawing.Color.LimeGreen
        Me.labタイムアウト.Location = New System.Drawing.Point(236, 78)
        Me.labタイムアウト.Name = "labタイムアウト"
        Me.labタイムアウト.Size = New System.Drawing.Size(84, 14)
        Me.labタイムアウト.TabIndex = 58
        Me.labタイムアウト.Text = "タイムアウト（秒）"
        '
        'texタイムア
        '
        Me.texタイムア.BackColor = System.Drawing.Color.Honeydew
        Me.texタイムア.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.texタイムア.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.texタイムア.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.texタイムア.Location = New System.Drawing.Point(320, 78)
        Me.texタイムア.MaxLength = 5
        Me.texタイムア.Name = "texタイムア"
        Me.texタイムア.Size = New System.Drawing.Size(50, 16)
        Me.texタイムア.TabIndex = 57
        Me.texタイムア.TabStop = False
        Me.texタイムア.Text = ""
        '
        'cbユーザー使用
        '
        Me.cbユーザー使用.ForeColor = System.Drawing.Color.LimeGreen
        Me.cbユーザー使用.Location = New System.Drawing.Point(14, 100)
        Me.cbユーザー使用.Name = "cbユーザー使用"
        Me.cbユーザー使用.Size = New System.Drawing.Size(238, 24)
        Me.cbユーザー使用.TabIndex = 3
        Me.cbユーザー使用.Text = "ユーザーＩＤおよびパスワードを使用する"
        '
        'cbプロキシ使用
        '
        Me.cbプロキシ使用.ForeColor = System.Drawing.Color.LimeGreen
        Me.cbプロキシ使用.Location = New System.Drawing.Point(14, 18)
        Me.cbプロキシ使用.Name = "cbプロキシ使用"
        Me.cbプロキシ使用.Size = New System.Drawing.Size(238, 24)
        Me.cbプロキシ使用.TabIndex = 0
        Me.cbプロキシ使用.Text = "プロキシサーバを使用する"
        '
        'texパスワー
        '
        Me.texパスワー.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.texパスワー.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.texパスワー.Location = New System.Drawing.Point(104, 152)
        Me.texパスワー.MaxLength = 20
        Me.texパスワー.Name = "texパスワー"
        Me.texパスワー.Size = New System.Drawing.Size(150, 23)
        Me.texパスワー.TabIndex = 5
        Me.texパスワー.TabStop = False
        Me.texパスワー.Text = ""
        '
        'texユーザー
        '
        Me.texユーザー.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.texユーザー.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.texユーザー.Location = New System.Drawing.Point(104, 124)
        Me.texユーザー.MaxLength = 20
        Me.texユーザー.Name = "texユーザー"
        Me.texユーザー.Size = New System.Drawing.Size(150, 23)
        Me.texユーザー.TabIndex = 4
        Me.texユーザー.TabStop = False
        Me.texユーザー.Text = ""
        '
        'tex設定プロキシポート番号
        '
        Me.tex設定プロキシポート番号.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex設定プロキシポート番号.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex設定プロキシポート番号.Location = New System.Drawing.Point(104, 74)
        Me.tex設定プロキシポート番号.MaxLength = 5
        Me.tex設定プロキシポート番号.Name = "tex設定プロキシポート番号"
        Me.tex設定プロキシポート番号.Size = New System.Drawing.Size(50, 23)
        Me.tex設定プロキシポート番号.TabIndex = 2
        Me.tex設定プロキシポート番号.Text = ""
        '
        'tex設定プロキシアドレス
        '
        Me.tex設定プロキシアドレス.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.tex設定プロキシアドレス.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.tex設定プロキシアドレス.Location = New System.Drawing.Point(104, 46)
        Me.tex設定プロキシアドレス.MaxLength = 45
        Me.tex設定プロキシアドレス.Name = "tex設定プロキシアドレス"
        Me.tex設定プロキシアドレス.Size = New System.Drawing.Size(268, 23)
        Me.tex設定プロキシアドレス.TabIndex = 1
        Me.tex設定プロキシアドレス.Text = ""
        '
        'labユーザーＩＤ
        '
        Me.labユーザーＩＤ.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labユーザーＩＤ.ForeColor = System.Drawing.Color.LimeGreen
        Me.labユーザーＩＤ.Location = New System.Drawing.Point(30, 128)
        Me.labユーザーＩＤ.Name = "labユーザーＩＤ"
        Me.labユーザーＩＤ.Size = New System.Drawing.Size(74, 14)
        Me.labユーザーＩＤ.TabIndex = 56
        Me.labユーザーＩＤ.Text = "ユーザーＩＤ"
        '
        'labパスワー
        '
        Me.labパスワー.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labパスワー.ForeColor = System.Drawing.Color.LimeGreen
        Me.labパスワー.Location = New System.Drawing.Point(30, 156)
        Me.labパスワー.Name = "labパスワー"
        Me.labパスワー.Size = New System.Drawing.Size(74, 14)
        Me.labパスワー.TabIndex = 55
        Me.labパスワー.Text = "パスワード"
        '
        'lab設定プロキシアドレス
        '
        Me.lab設定プロキシアドレス.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab設定プロキシアドレス.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab設定プロキシアドレス.Location = New System.Drawing.Point(20, 50)
        Me.lab設定プロキシアドレス.Name = "lab設定プロキシアドレス"
        Me.lab設定プロキシアドレス.Size = New System.Drawing.Size(84, 14)
        Me.lab設定プロキシアドレス.TabIndex = 52
        Me.lab設定プロキシアドレス.Text = "プロキシアドレス"
        '
        'lab設定プロキシポート番号
        '
        Me.lab設定プロキシポート番号.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lab設定プロキシポート番号.ForeColor = System.Drawing.Color.LimeGreen
        Me.lab設定プロキシポート番号.Location = New System.Drawing.Point(20, 78)
        Me.lab設定プロキシポート番号.Name = "lab設定プロキシポート番号"
        Me.lab設定プロキシポート番号.Size = New System.Drawing.Size(84, 14)
        Me.lab設定プロキシポート番号.TabIndex = 51
        Me.lab設定プロキシポート番号.Text = "ポート番号"
        '
        'btn設定
        '
        Me.btn設定.BackColor = System.Drawing.Color.PaleGreen
        Me.btn設定.ForeColor = System.Drawing.Color.Blue
        Me.btn設定.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btn設定.Location = New System.Drawing.Point(168, 186)
        Me.btn設定.Name = "btn設定"
        Me.btn設定.Size = New System.Drawing.Size(96, 36)
        Me.btn設定.TabIndex = 6
        Me.btn設定.Text = "設定"
        '
        'btn閉じる
        '
        Me.btn閉じる.BackColor = System.Drawing.Color.PaleGreen
        Me.btn閉じる.ForeColor = System.Drawing.Color.Red
        Me.btn閉じる.Location = New System.Drawing.Point(274, 186)
        Me.btn閉じる.Name = "btn閉じる"
        Me.btn閉じる.Size = New System.Drawing.Size(96, 36)
        Me.btn閉じる.TabIndex = 7
        Me.btn閉じる.TabStop = False
        Me.btn閉じる.Text = "キャンセル"
        '
        'panel7
        '
        Me.panel7.BackColor = System.Drawing.Color.FromArgb(CType(44, Byte), CType(241, Byte), CType(83, Byte))
        Me.panel7.Controls.Add(Me.labプロキシ設定タイトル)
        Me.panel7.Location = New System.Drawing.Point(0, 0)
        Me.panel7.Name = "panel7"
        Me.panel7.Size = New System.Drawing.Size(394, 26)
        Me.panel7.TabIndex = 15
        '
        'labプロキシ設定タイトル
        '
        Me.labプロキシ設定タイトル.BackColor = System.Drawing.Color.FromArgb(CType(44, Byte), CType(241, Byte), CType(83, Byte))
        Me.labプロキシ設定タイトル.Font = New System.Drawing.Font("MS UI Gothic", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labプロキシ設定タイトル.ForeColor = System.Drawing.Color.White
        Me.labプロキシ設定タイトル.Location = New System.Drawing.Point(12, 2)
        Me.labプロキシ設定タイトル.Name = "labプロキシ設定タイトル"
        Me.labプロキシ設定タイトル.Size = New System.Drawing.Size(264, 24)
        Me.labプロキシ設定タイトル.TabIndex = 0
        Me.labプロキシ設定タイトル.Text = "プロキシ設定"
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
        Me.Text = "is-2 プロキシ設定"
        Me.groupBox1.ResumeLayout(False)
        Me.panel7.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	Private Sub エンター移動(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		select case(e.KeyCode)
			'// [Enter]キーが押された時、次のコントロールへフォーカス移動
			case Keys.Enter:
				Me.SelectNextControl(Me.ActiveControl, true, true, true, true)
			'// [Esc]キーが押された時、フォームを閉じる
			case Keys.Escape:
				Close()
		end select
    End Sub
	Private Sub エンターキャンセル(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)  Handles MyBase.KeyPress
    	if e.KeyChar = Chr(13) then
			e.Handled = true
		end if
    End Sub

    Private Sub btn閉じる_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn閉じる.Click
        Me.DialogResult = Windows.Forms.DialogResult.No 
		Me.Close()
    End Sub

    Private Sub btn設定_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn設定.Click
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
        コントロール制御()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
		'//トリム
		tex設定プロキシアドレス.Text   = tex設定プロキシアドレス.Text.Trim()
		tex設定プロキシポート番号.Text = tex設定プロキシポート番号.Text.Trim()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
		texユーザー.Text               = texユーザー.Text.Trim()
		texパスワー.Text               = texパスワー.Text.Trim()
		texタイムア.Text               = texタイムア.Text.Trim()
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END

'//		if Not 必須チェック(tex設定プロキシアドレス,"プロキシアドレス") then return

		if Not 半角チェック(tex設定プロキシアドレス,"プロキシアドレス") then return
		if Not 半角チェック(tex設定プロキシポート番号,"ポート番号") then return
		if tex設定プロキシポート番号.Text.Length > 0 then
			if Not 数値チェック(tex設定プロキシポート番号,"ポート番号") then return
        end if
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
		if Not 半角チェック(texユーザー,"ユーザー") then return
		if Not 半角チェック(texパスワー,"パスワード") then return
		if Not 半角チェック(texタイムア,"タイムアウト") then return
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END

		sProxyAdrUserSet = tex設定プロキシアドレス.Text
        if tex設定プロキシポート番号.Text.Length = 0 then
            iProxyNoUserSet = 0
        else
            iProxyNoUserSet = integer.Parse(tex設定プロキシポート番号.Text)
        end if
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
        bProxyOnUserSet   = cbプロキシ使用.Checked
        bProxyIdOnUserSet = cbユーザー使用.Checked
		sProxyIdUserSet   = texユーザー.Text
		sProxyPaUserSet   = texパスワー.Text
        if texタイムア.Text.Length = 0 then
            iConnectTimeOut = 100
        else
            iConnectTimeOut = integer.Parse(texタイムア.Text)
        end if
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END

        Me.DialogResult = Windows.Forms.DialogResult.Yes
		Me.Close()
    End Sub

    Private Sub SetProxy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tex設定プロキシアドレス.Text = sProxyAdrUserSet
        if iProxyNoUserSet = 0 then
		    tex設定プロキシポート番号.Text = ""
        else 
    		tex設定プロキシポート番号.Text = iProxyNoUserSet.ToString()
        end if
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//		tex現行プロキシアドレス.Text   = sProxyAdr
'//     if iProxyNo = 0 then
'//		    tex現行プロキシポート番号.Text = ""
'//     else 
'//		    tex現行プロキシポート番号.Text = iProxyNo.ToString()
'//     end if
        cbプロキシ使用.Checked = bProxyOnUserSet
        cbユーザー使用.Checked = bProxyIdOnUserSet
        texユーザー.Text = sProxyIdUserSet
        texパスワー.Text = sProxyPaUserSet
        if iConnectTimeOut = 0 then
		    texタイムア.Text = ""
        else 
    		texタイムア.Text = iConnectTimeOut.ToString()
        end if
        コントロール制御()
    End Sub
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
    Private Sub コントロール制御()
        If cbプロキシ使用.Checked Then
            tex設定プロキシアドレス.Enabled   = True
            tex設定プロキシポート番号.Enabled = True
            cbユーザー使用.Enabled = True
            If cbユーザー使用.Checked Then
                texユーザー.Enabled = True
                texパスワー.Enabled = True
            Else
                texユーザー.Enabled = False
                texパスワー.Enabled = False
            End If
        Else
            tex設定プロキシアドレス.Enabled   = False
            tex設定プロキシポート番号.Enabled = False
            cbユーザー使用.Enabled = False
            texユーザー.Enabled = False
            texパスワー.Enabled = False
        End If
    End Sub
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END

    Private Sub SetProxy_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		tex設定プロキシアドレス.Text   = ""
		tex設定プロキシポート番号.Text = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 START
'//		tex現行プロキシアドレス.Text   = ""
'//		tex現行プロキシポート番号.Text = ""
        cbプロキシ使用.Checked         = False
        cbユーザー使用.Checked         = False
		texユーザー.Text               = ""
		texパスワー.Text               = ""
		texタイムア.Text               = ""
'// MOD 2010.06.21 東都）高木 プロキシ認証の追加 END
		tex設定プロキシアドレス.Focus()
    End Sub

	private function 必須チェック(tex as TextBox, name as string) as boolean
		if tex.Text.Length > 0 Then return true
		MessageBox.Show("必須項目(" + name + ")が入力されていません", _
			"入力チェック",MessageBoxButtons.OK)
		tex.Focus()
		return false
	end function

	private function ＳＪＩＳチェック(tex as TextBox, name as string, byref sUnicode as string, byref bSjis as byte()) as boolean
		'//逆変換してＳＪＩＳ文字をチェックする
		dim sRevUnicode as string = System.Text.Encoding.GetEncoding("shift-jis").GetString(bSjis)
		dim sErrChars as string = ""
		for iPos as integer = 0 To sUnicode.Length - 1
            if iPos >= sRevUnicode.Length then exit for
			if sUnicode.Chars(iPos) <> sRevUnicode.Chars(iPos) then
					sErrChars += sUnicode.Chars(iPos)
			end if
		next
		if sErrChars.Length > 0 then
			MessageBox.Show(name + "に使用できない文字があります\n" _
				+ "『" + sErrChars + "』", _
				"入力チェック",MessageBoxButtons.OK)
			tex.Focus()
			return false
		end if
		return true
	end function

	protected function 半角チェック(tex as TextBox, name as string) as boolean 
		dim sUnicode as string = tex.Text
		dim bSjis as byte() = System.Text.Encoding.GetEncoding("shift-jis").GetBytes(sUnicode)
		if Not ＳＪＩＳチェック(tex, name, sUnicode, bSjis) Then return false
		if bSjis.Length <> sUnicode.Length Then
			MessageBox.Show(name + "は半角文字で入力してください", _
				"入力チェック",MessageBoxButtons.OK)
			tex.Focus()
			return false
		End If

'// MOD 2011.09.22 東都）高木 記号チェック廃止 START
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
'//				MessageBox.Show(name + "に記号が入力されています","入力チェック",MessageBoxButtons.OK)
'//				tex.Focus()
'//				return false
'//			end if
'//		next
'// MOD 2011.09.22 東都）高木 記号チェック廃止 END
		return true
	end function
	protected function 数値チェック(tex as TextBox, name as string) as boolean
		try
			dim lChk as long = long.Parse(tex.Text.Replace(",",""))
			return true
		catch ex as Exception
			MessageBox.Show(name + "に数値が入力されていません","入力チェック",MessageBoxButtons.OK)
			tex.Focus()
			
			return false
		end try
	end function

    Private Sub cbプロキシ使用_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbプロキシ使用.CheckedChanged
        コントロール制御()
    End Sub

    Private Sub cbユーザー使用_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbユーザー使用.CheckedChanged
        コントロール制御()
    End Sub
End Class
