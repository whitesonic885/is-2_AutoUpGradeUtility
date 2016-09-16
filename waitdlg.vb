Public Class WaitDialog
  Inherits System.Windows.Forms.Form

  Private bAborting As Boolean = False   ' 中止フラグ
  Private bShowing As Boolean = False   ' ダイアログ表示中フラグ

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
  Public WithEvents progBarMeter As System.Windows.Forms.ProgressBar
  Public WithEvents labelProgress As System.Windows.Forms.Label
    Public WithEvents labelMainMsg As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.progBarMeter = New System.Windows.Forms.ProgressBar
        Me.labelProgress = New System.Windows.Forms.Label
        Me.labelMainMsg = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'progBarMeter
        '
        Me.progBarMeter.Location = New System.Drawing.Point(8, 64)
        Me.progBarMeter.Name = "progBarMeter"
        Me.progBarMeter.Size = New System.Drawing.Size(408, 23)
        Me.progBarMeter.TabIndex = 8
        '
        'labelProgress
        '
        Me.labelProgress.Location = New System.Drawing.Point(8, 40)
        Me.labelProgress.Name = "labelProgress"
        Me.labelProgress.Size = New System.Drawing.Size(408, 16)
        Me.labelProgress.TabIndex = 7
        Me.labelProgress.Text = "進捗メッセージ"
        Me.labelProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'labelMainMsg
        '
        Me.labelMainMsg.Location = New System.Drawing.Point(8, 16)
        Me.labelMainMsg.Name = "labelMainMsg"
        Me.labelMainMsg.Size = New System.Drawing.Size(408, 16)
        Me.labelMainMsg.TabIndex = 5
        Me.labelMainMsg.Text = "メイン・メッセージ"
        Me.labelMainMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'WaitDialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(424, 96)
        Me.Controls.Add(Me.progBarMeter)
        Me.Controls.Add(Me.labelProgress)
        Me.Controls.Add(Me.labelMainMsg)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WaitDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "起動中．．．"
        Me.ResumeLayout(False)

    End Sub

#End Region

  ' ShowDialogメソッドのシャドウ（WaitDialogクラスでは、ShowDialogメソッドは使用不可）
  Public Shadows Function ShowDialog() As System.Windows.Forms.DialogResult
    Debug.Assert(False, _
     "WaitDialogクラスはShowDialogメソッドを利用できません。" + vbCrLf + _
     "Showメソッドを使ってモードレス・ダイアログを構築してください。")
    Return DialogResult.Abort
  End Function

  ' Showメソッドのシャドウ
  Public Shadows Sub Show()
    ' 変数の初期化
    Me.DialogResult = DialogResult.OK
    Me.bAborting = False

    MyBase.Show()
    Me.bShowing = True
  End Sub

  ' Closeメソッドのシャドウ
  Public Shadows Sub Close()
    Me.bShowing = False
    MyBase.Close()
  End Sub

  ' キャンセル・ボタンが押されたときの処理
  ' 処理を途中でキャンセル（中断）する。
  Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ' 中止処理
    Abort()
  End Sub

  ' 中止（キャンセル）処理
  Private Sub Abort()
    Me.bAborting = True
    Me.DialogResult = DialogResult.Abort
  End Sub

  ' ダイアログが閉じられるときの処理
  ' 右上の［閉じる］ボタンが押された場合には、
  ' ［キャンセル］ボタンと同じように、処理を途中でキャンセル（中断）する。
  Private Sub WaitDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    If bShowing = True Then
      ' ダイアログ表示中なので中止（キャンセル）処理を実行
      Abort()
      ' まだダイアログは閉じない
      e.Cancel = True
    Else
      ' フォームは閉じられるところので素直に閉じる
      e.Cancel = False
    End If
  End Sub

  ' **************************************************************

  ' 処理がキャンセル（中止）されているかどうかを示す値を取得する。
  ' キャンセルされた場合はtrue。それ以外はfalse。
  Public ReadOnly Property IsAborting() As Boolean
    Get
      Return Me.bAborting
    End Get
  End Property

  ' メイン・メッセージのテキストを設定する。
  ' 処理の概要を表示する。
  ' 例えば、ファイルの転送を行っているなら、「ファイルを転送しています……」のように表示する。
  Public WriteOnly Property MainMsg() As String
    Set(ByVal Value As String)
      Me.labelMainMsg.Text = Value
    End Set
  End Property

  ' 進行状況メッセージのテキストを設定する。
  ' 処理の進行状況として、「何件分の何件が終わったのか」「全体の何％が終わったのか」などを表示する。
  Public WriteOnly Property ProgressMsg() As String
    Set(ByVal Value As String)
      Me.labelProgress.Text = Value
    End Set
  End Property

  ' 進行状況メーターの現在位置を設定する。
  ' 例えば、処理に200の工数があった場合、現在その200の工数の中のどの位置にいるかを示す値。
  ' 既定値は「0」。
  Public WriteOnly Property ProgressValue() As Integer
    Set(ByVal Value As Integer)
      Me.progBarMeter.Value = Value
    End Set
  End Property

  ' 進行状況メーターの範囲の最大値を設定する。
  ' 処理に200の工数があるなら「200」になる。
  ' 既定値は「100」。
  Public WriteOnly Property ProgressMax() As Integer
    Set(ByVal Value As Integer)
      Me.progBarMeter.Maximum = Value
    End Set
  End Property

  ' 進行状況メーターの範囲の最小値を設定する。
  ' 既定値は「0」。
  Public WriteOnly Property ProgressMin() As Integer
    Set(ByVal Value As Integer)
      Me.progBarMeter.Minimum = Value
    End Set
  End Property

  ' PerformStepメソッドを呼び出したときに、進行状況メーターの現在位置を進める量（ProgressValueの増分値）を設定する。
  ' 処理工数が200で、5つの工数が終わった段階で進行状況メーターを更新したい場合は「5」にする。
  ' 既定値は「10」。
  Public WriteOnly Property ProgressStep() As Integer
    Set(ByVal Value As Integer)
      Me.progBarMeter.Step = Value
    End Set
  End Property

  ' 進行状況メーターの現在位置（ProgressValue）をProgressStepプロパティの量だけ進める。
  Public Sub PerformStep()
    Me.progBarMeter.PerformStep()
  End Sub

End Class

