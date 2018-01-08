<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnGetSummaryInfo = New System.Windows.Forms.Button()
        Me.dataview = New System.Windows.Forms.DataGridView()
        Me.btnMigration = New System.Windows.Forms.Button()
        Me.status = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOutput = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ファイルToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOpenIni = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuReadIni = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOpenLog = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOpenStrage = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.dataview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetSummaryInfo
        '
        Me.btnGetSummaryInfo.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnGetSummaryInfo.Location = New System.Drawing.Point(12, 36)
        Me.btnGetSummaryInfo.Name = "btnGetSummaryInfo"
        Me.btnGetSummaryInfo.Size = New System.Drawing.Size(119, 30)
        Me.btnGetSummaryInfo.TabIndex = 0
        Me.btnGetSummaryInfo.Text = "サマリ情報取得"
        Me.btnGetSummaryInfo.UseVisualStyleBackColor = False
        '
        'dataview
        '
        Me.dataview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dataview.Location = New System.Drawing.Point(12, 72)
        Me.dataview.Name = "dataview"
        Me.dataview.RowTemplate.Height = 21
        Me.dataview.Size = New System.Drawing.Size(1097, 485)
        Me.dataview.TabIndex = 1
        '
        'btnMigration
        '
        Me.btnMigration.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnMigration.Location = New System.Drawing.Point(262, 36)
        Me.btnMigration.Name = "btnMigration"
        Me.btnMigration.Size = New System.Drawing.Size(119, 30)
        Me.btnMigration.TabIndex = 2
        Me.btnMigration.Text = "サマリ移行"
        Me.btnMigration.UseVisualStyleBackColor = False
        '
        'status
        '
        Me.status.AutoSize = True
        Me.status.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.status.Location = New System.Drawing.Point(402, 41)
        Me.status.Name = "status"
        Me.status.Size = New System.Drawing.Size(50, 16)
        Me.status.TabIndex = 3
        Me.status.Text = "status"
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnExit.Location = New System.Drawing.Point(1013, 36)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(96, 30)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "終了"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnOutput
        '
        Me.btnOutput.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnOutput.Location = New System.Drawing.Point(137, 36)
        Me.btnOutput.Name = "btnOutput"
        Me.btnOutput.Size = New System.Drawing.Size(119, 30)
        Me.btnOutput.TabIndex = 5
        Me.btnOutput.Text = "ファイル出力"
        Me.btnOutput.UseVisualStyleBackColor = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ファイルToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1121, 24)
        Me.MenuStrip1.TabIndex = 6
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ファイルToolStripMenuItem
        '
        Me.ファイルToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuOpenIni, Me.menuReadIni, Me.menuOpenLog, Me.menuOpenStrage})
        Me.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem"
        Me.ファイルToolStripMenuItem.Size = New System.Drawing.Size(51, 20)
        Me.ファイルToolStripMenuItem.Text = "ファイル"
        '
        'menuOpenIni
        '
        Me.menuOpenIni.Name = "menuOpenIni"
        Me.menuOpenIni.Size = New System.Drawing.Size(152, 22)
        Me.menuOpenIni.Text = "iniファイルを開く"
        '
        'menuReadIni
        '
        Me.menuReadIni.Name = "menuReadIni"
        Me.menuReadIni.Size = New System.Drawing.Size(152, 22)
        Me.menuReadIni.Text = "設定反映"
        '
        'menuOpenLog
        '
        Me.menuOpenLog.Name = "menuOpenLog"
        Me.menuOpenLog.Size = New System.Drawing.Size(152, 22)
        Me.menuOpenLog.Text = "logファイルを開く"
        '
        'menuOpenStrage
        '
        Me.menuOpenStrage.Name = "menuOpenStrage"
        Me.menuOpenStrage.Size = New System.Drawing.Size(152, 22)
        Me.menuOpenStrage.Text = "ストレージを開く"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(186, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(243, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1121, 565)
        Me.Controls.Add(Me.btnOutput)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.status)
        Me.Controls.Add(Me.btnMigration)
        Me.Controls.Add(Me.dataview)
        Me.Controls.Add(Me.btnGetSummaryInfo)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmMain"
        CType(Me.dataview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGetSummaryInfo As Button
    Friend WithEvents dataview As DataGridView
    Friend WithEvents btnMigration As Button
    Friend WithEvents status As Label
    Friend WithEvents btnExit As Button
    Friend WithEvents btnOutput As Button
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ファイルToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents menuOpenIni As ToolStripMenuItem
    Friend WithEvents menuReadIni As ToolStripMenuItem
    Friend WithEvents menuOpenLog As ToolStripMenuItem
    Friend WithEvents menuOpenStrage As ToolStripMenuItem
End Class
