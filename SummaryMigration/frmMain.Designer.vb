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
        CType(Me.dataview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnGetSummaryInfo
        '
        Me.btnGetSummaryInfo.Location = New System.Drawing.Point(12, 12)
        Me.btnGetSummaryInfo.Name = "btnGetSummaryInfo"
        Me.btnGetSummaryInfo.Size = New System.Drawing.Size(119, 30)
        Me.btnGetSummaryInfo.TabIndex = 0
        Me.btnGetSummaryInfo.Text = "サマリ情報取得"
        Me.btnGetSummaryInfo.UseVisualStyleBackColor = True
        '
        'dataview
        '
        Me.dataview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dataview.Location = New System.Drawing.Point(12, 48)
        Me.dataview.Name = "dataview"
        Me.dataview.RowTemplate.Height = 21
        Me.dataview.Size = New System.Drawing.Size(1097, 509)
        Me.dataview.TabIndex = 1
        '
        'btnMigration
        '
        Me.btnMigration.Location = New System.Drawing.Point(163, 12)
        Me.btnMigration.Name = "btnMigration"
        Me.btnMigration.Size = New System.Drawing.Size(119, 30)
        Me.btnMigration.TabIndex = 2
        Me.btnMigration.Text = "サマリ移行"
        Me.btnMigration.UseVisualStyleBackColor = True
        '
        'status
        '
        Me.status.AutoSize = True
        Me.status.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.status.Location = New System.Drawing.Point(321, 17)
        Me.status.Name = "status"
        Me.status.Size = New System.Drawing.Size(60, 19)
        Me.status.TabIndex = 3
        Me.status.Text = "status"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1121, 569)
        Me.Controls.Add(Me.status)
        Me.Controls.Add(Me.btnMigration)
        Me.Controls.Add(Me.dataview)
        Me.Controls.Add(Me.btnGetSummaryInfo)
        Me.Name = "frmMain"
        CType(Me.dataview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGetSummaryInfo As Button
    Friend WithEvents dataview As DataGridView
    Friend WithEvents btnMigration As Button
    Friend WithEvents status As Label
End Class
