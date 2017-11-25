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
        Me.SuspendLayout()
        '
        'btnGetSummaryInfo
        '
        Me.btnGetSummaryInfo.Location = New System.Drawing.Point(84, 170)
        Me.btnGetSummaryInfo.Name = "btnGetSummaryInfo"
        Me.btnGetSummaryInfo.Size = New System.Drawing.Size(148, 55)
        Me.btnGetSummaryInfo.TabIndex = 0
        Me.btnGetSummaryInfo.Text = "サマリ情報取得"
        Me.btnGetSummaryInfo.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.btnGetSummaryInfo)
        Me.Name = "frmMain"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnGetSummaryInfo As Button
End Class
