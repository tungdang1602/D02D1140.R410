<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F0052
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F0052))
        Me.tdbcTypeCodeID = New C1.Win.C1List.C1Combo()
        Me.lblTypeCodeID = New System.Windows.Forms.Label()
        Me.txtDescription_Type = New System.Windows.Forms.TextBox()
        Me.txtACodeID = New System.Windows.Forms.TextBox()
        Me.lblACodeID = New System.Windows.Forms.Label()
        Me.chkDisabled = New System.Windows.Forms.CheckBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tdbcTypeCodeID
        '
        Me.tdbcTypeCodeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcTypeCodeID.AllowColMove = False
        Me.tdbcTypeCodeID.AllowSort = False
        Me.tdbcTypeCodeID.AlternatingRows = True
        Me.tdbcTypeCodeID.AutoCompletion = True
        Me.tdbcTypeCodeID.AutoDropDown = True
        Me.tdbcTypeCodeID.Caption = ""
        Me.tdbcTypeCodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcTypeCodeID.ColumnWidth = 100
        Me.tdbcTypeCodeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcTypeCodeID.DisplayMember = "TypeCodeID"
        Me.tdbcTypeCodeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcTypeCodeID.DropDownWidth = 500
        Me.tdbcTypeCodeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcTypeCodeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcTypeCodeID.EmptyRows = True
        Me.tdbcTypeCodeID.ExtendRightColumn = True
        Me.tdbcTypeCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.Images.Add(CType(resources.GetObject("tdbcTypeCodeID.Images"), System.Drawing.Image))
        Me.tdbcTypeCodeID.Location = New System.Drawing.Point(120, 12)
        Me.tdbcTypeCodeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcTypeCodeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcTypeCodeID.MaxLength = 20
        Me.tdbcTypeCodeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcTypeCodeID.Name = "tdbcTypeCodeID"
        Me.tdbcTypeCodeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcTypeCodeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcTypeCodeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcTypeCodeID.TabIndex = 1
        Me.tdbcTypeCodeID.ValueMember = "TypeCodeID"
        Me.tdbcTypeCodeID.PropBag = resources.GetString("tdbcTypeCodeID.PropBag")
        '
        'lblTypeCodeID
        '
        Me.lblTypeCodeID.AutoSize = True
        Me.lblTypeCodeID.Location = New System.Drawing.Point(7, 17)
        Me.lblTypeCodeID.Name = "lblTypeCodeID"
        Me.lblTypeCodeID.Size = New System.Drawing.Size(101, 15)
        Me.lblTypeCodeID.TabIndex = 0
        Me.lblTypeCodeID.Text = "Mã loại phân tích"
        Me.lblTypeCodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDescription_Type
        '
        Me.txtDescription_Type.Enabled = False
        Me.txtDescription_Type.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.txtDescription_Type.Location = New System.Drawing.Point(253, 13)
        Me.txtDescription_Type.MaxLength = 250
        Me.txtDescription_Type.Name = "txtDescription_Type"
        Me.txtDescription_Type.ReadOnly = True
        Me.txtDescription_Type.Size = New System.Drawing.Size(215, 20)
        Me.txtDescription_Type.TabIndex = 2
        Me.txtDescription_Type.TabStop = False
        '
        'txtACodeID
        '
        Me.txtACodeID.BackColor = System.Drawing.SystemColors.Window
        Me.txtACodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtACodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtACodeID.Location = New System.Drawing.Point(120, 41)
        Me.txtACodeID.MaxLength = 20
        Me.txtACodeID.Name = "txtACodeID"
        Me.txtACodeID.Size = New System.Drawing.Size(128, 20)
        Me.txtACodeID.TabIndex = 4
        '
        'lblACodeID
        '
        Me.lblACodeID.AutoSize = True
        Me.lblACodeID.Location = New System.Drawing.Point(7, 46)
        Me.lblACodeID.Name = "lblACodeID"
        Me.lblACodeID.Size = New System.Drawing.Size(89, 15)
        Me.lblACodeID.TabIndex = 3
        Me.lblACodeID.Text = "Mã khoản mục"
        Me.lblACodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkDisabled
        '
        Me.chkDisabled.AutoSize = True
        Me.chkDisabled.Location = New System.Drawing.Point(309, 44)
        Me.chkDisabled.Name = "chkDisabled"
        Me.chkDisabled.Size = New System.Drawing.Size(109, 19)
        Me.chkDisabled.TabIndex = 5
        Me.chkDisabled.Text = "Không sử dụng"
        Me.chkDisabled.UseVisualStyleBackColor = True
        '
        'txtDescription
        '
        Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtDescription.Location = New System.Drawing.Point(120, 69)
        Me.txtDescription.MaxLength = 250
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(348, 20)
        Me.txtDescription.TabIndex = 7
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(6, 74)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(56, 15)
        Me.lblDescription.TabIndex = 6
        Me.lblDescription.Text = "Diễn giải"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(227, 102)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 8
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(391, 102)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(309, 102)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(76, 22)
        Me.btnNext.TabIndex = 9
        Me.btnNext.Text = "Nhập &tiếp"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'D02F0052
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 133)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.chkDisabled)
        Me.Controls.Add(Me.txtACodeID)
        Me.Controls.Add(Me.tdbcTypeCodeID)
        Me.Controls.Add(Me.lblTypeCodeID)
        Me.Controls.Add(Me.txtDescription_Type)
        Me.Controls.Add(Me.lblACodeID)
        Me.Controls.Add(Me.lblDescription)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F0052"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CËp nhËt mº ph¡n tÛch XDCB - D02F0052"
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents tdbcTypeCodeID As C1.Win.C1List.C1Combo
    Private WithEvents lblTypeCodeID As System.Windows.Forms.Label
    Private WithEvents txtDescription_Type As System.Windows.Forms.TextBox
    Private WithEvents txtACodeID As System.Windows.Forms.TextBox
    Private WithEvents lblACodeID As System.Windows.Forms.Label
    Private WithEvents chkDisabled As System.Windows.Forms.CheckBox
    Private WithEvents txtDescription As System.Windows.Forms.TextBox
    Private WithEvents lblDescription As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
End Class