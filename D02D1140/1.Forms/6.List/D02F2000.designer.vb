<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F2000
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F2000))
        Me.tdbg = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnsAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnsView = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnsEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnsDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator_Find = New System.Windows.Forms.ToolStripSeparator()
        Me.mnsFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnsListAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnsExportToExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnsImportData = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator_SysInfo = New System.Windows.Forms.ToolStripSeparator()
        Me.mnsSysInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.TableToolStrip = New System.Windows.Forms.ToolStrip()
        Me.tsbAdd = New System.Windows.Forms.ToolStripButton()
        Me.tsbView = New System.Windows.Forms.ToolStripButton()
        Me.tsbEdit = New System.Windows.Forms.ToolStripButton()
        Me.tsbDelete = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbFind = New System.Windows.Forms.ToolStripButton()
        Me.tsbListAll = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbExportToExcel = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbImportData = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSysInfo = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbClose = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsdActive = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsmAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmView = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmListAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator12 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmExportToExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmImportData = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmSysInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.chkShowDisabled = New System.Windows.Forms.CheckBox()
        Me.c1dateDate = New C1.Win.C1Input.C1DateEdit()
        Me.btnCodeID = New System.Windows.Forms.Button()
        Me.btnInformation = New System.Windows.Forms.Button()
        Me.grpFilter = New System.Windows.Forms.GroupBox()
        Me.txtCipNo = New System.Windows.Forms.TextBox()
        Me.lblCipNo = New System.Windows.Forms.Label()
        Me.txtCipName = New System.Windows.Forms.TextBox()
        Me.lblCipName = New System.Windows.Forms.Label()
        Me.btnFilter = New System.Windows.Forms.Button()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.TableToolStrip.SuspendLayout()
        CType(Me.c1dateDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'tdbg
        '
        Me.tdbg.AllowColMove = False
        Me.tdbg.AllowColSelect = False
        Me.tdbg.AllowFilter = False
        Me.tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg.AllowUpdate = False
        Me.tdbg.AlternatingRows = True
        Me.tdbg.ColumnFooters = True
        Me.tdbg.ContextMenuStrip = Me.ContextMenuStrip1
        Me.tdbg.EmptyRows = True
        Me.tdbg.ExtendRightColumn = True
        Me.tdbg.FilterBar = True
        Me.tdbg.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg.Images.Add(CType(resources.GetObject("tdbg.Images"), System.Drawing.Image))
        Me.tdbg.Location = New System.Drawing.Point(3, 123)
        Me.tdbg.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.HighlightRowRaiseCell
        Me.tdbg.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg.Name = "tdbg"
        Me.tdbg.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg.PrintInfo.PageSettings = CType(resources.GetObject("tdbg.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg.Size = New System.Drawing.Size(1012, 501)
        Me.tdbg.SplitDividerSize = New System.Drawing.Size(1, 1)
        Me.tdbg.TabAcrossSplits = True
        Me.tdbg.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg.TabIndex = 4
        Me.tdbg.Tag = "COL"
        Me.tdbg.PropBag = resources.GetString("tdbg.PropBag")
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnsAdd, Me.mnsView, Me.mnsEdit, Me.mnsDelete, Me.ToolStripSeparator_Find, Me.mnsFind, Me.mnsListAll, Me.ToolStripSeparator10, Me.mnsExportToExcel, Me.ToolStripSeparator8, Me.mnsImportData, Me.ToolStripSeparator_SysInfo, Me.mnsSysInfo})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(178, 226)
        '
        'mnsAdd
        '
        Me.mnsAdd.Name = "mnsAdd"
        Me.mnsAdd.Size = New System.Drawing.Size(177, 22)
        Me.mnsAdd.Text = "&Thêm"
        '
        'mnsView
        '
        Me.mnsView.Name = "mnsView"
        Me.mnsView.Size = New System.Drawing.Size(177, 22)
        Me.mnsView.Text = "Xe&m"
        '
        'mnsEdit
        '
        Me.mnsEdit.Name = "mnsEdit"
        Me.mnsEdit.Size = New System.Drawing.Size(177, 22)
        Me.mnsEdit.Text = "&Sửa"
        '
        'mnsDelete
        '
        Me.mnsDelete.Name = "mnsDelete"
        Me.mnsDelete.Size = New System.Drawing.Size(177, 22)
        Me.mnsDelete.Text = "&Xóa"
        '
        'ToolStripSeparator_Find
        '
        Me.ToolStripSeparator_Find.Name = "ToolStripSeparator_Find"
        Me.ToolStripSeparator_Find.Size = New System.Drawing.Size(174, 6)
        '
        'mnsFind
        '
        Me.mnsFind.Name = "mnsFind"
        Me.mnsFind.Size = New System.Drawing.Size(177, 22)
        Me.mnsFind.Text = "Tìm &kiếm"
        '
        'mnsListAll
        '
        Me.mnsListAll.Name = "mnsListAll"
        Me.mnsListAll.Size = New System.Drawing.Size(177, 22)
        Me.mnsListAll.Text = "&Liệt kê tất cả"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(174, 6)
        '
        'mnsExportToExcel
        '
        Me.mnsExportToExcel.Name = "mnsExportToExcel"
        Me.mnsExportToExcel.Size = New System.Drawing.Size(177, 22)
        Me.mnsExportToExcel.Text = "Xuất &Excel"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(174, 6)
        '
        'mnsImportData
        '
        Me.mnsImportData.Name = "mnsImportData"
        Me.mnsImportData.Size = New System.Drawing.Size(177, 22)
        Me.mnsImportData.Text = "Import"
        '
        'ToolStripSeparator_SysInfo
        '
        Me.ToolStripSeparator_SysInfo.Name = "ToolStripSeparator_SysInfo"
        Me.ToolStripSeparator_SysInfo.Size = New System.Drawing.Size(174, 6)
        '
        'mnsSysInfo
        '
        Me.mnsSysInfo.Name = "mnsSysInfo"
        Me.mnsSysInfo.Size = New System.Drawing.Size(177, 22)
        Me.mnsSysInfo.Text = "Thông tin &hệ thống"
        '
        'TableToolStrip
        '
        Me.TableToolStrip.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.TableToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbAdd, Me.tsbView, Me.tsbEdit, Me.tsbDelete, Me.ToolStripSeparator1, Me.tsbFind, Me.tsbListAll, Me.ToolStripSeparator11, Me.tsbExportToExcel, Me.ToolStripSeparator3, Me.tsbImportData, Me.ToolStripSeparator9, Me.tsbSysInfo, Me.ToolStripSeparator2, Me.tsbClose, Me.ToolStripSeparator4, Me.tsdActive})
        Me.TableToolStrip.Location = New System.Drawing.Point(0, 0)
        Me.TableToolStrip.Name = "TableToolStrip"
        Me.TableToolStrip.Size = New System.Drawing.Size(1018, 25)
        Me.TableToolStrip.TabIndex = 0
        Me.TableToolStrip.Text = "tbrTest"
        '
        'tsbAdd
        '
        Me.tsbAdd.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbAdd.Image = CType(resources.GetObject("tsbAdd.Image"), System.Drawing.Image)
        Me.tsbAdd.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAdd.Name = "tsbAdd"
        Me.tsbAdd.Size = New System.Drawing.Size(42, 22)
        Me.tsbAdd.Text = "Thêm"
        '
        'tsbView
        '
        Me.tsbView.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbView.Name = "tsbView"
        Me.tsbView.Size = New System.Drawing.Size(35, 22)
        Me.tsbView.Text = "Xem"
        '
        'tsbEdit
        '
        Me.tsbEdit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbEdit.Name = "tsbEdit"
        Me.tsbEdit.Size = New System.Drawing.Size(30, 22)
        Me.tsbEdit.Text = "Sửa"
        '
        'tsbDelete
        '
        Me.tsbDelete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbDelete.Name = "tsbDelete"
        Me.tsbDelete.Size = New System.Drawing.Size(31, 22)
        Me.tsbDelete.Text = "Xóa"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'tsbFind
        '
        Me.tsbFind.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbFind.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFind.Name = "tsbFind"
        Me.tsbFind.Size = New System.Drawing.Size(61, 22)
        Me.tsbFind.Text = "Tìm kiếm"
        '
        'tsbListAll
        '
        Me.tsbListAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbListAll.Name = "tsbListAll"
        Me.tsbListAll.Size = New System.Drawing.Size(45, 22)
        Me.tsbListAll.Text = "Liệt kê"
        '
        'ToolStripSeparator11
        '
        Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
        Me.ToolStripSeparator11.Size = New System.Drawing.Size(6, 25)
        '
        'tsbExportToExcel
        '
        Me.tsbExportToExcel.Name = "tsbExportToExcel"
        Me.tsbExportToExcel.Size = New System.Drawing.Size(64, 22)
        Me.tsbExportToExcel.Text = "Xuất &Excel"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'tsbImportData
        '
        Me.tsbImportData.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbImportData.Image = CType(resources.GetObject("tsbImportData.Image"), System.Drawing.Image)
        Me.tsbImportData.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbImportData.Name = "tsbImportData"
        Me.tsbImportData.Size = New System.Drawing.Size(23, 22)
        Me.tsbImportData.Text = "ToolStripButton1"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 25)
        '
        'tsbSysInfo
        '
        Me.tsbSysInfo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbSysInfo.Name = "tsbSysInfo"
        Me.tsbSysInfo.Size = New System.Drawing.Size(63, 22)
        Me.tsbSysInfo.Text = "Thông tin"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'tsbClose
        '
        Me.tsbClose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbClose.Name = "tsbClose"
        Me.tsbClose.Size = New System.Drawing.Size(40, 22)
        Me.tsbClose.Text = "Đóng"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'tsdActive
        '
        Me.tsdActive.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsdActive.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmAdd, Me.tsmView, Me.tsmEdit, Me.tsmDelete, Me.ToolStripSeparator5, Me.tsmFind, Me.tsmListAll, Me.ToolStripSeparator12, Me.tsmExportToExcel, Me.ToolStripSeparator7, Me.tsmImportData, Me.ToolStripSeparator6, Me.tsmSysInfo})
        Me.tsdActive.Image = CType(resources.GetObject("tsdActive.Image"), System.Drawing.Image)
        Me.tsdActive.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsdActive.Name = "tsdActive"
        Me.tsdActive.Size = New System.Drawing.Size(73, 22)
        Me.tsdActive.Text = "Thực hiện"
        '
        'tsmAdd
        '
        Me.tsmAdd.Name = "tsmAdd"
        Me.tsmAdd.Size = New System.Drawing.Size(177, 22)
        Me.tsmAdd.Text = "Thêm"
        '
        'tsmView
        '
        Me.tsmView.Name = "tsmView"
        Me.tsmView.Size = New System.Drawing.Size(177, 22)
        Me.tsmView.Text = "Xem"
        '
        'tsmEdit
        '
        Me.tsmEdit.Name = "tsmEdit"
        Me.tsmEdit.Size = New System.Drawing.Size(177, 22)
        Me.tsmEdit.Text = "Sửa"
        '
        'tsmDelete
        '
        Me.tsmDelete.Name = "tsmDelete"
        Me.tsmDelete.Size = New System.Drawing.Size(177, 22)
        Me.tsmDelete.Text = "Xóa"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(174, 6)
        '
        'tsmFind
        '
        Me.tsmFind.Name = "tsmFind"
        Me.tsmFind.Size = New System.Drawing.Size(177, 22)
        Me.tsmFind.Text = "Tìm kiếm"
        '
        'tsmListAll
        '
        Me.tsmListAll.Name = "tsmListAll"
        Me.tsmListAll.Size = New System.Drawing.Size(177, 22)
        Me.tsmListAll.Text = "Liệt kê tất cả"
        '
        'ToolStripSeparator12
        '
        Me.ToolStripSeparator12.Name = "ToolStripSeparator12"
        Me.ToolStripSeparator12.Size = New System.Drawing.Size(174, 6)
        '
        'tsmExportToExcel
        '
        Me.tsmExportToExcel.Name = "tsmExportToExcel"
        Me.tsmExportToExcel.Size = New System.Drawing.Size(177, 22)
        Me.tsmExportToExcel.Text = "Xuất &Excel"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(174, 6)
        '
        'tsmImportData
        '
        Me.tsmImportData.Name = "tsmImportData"
        Me.tsmImportData.Size = New System.Drawing.Size(177, 22)
        Me.tsmImportData.Text = "Import"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(174, 6)
        '
        'tsmSysInfo
        '
        Me.tsmSysInfo.Name = "tsmSysInfo"
        Me.tsmSysInfo.Size = New System.Drawing.Size(177, 22)
        Me.tsmSysInfo.Text = "Thông tin hệ thống"
        '
        'chkShowDisabled
        '
        Me.chkShowDisabled.AutoSize = True
        Me.chkShowDisabled.Location = New System.Drawing.Point(3, 629)
        Me.chkShowDisabled.Name = "chkShowDisabled"
        Me.chkShowDisabled.Size = New System.Drawing.Size(186, 17)
        Me.chkShowDisabled.TabIndex = 5
        Me.chkShowDisabled.Text = "Hiển thị danh mục không sử dụng"
        Me.chkShowDisabled.UseVisualStyleBackColor = True
        '
        'c1dateDate
        '
        Me.c1dateDate.AutoSize = False
        Me.c1dateDate.CustomFormat = "dd/MM/yyyy"
        Me.c1dateDate.EmptyAsNull = True
        Me.c1dateDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.c1dateDate.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateDate.Location = New System.Drawing.Point(333, 485)
        Me.c1dateDate.Name = "c1dateDate"
        Me.c1dateDate.Size = New System.Drawing.Size(100, 22)
        Me.c1dateDate.TabIndex = 6
        Me.c1dateDate.Tag = Nothing
        Me.c1dateDate.TrimStart = True
        Me.c1dateDate.Visible = False
        Me.c1dateDate.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'btnCodeID
        '
        Me.btnCodeID.Location = New System.Drawing.Point(778, 94)
        Me.btnCodeID.Name = "btnCodeID"
        Me.btnCodeID.Size = New System.Drawing.Size(111, 22)
        Me.btnCodeID.TabIndex = 2
        Me.btnCodeID.Text = "1. Mã phân tích"
        Me.btnCodeID.UseVisualStyleBackColor = True
        '
        'btnInformation
        '
        Me.btnInformation.Location = New System.Drawing.Point(895, 94)
        Me.btnInformation.Name = "btnInformation"
        Me.btnInformation.Size = New System.Drawing.Size(120, 22)
        Me.btnInformation.TabIndex = 3
        Me.btnInformation.Text = "2. Thông tin phụ"
        Me.btnInformation.UseVisualStyleBackColor = True
        '
        'grpFilter
        '
        Me.grpFilter.Controls.Add(Me.btnFilter)
        Me.grpFilter.Controls.Add(Me.txtCipName)
        Me.grpFilter.Controls.Add(Me.txtCipNo)
        Me.grpFilter.Controls.Add(Me.lblCipNo)
        Me.grpFilter.Controls.Add(Me.lblCipName)
        Me.grpFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpFilter.Location = New System.Drawing.Point(3, 31)
        Me.grpFilter.Name = "grpFilter"
        Me.grpFilter.Size = New System.Drawing.Size(1012, 55)
        Me.grpFilter.TabIndex = 1
        Me.grpFilter.TabStop = False
        Me.grpFilter.Text = "Tiêu thức lọc"
        '
        'txtCipNo
        '
        Me.txtCipNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCipNo.Location = New System.Drawing.Point(154, 25)
        Me.txtCipNo.MaxLength = 50
        Me.txtCipNo.Name = "txtCipNo"
        Me.txtCipNo.Size = New System.Drawing.Size(210, 20)
        Me.txtCipNo.TabIndex = 1
        '
        'lblCipNo
        '
        Me.lblCipNo.AutoSize = True
        Me.lblCipNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCipNo.Location = New System.Drawing.Point(6, 29)
        Me.lblCipNo.Name = "lblCipNo"
        Me.lblCipNo.Size = New System.Drawing.Size(96, 13)
        Me.lblCipNo.TabIndex = 0
        Me.lblCipNo.Text = "Mã XDCB có chứa"
        Me.lblCipNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCipName
        '
        Me.txtCipName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCipName.Location = New System.Drawing.Point(503, 25)
        Me.txtCipName.MaxLength = 500
        Me.txtCipName.Name = "txtCipName"
        Me.txtCipName.Size = New System.Drawing.Size(409, 20)
        Me.txtCipName.TabIndex = 3
        '
        'lblCipName
        '
        Me.lblCipName.AutoSize = True
        Me.lblCipName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCipName.Location = New System.Drawing.Point(371, 29)
        Me.lblCipName.Name = "lblCipName"
        Me.lblCipName.Size = New System.Drawing.Size(100, 13)
        Me.lblCipName.TabIndex = 2
        Me.lblCipName.Text = "Tên XDCB có chứa"
        Me.lblCipName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnFilter
        '
        Me.btnFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFilter.Location = New System.Drawing.Point(928, 24)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(76, 22)
        Me.btnFilter.TabIndex = 4
        Me.btnFilter.Text = "Lọc"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'D02F2000
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1018, 655)
        Me.Controls.Add(Me.grpFilter)
        Me.Controls.Add(Me.chkShowDisabled)
        Me.Controls.Add(Me.btnInformation)
        Me.Controls.Add(Me.TableToolStrip)
        Me.Controls.Add(Me.btnCodeID)
        Me.Controls.Add(Me.c1dateDate)
        Me.Controls.Add(Me.tdbg)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F2000"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Danh  móc mº c¤ng trØnh XDCB - D02F2000"
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.TableToolStrip.ResumeLayout(False)
        Me.TableToolStrip.PerformLayout()
        CType(Me.c1dateDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFilter.ResumeLayout(False)
        Me.grpFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents TableToolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbAdd As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbView As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbEdit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDelete As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbFind As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbListAll As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbSysInfo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbClose As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsdActive As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsmAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmFind As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmListAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmSysInfo As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents chkShowDisabled As System.Windows.Forms.CheckBox
    Private WithEvents c1dateDate As C1.Win.C1Input.C1DateEdit
    Private WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Private WithEvents mnsAdd As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents mnsView As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents mnsEdit As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents mnsDelete As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ToolStripSeparator_Find As System.Windows.Forms.ToolStripSeparator
    Private WithEvents mnsFind As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents mnsListAll As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ToolStripSeparator_SysInfo As System.Windows.Forms.ToolStripSeparator
    Private WithEvents mnsSysInfo As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents btnCodeID As System.Windows.Forms.Button
    Private WithEvents btnInformation As System.Windows.Forms.Button
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbImportData As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmImportData As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnsImportData As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents mnsExportToExcel As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents tsbExportToExcel As System.Windows.Forms.ToolStripButton
    Private WithEvents ToolStripSeparator12 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents tsmExportToExcel As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents grpFilter As System.Windows.Forms.GroupBox
    Private WithEvents btnFilter As System.Windows.Forms.Button
    Private WithEvents txtCipName As System.Windows.Forms.TextBox
    Private WithEvents txtCipNo As System.Windows.Forms.TextBox
    Private WithEvents lblCipNo As System.Windows.Forms.Label
    Private WithEvents lblCipName As System.Windows.Forms.Label
End Class