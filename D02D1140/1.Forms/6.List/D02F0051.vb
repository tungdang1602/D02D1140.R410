Public Class D02F0051
	Dim report As D99C2003
	Private _formIDPermission As String = "D02F0051"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property

    Dim dtGrid, dtCaptionCols As DataTable
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()


#Region "Const of tdbgMaster"
    Private Const COL_ACodeID As Integer = 0          ' Mã phân tích XDCB
    Private Const COL_Description As Integer = 1      ' Diễn giải
    Private Const COL_Disabled As Integer = 2         ' Không sử dụng
    Private Const COL_CreateDate As Integer = 3       ' CreateDate
    Private Const COL_CreateUserID As Integer = 4     ' CreateUserID
    Private Const COL_LastModifyDate As Integer = 5   ' LastModifyDate
    Private Const COL_LastModifyUserID As Integer = 6 ' LastModifyUserID
    Private Const COL_TypeCodeID As Integer = 7       ' TypeCodeID
#End Region

    Private Sub D02F0051_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
    End Sub

    Private Sub D02F0051_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        ResetColorGrid(tdbg)
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        LoadTDBC()
        tdbcTypeCodeID.Text = "%"
        Loadlanguage()
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_ma__phan_tich_XDCB_-_D02F0051") & UnicodeCaption(gbUnicode) 'Danh móc mº  ph¡n tÛch XDCB - D02F0051
        '================================================================ 
        lblTypeCodeID.Text = rl3("Loai_phan_tich") 'Loại phân tích
        '================================================================ 
        tdbcTypeCodeID.Columns("TypeCodeID").Caption = rl3("Ma") 'Mã
        tdbcTypeCodeID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbg.Columns("ACodeID").Caption = rl3("Ma_phan_tich_XDCB") 'Mã phân tích XDCB
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'KSD
        '================================================================ 
        chkShowDisabled.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    End Sub

    Private Sub LoadTDBC()
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, gbUnicode)

        Dim sSQL As String
        If geLanguage = EnumLanguage.Vietnamese Then
            sSQL = "Select 0 as DisplayOrder,'%' As TypeCodeID , " & sLanguage & " As Description" & vbCrLf
            sSQL &= "Union All" & vbCrLf
            sSQL &= "Select 1 as DisplayOrder,TypeCodeID, VieTypeCodeName" & sUnicode & " As Description " & vbCrLf
            sSQL &= "From D02T0040 WITH(NOLOCK)" & vbCrLf
            sSQL &= "Where Type = 'X' And Disabled = 0 " & vbCrLf
            sSQL &= "Order By DisplayOrder,TypeCodeID"
        Else
            sSQL = "Select 0 as DisplayOrder,'%' As TypeCodeID , " & sLanguage & " As Description" & vbCrLf
            sSQL &= "Union All" & vbCrLf
            sSQL &= "Select 1 as DisplayOrder,TypeCodeID, EngTypeCodeName" & sUnicode & " As Description " & vbCrLf
            sSQL &= "From D02T0040 WITH(NOLOCK)" & vbCrLf
            sSQL &= "Where Type = 'X' And Disabled = 0 " & vbCrLf
            sSQL &= "Order By DisplayOrder,TypeCodeID"
        End If

        dtGrid = ReturnDataTable(sSQL)
        LoadDataSource(tdbcTypeCodeID, dtGrid, gbUnicode)
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String = ""
        sSQL = "Select ACodeID, TypeCodeID, Type, " & vbCrLf
        sSQL &= "       Description" & UnicodeJoin(gbUnicode) & " As Description, " & vbCrLf
        sSQL &= "       Disabled, CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
        sSQL &= "From   D02T0041 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where  Type = 'X'" & vbCrLf

        If tdbcTypeCodeID.Text <> "%" Then
            sSQL &= "And TypeCodeID = " & SQLString(tdbcTypeCodeID.Text) & vbCrLf
        End If

        sSQL &= "Order By   ACodeID"
        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0

        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFilter = New System.Text.StringBuilder("")
            sFind = ""
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("ACodeID = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid(Optional ByVal bUseFilterBar As Boolean = False)
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled =0"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_ACodeID)
    End Sub

#Region "Events tdbcTypeCodeID with txtTypeCodeName"

    Private Sub tdbcTypeCodeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.Close
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        txtTypeCodeName.Text = tdbcTypeCodeID.Columns(1).Value.ToString
        LoadTDBGrid(True)
        'Dim sSQL As String = ""

        'If tdbcTypeCodeID.Text <> "%" Then
        '    sSQL = "select * from D02T0041 where Type = 'X' and TypeCodeID=" & SQLString(tdbcTypeCodeID.Text)
        '    sSQL &= "Order by AcodeID"
        '    dtGrid = ReturnDataTable(sSQL)
        '    LoadDataSource(tdbg, dtGrid, gbUnicode)
        'End If

        'CheckMenu(PARA_FormIDPermission, C1CommandHolder, tdbg.RowCount, gbEnabledUseFind, False)
    End Sub

    Private Sub tdbcTypeCodeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcTypeCodeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcTypeCodeID.Text = ""
            txtTypeCodeName.Text = ""
        End If
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
	Dim gbEnabledUseFind As Boolean = False
    'Cần sửa Tìm kiếm như sau:
	'Bỏ sự kiện Finder_FindClick.
	'Sửa tham số Me.Name -> Me
	'Phải tạo biến properties có tên chính xác strNewFind và strNewServer
	'Sửa gdtCaptionExcel thành dtCaptionCols: biến toàn cục trong form
	'Nếu có F12 dùng D09U1111 thì Sửa dtCaptionCols thành ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable)
    Private sFind As String = ""
	Public WriteOnly Property strNewFind() As String
		Set(ByVal Value As String)
			sFind = Value
			ReLoadTDBGrid()'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
		End Set
	End Property


    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim frm As New D02F0052
        With frm
            .TypeCodeID = tdbcTypeCodeID.Text
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If frm.SavedOK Then LoadTDBGrid(True, .ACodeID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim frm As New D02F0052
        With frm
            .TypeCodeID = tdbg.Columns(COL_TypeCodeID).Text
            .ACodeID = tdbg.Columns(COL_ACodeID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        Dim frm As New D02F0052
        With frm
            .TypeCodeID = tdbg.Columns(COL_TypeCodeID).Text
            .ACodeID = tdbg.Columns(COL_ACodeID).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
        End With
        If frm.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_ACodeID).Text)
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowDelete() Then Exit Sub

        Dim sSQL As String = ""
        sSQL = "Delete D02T0041 Where ACodeID= " & SQLString(tdbg.Columns(COL_ACodeID).Text) & vbCrLf
        sSQL &= "And TypeCodeID = " & SQLString(tdbg.Columns(COL_TypeCodeID).Text)
        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult Then
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Function AllowDelete() As Boolean
        Dim sTemp As String
        Dim sSQL As String = ""
        sTemp = tdbg.Columns(COL_TypeCodeID).Text
        sSQL = "Select Top 1 1 From D02T0001 WITH(NOLOCK) Where ACode" & sTemp.Substring(sTemp.Length - 2, 2) & "ID=" & SQLString(tdbg.Columns(COL_ACodeID).Text)
        If ExistRecord(sSQL) Then
            D99C0008.MsgCanNotDelete()
            Return False
        End If
        Return True
    End Function

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, tsmPrint.Click, mnsPrint.Click
        Me.Cursor = Cursors.WaitCursor

        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R0041"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rl3("Danh_muc_ma_XDCB") & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        sSQL = "select * from D02T0041 WITH(NOLOCK) where Type = 'X' and TypeCodeID like" & SQLString(tdbcTypeCodeID.Text) & vbCrLf
        sSQL &= "Order By AcodeID"
        sSQLSub = "Select * from D91T0025 WITH(NOLOCK) "
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        With report
            .OpenConnection(conn)
            If tdbcTypeCodeID.Text <> "%" Then
                .AddParameter("txtTitleReport", rl3("DANH_MUC_MA_XDCBV") & tdbcTypeCodeID.Text)
            Else
                .AddParameter("txtTitleReport", rl3("DANH_MUC_MA_XDCBV"))

            End If
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(dtGrid.DefaultView.ToTable)
            .PrintReport(sPathReport, sReportCaption)
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Grid"

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#End Region

End Class