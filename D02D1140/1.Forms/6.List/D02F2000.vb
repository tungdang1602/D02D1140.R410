'#-------------------------------------------------------------------------------------
'# Created Date: 04/09/2007 10:20:38 AM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 04/09/2007 10:20:38 AM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F2000
	Private _formIDPermission As String = "D02F2000"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
    End Property

    Private _savedOK As Boolean
    Public WriteOnly Property SavedOK() As Boolean
        Set(ByVal Value As Boolean)
            _savedOK = Value
        End Set
    End Property


#Region "Const of tdbg - Total of Columns: 58"
    Private Const COL_CipID As Integer = 0             ' CipID
    Private Const COL_CipNo As Integer = 1             ' Mã công trình
    Private Const COL_CipName As Integer = 2           ' Tên công trình
    Private Const COL_Description As Integer = 3       ' Diễn giải
    Private Const COL_AccountID As Integer = 4         ' TK phân bổ
    Private Const COL_StartDate As Integer = 5         ' Ngày bắt đầu
    Private Const COL_CompletionDate As Integer = 6    ' Ngày hoàn thành
    Private Const COL_SupplierID As Integer = 7        ' Mã NCC
    Private Const COL_SupplierName As Integer = 8      ' Tên NCC
    Private Const COL_ContractorID As Integer = 9      ' Mã nhà thầu
    Private Const COL_ContractorName As Integer = 10   ' Tên nhà thầu
    Private Const COL_StatusName As Integer = 11       ' Tình trạng
    Private Const COL_Disabled As Integer = 12         ' KSD
    Private Const COL_CreateUserID As Integer = 13     ' CreateUserID
    Private Const COL_CreateDate As Integer = 14       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 15 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 16   ' LastModifyDate
    Private Const COL_Status As Integer = 17           ' Status
    Private Const COL_X01ID As Integer = 18            ' X01ID
    Private Const COL_X02ID As Integer = 19            ' X02ID
    Private Const COL_X03ID As Integer = 20            ' X03ID
    Private Const COL_X04ID As Integer = 21            ' X04ID
    Private Const COL_X05ID As Integer = 22            ' X05ID
    Private Const COL_X06ID As Integer = 23            ' X06ID
    Private Const COL_X07ID As Integer = 24            ' X07ID
    Private Const COL_X08ID As Integer = 25            ' X08ID
    Private Const COL_X09ID As Integer = 26            ' X09ID
    Private Const COL_X10ID As Integer = 27            ' X10ID
    Private Const COL_CIPNum01 As Integer = 28         ' CIPNum01
    Private Const COL_CIPNum02 As Integer = 29         ' CIPNum02
    Private Const COL_CIPNum03 As Integer = 30         ' CIPNum03
    Private Const COL_CIPNum04 As Integer = 31         ' CIPNum04
    Private Const COL_CIPNum05 As Integer = 32         ' CIPNum05
    Private Const COL_CIPNum06 As Integer = 33         ' CIPNum06
    Private Const COL_CIPNum07 As Integer = 34         ' CIPNum07
    Private Const COL_CIPNum08 As Integer = 35         ' CIPNum08
    Private Const COL_CIPNum09 As Integer = 36         ' CIPNum09
    Private Const COL_CIPNum10 As Integer = 37         ' CIPNum10
    Private Const COL_CIPString01 As Integer = 38      ' CIPString01
    Private Const COL_CIPString02 As Integer = 39      ' CIPString02
    Private Const COL_CIPString03 As Integer = 40      ' CIPString03
    Private Const COL_CIPString04 As Integer = 41      ' CIPString04
    Private Const COL_CIPString05 As Integer = 42      ' CIPString05
    Private Const COL_CIPString06 As Integer = 43      ' CIPString06
    Private Const COL_CIPString07 As Integer = 44      ' CIPString07
    Private Const COL_CIPString08 As Integer = 45      ' CIPString08
    Private Const COL_CIPString09 As Integer = 46      ' CIPString09
    Private Const COL_CIPString10 As Integer = 47      ' CIPString10
    Private Const COL_CIPDate01 As Integer = 48        ' CIPDate01
    Private Const COL_CIPDate02 As Integer = 49        ' CIPDate02
    Private Const COL_CIPDate03 As Integer = 50        ' CIPDate03
    Private Const COL_CIPDate04 As Integer = 51        ' CIPDate04
    Private Const COL_CIPDate05 As Integer = 52        ' CIPDate05
    Private Const COL_CIPDate06 As Integer = 53        ' CIPDate06
    Private Const COL_CIPDate07 As Integer = 54        ' CIPDate07
    Private Const COL_CIPDate08 As Integer = 55        ' CIPDate08
    Private Const COL_CIPDate09 As Integer = 56        ' CIPDate09
    Private Const COL_CIPDate10 As Integer = 57        ' CIPDate10
#End Region



    Private dtGrid, dtCaptionCols As DataTable
    Private sAuditCode As String
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

    Private Sub D02F2000_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        ElseIf e.KeyCode = Keys.F5 Then
            btnFilter_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub D02F2000_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadInfoGeneral()
        InputbyUnicode(Me, gbUnicode)
        CheckIdTextBox(txtCipNo, 50) '13/10/2020, Chung Hoàng Khánh:id 144563-Thuận Minh - Thêm thông tin Lọc ở màn hình mã XDCB
        gbEnabledUseFind = False
        LoadTDBGridInformationCaption(tdbg, COL_CIPNum01, 1, True, gbUnicode)
        LoadTDBGridTypeCodeID(tdbg, COL_X01ID, 1, True, gbUnicode)
        tdbg.Columns(COL_StartDate).Editor = c1dateDate
        tdbg.Columns(COL_CompletionDate).Editor = c1dateDate
        InputDateInTrueDBGrid(tdbg, COL_CIPDate01, COL_CIPDate02, COL_CIPDate03, COL_CIPDate04, COL_CIPDate05, COL_CIPDate06, COL_CIPDate07, COL_CIPDate08, COL_CIPDate09, COL_CIPDate10)
        Loadlanguage()
        ResetColorGrid(tdbg, 0, 1)
        sAuditCode = "CIP"
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)

        btnCodeID_Click(Nothing, Nothing)
        ResetGrid()
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Danh_muc_XDCB") & " - " & Me.Name & UnicodeCaption(gbUnicode)  'rl3("Danh__muc_ma_cong_trinh_XDCB_-_D02F2000") & UnicodeCaption(gbUnicode) 'Danh  móc mº c¤ng trØnh XDCB - D02F2000
        '================================================================ 
        lblCipNo.Text = rL3("Ma_XDCB_co_chua") 'Mã XDCB có chứa
        lblCipName.Text = rL3("Ten_XDCB_co_chua") 'Tên XDCB có chứa
        '================================================================ 
        btnFilter.Text = rL3("Loc") & Space(1) & "(F5)" '&Lọc
        btnCodeID.Text = "1. " & rL3("Ma_phan_tich")
        btnInformation.Text = "2. " & rL3("Thong_tin_phu")
        '================================================================ 
        chkShowDisabled.Text = rL3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        '================================================================ 
        grpFilter.Text = rL3("Tieu_thuc_loc") 'Tiêu thức lọc
        '================================================================ 
        tdbg.Columns("CipNo").Caption = rL3("Ma_XDCB") 'Mã công trình
        tdbg.Columns("CipName").Caption = rL3("Ten_XDCB") 'Tên công trình
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("AccountID").Caption = rl3("TK_phan_bo") 'TK phân bổ
        tdbg.Columns("StartDate").Caption = rl3("Ngay_bat_dau") 'Ngày bắt đầu
        tdbg.Columns("CompletionDate").Caption = rL3("Ngay_hoan_thanh") 'Ngày hoàn thành
        tdbg.Columns("SupplierID").Caption = rL3("Ma_NCC") 'Mã NCC
        tdbg.Columns("SupplierName").Caption = rL3("Ten_NCC") 'Tên NCC
        tdbg.Columns("ContractorID").Caption = rL3("Ma_nha_thau") 'Mã nhà thầu
        tdbg.Columns("ContractorName").Caption = rL3("Ten__nha_thau") 'Tên  nhà thầu
        tdbg.Columns("StatusName").Caption = rL3("Tinh_trang") 'Tình trạng
        tdbg.Columns("Disabled").Caption = rL3("KSD") 'Không sử dụng
    End Sub

    '13/10/2020, Chung Hoàng Khánh:id 144563-Thuận Minh - Thêm thông tin Lọc ở màn hình mã XDCB
    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        btnFilter.Focus()
        If btnFilter.Focused = False Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        sFind = ""
        gbEnabledUseFind = False
        LoadTDBGrid()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        'Dim sSQL As New StringBuilder("")

        'sSQL.Append(" SELECT 	CipID, CipNo, RefCipID, Description" & UnicodeJoin(gbUnicode) & " As Description, CipName" & UnicodeJoin(gbUnicode) & " As CipName, AccountID, Quantity, Disabled, " & vbCrLf)
        'sSQL.Append("           Status,StartDate,CompletionDate,AssetID,CreateDate,CreateUserID, " & vbCrLf)
        'sSQL.Append("           LastModifyDate,LastModifyUserID,X01ID,X02ID,X03ID,X04ID,X05ID,X06ID," & vbCrLf)
        'sSQL.Append("           X07ID,X08ID,X09ID,X10ID,DivisionID, " & vbCrLf)
        'If gbUnicode Then
        '    sSQL.Append(" 	(CASE Status 	WHEN 0 THEN N'" & ConvertVniToUnicode(rl3("Moi_thiet_lap")) & "'" & vbCrLf)
        '    sSQL.Append(" 			        WHEN 1 THEN N'" & ConvertVniToUnicode(rl3("Dang_tap_hop")) & "'" & vbCrLf)
        '    sSQL.Append(" 			        WHEN 2 THEN N'" & ConvertVniToUnicode(rl3("Da_ket_thuc")) & "'" & vbCrLf)
        'Else
        '    sSQL.Append(" 	(CASE Status 	WHEN 0 THEN '" & rl3("Moi_thiet_lap") & "'" & vbCrLf)
        '    sSQL.Append(" 			        WHEN 1 THEN '" & rl3("Dang_tap_hop") & "'" & vbCrLf)
        '    sSQL.Append(" 			        WHEN 2 THEN '" & rl3("Da_ket_thuc") & "'" & vbCrLf)
        'End If
        'sSQL.Append(" 	END)    AS StatusName," & vbCrLf)
        'sSQL.Append("CIPString01" & UnicodeJoin(gbUnicode) & " as CIPString01," & vbCrLf)
        'sSQL.Append("CIPString02" & UnicodeJoin(gbUnicode) & " as CIPString02," & vbCrLf)
        'sSQL.Append("CIPString03" & UnicodeJoin(gbUnicode) & " as CIPString03," & vbCrLf)
        'sSQL.Append("CIPString04" & UnicodeJoin(gbUnicode) & " as CIPString04," & vbCrLf)
        'sSQL.Append("CIPString05" & UnicodeJoin(gbUnicode) & " as CIPString05," & vbCrLf)
        'sSQL.Append("CIPString06" & UnicodeJoin(gbUnicode) & " as CIPString06," & vbCrLf)
        'sSQL.Append("CIPString07" & UnicodeJoin(gbUnicode) & " as CIPString07," & vbCrLf)
        'sSQL.Append("CIPString08" & UnicodeJoin(gbUnicode) & " as CIPString08," & vbCrLf)
        'sSQL.Append("CIPString09" & UnicodeJoin(gbUnicode) & " as CIPString09," & vbCrLf)
        'sSQL.Append("CIPString10" & UnicodeJoin(gbUnicode) & " as CIPString10," & vbCrLf)
        'sSQL.Append("CIPNum01,CIPNum02, CIPNum03, CIPNum04, CIPNum05,CIPNum06, CIPNum07, CIPNum08, CIPNum09, CIPNum10,	CIPDate01, CIPDate02, CIPDate03, CIPDate04, CIPDate05,CIPDate06, CIPDate07, CIPDate08, CIPDate09, CIPDate10 " & vbCrLf)
        'sSQL.Append(" FROM D02T0100 WITH(NOLOCK)" & vbCrLf)
        'sSQL.Append(" WHERE DivisionID =" & SQLString(gsDivisionID) & vbCrLf)

        dtGrid = ReturnDataTable(SQLStoreD02P0025) '15/6/2017, 	Phạm Thị Thu: id 97556-Bổ sung điều kiện tìm kiếm tại Danh mục / Mã XDCB

        gbEnabledUseFind = dtGrid.Rows.Count > 0

        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
            sFilter = New System.Text.StringBuilder("")
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("CipNo = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid()
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
        '*Thay doi ngay 25/3/2013 theo ID 55186 bởi Văn Vinh
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1, , "D02F5810")
        'mnsImportData.Enabled = ReturnPermission("D02F5810") > 1 And gbClosed = False
        'tsmImportData.Enabled = ReturnPermission("D02F5810") > 1 And gbClosed = False
        'tsbImportData.Enabled = ReturnPermission("D02F5810") > 1 And gbClosed = False
        '******************************************************
        FooterTotalGrid(tdbg, COL_CipName)
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0025
    '# Created User: NGOCTHOAI
    '# Created Date: 15/06/2017 02:50:33
    '15/6/2017, 	Phạm Thị Thu: id 97556-Bổ sung điều kiện tìm kiếm tại Danh mục / Mã XDCB
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0025() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho luoi " & vbCrlf)
        sSQL &= "Exec D02P0025 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisonID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        '13/10/2020, Chung Hoàng Khánh:id 144563-Thuận Minh - Thêm thông tin Lọc ở màn hình mã XDCB
        sSQL &= SQLString(txtCipNo.Text) & COMMA 'CipNo, varchar[50], NOT NULL
        sSQL &= SQLStringUnicode(txtCipName.Text, gbUnicode, True) 'CipName, nvarchar[500], NOT NULL
        Return sSQL
    End Function

    Private Sub tsbImportData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbImportData.Click, mnsImportData.Click, tsmImportData.Click
        'Me.Cursor = Cursors.WaitCursor
        '.bSaved = False
        'Dim frm As New D80F2090
        'With frm
        '    .FormActive = "D80F2090"
        '    .FormPermission = "D02F2000"
        '    .ModuleID = "D02"
        '    .TransTypeID = "D02F2000" 'Theo TL phân tích
        '    .sFont = IIf(gbUnicode, "UNICODE", "VNI").ToString
        '    .ShowDialog()
        '    If .OutPut01 Then .bSaved = .OutPut01
        '    .Dispose()
        'End With
        'If .bSaved Then
        '    'Load lại dữ liệu
        '    LoadTDBGrid(True)
        'End If
        'Me.Cursor = Cursors.Default

        'Gọi form Import Data như sau:
        Me.Cursor = Cursors.WaitCursor
        If CallShowDialogD80F2090(D02, "D02F2000", "D02F2000") Then
            'Load lại dữ liệu
            LoadTDBGrid(True)
        End If
        Me.Cursor = Cursors.Default

    End Sub

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
        AddColVisible(tdbg, SPLIT1, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
        dtCaptionCols.Clear()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        'Incident 69247
        Dim f As New D02F2001
        With f
            .CipID = ""
            .CipNo = ""
            '.CIPAuto = D02Systems.CIPAuto
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If f.SavedOK Then LoadTDBGrid(True, .CipNo)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F2001
        With f
            .CipID = tdbg.Columns(COL_CipID).Text
            .CipNo = tdbg.Columns(COL_CipNo).Text
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        Dim f As New D02F2001
        With f
            .CipID = tdbg.Columns(COL_CipID).Text
            .CipNo = tdbg.Columns(COL_CipNo).Text
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
        End With
        If f.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_CipNo).Text)
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not CheckBeforeDelete() Then Exit Sub

        Dim sSQL As String = "Delete From D02T0100 Where CipID = " & SQLString(tdbg.Columns(COL_CipID).Text)
        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult = True Then
            'ExecuteAuditLog(sAuditCode, "03", tdbg.Columns(COL_CipNo).Text, tdbg.Columns(COL_CipName).Text)
            Lemon3.D91.RunAuditLog("02", sAuditCode, "03", tdbg.Columns(COL_CipNo).Text, tdbg.Columns(COL_CipName).Text)
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If

    End Sub

    Private Function CheckBeforeDelete() As Boolean
        If tdbg.Columns(COL_Status).Text <> "" Then
            If CInt(tdbg.Columns(COL_Status).Text) > 0 Then
                D99C0008.MsgCanNotDelete()
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
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

    Private Sub c1dateDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateDate.KeyDown
        'Fix: khi xóa giá trị sau đó nhấn TAB thì không giữ lại giá trị cũ
        Try
            If e.KeyCode = Keys.Tab Then
                'Chú ý: Nếu cột cuối cùng hiển thị là Date thì không cộng
                tdbg.Col = tdbg.Col + 1
                Exit Sub
            End If
        Catch ex As Exception
        End Try
    End Sub

#End Region

    Private Sub btnCodeID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCodeID.Click
        tdbg.Focus()
        VisibleColumns(1)
    End Sub

    Private Sub btnInformation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInformation.Click
        tdbg.Focus()
        VisibleColumns(2)
    End Sub

    Private Sub VisibleColumns(ByVal btn As Integer)
        Select Case btn
            Case 1
                btnInformation.Enabled = True
                btnCodeID.Enabled = False
                With tdbg.Splits(1).DisplayColumns
                    For i As Integer = COL_CIPNum01 To COL_CIPDate10
                        .Item(i).Visible = False
                    Next
                    For i As Integer = COL_X01ID To COL_X10ID
                        .Item(i).Visible = Convert.ToBoolean(tdbg.Columns(i).Tag)
                    Next
                End With
                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_CIPString01
            Case 2
                btnInformation.Enabled = False
                btnCodeID.Enabled = True
                With tdbg.Splits(1).DisplayColumns
                    For i As Integer = COL_CIPNum01 To COL_CIPDate10
                        .Item(i).Visible = Convert.ToBoolean(tdbg.Columns(i).Tag)
                    Next
                    For i As Integer = COL_X01ID To COL_X10ID
                        .Item(i).Visible = False
                    Next
                End With
                tdbg.Focus()
                tdbg.SplitIndex = SPLIT1
                tdbg.Col = COL_X01ID
        End Select
    End Sub
    Private ArrSpecVisiable(30) As Boolean
    Private Function LoadTDBGridInformationCaption(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_Spec01ID As Integer, ByVal Split As Integer, Optional ByVal IsVisibleColumn As Boolean = False, Optional ByVal bUnicode As Boolean = False) As Boolean
        Dim bUseSpec As Boolean = False
        Dim sSQL As String = SQLStoreD02P0015("D02T0100")
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        Dim iIndex As Integer = COL_Spec01ID
        Dim i As Integer
        If dt.Rows.Count > 0 Then
            For i = 0 To 29
                If (geLanguage = EnumLanguage.Vietnamese) Then
                    tdbg.Columns(iIndex).Caption = dt.Rows(i).Item("Data84").ToString
                Else
                    tdbg.Columns(iIndex).Caption = dt.Rows(i).Item("Data01").ToString
                End If
                tdbg.Columns(iIndex).Tag = (Convert.ToBoolean(dt.Rows(i).Item("Disabled")))
                If (L3Int(dt.Rows(i).Item("DataType").ToString) = 0) Then
                    tdbg.Columns(iIndex).NumberFormat = InsertFormat(L3Int(dt.Rows(i).Item("DecimalNum")))
                End If
                ArrSpecVisiable(iIndex - COL_Spec01ID) = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                If Not bUseSpec And Convert.ToBoolean(tdbg.Columns(iIndex).Tag) = True Then
                    bUseSpec = True
                End If
                tdbg.Splits(Split).DisplayColumns(iIndex).HeadingStyle.Font = FontUnicode(bUnicode, tdbg.Splits(Split).DisplayColumns(iIndex).HeadingStyle.Font.Style)
                If IsVisibleColumn Then ' Dùng cho lưới có nhiều nút
                    tdbg.Splits(Split).DisplayColumns(iIndex).Visible = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                End If
                iIndex += 1
            Next
        End If
        dt = Nothing
        Return bUseSpec
    End Function


    Private ArrSpecVisiable1(10) As Boolean
    Private Function LoadTDBGridTypeCodeID(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_X01 As Integer, ByVal Split As Integer, Optional ByVal IsVisibleColumn As Boolean = False, Optional ByVal bUnicode As Boolean = False) As Boolean
        Dim bUseSpec As Boolean = False
        Dim sSQL As String = "SELECT 	TypeCodeID,MaxLength, "
        If (geLanguage = EnumLanguage.Vietnamese) Then
            sSQL &= " VieTypeCodeNameU as Description, "
        Else
            sSQL &= " EngTypeCodeNameU as Description, "
        End If
        sSQL &= " Disabled"
        sSQL &= " FROM D02T0040 WITH(NOLOCK)"
        sSQL &= " WHERE 	Type = 'X' Order By TypeCodeID"
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        Dim iIndex As Integer = COL_X01
        Dim i As Integer
        If dt.Rows.Count > 0 Then
            For i = 0 To 9
                tdbg.Columns(iIndex).Caption = dt.Rows(i).Item("Description").ToString
                tdbg.Columns(iIndex).Tag = Not (Convert.ToBoolean(dt.Rows(i).Item("Disabled")))
                tdbg.Columns(iIndex).DataWidth = L3Int(dt.Rows(i).Item("MaxLength"))
                ArrSpecVisiable(iIndex - COL_X01) = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                If Not bUseSpec And Convert.ToBoolean(tdbg.Columns(iIndex).Tag) = True Then
                    bUseSpec = True
                End If

                If IsVisibleColumn Then ' Dùng cho lưới có nhiều nút
                    tdbg.Splits(Split).DisplayColumns(iIndex).Visible = Convert.ToBoolean(tdbg.Columns(iIndex).Tag)
                End If
                iIndex += 1
            Next
        End If
        dt = Nothing
        Return bUseSpec
    End Function

    Private Function InsertFormat(ByVal ONumber As Object) As String
        Dim iNumber As Int16 = Convert.ToInt16(ONumber)
        Dim sRet As String = "#,##0"
        If iNumber = 0 Then
        Else
            sRet &= "." & Strings.StrDup(iNumber, "0")
        End If
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0015
    '# Created User: Lê Đình Thái
    '# Created Date: 05/12/2011 02:46:31
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0015(ByVal sTable As String) As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0015 "
        sSQL &= SQLString(sTable) & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function
    Private Sub tdbg_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        '--- Chỉ cho nhập số
        Select Case tdbg.Col
            Case COL_CIPNum01, COL_CIPNum02, COL_CIPNum03, COL_CIPNum04, COL_CIPNum05, COL_CIPNum06, COL_CIPNum07, COL_CIPNum08, COL_CIPNum09, COL_CIPNum10
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_CIPDate01, COL_CIPDate02, COL_CIPDate03, COL_CIPDate04, COL_CIPDate05, COL_CIPDate06, COL_CIPDate07, COL_CIPDate08, COL_CIPDate09, COL_CIPDate10
                e.Handled = CheckKeyPress(e.KeyChar)
        End Select
    End Sub

    Private Sub tsmExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsmExportToExcel.Click, tsbExportToExcel.Click, mnsExportToExcel.Click
        'Lưới không có nút Hiển thị
        'Nếu lưới không có Group thì mở dòng code If dtCaptionCols Is Nothing Then
        'và truyền đối số cuối cùng là False vào hàm AddColVisible
        'If dtCaptionCols Is Nothing orelse dtCaptionCols.Rows.Count < 1 Then
        Dim arrColObligatory() As Integer = {}
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, arrColObligatory, , , gbUnicode)
        AddColVisible(tdbg, SPLIT1, Arr, arrColObligatory, , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        'Form trong DLL
        ''CallShowD99F2222(Me, ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable), dtFind, gsGroupColumns)'Nếu có sử dụng F12 cũ D09U1111
        CallShowD99F2222(Me, dtCaptionCols, dtGrid, gsGroupColumns)     
    End Sub



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AnchorForControl(EnumAnchorStyles.TopRight, btnCodeID, btnInformation)
        AnchorResizeColumnsGrid(EnumAnchorStyles.TopLeftRightBottom, tdbg)
        AnchorForControl(EnumAnchorStyles.BottomLeft, chkShowDisabled)
    End Sub
End Class