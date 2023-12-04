Imports System
Imports System.Text
Public Class D02F0052


    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Dim dt1 As DataTable

    Private _formCall As String = "" '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
    Public WriteOnly Property FormCall As String
        Set(ByVal Value As String)
            _formCall = Value
        End Set
    End Property
    
    Private _TypeCodeID As String = ""
    Public Property TypeCodeID() As String
        Get
            Return _TypeCodeID
        End Get
        Set(ByVal value As String)
            _TypeCodeID = value
        End Set
    End Property
    Private _ACodeID As String = ""
    Public Property ACodeID() As String
        Get
            Return _ACodeID
        End Get
        Set(ByVal value As String)
            _ACodeID = value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
	Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
	bLoadFormState = True
	LoadInfoGeneral()
            _FormState = value
            LoadTDBC()
            tdbcTypeCodeID.Text = _TypeCodeID
            If tdbcTypeCodeID.Text = "%" Then
                tdbcTypeCodeID.Text = ""

            End If
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    If _formCall = "D02F0005" Then '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
                        tdbcTypeCodeID.SelectedValue = _TypeCodeID
                        tdbcTypeCodeID.Enabled = False
                        chkDisabled.Visible = False
                    End If
                Case EnumFormState.FormEdit
                    tdbcTypeCodeID.Enabled = False
                    txtACodeID.Enabled = False
                    btnNext.Visible = False
                    btnSave.Enabled = True
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                Case EnumFormState.FormView
                    tdbcTypeCodeID.Enabled = False
                    txtACodeID.Enabled = False
                    LoadEdit()
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left

            End Select
        End Set
    End Property

    Private Sub LoadTDBC()
        Dim sSQl As String = ""
        If geLanguage = EnumLanguage.Vietnamese Then
            sSQl = "Select TypeCodeID, MaxLength, VieTypeCodeName" & UnicodeJoin(gbUnicode) & " As Description  " & vbCrLf
            sSQl &= "From D02T0040 WITH(NOLOCK) Where  Type='X' And Disabled = 0 Order By TypeCodeID"
        Else
            sSQl = "Select TypeCodeID, MaxLength, EngTypeCodeName" & UnicodeJoin(gbUnicode) & " As Description  " & vbCrLf
            sSQl &= "From D02T0040 WITH(NOLOCK) Where  Type='X' And Disabled = 0 Order By TypeCodeID"
        End If
        dt1 = ReturnDataTable(sSQl)
        LoadDataSource(tdbcTypeCodeID, dt1, gbUnicode)

    End Sub

#Region "Events tdbcTypeCodeID with txtDescription"

    Private Sub tdbcTypeCodeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.Close
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 Then
            tdbcTypeCodeID.Text = ""
            txtDescription_Type.Text = ""
        End If
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        txtDescription_Type.Text = tdbcTypeCodeID.Columns("Description").Value.ToString
        txtACodeID.MaxLength = CInt(tdbcTypeCodeID.Columns("MaxLength").Text)
    End Sub

    Private Sub tdbcTypeCodeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcTypeCodeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcTypeCodeID.Text = ""
            txtDescription_Type.Text = ""
        End If
    End Sub

#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0041
    '# Created User: Mỹ Vân
    '# Created Date: 23/10/2007 09:48:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0041() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0041(")
        sSQL.Append("ACodeID, TypeCodeID, Type, DescriptionU, Disabled, ")
        sSQL.Append("CreateDate, CreateUserID, LastModifyDate, LastModifyUserID")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtACodeID.Text) & COMMA) 'ACodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcTypeCodeID.Text) & COMMA) 'TypeCodeID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString("X") & COMMA) 'Type, varchar[1], NULL
        sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID)) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append(")")

        Return sSQL
    End Function
    
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0041
    '# Created User: Mỹ Vân
    '# Created Date: 23/10/2007 11:08:21
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0041() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0041 Set ")
		
        sSQL.Append("Type = " & SQLString("X") & COMMA) 'varchar[1], NULL
        sSQL.Append("DescriptionU = " & SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
        sSQL.Append(" Where ")
        sSQL.Append("ACodeID = " & SQLString(txtACodeID.Text) & " And ")
        sSQL.Append("TypeCodeID = " & SQLString(tdbcTypeCodeID.Text))

        Return sSQL
    End Function


    Private Function AllowSave() As Boolean
        If tdbcTypeCodeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Ma_loai_phan_tich"))
            tdbcTypeCodeID.Focus()
            Return False
        End If
        If txtACodeID.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ma_khoan_muc"))
            txtACodeID.Focus()
            Return False
        End If
        If _FormState = EnumFormState.FormAdd Then
            Dim s As String = ""
            s = "select Top 1 1 from D02T0041 WITH(NOLOCK) where AcodeID= " & SQLString(txtACodeID.Text) & vbCrLf
            s &= "and TypeCodeID= " & SQLString(tdbcTypeCodeID.Text)
            If ExistRecord(s) Then
                D99C0008.MsgDuplicatePKey()
                txtACodeID.Focus()
                Return False
            End If
        End If
        Return True
    End Function


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub
        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLInsertD02T0041.ToString() & vbCrLf)
            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T0041.ToString())
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _ACodeID = txtACodeID.Text
                    btnNext.Enabled = True
                    btnNext.Focus()

                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Sub LoadEdit()
        Dim sSQL As String = ""
        sSQL = "select * from D02T0041 WITH(NOLOCK) where ACodeID= " & SQLString(_ACodeID) & vbCrLf
        sSQL &= "and  TypeCodeID= " & SQLString(_TypeCodeID) & vbCrLf
        sSQL &= "and Type='X'"
        dt1 = ReturnDataTable(sSQL)
        If dt1.Rows.Count > 0 Then
            tdbcTypeCodeID.Text = dt1.Rows(0).Item("TypeCodeID").ToString
            txtACodeID.Text = dt1.Rows(0).Item("ACodeID").ToString
            chkDisabled.Checked = Convert.ToBoolean(dt1.Rows(0).Item("Disabled"))
            txtDescription.Text = dt1.Rows(0).Item("Description" & UnicodeJoin(gbUnicode)).ToString
        End If
    End Sub

    Private Sub D02F0052_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
            Exit Sub
        End If
    End Sub

    Private Sub D02F0052_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()
        InputbyUnicode(Me, gbUnicode)
        CheckIdTextBox(txtACodeID)
        SetResolutionForm(Me)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_ma_phan_tich_XDCB_-_D02F0052") & UnicodeCaption(gbUnicode) ' CËp nhËt mº ph¡n tÛch XDCB - D02F0052
        '================================================================ 
        lblTypeCodeID.Text = rl3("Ma_loai_phan_tich") 'Mã loại phân tích
        lblACodeID.Text = rl3("Ma_khoan_muc") 'Mã khoản mục
        lblDescription.Text = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        '================================================================ 
        tdbcTypeCodeID.Columns("TypeCodeID").Caption = rl3("Ma") 'Mã
        tdbcTypeCodeID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcTypeCodeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtACodeID.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnNext.Enabled = False
        btnSave.Enabled = True
        txtACodeID.Text = ""
        chkDisabled.Checked = False
        txtDescription.Text = ""
        txtACodeID.Focus()
    End Sub

End Class