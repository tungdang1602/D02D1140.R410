'#-------------------------------------------------------------------------------------
'# Created Date: 11/10/2014 9:20:43 AM
'# Created User: HUỲNH KHANH
'# Modify Date: 11/10/2014 9:20:43 AM
'# Modify User: HUỲNH KHANH
'# Description: 
'#-------------------------------------------------------------------------------------

Imports System
Public Class D02F0005


    Private _savedOk As Boolean
    Public ReadOnly Property SavedOk() As Boolean
        Get
            Return _savedOk
        End Get
    End Property

    Private _isAddNew As Boolean = False 'xác nhận có chọn thêm mớitiêu thức
    Public ReadOnly Property IsAddNew As Boolean
        Get
            Return _isAddNew
        End Get
    End Property
    
    Private _keyString As String = ""
    Public ReadOnly Property KeyString() As String
        Get
            Return _keyString
        End Get
    End Property

    Private _lastKey As Integer = 0
    Public ReadOnly Property LastKey() As Integer
        Get
            Return _lastKey
        End Get
    End Property

    Private _isAutoNum As Boolean = False
    Public ReadOnly Property IsAutoNum() As Boolean
        Get
            Return _isAutoNum
        End Get
    End Property

    Private _iD As String = ""
    Public ReadOnly Property  ID() As String
        Get
            Return _iD
        End Get
    End Property

    Private _dtXCode(9) As String
    Public Property dtXCode() As String()
        Get
            Return _dtXCode
        End Get
        Set(ByVal Value As String())
            _dtXCode = Value
        End Set
    End Property
    Private _dtXCodeValue(9) As String
    Public Property dtXCodeValue() As String()
        Get
            Return _dtXCodeValue
        End Get
        Set(ByVal Value As String())
            _dtXCodeValue = Value
        End Set
    End Property

    Private _iGEMethodID As String = ""
    Public Property IGEMethodID() As String
        Get
            Return _iGEMethodID
        End Get
        Set(ByVal Value As String)
            _iGEMethodID = Value
        End Set
    End Property

    Private _keyName As String
    Public ReadOnly Property KeyName() As String
        Get
            Return _keyName
        End Get
    End Property

    Private _formID As String = ""
    Public WriteOnly Property FormID() As String
        Set(ByVal Value As String)
            _formID = Value
        End Set
    End Property

    Private Sub D02F0005_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Me.Cursor = Cursors.WaitCursor
        LoadLanguage()
        LoadTDBCombo()
        LoadEdit()
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rl3("Tao_ma_XDCB_tu_dongF") & " - " & Me.Name & UnicodeCaption(gbUnicode) 'TÁo mº XDCB tø ¢èng
        '================================================================ 
        lblIGEMethodID.Text = rl3("Phuong_phap_tao_ma") 'Phương pháp tạo mã
        lblCaption01.Text = rl3("Tieu_thuc") & " 1" 'Tiêu thức 1
        lblCaption03.Text = rl3("Tieu_thuc") & " 3" 'Tiêu thức 3
        lblCaption02.Text = rl3("Tieu_thuc") & " 2" 'Tiêu thức 2
        lblCaption04.Text = rl3("Tieu_thuc") & " 4" 'Tiêu thức 4
        lblCaption05.Text = rl3("Tieu_thuc") & " 5" 'Tiêu thức 5
        lblCaption06.Text = rl3("Tieu_thuc") & " 6" 'Tiêu thức 6
        lblCaption07.Text = rl3("Tieu_thuc") & " 7" 'Tiêu thức 7
        lblCaption08.Text = rl3("Tieu_thuc") & " 8" 'Tiêu thức 8
        lblCaption09.Text = rl3("Tieu_thuc") & " 9" 'Tiêu thức 9
        lblCaption10.Text = rl3("Tieu_thuc") & " 10" 'Tiêu thức 10
        '================================================================ 
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnSave.Text = rL3("_Tao_ma") '&Tạo mã
        '================================================================ 
        grpCriterions.Text = rl3("Tieu_thuc_tao_ma") 'Tiêu thức tạo mã
        '================================================================ 
        tdbcIGEMethodID.Columns("IGEMethodID").Caption = rl3("Ma") 'Mã
        tdbcIGEMethodID.Columns("IGEMethodName").Caption = rl3("Ten") 'Tên
        tdbcCaption10.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption10.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption09.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption09.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption08.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption08.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption07.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption07.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption06.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption06.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption05.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption05.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption04.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption04.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption02.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption02.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption03.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption03.Columns("SelectionName").Caption = rl3("Ten") 'Tên
        tdbcCaption01.Columns("SelectionID").Caption = rl3("Ma") 'Mã
        tdbcCaption01.Columns("SelectionName").Caption = rl3("Ten") 'Tên
    End Sub


    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcIGEMethodID
        sSQL = " --Combo PP tao ma" & vbCrLf
        sSQL &= "SELECT IGEMethodID, IGEMethodName" & UnicodeJoin(gbUnicode) & " As IGEMethodName, "
        sSQL &= " Defaults, FormID" & vbCrLf
        sSQL &= " FROM D91T0045 WITH(NOLOCK)" & vbCrLf
        sSQL &= " WHERE ModuleID = '02'"
        sSQL &= " And Disabled = 0"
        sSQL &= " And FormID ='D02F0060'"
        sSQL &= " ORDER BY 	IGEMethodID"
        Dim dtIGEMethodID As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(tdbcIGEMethodID, dtIGEMethodID, gbUnicode)
        If dtIGEMethodID.Rows.Count > 0 Then
            If _iGEMethodID <> "" Then
                tdbcIGEMethodID.SelectedValue = _iGEMethodID
            Else
                Dim dr() As DataRow = dtIGEMethodID.Select("Defaults = 1")
                If dr.Length > 0 Then
                    tdbcIGEMethodID.SelectedValue = dr(0).Item("IGEMethodID").ToString
                End If
            End If
        End If
    End Sub

    Private Sub tdbcIGEMethodID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.SelectedValueChanged
        If tdbcIGEMethodID.SelectedValue Is Nothing Then
            txtIGEMethodName.Text = ""
            dtD02P2005 = Nothing
        Else
            txtIGEMethodName.Text = tdbcIGEMethodID.Columns(1).Value.ToString
            LoadCaptionSelectionID()
        End If
    End Sub

    Dim dtD02P2005 As DataTable
    Private Sub LoadCaptionSelectionID()
        dtD02P2005 = ReturnDataTable(SQLStoreD02P2005)
        Dim i As Integer = 0, j As Integer = 1
        ReDim _dtXCode(9)
        Dim sCaption As String = ""
        For i = 0 To dtD02P2005.Rows.Count - 1
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.grpCriterions.Controls("tdbcCaption" & j.ToString("00")), C1.Win.C1List.C1Combo)
            With dtD02P2005.Rows(i)
                If .Item("Caption").ToString.Length > 25 Then
                    sCaption = (.Item("Caption").ToString).Substring(0, 25) & "..."
                Else
                    sCaption = .Item("Caption").ToString
                End If
                Me.grpCriterions.Controls("lblCaption" & j.ToString("00")).Text = sCaption
                Me.grpCriterions.Controls("lblCaption" & j.ToString("00")).Font = FontUnicode(gbUnicode, Me.grpCriterions.Controls("lblCaption" & j.ToString("00")).Font.Style)
                tdbc.Enabled = True
                'Me.grpCriterions.Controls("lblXCode" & j.ToString("00")).Text = (.Item("XCodeTypeID").ToString).Substring(1, .Item("XCodeTypeID").ToString.Length - 2)
                _dtXCode(i) = (.Item("XCodeTypeID").ToString).Substring(1, .Item("XCodeTypeID").ToString.Length - 2)
                tdbc.Tag = _dtXCode(i)
                LoadtdbcSelectionID(tdbc, (.Item("XCodeTypeID").ToString).Substring(1, .Item("XCodeTypeID").ToString.Length - 2))
                If .Item("DefaultValues").ToString <> "" Then tdbc.SelectedValue = .Item("DefaultValues")
            End With
        Next
        For i = i To 9
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.grpCriterions.Controls("tdbcCaption" & j.ToString("00")), C1.Win.C1List.C1Combo)
            tdbc.Enabled = False
            tdbc.SelectedValue = "-1"
        Next

    End Sub

    Private Sub GetValueSelectionID()
        Dim i As Integer = 0, j As Integer = 1
        ReDim _dtXCodeValue(9)
        For i = 0 To dtD02P2005.Rows.Count - 1
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.grpCriterions.Controls("tdbcCaption" & j.ToString("00")), C1.Win.C1List.C1Combo)
            _dtXCodeValue(i) = ReturnValueC1Combo(tdbc)
        Next
    End Sub

    Private Sub LoadEdit()
        Dim i As Integer = 0, j As Integer = 1
        For i = 0 To 10 - 1
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.grpCriterions.Controls("tdbcCaption" & j.ToString("00")), C1.Win.C1List.C1Combo)
            If _dtXCodeValue(i) <> "" Then
                tdbc.SelectedValue = _dtXCodeValue(i)
            End If

        Next
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2005
    '# Created User: HUỲNH KHANH
    '# Created Date: 09/10/2014 02:25:15
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2005() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Load caption dong" & vbCrlf)
        sSQL &= "Exec D02P2005 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcIGEMethodID)) 'IGEMethodID, varchar[50], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2007
    '# Created User: HUỲNH KHANH
    '# Created Date: 09/10/2014 02:37:15
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2007(ByVal xCode As String) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho combo tieu thuc loc" & vbCrLf)
        sSQL &= "Exec D02P2007 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(xCode) 'XCodeTypeID, varchar[50], NOT NULL
        Return sSQL
    End Function

    Private Sub LoadtdbcSelectionID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal xCode As String)
        LoadDataSource(tdbc, SQLStoreD02P2007(xCode), gbUnicode) ', gbUnicode
    End Sub

    Private Sub D02F0005_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me, True)
        End Select
    End Sub

    Private Sub D02F0005_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'If Not _savedOk Then
        '    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
        'End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Chặn lỗi khi đang vi phạm trên lưới mà nhấn Alt + L
        btnSave.Focus()
        If btnSave.Focused = False Then Exit Sub
        '************************************
        'If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        GetValueSelectionID() ' luôn đặt trước InsertSelectionID
        Dim sSQL As String = ""
        sSQL &= SQLDeleteD91T1000() & vbCrLf
        sSQL &= InsertSelectionID() & vbCrLf
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            sSQL = SQLStoreD91P1000()
            Dim dtD91P1000 As DataTable = ReturnDataTable(sSQL)
            If dtD91P1000.Rows.Count > 0 Then
                With dtD91P1000.Rows(0)
                    Select Case Convert.ToInt16(.Item("Status"))
                        Case 0
                            _iD = .Item("ID").ToString
                            _keyString = .Item("KeyString").ToString
                            _keyName = .Item("Name").ToString
                            _lastKey = L3Int(.Item("LastKey"))
                            _isAutoNum = L3Bool(.Item("IsAutoNum"))
                            _iGEMethodID = ReturnValueC1Combo(tdbcIGEMethodID)

                        Case 1
                            D99C0008.MsgL3(ConvertVietwareFToUnicode(.Item("Message").ToString))
                    End Select
                End With
            End If
            'SaveOK()
            _savedOk = True
            btnClose.Enabled = True
            btnClose.Focus()
            Me.Close()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Function AllowSave() As Boolean
        If tdbcIGEMethodID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Phuong_phap_tao_ma"))
            tdbcIGEMethodID.Focus()
            Return False
        End If
        If tdbcCaption01.Enabled AndAlso tdbcCaption01.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption01.Text), lblCaption01.Text).ToString)
            tdbcCaption01.Focus()
            Return False
        End If
        If tdbcCaption06.Enabled AndAlso tdbcCaption06.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption06.Text), lblCaption06.Text).ToString)
            tdbcCaption06.Focus()
            Return False
        End If
        If tdbcCaption02.Enabled AndAlso tdbcCaption02.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption02.Text), lblCaption02.Text).ToString)
            tdbcCaption02.Focus()
            Return False
        End If
        If tdbcCaption07.Enabled AndAlso tdbcCaption07.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption07.Text), lblCaption07.Text).ToString)
            tdbcCaption07.Focus()
            Return False
        End If
        If tdbcCaption03.Enabled AndAlso tdbcCaption03.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption03.Text), lblCaption03.Text).ToString)
            tdbcCaption03.Focus()
            Return False
        End If
        If tdbcCaption08.Enabled AndAlso tdbcCaption08.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption08.Text), lblCaption08.Text).ToString)
            tdbcCaption08.Focus()
            Return False
        End If
        If tdbcCaption04.Enabled AndAlso tdbcCaption04.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption04.Text), lblCaption04.Text).ToString)
            tdbcCaption04.Focus()
            Return False
        End If
        If tdbcCaption09.Enabled AndAlso tdbcCaption09.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption09.Text), lblCaption09.Text).ToString)
            tdbcCaption09.Focus()
            Return False
        End If
        If tdbcCaption05.Enabled AndAlso tdbcCaption05.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption05.Text), lblCaption05.Text).ToString)
            tdbcCaption05.Focus()
            Return False
        End If
        If tdbcCaption10.Enabled AndAlso tdbcCaption10.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(IIf(gbUnicode = False, ConvertVniToUnicode(lblCaption10.Text), lblCaption10.Text).ToString)
            tdbcCaption10.Focus()
            Return False
        End If
        Return True
    End Function


    Private Sub SetBackColorObligatory()
        tdbcIGEMethodID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption01.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption06.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption02.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption07.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption03.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption08.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption04.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption09.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption05.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCaption10.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 09/10/2014 03:42:27
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xoa bang tam" & vbCrlf)
        sSQL &= "Delete From D91T1000"
        sSQL &= " Where "
        sSQL &= "UserID = " & SQLString(gsUserID) & " And "
        sSQL &= "HostID = " & SQLString(My.Computer.Name) & " And "
        sSQL &= "FormID = " & SQLString("D02F0060") & " And "
        sSQL &= "ModuleID = " & SQLString("02")
        Return sSQL
    End Function

    Private Function InsertSelectionID() As String
        Dim sSQL As String = ""
        sSQL &= "--Insert vao bang tam" & vbCrLf
        Dim i As Integer = 0, j As Integer = 1
        For i = 0 To 10
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.grpCriterions.Controls("tdbcCaption" & j.ToString("00")), C1.Win.C1List.C1Combo)
            If tdbc.Enabled Then
                sSQL &= SQLInsertD91T1000(tdbc, "{" & _dtXCode(i) & "}").ToString & vbCrLf
                'Else
                '    sSQL &= SQLInsertD91T1000(tdbc, "").ToString & vbCrLf
            End If
        Next
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 09/10/2014 03:44:11
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1000(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal xCode As String) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Insert bang tam" & vbCrLf)
        sSQL.Append("Insert Into D91T1000(")
        sSQL.Append("UserID, HostID, ModuleID, FormID, Code, ")
        sSQL.Append("SelectionID, SelectionNameU")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[50], NOT NULL
        sSQL.Append(SQLString(My.Computer.Name) & COMMA) 'HostID, varchar[50], NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[50], NOT NULL
        sSQL.Append(SQLString("D02F0060") & COMMA) 'FormID, varchar[50], NOT NULL
        sSQL.Append(SQLString(xCode) & COMMA) 'Code, varchar[500], NOT NULL
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbc)) & COMMA) 'SelectionID, varchar[50], NOT NULL
        sSQL.Append(SQLStringUnicode(ReturnValueC1Combo(tdbc, "SelectionName"), gbUnicode, True)) 'SelectionNameU, nvarchar[1000], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P1000
    '# Created User: HUỲNH KHANH
    '# Created Date: 09/10/2014 04:17:05
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P1000() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Thuc thi store D91P1000" & vbCrlf)
        sSQL &= "Exec D91P1000 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString("D02F0060") & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcIGEMethodID)) & COMMA 'IGEMethodID, varchar[50], NOT NULL
        sSQL &= SQLNumber("0") & COMMA 'AutoCreateName, tinyint, NOT NULL
        sSQL &= SQLNumber("50") & COMMA 'Length, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber("0") & COMMA 'IsD07F0011, tinyint, NOT NULL
        sSQL &= SQLNumber("0") 'NewLastKey, int, NOT NULL
        Return sSQL
    End Function


#Region "Events tdbcIGEMethodID with txtIGEMethodName"

    Private Sub tdbcIGEMethodID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcIGEMethodID.LostFocus
        If tdbcIGEMethodID.FindStringExact(tdbcIGEMethodID.Text) = -1 Then
            tdbcIGEMethodID.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcCaption01"

    Private Sub tdbcCaption01_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption01.LostFocus
        If tdbcCaption01.FindStringExact(tdbcCaption01.Text) = -1 Then tdbcCaption01.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption02"

    Private Sub tdbcCaption02_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption02.LostFocus
        If tdbcCaption02.FindStringExact(tdbcCaption02.Text) = -1 Then tdbcCaption02.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption03"

    Private Sub tdbcCaption03_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption03.LostFocus
        If tdbcCaption03.FindStringExact(tdbcCaption03.Text) = -1 Then tdbcCaption03.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption04"

    Private Sub tdbcCaption04_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption04.LostFocus
        If tdbcCaption04.FindStringExact(tdbcCaption04.Text) = -1 Then tdbcCaption04.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption05"

    Private Sub tdbcCaption05_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption05.LostFocus
        If tdbcCaption05.FindStringExact(tdbcCaption05.Text) = -1 Then tdbcCaption05.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption06"

    Private Sub tdbcCaption06_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption06.LostFocus
        If tdbcCaption06.FindStringExact(tdbcCaption06.Text) = -1 Then tdbcCaption06.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption07"

    Private Sub tdbcCaption07_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption07.LostFocus
        If tdbcCaption07.FindStringExact(tdbcCaption07.Text) = -1 Then tdbcCaption07.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption08"

    Private Sub tdbcCaption08_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption08.LostFocus
        If tdbcCaption08.FindStringExact(tdbcCaption08.Text) = -1 Then tdbcCaption08.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption09"

    Private Sub tdbcCaption09_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption09.LostFocus
        If tdbcCaption09.FindStringExact(tdbcCaption09.Text) = -1 Then tdbcCaption09.Text = ""
    End Sub

#End Region

#Region "Events tdbcCaption10"

    Private Sub tdbcCaption10_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCaption10.LostFocus
        If tdbcCaption10.FindStringExact(tdbcCaption10.Text) = -1 Then tdbcCaption10.Text = ""
    End Sub

#End Region


    Private Sub tdbcCaption01_SelectedValueChanged(sender As Object, e As EventArgs) Handles tdbcCaption01.SelectedValueChanged, tdbcCaption02.SelectedValueChanged, tdbcCaption03.SelectedValueChanged, tdbcCaption04.SelectedValueChanged, tdbcCaption05.SelectedValueChanged, tdbcCaption06.SelectedValueChanged, tdbcCaption07.SelectedValueChanged, tdbcCaption08.SelectedValueChanged, tdbcCaption09.SelectedValueChanged, tdbcCaption10.SelectedValueChanged
        '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        If tdbc.Text = "+" Then
            AddNewD02F0052(tdbc)
        End If
    End Sub

    '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
    Private Sub AddNewD02F0052(tdbc As C1.Win.C1List.C1Combo)
        If ReturnPermission("D02F0051") <= 1 Then
            D99C0008.MsgL3(rL3("Ban_khong_co_quyen_them_moi"), L3MessageBoxIcon.Information)
            tdbc.Text = ""
            Exit Sub
        End If
        Dim sValue As String = ""
        Dim frm As New D02F0052
        With frm
            .FormCall = "D02F0005"
            .TypeCodeID = ReturnValueC1Combo(tdbc, "TypeCodeID")
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If frm.SavedOK Then
                _isAddNew = True
                sValue = .ACodeID
            End If
            .Dispose()
        End With

        If sValue <> "" Then LoadtdbcSelectionID(tdbc, tdbc.Tag.ToString)
        tdbc.Text = sValue
    End Sub

End Class