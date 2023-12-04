'#-------------------------------------------------------------------------------------
'# Created Date: 04/09/2007 11:02:55 AM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 04/09/2007 11:02:55 AM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F2001

    Private _savedOK As Boolean = False
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private sID As String = ""
    Private sKeyString As String = ""
    Private sLastKey As Integer = 0
    Private bIsAutoNum As Boolean = True
    Private dtXCodeValue(9) As String
    Private dtXCode(9) As String
    Private sKeyName As String

    Dim clsFilterCombo As Lemon3.Controls.FilterCombo

    Private _cipID As String = ""
    Public Property CipID() As String
        Get
            Return _cipID
        End Get
        Set(ByVal value As String)
            _cipID = value
        End Set
    End Property

    Private _cipNo As String = ""
    Public Property CipNo() As String
        Get
            Return _cipNo
        End Get
        Set(ByVal value As String)
            _cipNo = value
        End Set
    End Property

    Private _auditCode As String = "CIP"
    Public WriteOnly Property sAuditCode() As String
        Set(ByVal value As String)
            _auditCode = value
        End Set
    End Property

    Private _iGEMethodID As String = ""
    Public WriteOnly Property IGEMethodID() As String
        Set(ByVal Value As String)
            _iGEMethodID = Value
        End Set
    End Property

    Private _cIPAuto As Integer
    Private _d54ProjectID As String = ""
    Public WriteOnly Property D54ProjectID() As String
        Set(ByVal Value As String)
            _d54ProjectID = Value
        End Set
    End Property

    Private _d27PropertyProductID As String = ""
    Public WriteOnly Property D27PropertyProductID() As String
        Set(ByVal Value As String)
            _d27PropertyProductID = Value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            bLoadFormState = True
            LoadInfoGeneral()
            _CIPOutputOrder = D02Systems.CIPOutputOrder
            _CIPOutputLength = D02Systems.CIPOutputLength
            _CIPSeperated = D02Systems.CIPSeperated
            _CIPSeperator = D02Systems.CIPSeperator
            _cIPAuto = D02Systems.CIPAuto
            _FormState = value

            '26/6/2017, Phạm Thị Thu: id 97566-Bổ sung điều kiện tìm kiếm tại Danh mục / Mã XDCB
            clsFilterCombo = New Lemon3.Controls.FilterCombo
            clsFilterCombo.CheckD91 = True
            clsFilterCombo.AddPairObject(tdbcContractorOTID, tdbcContractorID)
            clsFilterCombo.AddPairObject(tdbcSupplierOTID, tdbcSupplierID)
            clsFilterCombo.UseFilterComboObjectID()

            LoadCaption()
            LoadTDBCombo()
            GetAutoCIPInfo()
            LoadLabelACode()
            cboStatus.Enabled = False
            Select Case _FormState
                Case EnumFormState.FormAdd
                    If _cIPAuto = 1 Then
                        Dim bIsAddNew As Boolean = False
                        Dim f As New D02F0005
                        With f
                            .ShowInTaskbar = True
                            .ShowDialog()
                            If .SavedOk Then
                                sID = f.ID
                                sKeyString = f.KeyString
                                sKeyName = f.KeyName
                                sLastKey = f.LastKey
                                bIsAutoNum = f.IsAutoNum
                                _iGEMethodID = f.IGEMethodID
                                dtXCodeValue = f.dtXCodeValue
                                dtXCode = f.dtXCode
                                '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
                                bIsAddNew = .IsAddNew
                                If bIsAddNew Then LoadTDBCACodeID()

                                LoadDefaultTDBCACodeID() ' Default giá trị được lấy từ
                            End If
                            .Dispose()
                        End With
                    End If
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    LoadAdd()
                Case EnumFormState.FormEdit
                    tdbcCIPS1ID.Enabled = False
                    tdbcCIPS2ID.Enabled = False
                    tdbcCIPS3ID.Enabled = False
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                    'VisibleIGEMethodID()
                Case EnumFormState.FormView
                    tdbcCIPS1ID.Enabled = False
                    tdbcCIPS2ID.Enabled = False
                    tdbcCIPS3ID.Enabled = False
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                    'VisibleIGEMethodID()
            End Select
        End Set
    End Property

    Private Sub LoadDefaultTDBCACodeID()
        Dim i As Integer = 0, j As Integer = 1
        For i = 0 To 10
            j = i + 1
            If j > 10 Then Exit For
            Dim tdbc As C1.Win.C1List.C1Combo = CType(Me.TabPage2.Controls("tdbcACodeID" & j.ToString), C1.Win.C1List.C1Combo)
            tdbc.SelectedValue = "-1"
            If dtXCodeValue(i) IsNot Nothing Then
                For k As Integer = 0 To dtXCodeValue.Length - 1
                    If dtXCode(k) IsNot Nothing AndAlso dtXCode(k).Contains(j.ToString("00")) Then
                        tdbc.SelectedValue = dtXCodeValue(k)
                    End If
                Next

            End If
        Next
        If sID <> "" Then
            txtCipNo.Text = sID
            ReadOnlyControl(txtCipNo)
        Else
            txtCipNo.Text = ""
            UnReadOnlyControl(True, txtCipNo)
        End If

        If sKeyName <> "" Then
            txtCipName.Text = sKeyName
        Else
            txtCipName.Text = ""
        End If
    End Sub


    Private dtLabel As DataTable
    Dim dtObject As DataTable

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F2001_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If _savedOK Then
            D99C0007.SaveOthersSetting(D02, EXECHILD, "Output01", "1")
            D99C0007.SaveOthersSetting(D02, EXECHILD, "Output02", _cipID)
            D99C0007.SaveOthersSetting(D02, EXECHILD, "Output03", txtCipNo.Text)
        End If
    End Sub

    Private Sub D02F2001_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (txtCipNo.Text <> "" Or txtCipName.Text <> "" Or tdbcAccountID.Text <> "") Then
                If Not _savedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F2001_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        ElseIf e.Alt = True And (e.KeyCode = Keys.D1 Or e.KeyCode = Keys.NumPad1) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage1
            Application.DoEvents()
            If D02Systems.CIPAuto = 2 Then
                tdbcCIPS1ID.Focus()
            Else
                txtCipNo.Focus() 'txtCipNo.Focus()
            End If
        ElseIf e.Alt = True And (e.KeyCode = Keys.D2 Or e.KeyCode = Keys.NumPad2) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage2
            Application.DoEvents()
            tdbcACodeID1.Focus()
        ElseIf e.Alt = True And (e.KeyCode = Keys.D3 Or e.KeyCode = Keys.NumPad3) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage3
            txtFAString01.Focus()
            Application.DoEvents()
        ElseIf e.KeyCode = Keys.F5 Then
            Me.btnF5_Click(Nothing, Nothing)
        End If
    End Sub

    'Private Sub VisibleIGEMethodID()
    '    ' ẩn khí xem sửa
    '    tdbcIGEMethodID.Visible = False
    '    lblIGEMethodID.Visible = False
    '    lblCipNo.Left = lblIGEMethodID.Left
    '    txtCipNo.Left = txtDescription.Left
    '    txtCipNo.Width = tdbcIGEMethodID.Width
    'End Sub

    Private Sub D02F2001_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()
        CheckIdTextBox(txtCipNo)
        InputbyUnicode(Me, gbUnicode)
        If _d27PropertyProductID <> "" Then 'Incident 71699 trường hợp thêm sửa được gọi từ D27F1514
            txtProjectID.Text = _d54ProjectID
            txtPropertyProductID.Text = _d27PropertyProductID
        End If
        SetResolutionForm(Me)
    End Sub

    Private Sub LoadAdd()
        btnK.Enabled = _cIPAuto = 1 And bIsAutoNum And _FormState = EnumFormState.FormAdd
        btnF5.Enabled = _cIPAuto = 1 And _FormState = EnumFormState.FormAdd
        chkDisabled.Checked = False
        tdbcAccountID.Text = ""
        cboStatus.SelectedIndex = 0
        txtDescription.Text = ""
        tdbcContractorOTID.SelectedValue = ""
        LoadtdbcContractorID("-1")
        tdbcContractorID.SelectedValue = ""
        tdbcSupplierOTID.SelectedValue = ""
        LoadtdbcSupplierID("-1")
        tdbcSupplierID.SelectedValue = ""
        txtCipCost.Text = ""
        LoadDefault()
        If _FormState = EnumFormState.FormAdd Then
            If _d27PropertyProductID <> "" Then 'Incident 71699 trường hợp thêm sửa được gọi từ D27F1514
                lblProjectID.Visible = True
                txtPropertyProductID.Visible = True
            Else
                lblProjectID.Visible = False
                txtPropertyProductID.Visible = False
            End If
        End If
        
    End Sub

    Private Sub LoadMaster()
        Dim sSQL As String = ""
        If _d27PropertyProductID <> "" Then
            _cipID = ReturnScalar("Select cipId from D02T0100 where D27PropertyProductID= " & SQLString(_d27PropertyProductID) & "and D54ProjectID = " & SQLString(_d54ProjectID))
        End If
        sSQL &= "Select * "
        sSQL &= " From D02T0100 WITH(NOLOCK) Where CipID = " & SQLString(_cipID)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                txtCipNo.Text = .Item("CipNo").ToString
                txtCipName.Text = .Item("CipName" & UnicodeJoin(gbUnicode)).ToString
                tdbcAccountID.Text = .Item("AccountID").ToString
                If Not IsDBNull(.Item("StartDate")) Then c1dateStartDate.Value = CDate(.Item("StartDate"))
                If Not IsDBNull(.Item("CompletionDate")) Then c1dateCompletionDate.Value = CDate(.Item("CompletionDate"))
                If CInt(.Item("Status")) = 0 Then
                    cboStatus.SelectedIndex = 0
                ElseIf CInt(.Item("Status")) = 1 Then
                    cboStatus.SelectedIndex = 1
                    tdbcAccountID.Enabled = False
                ElseIf CInt(.Item("Status")) = 2 Then
                    cboStatus.SelectedIndex = 2
                    tdbcAccountID.Enabled = False
                End If
                txtDescription.Text = .Item("Description" & UnicodeJoin(gbUnicode)).ToString
                tdbcACodeID1.Text = .Item("X01ID").ToString
                tdbcACodeID2.Text = .Item("X02ID").ToString
                tdbcACodeID3.Text = .Item("X03ID").ToString
                tdbcACodeID4.Text = .Item("X04ID").ToString
                tdbcACodeID5.Text = .Item("X05ID").ToString
                tdbcACodeID6.Text = .Item("X06ID").ToString
                tdbcACodeID7.Text = .Item("X07ID").ToString
                tdbcACodeID8.Text = .Item("X08ID").ToString
                tdbcACodeID9.Text = .Item("X09ID").ToString
                tdbcACodeID10.Text = .Item("X10ID").ToString
                chkDisabled.Checked = Convert.ToBoolean(IIf(.Item("Disabled").ToString = "1", True, False))

                tdbcContractorOTID.SelectedValue = .Item("ContractorOTID").ToString
                LoadtdbcContractorID(tdbcContractorOTID.Text)
                tdbcContractorID.SelectedValue = .Item("ContractorID").ToString
                tdbcSupplierOTID.SelectedValue = .Item("SupplierOTID").ToString
                LoadtdbcSupplierID(tdbcSupplierOTID.Text)
                tdbcSupplierID.SelectedValue = .Item("SupplierID").ToString
                txtCipCost.Text = SQLNumber(.Item("CipCost").ToString, DxxFormat.DecimalPlaces)
                c1dateExpStartDate.Value = .Item("ExpStartDate").ToString
                c1dateExpEndDate.Value = .Item("ExpEndDate").ToString

                txtFAString01.Text = .Item("CIPString01" & UnicodeJoin(gbUnicode)).ToString
                txtFAString02.Text = .Item("CIPString02" & UnicodeJoin(gbUnicode)).ToString
                txtFAString03.Text = .Item("CIPString03" & UnicodeJoin(gbUnicode)).ToString
                txtFAString04.Text = .Item("CIPString04" & UnicodeJoin(gbUnicode)).ToString
                txtFAString05.Text = .Item("CIPString05" & UnicodeJoin(gbUnicode)).ToString
                txtFAString06.Text = .Item("CIPString06" & UnicodeJoin(gbUnicode)).ToString
                txtFAString07.Text = .Item("CIPString07" & UnicodeJoin(gbUnicode)).ToString
                txtFAString08.Text = .Item("CIPString08" & UnicodeJoin(gbUnicode)).ToString
                txtFAString09.Text = .Item("CIPString09" & UnicodeJoin(gbUnicode)).ToString
                txtFAString10.Text = .Item("CIPString10" & UnicodeJoin(gbUnicode)).ToString
                txtFANum01.Text = Double.Parse((.Item("CIPNum01").ToString())).ToString
                txtFANum02.Text = Double.Parse((.Item("CIPNum02").ToString())).ToString
                txtFANum03.Text = Double.Parse((.Item("CIPNum03").ToString())).ToString
                txtFANum04.Text = Double.Parse((.Item("CIPNum04").ToString())).ToString
                txtFANum05.Text = Double.Parse((.Item("CIPNum05").ToString())).ToString
                txtFANum06.Text = Double.Parse((.Item("CIPNum06").ToString())).ToString
                txtFANum07.Text = Double.Parse((.Item("CIPNum07").ToString())).ToString
                txtFANum08.Text = Double.Parse((.Item("CIPNum08").ToString())).ToString
                txtFANum09.Text = Double.Parse((.Item("CIPNum09").ToString())).ToString
                txtFANum10.Text = Double.Parse((.Item("CIPNum10").ToString())).ToString
                txtFANum01.Text = SQLNumber(txtFANum01.Text, InsertFormat(txtFANum01.Tag))
                txtFANum02.Text = SQLNumber(txtFANum02.Text, InsertFormat(txtFANum02.Tag))
                txtFANum03.Text = SQLNumber(txtFANum03.Text, InsertFormat(txtFANum03.Tag))
                txtFANum04.Text = SQLNumber(txtFANum04.Text, InsertFormat(txtFANum04.Tag))
                txtFANum05.Text = SQLNumber(txtFANum05.Text, InsertFormat(txtFANum05.Tag))
                txtFANum06.Text = SQLNumber(txtFANum06.Text, InsertFormat(txtFANum06.Tag))
                txtFANum07.Text = SQLNumber(txtFANum07.Text, InsertFormat(txtFANum07.Tag))
                txtFANum08.Text = SQLNumber(txtFANum08.Text, InsertFormat(txtFANum08.Tag))
                txtFANum09.Text = SQLNumber(txtFANum09.Text, InsertFormat(txtFANum09.Tag))
                txtFANum10.Text = SQLNumber(txtFANum10.Text, InsertFormat(txtFANum10.Tag))
                c1dateFADate01.Value = .Item("CIPDate01").ToString()
                c1dateFADate02.Value = .Item("CIPDate02").ToString()
                c1dateFADate03.Value = .Item("CIPDate03").ToString()
                c1dateFADate04.Value = .Item("CIPDate04").ToString()
                c1dateFADate05.Value = .Item("CIPDate05").ToString()
                c1dateFADate06.Value = .Item("CIPDate06").ToString()
                c1dateFADate07.Value = .Item("CIPDate07").ToString()
                c1dateFADate08.Value = .Item("CIPDate08").ToString()
                c1dateFADate09.Value = .Item("CIPDate09").ToString()
                c1dateFADate10.Value = .Item("CIPDate10").ToString()
                tdbcCIPEmployeeID.Text = .Item("CIPEmployeeID").ToString()
                tdbcCIPObjectTypeID.Text = .Item("CIPObjectTypeID").ToString()
                tdbcCIPObjectID.Text = .Item("CIPObjectID").ToString()
                If _d27PropertyProductID = "" Then 'trường hợp sửa được gọi từ D02
                    txtProjectID.Text = .Item("D54ProjectID").ToString()
                    txtPropertyProductID.Text = .Item("D27PropertyProductID").ToString()
                End If
            End With
        End If
    End Sub

    Private Sub LoadEdit()
        txtCipNo.Enabled = False
        btnK.Enabled = False
        btnF5.Enabled = False
        LoadMaster()
    End Sub

    Private Sub LoadDefault()
        c1dateStartDate.Value = Date.Today
        c1dateCompletionDate.Value = Date.Today
        c1dateExpStartDate.Value = Date.Today
        c1dateExpEndDate.Value = Date.Today
        'gán cặc định
        'Dim dr() As DataRow = dtIGEMethodID.Select("Defaults = 1", dtIGEMethodID.DefaultView.Sort)
        'If dr.Length > 0 Then
        '    tdbcIGEMethodID.SelectedValue = dr(0).Item("IGEMethodID").ToString
        'End If
    End Sub

    Private Sub SetBackColorObligatory()
        txtCipNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtCipName.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private dtObjectID As DataTable
    Private dtIGEMethodID As DataTable

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcIGEMethodID
        'sSQL = " --Combo PP tao ma" & vbCrLf
        'sSQL &= "SELECT IGEMethodID, IGEMethodName" & UnicodeJoin(gbUnicode) & " As IGEMethodName, "
        'sSQL &= " Defaults, FormID" & vbCrLf
        'sSQL &= " FROM D91T0045 WITH(NOLOCK)" & vbCrLf
        'sSQL &= " WHERE 		ModuleID = '02'"
        'sSQL &= " And Disabled = 0"
        'sSQL &= " And FormID ='D02F0060'"
        'sSQL &= " And (DivisionID = " & SQLString(gsDivisionID) & " Or DivisionID = '' )" & vbCrLf
        'sSQL &= " ORDER BY 	IGEMethodID"
        'dtIGEMethodID = ReturnDataTable(sSQL)
        'LoadDataSource(tdbcIGEMethodID, dtIGEMethodID, gbUnicode)

        'Load tdbcAccountID
        LoadAccountID(tdbcAccountID, "AccountStatus = 0 And GroupID = '9'", gbUnicode)

        'Load tdbcACodeID
        LoadTDBCACodeID()

        cboStatus.Items.Add(rL3("Moi_thiet_lap"))
        cboStatus.Items.Add(rL3("Dang_tap_hop"))
        cboStatus.Items.Add(rL3("Da_ket_thuc"))

        'Load tdbcContractorOTID
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " As ObjectName, ObjectTypeID From Object WITH(NOLOCK) Where Disabled = 0"
        dtObject = ReturnDataTable(sSQL)
        LoadObjectTypeID(tdbcContractorOTID, gbUnicode)

        'Load tdbcContractorID
        LoadtdbcContractorID("-1")

        'Load tdbcSupplierOTID
        LoadObjectTypeID(tdbcSupplierOTID, gbUnicode)

        'Load tdbcSupplierID
        LoadtdbcSupplierID("-1")
        'Load tdbcObjectTypeID
        Dim dtObjectTypeID As DataTable = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcCIPObjectTypeID, dtObjectTypeID, gbUnicode)
        'Load tdbcEmployeeID
        sSQL = "Select ObjectID as EmployeeID, ObjectName" & UnicodeJoin(gbUnicode) & " As EmployeeName From Object WITH(NOLOCK) Where ObjectTypeID = 'NV' Order by ObjectID"
        LoadDataSource(tdbcCIPEmployeeID, sSQL, gbUnicode)
        'Load tdbdOjectID
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " As ObjectName, ObjectTypeID From Object WITH(NOLOCK) Where Disabled = 0  order by ObjectID" ' and ObjectTypeID=" & SQLString(ID)
        dtObjectID = ReturnDataTable(sSQL)

        'Load tdbcAssetS1ID
        LoadTdbcCIP1()
        'Load tdbcAssetS2ID
        LoadTdbcCIP2()
        'Load tdbcAssetS3ID
        LoadTdbcCIP3()
    End Sub

    Private Sub LoadTDBCACodeID()
        'Load tdbcACodeID
        Dim sSQL As String = ""
        Dim dtACodeID As DataTable
        sSQL = "Select ACodeID, Description" & UnicodeJoin(gbUnicode) & " As Description, TypeCodeID From D02T0041 WITH(NOLOCK) Where Disabled = 0 Order By ACodeID"
        dtACodeID = ReturnDataTable(sSQL)
        LoadDataSource(tdbcACodeID1, ReturnTableFilter(dtACodeID, "TypeCodeID='X01'", True), gbUnicode)
        LoadDataSource(tdbcACodeID2, ReturnTableFilter(dtACodeID, "TypeCodeID='X02'", True), gbUnicode)
        LoadDataSource(tdbcACodeID3, ReturnTableFilter(dtACodeID, "TypeCodeID='X03'", True), gbUnicode)
        LoadDataSource(tdbcACodeID4, ReturnTableFilter(dtACodeID, "TypeCodeID='X04'", True), gbUnicode)
        LoadDataSource(tdbcACodeID5, ReturnTableFilter(dtACodeID, "TypeCodeID='X05'", True), gbUnicode)
        LoadDataSource(tdbcACodeID6, ReturnTableFilter(dtACodeID, "TypeCodeID='X06'", True), gbUnicode)
        LoadDataSource(tdbcACodeID7, ReturnTableFilter(dtACodeID, "TypeCodeID='X07'", True), gbUnicode)
        LoadDataSource(tdbcACodeID8, ReturnTableFilter(dtACodeID, "TypeCodeID='X08'", True), gbUnicode)
        LoadDataSource(tdbcACodeID9, ReturnTableFilter(dtACodeID, "TypeCodeID='X09'", True), gbUnicode)
        LoadDataSource(tdbcACodeID10, ReturnTableFilter(dtACodeID, "TypeCodeID='X10'", True), gbUnicode)
    End Sub

    Private Sub LoadtdbcContractorID(ByVal sObjectTypeID As String)
        If clsFilterCombo.IsNewFilter Then
            tdbcContractorID.Splits(0).DisplayColumns("ObjectTypeID").Visible = (sObjectTypeID = "" Or sObjectTypeID = "-1")
            If sObjectTypeID = "" Or sObjectTypeID = "-1" Then
                LoadDataSource(tdbcContractorID, dtObject.DefaultView.ToTable, gbUnicode)
            Else
                LoadDataSource(tdbcContractorID, ReturnTableFilter(dtObject, "ObjectTypeID = " & SQLString(sObjectTypeID), True), gbUnicode)
            End If
        Else
            LoadDataSource(tdbcContractorID, ReturnTableFilter(dtObject, "ObjectTypeID = " & SQLString(sObjectTypeID), True), gbUnicode)
        End If
    End Sub


    Private Sub LoadtdbcSupplierID(ByVal sObjectTypeID As String)

        If clsFilterCombo.IsNewFilter Then
            tdbcSupplierID.Splits(0).DisplayColumns("ObjectTypeID").Visible = (sObjectTypeID = "" Or sObjectTypeID = "-1")
            If sObjectTypeID = "" Or sObjectTypeID = "-1" Then
                LoadDataSource(tdbcSupplierID, dtObject.DefaultView.ToTable, gbUnicode)
            Else
                LoadDataSource(tdbcSupplierID, ReturnTableFilter(dtObject, "ObjectTypeID = " & SQLString(sObjectTypeID), True), gbUnicode)
            End If
        Else
            LoadDataSource(tdbcSupplierID, ReturnTableFilter(dtObject, "ObjectTypeID = " & SQLString(sObjectTypeID), True), gbUnicode)
        End If
    End Sub

    Private Sub LoadLabelACode()
        Dim sSQL As String = ""
        sSQL = "Select TypeCodeID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCodeName", "EngTypeCodeName").ToString() & UnicodeJoin(gbUnicode) & " as Description, Disabled" & vbCrLf
        sSQL &= "From D02T0040 WITH(NOLOCK) Where Type = 'X' Order By TypeCodeID"
        dtLabel = ReturnDataTable(sSQL)

        If dtLabel.Rows.Count > 0 Then
            With dtLabel
                lblACodeID1.Text = .Rows(0).Item("Description").ToString
                lblACodeID2.Text = .Rows(1).Item("Description").ToString
                lblACodeID3.Text = .Rows(2).Item("Description").ToString
                lblACodeID4.Text = .Rows(3).Item("Description").ToString
                lblACodeID5.Text = .Rows(4).Item("Description").ToString
                lblACodeID6.Text = .Rows(5).Item("Description").ToString
                lblACodeID7.Text = .Rows(6).Item("Description").ToString
                lblACodeID8.Text = .Rows(7).Item("Description").ToString
                lblACodeID9.Text = .Rows(8).Item("Description").ToString
                lblACodeID10.Text = .Rows(9).Item("Description").ToString

                lblACodeID1.Font = FontUnicode(gbUnicode)
                lblACodeID2.Font = FontUnicode(gbUnicode)
                lblACodeID3.Font = FontUnicode(gbUnicode)
                lblACodeID4.Font = FontUnicode(gbUnicode)
                lblACodeID5.Font = FontUnicode(gbUnicode)
                lblACodeID6.Font = FontUnicode(gbUnicode)
                lblACodeID7.Font = FontUnicode(gbUnicode)
                lblACodeID8.Font = FontUnicode(gbUnicode)
                lblACodeID9.Font = FontUnicode(gbUnicode)
                lblACodeID10.Font = FontUnicode(gbUnicode)

                tdbcACodeID1.Enabled = CBool(IIf(CInt(.Rows(0).Item("Disabled")) = 0, True, False))
                tdbcACodeID2.Enabled = CBool(IIf(CInt(.Rows(1).Item("Disabled")) = 0, True, False))
                tdbcACodeID3.Enabled = CBool(IIf(CInt(.Rows(2).Item("Disabled")) = 0, True, False))
                tdbcACodeID4.Enabled = CBool(IIf(CInt(.Rows(3).Item("Disabled")) = 0, True, False))
                tdbcACodeID5.Enabled = CBool(IIf(CInt(.Rows(4).Item("Disabled")) = 0, True, False))
                tdbcACodeID6.Enabled = CBool(IIf(CInt(.Rows(5).Item("Disabled")) = 0, True, False))
                tdbcACodeID7.Enabled = CBool(IIf(CInt(.Rows(6).Item("Disabled")) = 0, True, False))
                tdbcACodeID8.Enabled = CBool(IIf(CInt(.Rows(7).Item("Disabled")) = 0, True, False))
                tdbcACodeID9.Enabled = CBool(IIf(CInt(.Rows(8).Item("Disabled")) = 0, True, False))
                tdbcACodeID10.Enabled = CBool(IIf(CInt(.Rows(9).Item("Disabled")) = 0, True, False))
            End With
        End If
    End Sub
#Region "Events tdbcIGEMethodID"

    ' update 5/9/2013 id 58781
    'Private Sub tdbcIGEMethodID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If tdbcIGEMethodID.FindStringExact(tdbcIGEMethodID.Text) = -1 Then tdbcIGEMethodID.Text = ""
    '    If tdbcIGEMethodID.Text = "" Then
    '        UnReadOnlyControl(True, txtCipNo)
    '        txtCipNo.Text = ""
    '        txtCipNo.Focus()
    '        Exit Sub
    '    End If
    'End Sub

    ' update 5/9/2013 id 58781
    'Dim sLastKey As String = ""
    'Dim sKeyString As String = ""
    ''   Dim bObjectChanged As Boolean = False ' biến lư trạng thái có thay đổi giá trị combo DT hay không???

    'Private Sub tdbcIGEMethodID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    ReadOnlyControl(txtCipNo)

    '    Dim sSQL As String = ""
    '    sSQL = "--Xoa du lieu" & vbCrLf
    '    sSQL &= "DELETE D91T1000"
    '    sSQL &= " WHERE   FormID  = 'D02F0060'"
    '    sSQL &= " AND ModuleID = '02'"
    '    sSQL &= " AND UserID  = " & SQLString(gsUserID)
    '    sSQL &= " AND HostID =  " & SQLString(My.Computer.Name) & vbCrLf
    '    sSQL &= SQLInsertD91T1000.ToString & vbCrLf
    '    sSQL &= SQLStoreD91P1000()
    '    Dim dt As DataTable = ReturnDataTable(sSQL)
    '    If dt.Rows.Count > 0 Then
    '        sLastKey = dt.Rows(0).Item("LastKey").ToString
    '        sKeyString = dt.Rows(0).Item("KeyString").ToString
    '        If dt.Rows(0).Item("Status").ToString = "1" Then
    '            D99C0008.MsgL3(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString))
    '        Else
    '            txtCipNo.Text = dt.Rows(0).Item("ID").ToString
    '        End If
    '    End If
    '    sSQL = "--Xoa du lieu" & vbCrLf
    '    sSQL &= "DELETE D91T1000"
    '    sSQL &= " WHERE   FormID  = 'D02F0060'"
    '    sSQL &= " AND ModuleID = '02'"
    '    sSQL &= " AND UserID  = " & SQLString(gsUserID)
    '    sSQL &= " AND HostID =  " & SQLString(My.Computer.Name)
    '    ExecuteSQL(sSQL)
    'End Sub
#End Region

#Region "Events tdbcContractorOTID load tdbcContractorID with txtContractorName"

    Private Sub tdbcContractorOTID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcContractorOTID.GotFocus
        'Dùng phím Enter
        tdbcContractorOTID.Tag = tdbcContractorOTID.Text
    End Sub

    Private Sub tdbcContractorOTID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcContractorOTID.MouseDown
        'Di chuyển chuột
        tdbcContractorOTID.Tag = tdbcContractorOTID.Text
    End Sub

    Private Sub tdbcContractorOTID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcContractorOTID.SelectedValueChanged
        If tdbcContractorOTID.SelectedValue Is Nothing Then
            tdbcContractorOTID.Text = ""
            tdbcContractorID.Text = ""
            txtContractorName.Text = ""
            Exit Sub
        End If
        LoadtdbcContractorID(tdbcContractorOTID.Text)
    End Sub

    Private Sub tdbcContractorOTID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcContractorOTID.LostFocus
        If tdbcContractorOTID.Tag Is Nothing Then Exit Sub
        If tdbcContractorOTID.Tag.ToString = tdbcContractorOTID.Text Then
            If clsFilterCombo.IsNewFilter And tdbcContractorOTID.FindStringExact(tdbcContractorOTID.Text) = -1 Then
                LoadtdbcContractorID("-1")
            End If
            Exit Sub
        End If

        If tdbcContractorOTID.FindStringExact(tdbcContractorOTID.Text) = -1 OrElse tdbcContractorOTID.SelectedValue Is Nothing Then
            tdbcContractorOTID.Text = ""
            LoadtdbcContractorID("-1")
            tdbcContractorID.Text = ""
            txtContractorName.Text = ""
            Exit Sub
        End If
        tdbcContractorID.Text = ""
        txtContractorName.Text = ""
    End Sub

    Private Sub tdbcContractorID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcContractorID.SelectedValueChanged
        If tdbcContractorID.SelectedValue Is Nothing Then
            txtContractorName.Text = ""
        Else
            txtContractorName.Text = tdbcContractorID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcContractorID_Validated(sender As Object, e As EventArgs) Handles tdbcContractorID.Validated
        clsFilterCombo.FilterCombo(tdbcContractorID, e)

        If tdbcContractorID.FindStringExact(tdbcContractorID.Text) = -1 Then
            tdbcContractorID.Text = ""
            txtContractorName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcSupplierOTID load tdbcSupplierID with txtSupplierName"

    Private Sub tdbcSupplierOTID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.GotFocus
        'Dùng phím Enter
        tdbcSupplierOTID.Tag = tdbcSupplierOTID.Text
    End Sub

    Private Sub tdbcSupplierOTID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcSupplierOTID.MouseDown
        'Di chuyển chuột
        tdbcSupplierOTID.Tag = tdbcSupplierOTID.Text
    End Sub

    Private Sub tdbcSupplierOTID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.SelectedValueChanged
        If tdbcSupplierOTID.SelectedValue Is Nothing Then
            tdbcSupplierOTID.Text = ""
            tdbcSupplierID.Text = ""
            txtSupplierName.Text = ""
            Exit Sub
        End If
        LoadtdbcSupplierID(tdbcSupplierOTID.Text)
    End Sub

    Private Sub tdbcSupplierOTID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierOTID.LostFocus
        If tdbcSupplierOTID.Tag Is Nothing Then Exit Sub
        If tdbcSupplierOTID.Tag.ToString = tdbcSupplierOTID.Text Then
            If clsFilterCombo.IsNewFilter And tdbcSupplierOTID.FindStringExact(tdbcSupplierOTID.Text) = -1 Then
                LoadtdbcSupplierID("-1")
            End If
            Exit Sub
        End If

        If tdbcSupplierOTID.FindStringExact(tdbcSupplierOTID.Text) = -1 OrElse tdbcSupplierOTID.SelectedValue Is Nothing Then
            tdbcSupplierOTID.Text = ""
            LoadtdbcSupplierID("-1")
            tdbcSupplierID.Text = ""
            txtSupplierName.Text = ""
            Exit Sub
        End If
        tdbcSupplierID.Text = ""
        txtSupplierName.Text = ""
    End Sub

    Private Sub tdbcSupplierID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSupplierID.SelectedValueChanged
        If tdbcSupplierID.SelectedValue Is Nothing Then
            txtSupplierName.Text = ""
        Else
            txtSupplierName.Text = tdbcSupplierID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcSupplierID_Validated(sender As Object, e As EventArgs) Handles tdbcSupplierID.Validated
        clsFilterCombo.FilterCombo(tdbcSupplierID, e)

        If tdbcSupplierID.FindStringExact(tdbcSupplierID.Text) = -1 Then
            tdbcSupplierID.Text = ""
            txtSupplierName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcAccountID"

    Private Sub tdbcAccountID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.LostFocus
        If tdbcAccountID.FindStringExact(tdbcAccountID.Text) = -1 Then tdbcAccountID.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID1"

    Private Sub tdbcACodeID1_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID1.Close
        If tdbcACodeID1.FindStringExact(tdbcACodeID1.Text) = -1 Then tdbcACodeID1.Text = ""
    End Sub

    Private Sub tdbcACodeID1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID1.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID1.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID2"

    Private Sub tdbcACodeID2_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID2.Close
        If tdbcACodeID2.FindStringExact(tdbcACodeID2.Text) = -1 Then tdbcACodeID2.Text = ""
    End Sub

    Private Sub tdbcACodeID2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID2.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID2.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID3"

    Private Sub tdbcACodeID3_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID3.Close
        If tdbcACodeID3.FindStringExact(tdbcACodeID3.Text) = -1 Then tdbcACodeID3.Text = ""
    End Sub

    Private Sub tdbcACodeID3_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID3.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID3.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID4"

    Private Sub tdbcACodeID4_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID4.Close
        If tdbcACodeID4.FindStringExact(tdbcACodeID4.Text) = -1 Then tdbcACodeID4.Text = ""
    End Sub

    Private Sub tdbcACodeID4_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID4.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID4.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID5"

    Private Sub tdbcACodeID5_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID5.Close
        If tdbcACodeID5.FindStringExact(tdbcACodeID5.Text) = -1 Then tdbcACodeID5.Text = ""
    End Sub

    Private Sub tdbcACodeID5_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID5.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID5.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID6"

    Private Sub tdbcACodeID6_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID6.Close
        If tdbcACodeID6.FindStringExact(tdbcACodeID6.Text) = -1 Then tdbcACodeID6.Text = ""
    End Sub

    Private Sub tdbcACodeID6_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID6.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID6.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID7"

    Private Sub tdbcACodeID7_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID7.Close
        If tdbcACodeID7.FindStringExact(tdbcACodeID7.Text) = -1 Then tdbcACodeID7.Text = ""
    End Sub

    Private Sub tdbcACodeID7_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID7.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID7.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID8"

    Private Sub tdbcACodeID8_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID8.Close
        If tdbcACodeID8.FindStringExact(tdbcACodeID8.Text) = -1 Then tdbcACodeID8.Text = ""
    End Sub

    Private Sub tdbcACodeID8_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID8.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID8.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID9"

    Private Sub tdbcACodeID9_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID9.Close
        If tdbcACodeID9.FindStringExact(tdbcACodeID9.Text) = -1 Then tdbcACodeID9.Text = ""
    End Sub

    Private Sub tdbcACodeID9_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID9.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID9.Text = ""
    End Sub

#End Region

#Region "Events tdbcACodeID10"

    Private Sub tdbcACodeID10_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcACodeID10.Close
        If tdbcACodeID10.FindStringExact(tdbcACodeID10.Text) = -1 Then tdbcACodeID10.Text = ""
    End Sub

    Private Sub tdbcACodeID10_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcACodeID10.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcACodeID10.Text = ""
    End Sub

#End Region

    Private Function AllowSave() As Boolean
        If D02Systems.CIPAuto = 2 Then
            If tdbcCIPS1ID.Enabled And tdbcCIPS1ID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Ma_XDCB") & " 1")
                tdbcCIPS1ID.Focus()
                Return False
            End If
            If tdbcCIPS2ID.Enabled And tdbcCIPS2ID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Ma_XDCB") & " 2")
                tdbcCIPS2ID.Focus()
                Return False
            End If
            If tdbcCIPS3ID.Enabled And tdbcCIPS3ID.Text = "" Then
                D99C0008.MsgNotYetChoose(rL3("Ma_XDCB") & " 3")
                tdbcCIPS3ID.Focus()
                Return False
            End If
        End If
        If txtCipNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ma_XDCB"))
            tabMain.SelectedIndex = 0
            If D02Systems.CIPAuto = 2 Then
                tdbcCIPS1ID.Focus()
            Else
                txtCipNo.Focus() 'txtCipNo.Focus()
            End If
            Return False
        End If
        If txtCipName.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ten_XDCB"))
            tabMain.SelectedIndex = 0
            txtCipName.Focus()
            Return False
        End If
        If tdbcAccountID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_tai_khoan"))
            tabMain.SelectedIndex = 0
            tdbcAccountID.Focus()
            Return False
        End If
        If _FormState = EnumFormState.FormAdd Then
            If IsExistKey("D02T0100", "CipNo", txtCipNo.Text) Then
                D99C0008.MsgDuplicatePKey()
                tabMain.SelectedIndex = 0
                If D02Systems.CIPAuto = 2 Then
                    tdbcCIPS1ID.Focus()
                Else
                    txtCipNo.Focus() 'txtCipNo.Focus()
                End If
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
        _savedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                If D02Systems.CIPAuto = 2 Then '4/5/2017, 	Trần Hoàng Anh: id 96917-Lỗi không lưu Last Key khi tạo mã XDCB tự động 
                    _S1 = IIf(IsDBNull(tdbcCIPS1ID.Text) Or tdbcCIPS1ID.Text = "<<", "", tdbcCIPS1ID.Text).ToString
                    _S2 = IIf(IsDBNull(tdbcCIPS2ID.Text) Or tdbcCIPS2ID.Text = "<<", "", tdbcCIPS2ID.Text).ToString
                    _S3 = IIf(IsDBNull(tdbcCIPS3ID.Text) Or tdbcCIPS3ID.Text = "<<", "", tdbcCIPS3ID.Text).ToString

                    SaveLastKey1(_S1, _S2, _S3, _CIPOutputOrder, _CIPOutputLength, _CIPSeperator, True, _CIPTableName)
                End If

                sSQL.Append(SQLInsertD02T0100)
                'Lưu LastKey của Số phiếu xuống Database (gọi hàm CreateIGEVoucherNo bật cờ True)
                'Kiểm tra trùng Số phiếu (gọi hàm CheckDuplicateVoucherNo)

            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T0100)
        End Select
        If _iGEMethodID <> "" Then ' nếu có chọn Mã PP tự động mới thực hiện đoạn lệnh update Lastkey, ngược lại thì không.
            If ExistRecord("select top 1 1 from D91T1001 where KeyString = " & SQLString(sKeyString) & " And ModuleID = " & SQLString("02") & " And FormID = " & SQLString("D02F0060")) Then
                sSQL.Append(SQLUpdateD91T1001().ToString)
            Else
                sSQL.Append(SQLInsertD91T1001().ToString)
            End If

        End If
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _cipNo = txtCipNo.Text
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    'ExecuteAuditLog(_auditCode, "02", txtCipNo.Text, txtCipName.Text)
                    Lemon3.D91.RunAuditLog("02", _auditCode, "02", txtCipNo.Text, txtCipName.Text)
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Sub SaveLastKey1(ByVal strS1 As String, ByVal strS2 As String, ByVal strS3 As String, ByVal sOutputOrder As String, ByVal iOutputLength As Integer, ByVal sSeperator As String, ByVal bFlagSave As Boolean, ByVal sTableName As String)
        Dim iOutputOrder As Integer = -1
        Select Case sOutputOrder
            Case "NSSS"
                iOutputOrder = D99D0041.OutOrderEnum.lmNSSS
            Case "SNSS"
                iOutputOrder = D99D0041.OutOrderEnum.lmSNSS
            Case "SSNS"
                iOutputOrder = D99D0041.OutOrderEnum.lmSSNS
            Case "SSSN"
                iOutputOrder = D99D0041.OutOrderEnum.lmSSSN
        End Select
        CreateIGEVoucherNo(strS1, strS2, strS3, CType(iOutputOrder, D99D0041.OutOrderEnum), iOutputLength, sSeperator, True, sTableName)
    End Sub

    Private Function SQLInsertD02T0100() As StringBuilder
        _cipID = CreateIGE("D02T0100", "CipID", "02", "CI", gsStringKey)

        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0100(")
        sSQL.Append("CipID, CipNo, DescriptionU, CipNameU, ")
        sSQL.Append("AccountID, Disabled, Status, StartDate, ")
        sSQL.Append("CompletionDate, CreateDate, CreateUserID, LastModifyDate, ")
        sSQL.Append("LastModifyUserID, X01ID, X02ID, X03ID, X04ID, ")
        sSQL.Append("X05ID, X06ID, X07ID, X08ID, X09ID, ")
        sSQL.Append("X10ID, DivisionID, ContractorOTID, ContractorID, SupplierOTID, ")
        sSQL.Append("SupplierID, ExpStartDate, ExpEndDate, CipCost, ")
        sSQL.Append("CIPNum01, CIPNum02, CIPNum03, CIPNum04, ")
        sSQL.Append("CIPNum05, CIPNum06, CIPNum07, CIPNum08, CIPNum09, ")
        sSQL.Append("CIPNum10, CIPDate01, CIPDate02, CIPDate03, CIPDate04, ")
        sSQL.Append("CIPDate05, CIPDate06, CIPDate07, CIPDate08, CIPDate09, ")
        sSQL.Append("CIPDate10,")
        sSQL.Append("CIPString01U, CIPString02U, CIPString03U, CIPString04U,CIPString05U, CIPString06U, CIPString07U, CIPString08U, CIPString09U, ")
        sSQL.Append("CIPString10U,CIPObjectTypeID,CIPObjectID,CIPEmployeeID,D54ProjectID , D27PropertyProductID ")

        sSQL.Append(") Values(")
        sSQL.Append(SQLString(_cipID) & COMMA) 'CipID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(txtCipNo.Text) & COMMA) 'CipNo, varchar[20], NULL
        sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
        sSQL.Append(SQLStringUnicode(txtCipName.Text, gbUnicode, True) & COMMA) 'CipName, varchar[250], NULL
        sSQL.Append(SQLString(tdbcAccountID.Text) & COMMA) 'AccountID, varchar[20], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NULL
        sSQL.Append(SQLNumber(cboStatus.SelectedIndex) & COMMA) 'Status, tinyint, NULL
        sSQL.Append(SQLDateSave(c1dateStartDate.Value) & COMMA) 'StartDate, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateCompletionDate.Value) & COMMA) 'CompletionDate, datetime, NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID1.Text) & COMMA) 'X01ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID2.Text) & COMMA) 'X02ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID3.Text) & COMMA) 'X03ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID4.Text) & COMMA) 'X04ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID5.Text) & COMMA) 'X05ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID6.Text) & COMMA) 'X06ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID7.Text) & COMMA) 'X07ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID8.Text) & COMMA) 'X08ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID9.Text) & COMMA) 'X09ID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcACodeID10.Text) & COMMA) 'X10ID, varchar[20], NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcContractorOTID.Text) & COMMA) 'ContractorOTID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcContractorID.Text) & COMMA) 'ContractorID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcSupplierOTID.Text) & COMMA) 'SupplierOTID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcSupplierID.Text) & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLDateSave(c1dateExpStartDate.Value) & COMMA) 'ExpStartDate, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateExpEndDate.Value) & COMMA) 'ExpEndDate, datetime, NULL
        sSQL.Append(SQLMoney(txtCipCost.Text, DxxFormat.DecimalPlaces) & COMMA) 'CipCost, decimal, NOT NULL

        sSQL.Append(SQLMoney(txtFANum01.Text) & COMMA) 'CIPNum01, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum02.Text) & COMMA) 'CIPNum02, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum03.Text) & COMMA) 'CIPNum03, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum04.Text) & COMMA) 'CIPNum04, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum05.Text) & COMMA) 'CIPNum05, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum06.Text) & COMMA) 'CIPNum06, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum07.Text) & COMMA) 'CIPNum07, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum08.Text) & COMMA) 'CIPNum08, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum09.Text) & COMMA) 'CIPNum09, money, NOT NULL
        sSQL.Append(SQLMoney(txtFANum10.Text) & COMMA) 'CIPNum10, money, NOT NULL
        sSQL.Append(SQLDateSave(c1dateFADate01.Text) & COMMA) 'CIPDate01, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate02.Text) & COMMA) 'CIPDate02, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate03.Text) & COMMA) 'CIPDate03, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate04.Text) & COMMA) 'CIPDate04, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate05.Text) & COMMA) 'CIPDate05, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate06.Text) & COMMA) 'CIPDate06, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate07.Text) & COMMA) 'CIPDate07, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate08.Text) & COMMA) 'CIPDate08, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate09.Text) & COMMA) 'CIPDate09, datetime, NULL
        sSQL.Append(SQLDateSave(c1dateFADate10.Text) & COMMA) 'CIPDate10, datetime, NULL
        sSQL.Append(SQLStringUnicode(txtFAString01.Text, gbUnicode, True) & COMMA) 'CIPString01U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString02.Text, gbUnicode, True) & COMMA) 'CIPString02U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString03.Text, gbUnicode, True) & COMMA) 'CIPString03U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString04.Text, gbUnicode, True) & COMMA) 'CIPString04U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString05.Text, gbUnicode, True) & COMMA) 'CIPString05U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString06.Text, gbUnicode, True) & COMMA) 'CIPString06U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString07.Text, gbUnicode, True) & COMMA) 'CIPString07U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString08.Text, gbUnicode, True) & COMMA) 'CIPString08U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString09.Text, gbUnicode, True) & COMMA) 'CIPString09U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode(txtFAString10.Text, gbUnicode, True) & COMMA) 'CIPString10U, varchar[1000], NOT NULL
        sSQL.Append(SQLString(tdbcCIPObjectTypeID.Text) & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcCIPObjectID.Text) & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcCIPEmployeeID.Text) & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLString(txtProjectID.Text) & COMMA) 'ProjectID, varchar[50], NOT NULL
        sSQL.Append(SQLString(txtPropertyProductID.Text)) 'PropertyProductID, varchar[50], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function


    Private Function SQLUpdateD02T0100() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("DescriptionU = " & SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("CipNameU = " & SQLStringUnicode(txtCipName.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("AccountID = " & SQLString(tdbcAccountID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("CIPObjectTypeID = " & SQLString(tdbcCIPObjectTypeID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("CIPObjectID = " & SQLString(tdbcCIPObjectID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("CIPEmployeeID = " & SQLString(tdbcCIPEmployeeID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'tinyint, NULL
        sSQL.Append("StartDate = " & SQLDateSave(c1dateStartDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("CompletionDate = " & SQLDateSave(c1dateCompletionDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("X01ID = " & SQLString(tdbcACodeID1.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X02ID = " & SQLString(tdbcACodeID2.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X03ID = " & SQLString(tdbcACodeID3.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X04ID = " & SQLString(tdbcACodeID4.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X05ID = " & SQLString(tdbcACodeID5.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X06ID = " & SQLString(tdbcACodeID6.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X07ID = " & SQLString(tdbcACodeID7.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X08ID = " & SQLString(tdbcACodeID8.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X09ID = " & SQLString(tdbcACodeID9.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("X10ID = " & SQLString(tdbcACodeID10.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("ContractorOTID = " & SQLString(tdbcContractorOTID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ContractorID = " & SQLString(tdbcContractorID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("SupplierOTID = " & SQLString(tdbcSupplierOTID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("SupplierID = " & SQLString(tdbcSupplierID.Text) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("ExpStartDate = " & SQLDateSave(c1dateExpStartDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("ExpEndDate = " & SQLDateSave(c1dateExpEndDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("CipCost = " & SQLMoney(txtCipCost.Text, DxxFormat.DecimalPlaces) & COMMA) 'decimal, NOT NULL
        sSQL.Append("CIPNum01 = " & SQLMoney(txtFANum01.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum02 = " & SQLMoney(txtFANum02.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum03 = " & SQLMoney(txtFANum03.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum04 = " & SQLMoney(txtFANum04.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum05 = " & SQLMoney(txtFANum05.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum06 = " & SQLMoney(txtFANum06.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum07 = " & SQLMoney(txtFANum07.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum08 = " & SQLMoney(txtFANum08.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum09 = " & SQLMoney(txtFANum09.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPNum10 = " & SQLMoney(txtFANum10.Text) & COMMA) 'money, NOT NULL
        sSQL.Append("CIPDate01 = " & SQLDateSave(c1dateFADate01.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate02 = " & SQLDateSave(c1dateFADate02.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate03 = " & SQLDateSave(c1dateFADate03.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate04 = " & SQLDateSave(c1dateFADate04.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate05 = " & SQLDateSave(c1dateFADate05.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate06 = " & SQLDateSave(c1dateFADate06.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate07 = " & SQLDateSave(c1dateFADate07.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate08 = " & SQLDateSave(c1dateFADate08.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate09 = " & SQLDateSave(c1dateFADate09.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPDate10 = " & SQLDateSave(c1dateFADate10.Text) & COMMA) 'datetime, NULL
        sSQL.Append("CIPString01U = " & SQLStringUnicode(txtFAString01.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString02U = " & SQLStringUnicode(txtFAString02.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString03U = " & SQLStringUnicode(txtFAString03.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString04U = " & SQLStringUnicode(txtFAString04.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString05U = " & SQLStringUnicode(txtFAString05.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString06U = " & SQLStringUnicode(txtFAString06.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString07U = " & SQLStringUnicode(txtFAString07.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString08U = " & SQLStringUnicode(txtFAString08.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString09U = " & SQLStringUnicode(txtFAString09.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("CIPString10U = " & SQLStringUnicode(txtFAString10.Text, gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
        sSQL.Append("D54ProjectID = " & SQLString(txtProjectID.Text) & COMMA) 'varchar[50], NOT NULL
        sSQL.Append("D27PropertyProductID = " & SQLString(txtPropertyProductID.Text)) 'varchar[50], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("CipID = " & SQLString(_cipID))

        Return sSQL
    End Function


    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        _FormState = EnumFormState.FormAdd
        tdbcCIPS1ID.SelectedValue = "-1"
        tdbcCIPS2ID.SelectedValue = "-1"
        tdbcCIPS3ID.SelectedValue = "-1"
        If D02Options.SaveLastRecent Then
            ' LoadAdd()
            Application.DoEvents()
            tabMain.SelectedTab = TabPage1
            Application.DoEvents()
            If D02Systems.CIPAuto = 2 Then
                tdbcCIPS1ID.Focus()
            Else
                txtCipNo.Focus() 'txtCipNo.Focus()
            End If
            txtCipName.Text = ""
            tdbcAccountID.Text = ""
            txtCipNo.Text = ""
            btnNext.Enabled = False
            btnSave.Enabled = True
        Else
            tdbcACodeID1.Text = ""
            tdbcACodeID2.Text = ""
            tdbcACodeID3.Text = ""
            tdbcACodeID4.Text = ""
            tdbcACodeID5.Text = ""
            tdbcACodeID6.Text = ""
            tdbcACodeID7.Text = ""
            tdbcACodeID8.Text = ""
            tdbcACodeID9.Text = ""
            tdbcACodeID10.Text = ""
            tdbcCIPEmployeeID.Text = ""
            tdbcCIPObjectID.Text = ""
            tdbcCIPObjectTypeID.Text = ""
            btnNext.Enabled = False
            btnSave.Enabled = True
            ClearText(Me)
            LoadAdd()
            Application.DoEvents()
            tabMain.SelectedTab = TabPage1
            Application.DoEvents()
            If D02Systems.CIPAuto = 2 Then
                tdbcCIPS1ID.Focus()
            Else
                txtCipNo.Focus() 'txtCipNo.Focus()
            End If  '  txtCipNo.Focus()
        End If

        'If D02Systems.CIPAuto = 2 Then
        '    GetAutoCIPInfo()
        'Else

        '    Dim f As New D02F0005()
        '    With f
        '        .FormID = "D02F2001"
        '        .IGEMethodID = _iGEMethodID
        '        If dtXCodeValue IsNot Nothing Then
        '            .dtXCodeValue = dtXCodeValue
        '        End If
        '        .ShowInTaskbar = True
        '        .ShowDialog()
        '        If .SavedOk Then
        '            sID = .ID
        '            sKeyString = .KeyString
        '            sKeyName = .KeyName
        '            sLastKey = .LastKey
        '            bIsAutoNum = .IsAutoNum
        '            _iGEMethodID = .IGEMethodID
        '            dtXCodeValue = .dtXCodeValue
        '            dtXCode = .dtXCode
        '            LoadDefaultTDBCACodeID() ' Default giá trị đư
        '        End If
        '    End With
        '    f.Dispose()
        'End If

        '20/7/2017, Hồng Thanh Long: id 100003-ER - Lỗi tạo mã XDCB không thiết lập tạo mã tự động
        GetAutoCIPInfo()
        If _cIPAuto = 1 Then
            Dim bIsAddNew As Boolean = False
            Dim f As New D02F0005
            With f
                .ShowInTaskbar = True
                .ShowDialog()
                If .SavedOk Then
                    sID = f.ID
                    sKeyString = f.KeyString
                    sKeyName = f.KeyName
                    sLastKey = f.LastKey
                    bIsAutoNum = f.IsAutoNum
                    _iGEMethodID = f.IGEMethodID
                    dtXCodeValue = f.dtXCodeValue
                    dtXCode = f.dtXCode
                    '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
                    bIsAddNew = .IsAddNew
                    If bIsAddNew Then LoadTDBCACodeID()

                    LoadDefaultTDBCACodeID() ' Default giá trị được lấy từ
                End If
                .Dispose()
            End With
        End If

    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Cap_nhat_XDCB_-_D02F2001") & UnicodeCaption(gbUnicode) ' rL3("Cap_nhat_ma_cong_trinh_XDCB_-_D02F2001") & UnicodeCaption(gbUnicode) 'CËp nhËt mº c¤ng trØnh XDCB - D02F2001
        '================================================================ 
        'lblIGEMethodID.Text = rl3("PP_tao_ma") 'PP tạo mã    ' update 5/9/2013 id 58781
        lblCipNo.Text = rL3("Ma_XDCB") ' rL3("Ma_cong_trinh")
        lblCipName.Text = rL3("Ten_XDCB") 'rL3("Ten_cong_trinh") 'Tên công trình
        lblAccountID.Text = rL3("Ma_tai_khoan") 'Mã tài khoản
        lblteStartDate.Text = rL3("Ngay_bat_dau") 'Ngày bắt đầu
        lblteCompletionDate.Text = rL3("Ngay_hoan_thanh") 'Ngày hoàn thành
        lblDescription.Text = rL3("Dien_giai") 'Diễn giải
        lblStatus.Text = rL3("Tinh_trang") 'Tình trạng        
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        lblEmployeeID.Text = rL3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblProjectID.Text = rL3("Ma_BDS") 'Mã BĐS
        '================================================================ 
        btnSave.Text = rL3("_Luu") '&Lưu
        btnNext.Text = rL3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rL3("Khong_su_dung") 'Không sử dụng
        '================================================================ 
        TabPage1.Text = "1. " & rL3("Thong_tin_chung") '1. Thông tin chung
        TabPage2.Text = "2. " & rL3("Ma_phan_tich") '2. Mã phân tích
        TabPage3.Text = "3. " & rL3("Thong_tin_phu")
        '================================================================ 

        tdbcAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbcAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbcACodeID10.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID10.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID9.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID9.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID8.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID8.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID7.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID7.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID6.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID6.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID5.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID5.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID4.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID4.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID3.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID3.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID2.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID2.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcACodeID1.Columns("ACodeID").Caption = rL3("Ma") 'Mã
        tdbcACodeID1.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbcCIPEmployeeID.Columns("EmployeeID").Caption = rL3("Ma") 'Mã
        tdbcCIPEmployeeID.Columns("EmployeeName").Caption = rL3("Ten") 'Tên
        tdbcCIPObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcCIPObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbcCIPObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcCIPObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Dien_giai") 'Diễn giải

        lblContractorOTID.Text = rL3("Nha_thau")
        lblSupplierOTID.Text = rL3("Nha_cung_cap")
        lblCipCost.Text = rL3("Gia_tri_hang_muc") 'Giá trị hạng mục
        lblExpStartDate.Text = rL3("Ngay_du_kien_bat_dau")
        lblExpEndDate.Text = rL3("Ngay_du_kien_ket_thuc")
        lblteCompletionDate.Text = rL3("Ngay_ket_thuc")
        tdbcContractorOTID.Columns("ObjectTypeID").Caption = rL3("Ma")
        tdbcContractorOTID.Columns("ObjectTypeName").Caption = rL3("Ten")
        tdbcContractorID.Columns("ObjectID").Caption = rL3("Ma")
        tdbcContractorID.Columns("ObjectName").Caption = rL3("Ten")
        tdbcSupplierOTID.Columns("ObjectTypeID").Caption = rL3("Ma")
        tdbcSupplierOTID.Columns("ObjectTypeName").Caption = rL3("Ten")
        tdbcSupplierID.Columns("ObjectID").Caption = rL3("Ma")
        tdbcSupplierID.Columns("ObjectName").Caption = rL3("Ten")

        '================================================================ 
        tdbcCIPS1ID.Columns("CIPS1ID").Caption = rL3("Ma_phan_loai") & " 1" 'Mã phân loại 1
        tdbcCIPS1ID.Columns("CIPS1Name").Caption = rL3("Ten_phan_loai") & " 1" 'Tên phân loại 1
        tdbcCIPS2ID.Columns("CIPS2ID").Caption = rL3("Ma_phan_loai") & " 2" 'Mã phân loại 2
        tdbcCIPS2ID.Columns("CIPS2Name").Caption = rL3("Ten_phan_loai") & " 2" 'Tên phân loại 2
        tdbcCIPS3ID.Columns("CIPS3ID").Caption = rL3("Ma_phan_loai") & " 3" 'Mã phân loại 3
        tdbcCIPS3ID.Columns("CIPS3Name").Caption = rL3("Ten_phan_loai") & " 3" 'Tên phân loại 3
    End Sub

    Private Sub c1dateStartDate_DropDownClosed(ByVal sender As Object, ByVal e As C1.Win.C1Input.DropDownClosedEventArgs) Handles c1dateStartDate.DropDownClosed
        If c1dateStartDate.Text <> "" And c1dateCompletionDate.Text <> "" Then
            If CDate(c1dateStartDate.Text) > CDate(c1dateCompletionDate.Text) Then
                D99C0008.MsgL3(rL3("Khoang_ngay_khong_hop_le"))
                c1dateStartDate.Focus()
            End If
        End If
    End Sub

    Private Sub c1dateStartDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles c1dateStartDate.Validating
        If c1dateStartDate.Text <> "" And c1dateCompletionDate.Text <> "" Then
            If CDate(c1dateStartDate.Text) > CDate(c1dateCompletionDate.Text) Then
                D99C0008.MsgL3(rL3("Khoang_ngay_khong_hop_le"))
                c1dateStartDate.Focus()
            End If
        End If
    End Sub

    Private Sub c1dateCompletionDate_DropDownClosed(ByVal sender As Object, ByVal e As C1.Win.C1Input.DropDownClosedEventArgs) Handles c1dateCompletionDate.DropDownClosed
        If c1dateStartDate.Text <> "" And c1dateCompletionDate.Text <> "" Then
            If CDate(c1dateStartDate.Text) > CDate(c1dateCompletionDate.Text) Then
                D99C0008.MsgL3(rL3("Khoang_ngay_khong_hop_le"))
                c1dateCompletionDate.Focus()
            End If
        End If
    End Sub

    Private Sub c1dateCompletionDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles c1dateCompletionDate.Validating
        If c1dateStartDate.Text <> "" And c1dateCompletionDate.Text <> "" Then
            If CDate(c1dateStartDate.Text) > CDate(c1dateCompletionDate.Text) Then
                D99C0008.MsgL3(rL3("Khoang_ngay_khong_hop_le"))
                c1dateCompletionDate.Focus()
            End If
        End If
    End Sub

    Private Sub txtCipCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCipCost.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtCipCost_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCipCost.LostFocus
        If Not L3IsNumeric(txtCipCost.Text) Then
            txtCipCost.Text = SQLNumber("0", DxxFormat.DecimalPlaces)
        Else
            If Number(txtCipCost.Text) > MaxMoney Or Number(txtCipCost.Text) < MinMoney Then
                txtCipCost.Text = SQLNumber("0", DxxFormat.DecimalPlaces)
            Else
                txtCipCost.Text = SQLNumber(txtCipCost.Text, DxxFormat.DecimalPlaces)
            End If
        End If
    End Sub


    Private Sub LoadCaption()
        Dim bUseSpec As Boolean = False
        Dim sSQL As String = SQLStoreD02P0015()
        Dim dt As New DataTable
        dt = ReturnDataTable(sSQL)
        If (dt.Rows.Count > 0) Then
            Str01.Font = FontUnicode(gbUnicode)
            Str02.Font = FontUnicode(gbUnicode)
            Str03.Font = FontUnicode(gbUnicode)
            Str04.Font = FontUnicode(gbUnicode)
            Str05.Font = FontUnicode(gbUnicode)
            Str06.Font = FontUnicode(gbUnicode)
            Str07.Font = FontUnicode(gbUnicode)
            Str08.Font = FontUnicode(gbUnicode)
            Str09.Font = FontUnicode(gbUnicode)
            Str10.Font = FontUnicode(gbUnicode)

            Num01.Font = FontUnicode(gbUnicode)
            Num02.Font = FontUnicode(gbUnicode)
            Num03.Font = FontUnicode(gbUnicode)
            Num04.Font = FontUnicode(gbUnicode)
            Num05.Font = FontUnicode(gbUnicode)
            Num06.Font = FontUnicode(gbUnicode)
            Num07.Font = FontUnicode(gbUnicode)
            Num08.Font = FontUnicode(gbUnicode)
            Num09.Font = FontUnicode(gbUnicode)
            Num10.Font = FontUnicode(gbUnicode)
            Date01.Font = FontUnicode(gbUnicode)
            Date02.Font = FontUnicode(gbUnicode)
            Date03.Font = FontUnicode(gbUnicode)
            Date04.Font = FontUnicode(gbUnicode)
            Date05.Font = FontUnicode(gbUnicode)
            Date06.Font = FontUnicode(gbUnicode)
            Date07.Font = FontUnicode(gbUnicode)
            Date08.Font = FontUnicode(gbUnicode)
            Date09.Font = FontUnicode(gbUnicode)
            Date10.Font = FontUnicode(gbUnicode)

            If geLanguage = EnumLanguage.Vietnamese Then
                Str01.Text = dt.Rows(10).Item("Data84").ToString
            Else
                Str01.Text = dt.Rows(10).Item("Data01").ToString
            End If
            Str01.Tag = dt.Rows(10).Item("DecimalNum")
            txtFAString01.Enabled = L3Bool(dt.Rows(10).Item("Disabled"))
            txtFAString01.MaxLength = L3Int(dt.Rows(10).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str02.Text = dt.Rows(11).Item("Data84").ToString
            Else
                Str02.Text = dt.Rows(11).Item("Data01").ToString
            End If
            txtFAString02.Enabled = L3Bool(dt.Rows(11).Item("Disabled"))
            txtFAString02.MaxLength = L3Int(dt.Rows(11).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str03.Text = dt.Rows(12).Item("Data84").ToString
            Else
                Str03.Text = dt.Rows(12).Item("Data01").ToString
            End If
            txtFAString03.Enabled = L3Bool(dt.Rows(12).Item("Disabled"))
            txtFAString03.MaxLength = L3Int(dt.Rows(12).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str04.Text = dt.Rows(13).Item("Data84").ToString
            Else
                Str04.Text = dt.Rows(13).Item("Data01").ToString
            End If
            txtFAString04.Enabled = L3Bool(dt.Rows(13).Item("Disabled"))
            txtFAString04.MaxLength = L3Int(dt.Rows(13).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str05.Text = dt.Rows(14).Item("Data84").ToString
            Else
                Str05.Text = dt.Rows(14).Item("Data01").ToString
            End If
            txtFAString05.Enabled = L3Bool(dt.Rows(14).Item("Disabled"))
            txtFAString05.MaxLength = L3Int(dt.Rows(14).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str06.Text = dt.Rows(15).Item("Data84").ToString
            Else
                Str06.Text = dt.Rows(15).Item("Data01").ToString
            End If
            txtFAString06.Enabled = L3Bool(dt.Rows(15).Item("Disabled"))
            txtFAString06.MaxLength = L3Int(dt.Rows(15).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str07.Text = dt.Rows(16).Item("Data84").ToString
            Else
                Str07.Text = dt.Rows(16).Item("Data01").ToString
            End If
            txtFAString07.Enabled = L3Bool(dt.Rows(16).Item("Disabled"))
            txtFAString07.MaxLength = L3Int(dt.Rows(16).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str08.Text = dt.Rows(17).Item("Data84").ToString
            Else
                Str08.Text = dt.Rows(17).Item("Data01").ToString
            End If
            txtFAString08.Enabled = L3Bool(dt.Rows(17).Item("Disabled"))
            txtFAString08.MaxLength = L3Int(dt.Rows(17).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str09.Text = dt.Rows(18).Item("Data84").ToString
            Else
                Str09.Text = dt.Rows(18).Item("Data01").ToString
            End If
            txtFAString09.Enabled = L3Bool(dt.Rows(18).Item("Disabled"))
            txtFAString09.MaxLength = L3Int(dt.Rows(18).Item("DecimalNum"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Str10.Text = dt.Rows(19).Item("Data84").ToString
            Else
                Str10.Text = dt.Rows(19).Item("Data01").ToString
            End If
            txtFAString10.Enabled = L3Bool(dt.Rows(19).Item("Disabled"))
            txtFAString10.MaxLength = L3Int(dt.Rows(19).Item("DecimalNum"))

            If geLanguage = EnumLanguage.Vietnamese Then
                Num01.Text = dt.Rows(0).Item("Data84").ToString
            Else
                Num01.Text = dt.Rows(0).Item("Data01").ToString
            End If
            txtFANum01.Enabled = L3Bool(dt.Rows(0).Item("Disabled"))
            txtFANum01.Tag = dt.Rows(0).Item("DecimalNum")
            If geLanguage = EnumLanguage.Vietnamese Then
                Num02.Text = dt.Rows(1).Item("Data84").ToString
            Else
                Num02.Text = dt.Rows(1).Item("Data01").ToString
            End If
            txtFANum02.Enabled = L3Bool(dt.Rows(1).Item("Disabled"))
            txtFANum02.Tag = dt.Rows(1).Item("DecimalNum")
            If geLanguage = EnumLanguage.Vietnamese Then
                Num03.Text = dt.Rows(2).Item("Data84").ToString
            Else
                Num03.Text = dt.Rows(2).Item("Data01").ToString
            End If
            txtFANum03.Tag = dt.Rows(2).Item("DecimalNum")
            txtFANum03.Enabled = L3Bool(dt.Rows(2).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num04.Text = dt.Rows(3).Item("Data84").ToString
            Else
                Num04.Text = dt.Rows(3).Item("Data01").ToString
            End If
            txtFANum04.Tag = dt.Rows(3).Item("DecimalNum")
            txtFANum04.Enabled = L3Bool(dt.Rows(3).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num05.Text = dt.Rows(4).Item("Data84").ToString
            Else
                Num05.Text = dt.Rows(4).Item("Data01").ToString
            End If
            txtFANum05.Tag = dt.Rows(4).Item("DecimalNum")
            txtFANum05.Enabled = L3Bool(dt.Rows(4).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num06.Text = dt.Rows(5).Item("Data84").ToString
            Else
                Num06.Text = dt.Rows(5).Item("Data01").ToString
            End If
            txtFANum06.Tag = dt.Rows(5).Item("DecimalNum")
            txtFANum06.Enabled = L3Bool(dt.Rows(5).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num07.Text = dt.Rows(6).Item("Data84").ToString
            Else
                Num07.Text = dt.Rows(6).Item("Data01").ToString
            End If
            txtFANum07.Tag = dt.Rows(6).Item("DecimalNum")
            txtFANum07.Enabled = L3Bool(dt.Rows(6).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num08.Text = dt.Rows(7).Item("Data84").ToString
            Else
                Num08.Text = dt.Rows(7).Item("Data01").ToString
            End If
            txtFANum08.Tag = dt.Rows(7).Item("DecimalNum")
            txtFANum08.Enabled = L3Bool(dt.Rows(7).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num09.Text = dt.Rows(8).Item("Data84").ToString
            Else
                Num09.Text = dt.Rows(8).Item("Data01").ToString
            End If
            txtFANum09.Tag = dt.Rows(8).Item("DecimalNum")
            txtFANum09.Enabled = L3Bool(dt.Rows(8).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Num10.Text = dt.Rows(9).Item("Data84").ToString
            Else
                Num10.Text = dt.Rows(9).Item("Data01").ToString
            End If
            txtFANum10.Tag = dt.Rows(9).Item("DecimalNum")
            txtFANum10.Enabled = L3Bool(dt.Rows(9).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date01.Text = dt.Rows(20).Item("Data84").ToString
            Else
                Date01.Text = dt.Rows(20).Item("Data01").ToString
            End If
            c1dateFADate01.Enabled = L3Bool(dt.Rows(20).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date02.Text = dt.Rows(21).Item("Data84").ToString
            Else
                Date02.Text = dt.Rows(21).Item("Data01").ToString
            End If
            c1dateFADate02.Enabled = L3Bool(dt.Rows(21).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date03.Text = dt.Rows(22).Item("Data84").ToString
            Else
                Date03.Text = dt.Rows(22).Item("Data01").ToString
            End If
            c1dateFADate03.Enabled = L3Bool(dt.Rows(22).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date04.Text = dt.Rows(23).Item("Data84").ToString
            Else
                Date04.Text = dt.Rows(23).Item("Data01").ToString
            End If
            c1dateFADate04.Enabled = L3Bool(dt.Rows(23).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date05.Text = dt.Rows(24).Item("Data84").ToString
            Else
                Date05.Text = dt.Rows(24).Item("Data01").ToString
            End If
            c1dateFADate05.Enabled = L3Bool(dt.Rows(24).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date06.Text = dt.Rows(25).Item("Data84").ToString
            Else
                Date06.Text = dt.Rows(25).Item("Data01").ToString
            End If
            c1dateFADate06.Enabled = L3Bool(dt.Rows(25).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date07.Text = dt.Rows(26).Item("Data84").ToString
            Else
                Date07.Text = dt.Rows(26).Item("Data01").ToString
            End If
            c1dateFADate07.Enabled = L3Bool(dt.Rows(26).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date08.Text = dt.Rows(27).Item("Data84").ToString
            Else
                Date08.Text = dt.Rows(27).Item("Data01").ToString
            End If
            c1dateFADate08.Enabled = L3Bool(dt.Rows(27).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date09.Text = dt.Rows(28).Item("Data84").ToString
            Else
                Date09.Text = dt.Rows(28).Item("Data01").ToString
            End If
            c1dateFADate09.Enabled = L3Bool(dt.Rows(28).Item("Disabled"))
            If geLanguage = EnumLanguage.Vietnamese Then
                Date10.Text = dt.Rows(29).Item("Data84").ToString
            Else
                Date10.Text = dt.Rows(29).Item("Data01").ToString
            End If
            c1dateFADate10.Enabled = L3Bool(dt.Rows(29).Item("Disabled"))
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0015
    '# Created User: Lê Đình Thái
    '# Created Date: 08/11/2011 03:37:03
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0015() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0015 "
        sSQL &= SQLString("D02T0100") & COMMA 'TableName, nvarchar, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, nvarchar, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    Private Sub txtFANum01_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum01.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum02_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum02.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum03_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum03.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum04_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum04.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum05_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum05.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum06_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum06.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum07_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum07.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum08_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum08.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum09_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum09.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub
    Private Sub txtFANum10_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFANum10.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Function InsertFormat(ByVal ONumber As Object) As String
        Dim iNumber As Int16 = Convert.ToInt16(ONumber)
        Dim sRet As String = "#,##0"
        If iNumber = 0 Then
        Else
            sRet &= "." & Strings.StrDup(iNumber, "0")
        End If
        Return sRet
    End Function

    Private Sub txtFANum01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum01.LostFocus
        txtFANum01.Text = SQLNumber(IIf(txtFANum01.Text = "", 0, txtFANum01.Text), InsertFormat(txtFANum01.Tag))
    End Sub
    Private Sub txtFANum02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum02.LostFocus
        txtFANum02.Text = SQLNumber(IIf(txtFANum02.Text = "", 0, txtFANum02.Text), InsertFormat(txtFANum02.Tag))
    End Sub
    Private Sub txtFANum03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum03.LostFocus
        txtFANum03.Text = SQLNumber(IIf(txtFANum03.Text = "", 0, txtFANum03.Text), InsertFormat(txtFANum03.Tag))
    End Sub
    Private Sub txtFANum04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum04.LostFocus
        txtFANum04.Text = SQLNumber(IIf(txtFANum04.Text = "", 0, txtFANum04.Text), InsertFormat(txtFANum04.Tag))
    End Sub
    Private Sub txtFANum05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum05.LostFocus
        txtFANum05.Text = SQLNumber(IIf(txtFANum05.Text = "", 0, txtFANum05.Text), InsertFormat(txtFANum05.Tag))
    End Sub
    Private Sub txtFANum06_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum06.LostFocus
        txtFANum06.Text = SQLNumber(IIf(txtFANum06.Text = "", 0, txtFANum06.Text), InsertFormat(txtFANum06.Tag))
    End Sub
    Private Sub txtFANum07_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum07.LostFocus
        txtFANum07.Text = SQLNumber(IIf(txtFANum07.Text = "", 0, txtFANum07.Text), InsertFormat(txtFANum07.Tag))
    End Sub
    Private Sub txtFANum08_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum08.LostFocus
        txtFANum08.Text = SQLNumber(IIf(txtFANum08.Text = "", 0, txtFANum08.Text), InsertFormat(txtFANum08.Tag))
    End Sub
    Private Sub txtFANum09_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum09.LostFocus
        txtFANum09.Text = SQLNumber(IIf(txtFANum09.Text = "", 0, txtFANum09.Text), InsertFormat(txtFANum09.Tag))
    End Sub
    Private Sub txtFANum10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFANum10.LostFocus
        txtFANum10.Text = SQLNumber(IIf(txtFANum10.Text = "", 0, txtFANum10.Text), InsertFormat(txtFANum10.Tag))
    End Sub



    Private Sub txtFAString10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString10.KeyPress
        If (txtFAString10.MaxLength = 0) Then
            e.Handled = True

        End If
    End Sub

    Private Sub txtFAString9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString09.KeyPress
        If (txtFAString09.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString01.KeyPress
        If (txtFAString01.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString02.KeyPress
        If (txtFAString02.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString03.KeyPress
        If (txtFAString03.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString04.KeyPress
        If (txtFAString04.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString05.KeyPress
        If (txtFAString05.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString06.KeyPress
        If (txtFAString06.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString07.KeyPress
        If (txtFAString07.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFAString8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFAString08.KeyPress
        If (txtFAString08.MaxLength = 0) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub

    Private Sub LoadtdbcObjectID(ByVal ID As String)
        LoadDataSource(tdbcCIPObjectID, ReturnTableFilter(dtObjectID, "ObjectTypeID = " & SQLString(ID), True), gbUnicode)
    End Sub

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPObjectTypeID.Close
        If tdbcCIPObjectTypeID.FindStringExact(tdbcCIPObjectTypeID.Text) = -1 Then tdbcCIPObjectTypeID.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPObjectTypeID.SelectedValueChanged
        If Not (tdbcCIPObjectTypeID.Tag Is Nothing OrElse tdbcCIPObjectTypeID.Tag.ToString = "") Then
            tdbcCIPObjectTypeID.Tag = ""
            Exit Sub
        End If
        If tdbcCIPObjectTypeID.SelectedValue Is Nothing Then
            LoadtdbcObjectID("-1")
            Exit Sub
        End If
        LoadtdbcObjectID(tdbcCIPObjectTypeID.SelectedValue.ToString())
        tdbcCIPObjectID.Text = ""
        txtObjectName.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPObjectTypeID.KeyDown
        If e.Alt = True Then
            tdbcCIPObjectTypeID.AutoDropDown = False
        Else
            tdbcCIPObjectTypeID.AutoDropDown = True
        End If
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCIPObjectTypeID.Text = ""
    End Sub

    Private Sub tdbcObjectID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPObjectID.Close
        If tdbcCIPObjectID.FindStringExact(tdbcCIPObjectID.Text) = -1 Then
            tdbcCIPObjectID.Text = ""
            txtObjectName.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPObjectID.SelectedValueChanged
        txtObjectName.Text = tdbcCIPObjectID.Columns(1).Value.ToString()
    End Sub

    Private Sub tdbcObjectID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPObjectID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcCIPObjectID.Text = ""
            txtObjectName.Text = ""
        End If

        'Dim sKeyID As String
        If e.KeyCode = Keys.F2 Then
            'Dim f As New D91F6010
            'f.FormPermision = "D91F6010"
            'f.InListID = "2"
            'f.InWhere = " ObjectTypeID = " & SQLString(tdbcCIPObjectTypeID.Text)
            'f.WhereValue = ""
            'f.ShowDialog()
            'sKeyID = f.OutPut01
            'f.Dispose()
            'If sKeyID <> "" Then
            '    tdbcCIPObjectID.SelectedValue = sKeyID
            '    tdbcCIPObjectID.Focus()
            'End If

            Try
                Dim arrPro() As StructureProperties = Nothing
                SetProperties(arrPro, "InListID", "2")
                SetProperties(arrPro, "InWhere", "ObjectTypeID = " & SQLString(tdbcCIPObjectTypeID.Text))
                SetProperties(arrPro, "WhereValue", "")
                Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
                Dim sKey As String = GetProperties(frm, "Output01").ToString
                If sKey <> "" Then
                    'Load dữ liệu
                    tdbcCIPObjectID.SelectedValue = sKey
                    tdbcCIPObjectID.Focus()
                End If
            Catch ex As Exception
                D99C0008.MsgL3(ex.Message)
            End Try
        End If
    End Sub

#End Region

#Region "Events tdbcEmployeeID with txtEmployeeName"

    Private Sub tdbcEmployeeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCIPEmployeeID.Close
        If tdbcCIPEmployeeID.FindStringExact(tdbcCIPEmployeeID.Text) = -1 Then
            tdbcCIPEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If
    End Sub

    Private Sub tdbcEmployeeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCIPEmployeeID.SelectedValueChanged
        txtEmployeeName.Text = tdbcCIPEmployeeID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcEmployeeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPEmployeeID.KeyDown
        'Dim sKeyID As String

        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcCIPEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If

        If e.KeyCode = Keys.F2 Then
            'Dim f As New D91F6010
            'f.FormPermision = "D91F6010"
            'f.InListID = "2"
            'f.InWhere = " ObjectTypeID ='NV' "
            'f.WhereValue = ""
            'f.ShowDialog()
            'sKeyID = f.OutPut01
            'f.Dispose()
            'If sKeyID <> "" Then
            '    tdbcCIPEmployeeID.SelectedValue = sKeyID
            '    tdbcCIPEmployeeID.Focus()
            'End If

            Try
                Dim arrPro() As StructureProperties = Nothing
                SetProperties(arrPro, "InListID", "2")
                SetProperties(arrPro, "InWhere", "ObjectTypeID = 'NV'")
                SetProperties(arrPro, "WhereValue", "")
                Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
                Dim sKey As String = GetProperties(frm, "Output01").ToString
                If sKey <> "" Then
                    'Load dữ liệu
                    tdbcCIPEmployeeID.SelectedValue = sKey
                    tdbcCIPEmployeeID.Focus()
                End If
            Catch ex As Exception
                D99C0008.MsgL3(ex.Message)
            End Try
        End If
    End Sub

#End Region

    Private Sub D02F2001_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub


    Private Sub D02F2001_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If _FormState = EnumFormState.FormAdd Then
            If D02Systems.CIPAuto = 2 Then
                tdbcCIPS1ID.Focus()
            Else
                txtCipNo.Focus() 'txtCipNo.Focus()
            End If
        Else
            txtCipName.Focus()
        End If

    End Sub


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1000
    '# Created User: Trần Hoàng Nhân
    '# Created Date: 14/11/2012 10:02:26
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Insert du lieu " & vbCrLf)
        sSQL.Append("Insert Into D91T1000(")
        sSQL.Append("UserID, HostID, ModuleID, FormID ")
        sSQL.Append(") Values(" & vbCrLf)
        sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[50], NOT NULL
        sSQL.Append(SQLString(My.Computer.Name) & COMMA) 'HostID, varchar[50], NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[50], NOT NULL
        sSQL.Append(SQLString("D02F0060")) 'FormID, varchar[50], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P1000
    '# Created User: Trần Hoàng Nhân
    '# Created Date: 14/11/2012 10:06:13
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P1000() As String
    '    Dim sSQL As String = ""
    '    sSQL &= ("-- Sinh Ma cap nhat" & vbCrlf)
    '    sSQL &= "Exec D91P1000 "
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
    '    sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
    '    sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
    '    sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
    '    sSQL &= SQLString("D02F0060") & COMMA 'FormID, varchar[20], NOT NULL
    '    sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[20], NOT NULL
    '    sSQL &= SQLString(tdbcIGEMethodID.Text) & COMMA 'IGEMethodID, varchar[50], NOT NULL
    '    sSQL &= SQLNumber(1) & COMMA 'AutoCreateName, tinyint, NOT NULL
    '    sSQL &= SQLNumber(20) & COMMA 'Length, tinyint, NOT NULL
    '    sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL

    '    Return sSQL
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T1001
    '# Created User: Trần Hoàng Nhân
    '# Created Date: 16/11/2012 04:05:12
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Update KeyString & Lastkey" & vbCrlf)
        sSQL.Append("Insert Into D91T1001(")
        sSQL.Append("KeyString, LastKey, ModuleID, FormID")
        sSQL.Append(") Values(" & vbCrlf)
        sSQL.Append(SQLString(sKeyString) & COMMA) 'KeyString, varchar[250], NOT NULL
        sSQL.Append(SQLNumber(sLastKey) & COMMA) 'LastKey, int, NOT NULL
        sSQL.Append(SQLString("02") & COMMA) 'ModuleID, varchar[20], NOT NULL
        sSQL.Append(SQLString("D02F0060")) 'FormID, varchar[20], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD91T1001
    '# Created User: Trần Hoàng Nhân
    '# Created Date: 16/11/2012 04:06:55
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD91T1001() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("-- Update KeyString & Lastkey" & vbCrlf)
        sSQL.Append("Update D91T1001 Set ")
        sSQL.Append("LastKey = " & SQLNumber(sLastKey)) 'varchar[250], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("KeyString = " & SQLString(sKeyString)) 'int, NOT NULL
        sSQL.Append(" And ModuleID = " & SQLString("02")) 'varchar[20], NOT NULL
        sSQL.Append(" And FormID = " & SQLString("D02F0060")) 'varchar[20], NOT NULL

        Return sSQL
    End Function



    Private Sub btnF5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnF5.Click
        Dim bIsAddNew As Boolean
        Dim f As New D02F0005()
        With f
            .FormID = "D02F2001"
            .IGEMethodID = _iGEMethodID
            If dtXCodeValue IsNot Nothing Then
                .dtXCodeValue = dtXCodeValue
            End If
            .ShowInTaskbar = True
            .ShowDialog()
            If .SavedOk Then
                sID = .ID
                sKeyString = .KeyString
                sKeyName = .KeyName
                sLastKey = .LastKey
                bIsAutoNum = .IsAutoNum
                _iGEMethodID = .IGEMethodID
                dtXCodeValue = .dtXCodeValue
                dtXCode = .dtXCode
                '5/11/2019, Đặng Ngọc Tài:id 126370-Thêm nhanh tiêu thức tạo mã Xây Dựng cơ bản
                bIsAddNew = .IsAddNew
                If bIsAddNew Then LoadTDBCACodeID()

                LoadDefaultTDBCACodeID() ' Default giá trị đư
            End If
        End With
        f.Dispose()
    End Sub

    Private Sub btnK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnK.Click
        Dim form As New D02F5705()
        With form
            .LastKey = sLastKey
            .IGEMethodID = _iGEMethodID
            .ShowDialog()
            If .SavedOk Then
                If .ID <> "" Then
                    txtCipNo.Text = .ID
                    sLastKey = .LastKey
                End If
            End If
        End With
        form.Dispose()
    End Sub


    Private Sub GetAutoCIPInfo()
        If D02Systems.CIPAuto = 0 Then
            tdbcCIPS1ID.Visible = False
            tdbcCIPS2ID.Visible = False
            tdbcCIPS3ID.Visible = False
            txtCipNo.ReadOnly = False
            txtCIPNo.TabStop = True
            btnK.Enabled = False
            btnF5.Visible = False
            pnlCode.Location = tdbcCIPS1ID.Location
        ElseIf D02Systems.CIPAuto = 1 Then
            tdbcCIPS1ID.Visible = False
            tdbcCIPS2ID.Visible = False
            tdbcCIPS3ID.Visible = False
            txtCipNo.ReadOnly = True
            txtCipNo.TabStop = False
            btnK.Enabled = True
            btnF5.Visible = True
            pnlCode.Location = tdbcCIPS1ID.Location

            
        Else
            'tdbcCIPS1ID.Visible = True
            'tdbcCIPS2ID.Visible = True
            'tdbcCIPS3ID.Visible = True
            txtCipNo.ReadOnly = True
            txtCipNo.TabStop = False
            btnK.Enabled = True
            btnF5.Visible = False
            If D02Systems.CIPS1Enabled Then
                tdbcCIPS1ID.Enabled = True
                tdbcCIPS1ID.Text = D02Systems.CIPS1Default
                tdbcCIPS1ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcCIPS1ID.Enabled = False
            End If
            If D02Systems.CIPS2Enabled Then
                tdbcCIPS2ID.Enabled = True
                tdbcCIPS2ID.Text = D02Systems.CIPS2Default
                tdbcCIPS2ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcCIPS2ID.Enabled = False
            End If
            If D02Systems.CIPS3Enabled Then
                tdbcCIPS3ID.Enabled = True
                tdbcCIPS3ID.Text = D02Systems.CIPS3Default
                tdbcCIPS3ID_SelectedValueChanged(Nothing, Nothing)
            Else
                tdbcCIPS3ID.Enabled = False
            End If
        End If
    End Sub

    Private _S1 As String = ""
    Private _S2 As String = ""
    Private _S3 As String = ""
    Private _CIPOutputOrder As String = ""
    Private _CIPOutputLength As Integer
    Private _CIPSeperated As Boolean
    Private _CIPSeperator As String = ""
    Private _CIPTableName As String = "D02T0100"

#Region "CIPS1ID"
    Private Sub tdbcCIPS1ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPS1ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If D02Systems.CIPAuto = 2 Then
                gnNewLastKey = 0
                _S1 = IIf(IsDBNull(tdbcCIPS1ID.Text) Or tdbcCIPS1ID.Text = "<<", "", tdbcCIPS1ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _CIPOutputOrder, _CIPOutputLength, _CIPSeperator, txtCipNo, False, _CIPTableName)
            End If
        End If
    End Sub

    Private Sub tdbcCIPS1ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPS1ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCIPS1ID.Text = ""
    End Sub

    Private Sub tdbcCIPS1ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCIPS1ID.Close
        If tdbcCIPS1ID.FindStringExact(tdbcCIPS1ID.Text) = -1 Then
            tdbcCIPS1ID.Text = ""
            Exit Sub
        End If
        If tdbcCIPS1ID.Text = "<<" Then
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "IndexTab", 0)
            CallFormShow(Me, "D02D1240", "D02F3001", arrPro)
            LoadTdbcCIP1()
            tdbcCIPS1ID.Text = "<<"
        End If
    End Sub
#End Region

#Region "CIPS2ID"
    Private Sub tdbcCIPS2ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPS2ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If D02Systems.CIPAuto = 2 Then
                gnNewLastKey = 0
                _S2 = IIf(IsDBNull(tdbcCIPS2ID.Text) Or tdbcCIPS2ID.Text = "<<", "", tdbcCIPS2ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _CIPOutputOrder, _CIPOutputLength, _CIPSeperator, txtCipNo, False, _CIPTableName)
            End If
        End If
    End Sub

    Private Sub tdbcCIPS2ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPS2ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCIPS2ID.Text = ""
    End Sub

    Private Sub tdbcCIPS2ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCIPS2ID.Close
        If tdbcCIPS2ID.FindStringExact(tdbcCIPS2ID.Text) = -1 Then
            tdbcCIPS2ID.Text = ""
            Exit Sub
        End If
        If tdbcCIPS2ID.Text = "<<" Then
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "IndexTab", 0)
            CallFormShow(Me, "D02D1240", "D02F3001", arrPro)
            LoadTdbcCIP1()
            tdbcCIPS2ID.Text = "<<"
        End If
    End Sub
#End Region


#Region "CIPS3ID"

    Private Sub tdbcCIPS3ID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCIPS3ID.SelectedValueChanged
        If _FormState <> EnumFormState.FormEdit Then
            If D02Systems.CIPAuto = 2 Then
                gnNewLastKey = 0
                _S3 = IIf(IsDBNull(tdbcCIPS3ID.Text) Or tdbcCIPS3ID.Text = "<<", "", tdbcCIPS3ID.Text).ToString
                D02X0002.GetNewVoucherNo(_S1, _S2, _S3, _CIPOutputOrder, _CIPOutputLength, _CIPSeperator, txtCipNo, False, _CIPTableName)
            End If
        End If
    End Sub

    Private Sub tdbcCIPS3ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCIPS3ID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCIPS3ID.Text = ""
    End Sub

    Private Sub tdbcCIPS3ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCIPS3ID.Close
        If tdbcCIPS3ID.FindStringExact(tdbcCIPS3ID.Text) = -1 Then
            tdbcCIPS3ID.Text = ""
            Exit Sub
        End If
        If tdbcCIPS3ID.Text = "<<" Then
            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "IndexTab", 0)
            CallFormShow(Me, "D02D1240", "D02F3001", arrPro)
            LoadTdbcCIP1()
            tdbcCIPS3ID.Text = "<<"
        End If
    End Sub
#End Region


#Region "LoadCIPCombo"

    Private Sub LoadTdbcCIP1()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        If gsLanguage = "84" Then
            sSQL = "Select     '<<' as CIPS1ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' As CIPS1Name " & vbCrLf
            sSQL &= "Union" & vbCrLf
            sSQL &= "Select     AssetS1ID as CIPS1ID, AssetS1Name" & sUnicode & " As CIPS1Name" & vbCrLf
            sSQL &= "From       D02T1000 WITH(NOLOCK)" & vbCrLf
            sSQL &= "Where      Disabled = 0" & vbCrLf
            sSQL &= "Order by   CIPS1ID" & vbCrLf
        Else
            sSQL = "Select     '<<' as CIPS1ID, 'Add' as CIPS1Name" & vbCrLf
            sSQL &= "Union" & vbCrLf
            sSQL &= "Select     AssetS1ID as CIPS1ID, AssetS1Name" & sUnicode & " As CIPS1Name" & vbCrLf
            sSQL &= "From       D02T1000 WITH(NOLOCK) " & vbCrLf
            sSQL &= "Where      Disabled = 0" & vbCrLf
            sSQL &= "Order by   CIPS1ID" & vbCrLf
        End If
        LoadDataSource(tdbcCIPS1ID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTdbcCIP2()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        If gsLanguage = "84" Then
            sSQL = "Select '<<' as CIPS2ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' as CIPS2Name Union Select AssetS2ID as CIPS2ID, AssetS2Name" & sUnicode & " As CIPS2Name From D02T2000 WITH(NOLOCK) Where Disabled=0 Order by CIPS2ID"
        Else
            sSQL = "Select '<<' as CIPS2ID, 'Add' as CIPS2Name Union Select AssetS2ID as CIPS2ID, AssetS2Name" & sUnicode & " As CIPS2Name From D02T2000 WITH(NOLOCK) Where Disabled=0 Order by CIPS2ID"
        End If
        LoadDataSource(tdbcCIPS2ID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTdbcCIP3()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        If gsLanguage = "84" Then
            sSQL = "Select '<<' as CIPS3ID, " & IIf(gbUnicode, "N'<Thêm mới", "'<Theâm môùi").ToString() & ">' as CIPS3Name Union Select AssetS3ID as CIPS3ID, AssetS3Name" & sUnicode & " As CIPS3Name From D02T3000 WITH(NOLOCK) Where Disabled=0 Order by CIPS3ID"
        Else
            sSQL = "Select '<<' as CIPS3ID, 'Add' as CIPS3Name Union Select AssetS3ID as CIPS3ID, AssetS3Name" & sUnicode & " As CIPS3Name From D02T3000 WITH(NOLOCK) Where Disabled=0 Order by  CIPS3ID"
        End If
        LoadDataSource(tdbcCIPS3ID, sSQL, gbUnicode)
    End Sub
#End Region


 

End Class