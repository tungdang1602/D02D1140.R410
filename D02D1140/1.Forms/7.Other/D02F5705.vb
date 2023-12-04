'#----------------------------------------------------------
'#Title: D99F1111
'#CreateUser:Nguyễn Thị Minh Hßa
'#CreateDate: 18/10/05
'#ModifiedUser: Nguyễn Thị Minh Hßa
'#ModifiedDate: 25/03/2008
'#Description: Form tạo mã tự động cho chứng từ
'#----------------------------------------------------------
Imports C1.Win
Imports C1.Win.C1List
Imports System.Data.SqlClient
Imports System

Public Class D02F5705
    Inherits System.Windows.Forms.Form

    Friend WithEvents txtAutoNum As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Private WithEvents grp1 As System.Windows.Forms.GroupBox

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F5705))
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.txtAutoNum = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.grp1 = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(126, 97)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(70, 24)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Đó&ng"
        '
        'btnOK
        '
        Me.btnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(33, 97)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(70, 24)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "Đồn&g ý"
        '
        'txtAutoNum
        '
        Me.txtAutoNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAutoNum.Location = New System.Drawing.Point(33, 45)
        Me.txtAutoNum.Name = "txtAutoNum"
        Me.txtAutoNum.Size = New System.Drawing.Size(163, 20)
        Me.txtAutoNum.TabIndex = 6
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.Location = New System.Drawing.Point(30, 13)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(181, 20)
        Me.lbl1.TabIndex = 7
        Me.lbl1.Text = "Chỉ số mã tự động hiện tại"
        '
        'grp1
        '
        Me.grp1.Location = New System.Drawing.Point(1, 84)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(236, 7)
        Me.grp1.TabIndex = 8
        Me.grp1.TabStop = False
        '
        'D02F5705
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(235, 130)
        Me.Controls.Add(Me.grp1)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.txtAutoNum)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F5705"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "˜Æt chÙ sç mº tø ¢èng"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private _iGEMethodID As String = ""
    Public WriteOnly Property IGEMethodID() As String
        Set(ByVal Value As String)
            _iGEMethodID = Value
        End Set
    End Property

    Private _formID As String = "D02F0060"
    Public WriteOnly Property FormID() As String
        Set(ByVal Value As String)
            _formID = Value
        End Set
    End Property

    Private _moduleID As String = "02"
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

    Private _savedOK As Boolean = False
    Public ReadOnly Property SavedOk() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private _iD As String = ""
    Public ReadOnly Property ID() As String
        Get
            Return _iD
        End Get
    End Property


    Private _lastKey As Integer
    Public Property  LastKey() As Integer
        Get
            Return _lastKey
        End Get
        Set(ByVal Value As Integer)
            _lastKey = Value
        End Set
    End Property

    Private Sub D07F5705_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	LoadInfoGeneral()
        LoadLanguage()      
        txtAutoNum.Text = _lastKey.ToString
        SetResolutionForm(Me)

    End Sub

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rL3("Dat_chi_so_ma_tu_dong") ' & " - " & Me.Name & UnicodeCaption(gbUnicode) '˜Æt chÙ sç mº tø ¢èng
        '================================================================ 
        lbl1.Text = rL3("Chi_so_ma_tu_dong_hien_tai") 'Chỉ số mã tự động hiện tại:
        '================================================================ 
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        btnOK.Text = rL3("Don_g_y") 'Đồn&g ý
        '================================================================ 
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        _savedOK = False
        Me.Close()
    End Sub

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
        sSQL &= SQLString(_formID) & COMMA 'FormID, varchar[20], NOT NULL
        sSQL &= SQLString(_moduleID) & COMMA 'ModuleID, varchar[20], NOT NULL
        sSQL &= SQLString(_iGEMethodID) & COMMA 'IGEMethodID, varchar[50], NOT NULL
        sSQL &= SQLNumber("0") & COMMA 'AutoCreateName, tinyint, NOT NULL
        sSQL &= SQLNumber("50") & COMMA 'Length, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLNumber("1") & COMMA 'IsD07F0011, tinyint, NOT NULL
        sSQL &= SQLNumber(txtAutoNum.Text) 'NewLastKey, int, NOT NULL
        Return sSQL
    End Function


    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If txtAutoNum.Text = "" Then
            If geLanguage = EnumLanguage.English Then
                D99C0008.MsgNotYetEnter("Code auto")
            Else
                D99C0008.MsgNotYetEnter("Chỉ số tự động")
            End If
            txtAutoNum.Focus()
            Exit Sub
        End If

        _savedOK = True
        Dim dtD91P1000 As DataTable = ReturnDataTable(SQLStoreD91P1000)
        If dtD91P1000.Rows.Count > 0 Then
            With dtD91P1000.Rows(0)
                Select Case Convert.ToInt16(.Item("Status"))
                    Case 0
                        _iD = .Item("ID").ToString
                        _lastKey = L3Int(.Item("LastKey").ToString)
                    Case 1                      
                        D99C0008.MsgL3(ConvertVietwareFToUnicode(.Item("Message").ToString))
                End Select
            End With
        End If
        dtD91P1000.Dispose()
        Me.Close()
    End Sub

    Private Sub txtLastKey_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAutoNum.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

End Class
