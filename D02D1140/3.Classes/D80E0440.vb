Imports System

''' <summary>
''' Các màn hình của exe con D80E0440
''' </summary>
Public Enum D80E0440Form
    ''' <summary>
    ''' D80F2090: Nhập dữ liệu từ Excel
    ''' </summary>
    D80F2090 = 1
End Enum

Public Class D80E0440
    Private Const EXEMODULE As String = "D80"
    Private Const EXECHILD As String = "D80E0440"
    Private sLanguage As String

    ''' <summary>
    ''' Khởi tạo exe con D80E0440
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String)
        sLanguage = Language
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", Server, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", Database, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", UserDatabaseID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", Password, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", UserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", Language)
'Update 20/10/2010: Lưu biến nhập liệu Unicode
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "CodeTable", gbUnicode.ToString)
    End Sub

    ''' <summary>
    ''' Khởi tạo exe con D80E0440
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    ''' <param name="DivisionID">Đơn vị hiện tại</param>
    ''' <param name="TranMonth">Tháng kế toán hiện tại</param>
    ''' <param name="TranYear">Năm kế toán hiện tại</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer)
        Me.New(Server, Database, UserDatabaseID, Password, UserID, Language)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", DivisionID)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", TranMonth.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", TranYear.ToString)
    End Sub

    ''' <summary>
    ''' Màn hình cần hiển thị cho exe con
    ''' </summary>
    Public WriteOnly Property FormActive() As D80E0440Form
        Set(ByVal Value As D80E0440Form)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", [Enum].GetName(GetType(D80E0440Form), Value))
        End Set
    End Property

    ''' <summary>
    ''' Form Phân quyền cho màn hình được gọi
    ''' </summary>
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", Value.ToString)
        End Set
    End Property

    ''' <summary>
    ''' Trạng thái Form màn hình : AddNew , Edit or View
    ''' </summary>
    Public WriteOnly Property FormStatus() As EnumFormState
        Set(ByVal Value As EnumFormState)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl02", CByte(Value).ToString)
            'D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl02", [Enum].GetName(GetType(EnumFormState), Value))
        End Set
    End Property

    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ModuleID", Value.ToString)
        End Set
    End Property

    Public WriteOnly Property TransTypeID() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID01", Value.ToString)
        End Set
    End Property

    Public WriteOnly Property sFont() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID02", Value.ToString)
        End Set
    End Property

    ''' <summary>
    ''' Kết quả trả về
    ''' </summary>
    Public ReadOnly Property Output01() As Boolean
        Get
            Return CBool(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output01", "False"))
        End Get
    End Property

    ''' <summary>
    ''' Thực thi exe con
    ''' </summary>
    Public Sub Run()
        If Not ExistFile(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        SaveRunningExeSettings("D10E2040", EXECHILD)
        Process.Start(pInfo)
    End Sub

    ''' <summary>
    ''' Kiểm tra tồn tại exe con không ?
    ''' </summary>
    Private Function ExistFile(ByVal Path As String) As Boolean
        If System.IO.File.Exists(Path) Then Return True
        If sLanguage = "0" Then
            D99C0008.MsgL3("Không tồn tại file " & EXECHILD & ".exe")
        Else
            D99C0008.MsgL3("Not exist file" & EXECHILD & ".exe")
        End If
        Return False
    End Function

End Class
