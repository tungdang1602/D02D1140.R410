Imports System.Text
Public Class D02F0050

    Dim dt1 As DataTable
    Dim sLanguage As Boolean


    Private Sub LoadForm()
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, gbUnicode)

        Dim sSQL As String = ""
        If geLanguage = EnumLanguage.Vietnamese Then

            sSQL = "Select TypeCodeID, Disabled, MaxLength, VieTypeCodeName" & sUnicode & " As Description" & vbCrLf
            sSQL &= "From D02T0040 WITH(NOLOCK)  " & vbCrLf
            sSQL &= "Where Type = 'X' " & vbCrLf
            sSQL &= "Order by TypeCodeID"
            dt1 = ReturnDataTable(sSQL)
            If dt1.Rows.Count > 0 Then
                txtTypeCodeID_0.Text = dt1.Rows(0).Item("TypeCodeID").ToString
                txtDescription_0.Text = dt1.Rows(0).Item("Description").ToString
                chkDisabled_0.Checked = Convert.ToBoolean(dt1.Rows(0).Item("Disabled"))
                txtMaxLength_0.Text = dt1.Rows(0).Item("MaxLength").ToString

                txtTypeCodeID_1.Text = dt1.Rows(1).Item("TypeCodeID").ToString
                txtDescription_1.Text = dt1.Rows(1).Item("Description").ToString
                chkDisabled_1.Checked = Convert.ToBoolean(dt1.Rows(1).Item("Disabled"))
                txtMaxLength_1.Text = dt1.Rows(1).Item("MaxLength").ToString

                txtTypeCodeID_2.Text = dt1.Rows(2).Item("TypeCodeID").ToString
                txtDescription_2.Text = dt1.Rows(2).Item("Description").ToString
                chkDisabled_2.Checked = Convert.ToBoolean(dt1.Rows(2).Item("Disabled"))
                txtMaxLength_2.Text = dt1.Rows(2).Item("MaxLength").ToString

                txtTypeCodeID_3.Text = dt1.Rows(3).Item("TypeCodeID").ToString
                txtDescription_3.Text = dt1.Rows(3).Item("Description").ToString
                chkDisabled_3.Checked = Convert.ToBoolean(dt1.Rows(3).Item("Disabled"))
                txtMaxLength_3.Text = dt1.Rows(3).Item("MaxLength").ToString

                txtTypeCodeID_4.Text = dt1.Rows(4).Item("TypeCodeID").ToString
                txtDescription_4.Text = dt1.Rows(4).Item("Description").ToString
                chkDisabled_4.Checked = Convert.ToBoolean(dt1.Rows(4).Item("Disabled"))
                txtMaxLength_4.Text = dt1.Rows(4).Item("MaxLength").ToString

                txtTypeCodeID_5.Text = dt1.Rows(5).Item("TypeCodeID").ToString
                txtDescription_5.Text = dt1.Rows(5).Item("Description").ToString
                chkDisabled_5.Checked = Convert.ToBoolean(dt1.Rows(5).Item("Disabled"))
                txtMaxLength_5.Text = dt1.Rows(5).Item("MaxLength").ToString

                txtTypeCodeID_6.Text = dt1.Rows(6).Item("TypeCodeID").ToString
                txtDescription_6.Text = dt1.Rows(6).Item("Description").ToString
                chkDisabled_6.Checked = Convert.ToBoolean(dt1.Rows(6).Item("Disabled"))
                txtMaxLength_6.Text = dt1.Rows(6).Item("MaxLength").ToString

                txtTypeCodeID_7.Text = dt1.Rows(7).Item("TypeCodeID").ToString
                txtDescription_7.Text = dt1.Rows(7).Item("Description").ToString
                chkDisabled_7.Checked = Convert.ToBoolean(dt1.Rows(7).Item("Disabled"))
                txtMaxLength_7.Text = dt1.Rows(7).Item("MaxLength").ToString

                txtTypeCodeID_8.Text = dt1.Rows(8).Item("TypeCodeID").ToString
                txtDescription_8.Text = dt1.Rows(8).Item("Description").ToString
                chkDisabled_8.Checked = Convert.ToBoolean(dt1.Rows(8).Item("Disabled"))
                txtMaxLength_8.Text = dt1.Rows(8).Item("MaxLength").ToString

                txtTypeCodeID_9.Text = dt1.Rows(9).Item("TypeCodeID").ToString
                txtDescription_9.Text = dt1.Rows(9).Item("Description").ToString
                chkDisabled_9.Checked = Convert.ToBoolean(dt1.Rows(9).Item("Disabled"))
                txtMaxLength_9.Text = dt1.Rows(9).Item("MaxLength").ToString
            End If
        ElseIf geLanguage = EnumLanguage.English Then
            sSQL = "Select TypeCodeID, Disabled, MaxLength, EngTypeCodeName" & sUnicode & " As Description" & vbCrLf
            sSQL &= "From D02T0040 WITH(NOLOCK)  " & vbCrLf
            sSQL &= "Where Type = 'X' " & vbCrLf
            sSQL &= "Order by TypeCodeID"
            dt1 = ReturnDataTable(sSQL)
            If dt1.Rows.Count > 0 Then
                txtTypeCodeID_0.Text = dt1.Rows(0).Item("TypeCodeID").ToString
                txtDescription_0.Text = dt1.Rows(0).Item("Description").ToString
                chkDisabled_0.Checked = Convert.ToBoolean(dt1.Rows(0).Item("Disabled"))
                txtMaxLength_0.Text = dt1.Rows(0).Item("MaxLength").ToString

                txtTypeCodeID_1.Text = dt1.Rows(1).Item("TypeCodeID").ToString
                txtDescription_1.Text = dt1.Rows(1).Item("Description").ToString
                chkDisabled_1.Checked = Convert.ToBoolean(dt1.Rows(1).Item("Disabled"))
                txtMaxLength_1.Text = dt1.Rows(1).Item("MaxLength").ToString

                txtTypeCodeID_2.Text = dt1.Rows(2).Item("TypeCodeID").ToString
                txtDescription_2.Text = dt1.Rows(2).Item("Description").ToString
                chkDisabled_2.Checked = Convert.ToBoolean(dt1.Rows(2).Item("Disabled"))
                txtMaxLength_2.Text = dt1.Rows(2).Item("MaxLength").ToString

                txtTypeCodeID_3.Text = dt1.Rows(3).Item("TypeCodeID").ToString
                txtDescription_3.Text = dt1.Rows(3).Item("Description").ToString
                chkDisabled_3.Checked = Convert.ToBoolean(dt1.Rows(3).Item("Disabled"))
                txtMaxLength_3.Text = dt1.Rows(3).Item("MaxLength").ToString

                txtTypeCodeID_4.Text = dt1.Rows(4).Item("TypeCodeID").ToString
                txtDescription_4.Text = dt1.Rows(4).Item("Description").ToString
                chkDisabled_4.Checked = Convert.ToBoolean(dt1.Rows(4).Item("Disabled"))
                txtMaxLength_4.Text = dt1.Rows(4).Item("MaxLength").ToString

                txtTypeCodeID_5.Text = dt1.Rows(5).Item("TypeCodeID").ToString
                txtDescription_5.Text = dt1.Rows(5).Item("Description").ToString
                chkDisabled_5.Checked = Convert.ToBoolean(dt1.Rows(5).Item("Disabled"))
                txtMaxLength_5.Text = dt1.Rows(5).Item("MaxLength").ToString

                txtTypeCodeID_6.Text = dt1.Rows(6).Item("TypeCodeID").ToString
                txtDescription_6.Text = dt1.Rows(6).Item("Description").ToString
                chkDisabled_6.Checked = Convert.ToBoolean(dt1.Rows(6).Item("Disabled"))
                txtMaxLength_6.Text = dt1.Rows(6).Item("MaxLength").ToString

                txtTypeCodeID_7.Text = dt1.Rows(7).Item("TypeCodeID").ToString
                txtDescription_7.Text = dt1.Rows(7).Item("Description").ToString
                chkDisabled_7.Checked = Convert.ToBoolean(dt1.Rows(7).Item("Disabled"))
                txtMaxLength_7.Text = dt1.Rows(7).Item("MaxLength").ToString

                txtTypeCodeID_8.Text = dt1.Rows(8).Item("TypeCodeID").ToString
                txtDescription_8.Text = dt1.Rows(8).Item("Description").ToString
                chkDisabled_8.Checked = Convert.ToBoolean(dt1.Rows(8).Item("Disabled"))
                txtMaxLength_8.Text = dt1.Rows(8).Item("MaxLength").ToString

                txtTypeCodeID_9.Text = dt1.Rows(9).Item("TypeCodeID").ToString
                txtDescription_9.Text = dt1.Rows(9).Item("Description").ToString
                chkDisabled_9.Checked = Convert.ToBoolean(dt1.Rows(9).Item("Disabled"))
                txtMaxLength_9.Text = dt1.Rows(9).Item("MaxLength").ToString
            End If

        End If

    End Sub

    Private Function AllowSave() As Boolean
        Dim i As Integer = 0
        For i = 0 To 9
            Select Case i
                Case 0
                    If CInt(txtMaxLength_0.Text) < 1 Or CInt(txtMaxLength_0.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_0.Text = rl3("1")
                        txtMaxLength_0.Focus()
                        Return False

                    End If
                Case 1
                    If CInt(txtMaxLength_1.Text) < 1 Or CInt(txtMaxLength_1.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_1.Text = rl3("1")
                        txtMaxLength_1.Focus()
                        Return False



                    End If
                Case 2
                    If CInt(txtMaxLength_2.Text) < 1 Or CInt(txtMaxLength_2.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_2.Text = rl3("1")
                        txtMaxLength_2.Focus()
                        Return False
                    End If
                Case 3
                    If CInt(txtMaxLength_3.Text) < 1 Or CInt(txtMaxLength_3.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_3.Text = rl3("1")
                        txtMaxLength_3.Focus()
                        Return False
                    End If
                Case 4
                    If CInt(txtMaxLength_4.Text) < 1 Or CInt(txtMaxLength_4.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_4.Text = rl3("1")
                        txtMaxLength_4.Focus()
                        Return False
                    End If
                Case 5
                    If CInt(txtMaxLength_5.Text) < 1 Or CInt(txtMaxLength_5.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_5.Text = rl3("1")
                        txtMaxLength_5.Focus()
                        Return False
                    End If
                Case 6
                    If CInt(txtMaxLength_6.Text) < 1 Or CInt(txtMaxLength_6.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_6.Text = rl3("1")
                        txtMaxLength_6.Focus()
                        Return False
                    End If
                Case 7
                    If CInt(txtMaxLength_7.Text) < 1 Or CInt(txtMaxLength_7.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_7.Text = rl3("1")
                        txtMaxLength_7.Focus()
                        Return False
                    End If
                Case 8
                    If CInt(txtMaxLength_8.Text) < 1 Or CInt(txtMaxLength_8.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_8.Text = rl3("1")
                        txtMaxLength_8.Focus()
                        Return False
                    End If
                Case 9
                    If CInt(txtMaxLength_9.Text) < 1 Or CInt(txtMaxLength_9.Text) > 20 Then
                        D99C0008.MsgL3(rl3("Chieu_dai_khong_hop_le"))
                        txtMaxLength_9.Text = rl3("1")
                        txtMaxLength_9.Focus()
                        Return False
                    End If

            End Select

        Next

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

        sSQL.Append(SQLUpdateD02T0040.ToString)

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnClose.Enabled = True
            btnSave.Enabled = True

        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0040
    '# Created User: Mỹ Vân
    '# Created Date: 22/10/2007 11:57:05
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0040() As StringBuilder
        Dim sSQL As New StringBuilder
        Dim i As Integer = 0

        For i = 0 To 9
            sSQL.Append("Update D02T0040 Set ")
            sSQL.Append("Type = " & SQLString("X") & COMMA) 'varchar[1], NULL
            sSQL.Append("FieldName = " & SQLString("") & COMMA) 'varchar[20], NULL
            If geLanguage = EnumLanguage.Vietnamese Then
                Select Case i
                    Case 0
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_0.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_0.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_0.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_0.Text))
                    Case 1
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_1.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_1.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_1.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_1.Text))
                    Case 2
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_2.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_2.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_2.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_2.Text))
                    Case 3
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_3.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_3.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_3.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_3.Text))
                    Case 4
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_4.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_4.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_4.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_4.Text))
                    Case 5
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_5.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_5.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_5.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_5.Text))
                    Case 6
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_6.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_6.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_6.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_6.Text))
                    Case 7
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_7.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_7.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_7.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_7.Text))
                    Case 8
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_8.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_8.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_8.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_8.Text))
                    Case 9
                        sSQL.Append("VieTypeCodeNameU = " & SQLStringUnicode(txtDescription_9.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_9.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_9.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_9.Text))

                End Select
            Else
                Select Case i
                    Case 0

                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_0.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_0.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_0.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_0.Text))
                    Case 1
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_1.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_1.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_1.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_1.Text))
                    Case 2
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_2.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_2.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_2.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_2.Text))
                    Case 3
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_3.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_3.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_3.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_3.Text))
                    Case 4
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_4.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_4.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_4.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_4.Text))
                    Case 5
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_5.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_5.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_5.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_5.Text))
                    Case 6
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_6.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_6.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_6.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_6.Text))
                    Case 7
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_7.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_7.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_7.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_7.Text))
                    Case 8
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_8.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_8.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_8.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_8.Text))
                    Case 9
                        sSQL.Append("EngTypeCodeNameU = " & SQLStringUnicode(txtDescription_9.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
                        sSQL.Append("Disabled = " & SQLNumber(chkDisabled_9.Checked) & COMMA) 'bit, NOT NULL
                        sSQL.Append("MaxLength = " & SQLNumber(txtMaxLength_9.Text) & COMMA) 'tinyint, NOT NULL
                        sSQL.Append("CreateDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("CreateUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
                        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID)) 'varchar[20], NULL
                        sSQL.Append(" Where ")
                        sSQL.Append("TypeCodeID = " & SQLString(txtTypeCodeID_9.Text))
                End Select
            End If
        Next
        Return sSQL
    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtMaxLength_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_0.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_1.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_2.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_3.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_4.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_5.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_6.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_7.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_8.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub txtMaxLength_9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxLength_9.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
    End Sub

    Private Sub D02F0050_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
            Exit Sub
        End If
    End Sub

    Private Sub D02F0050_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        LoadForm()
        Loadlanguage()
        If gsLanguage = "01" Then
            lblDisabled.Left = lblDisabled.Left + 15
        End If
        btnSave.Enabled = ReturnPermission(Me.Name) > EnumPermission.View
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Ma_loai_phan_tich_XDCB_-_D02F0050") & UnicodeCaption(gbUnicode)  'Mº loÁi ph¡n tÛch XDCB - D02F0050
        '================================================================ 
        lblID.Text = rl3("Ma") 'Mã
        lblName.Text = rl3("Ten_loai_ma_phan_tich") 'Tên loại mã phân tích
        lblDisabled.Text = rl3("KSD") 'KSD
        lblLength.Text = rl3("Chieu_dai") 'Chiều dài (<=20)
        '================================================================ 
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnSave.Text = rl3("_Luu") '&Lưu
        '================================================================ 
    End Sub

End Class