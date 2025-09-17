Imports Microsoft.VisualBasic
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration

Public Class datacenter

    Dim constringData As String = System.Configuration.ConfigurationManager.ConnectionStrings("CENTRALDBTESTConnectionString").ConnectionString
    'Dim constringData As String = System.Configuration.ConfigurationManager.ConnectionStrings("DataConnectionLive").ConnectionString
    Dim IsCipher As String = System.Configuration.ConfigurationManager.AppSettings.Item("Cipher")
    Dim conn As New SqlConnection(forCon())
    Dim cmd As New SqlCommand()

    Public Function GetDataSet(query As String, dt As String) As NatureOfRequest

        Dim dsDeclaredDataSet As New NatureOfRequest
        Dim a As SqlDataAdapter
        a = New SqlDataAdapter(query, conn)
        a.Fill(dsDeclaredDataSet, dt)
        Return dsDeclaredDataSet

    End Function

    Public Function GetDataSetDSSDF(query As String, dt As String) As DSSDF

        Dim dsDeclaredDataSet As New DSSDF
        Dim a As SqlDataAdapter
        a = New SqlDataAdapter(query, conn)
        a.Fill(dsDeclaredDataSet, dt)
        Return dsDeclaredDataSet


        'Dim conString As String = "Server=192.168.0.15;Database=CENTRALDBTEST;Uid=sa;Pwd=abcd@1234;"
        'Dim cmd As New SqlCommand(query)
        'Using con As New SqlConnection(conString)
        '    Using sda As New SqlDataAdapter()
        '        cmd.Connection = con
        '        sda.SelectCommand = cmd
        '        Using dsDeclaredDataSet As New DSSDF
        '            sda.Fill(dsDeclaredDataSet, dt) 'datatable name
        '            Return dsDeclaredDataSet
        '        End Using
        '    End Using
        'End Using

    End Function

    Public Function GetRecord(strFields, strTable, strCondition) As DataTable
        ' conn = New SqlConnection(constringData)
        GetRecord = New DataTable()
        Dim a As SqlDataAdapter
        a = New SqlDataAdapter("SELECT " & strFields & " from " & strTable & strCondition, conn)
        a.Fill(GetRecord)
    End Function

    Public Function getProjectObject(DCFNo As String, ObjectClass As String, ObjectType As String) As DataTable
        getProjectObject = New DataTable
        getProjectObject = GetRecord("ObjectName,Remarks,isnull(FromDB,'') FromDB", "ProjectObject", " where ObjectClass='" & ObjectClass & "' and ObjectType='" & ObjectType & "' and DCFNo='" & DCFNo & "'")
    End Function


    Public Function ExecQuery(ByVal sSql As String)
        Try
            conn.Open()
            cmd = New SqlCommand(sSql, conn)
            cmd.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function GetRecordII(ByVal sSql As String) As DataTable
        Try
            GetRecordII = New DataTable()
            Dim a As SqlDataAdapter
            a = New SqlDataAdapter(sSql, conn)
            a.Fill(GetRecordII)
        Catch ex As Exception
            MessageBox.Show(ex.Message & "asdasda")
        End Try
    End Function




    Function forCon()
        'Dim line As String

        'Dim path As String = AppDomain.CurrentDomain.BaseDirectory
        'path = path.Replace("\Proto1.1\bin\Debug\", "\SLS.config")
        'path = path.Replace("\", "\\")
        'Using reader As StreamReader = New StreamReader(path)
        '    line = reader.ReadLine
        '    line = reader.ReadLine
        'End Using

        'If Left(line, 15) = "FMISCONNECTION=" Then
        '    line = line.Substring(15)
        'Else
        '    line = ""
        'End If

        'Return decry(line)

        SetSecurity()
        If IsCipher = "1" Then
            Try
                constringData = DeCipherText(constringData)
            Catch ex As Exception
                MsgBox("Connection Error. Please contact Administrator.", MsgBoxStyle.OkOnly + vbCritical, "SYSTEM MESSAGE")
            End Try

        End If
        Return constringData
    End Function

    'Sub populatecombo(combo As ComboBox, type As String, Optional condition As String = "")
    '    combo.Items.Clear()

    '    If combo.Name = "cmbFilter" Then
    '        combo.Items.Add("All")
    '    End If

    '    If condition <> "" Then
    '        condition = "and " & condition
    '    End If

    '    Dim dt = New DataTable
    '    dt = cls.GetRecord("concat(rtrim(code),' - ',description) as code", "Reference", " where type='" & type & "' and active='1' " & condition & " order by description asc")

    '    For i As Integer = 0 To dt.Rows.Count - 1
    '        combo.Items.Add(dt.Rows(i)("code").ToString())
    '    Next
    'End Sub

    Function getcode(code As String) As String
        getcode = code

        If code <> "All" Then
            getcode = Split(code, " - ")(0)
        End If

        Return Trim(getcode)

    End Function

    Function getFullName(uid As String) As String
        Dim dt = New DataTable
        dt = GetRecord("UserName", "useraccess", " where UserID='" & uid & "'")
        getFullName = ""
        If dt.Rows.Count > 0 Then
            getFullName = dt.Rows(0)("UserName")
        End If
    End Function

    Function getDCFApprove(dcfNo As String) As String
        Dim dt = New DataTable
        dt = GetRecord("top 1 UserID", "ProjectStatusHist", " where DCFNo='" & dcfNo & "' and ProjectStatus='T' order by ProjectDate desc")
        Try
            getDCFApprove = getUserIDName(dt.Rows(0)("userid").ToString(), "2")
        Catch ex As Exception
            getDCFApprove = ""
        End Try

    End Function

    Function getDCFQA(dcfno As String) As String
        Dim dt = New DataTable
        dt = GetRecord("top 1 UserID", "ProjectStatusHist", " where DCFNo='" & dcfno & "' and ProjectStatus='F' order by ProjectDate desc")
        Try
            getDCFQA = getUserIDName(dt.Rows(0)("userid").ToString(), "2")
        Catch ex As Exception
            getDCFQA = ""
        End Try
    End Function

    Function getDCFQADate(dcfno As String) As String
        Dim dt = New DataTable
        dt = GetRecord("top 1 cast(ProjectDate as date) ProjectDate", "ProjectStatusHist", " where DCFNo='" & dcfno & "' and ProjectStatus='F' order by ProjectDate desc")
        Try
            getDCFQADate = dateToWord(dt.Rows(0)("ProjectDate").ToString())
        Catch ex As Exception
            getDCFQADate = ""
        End Try
    End Function



    Function getDCFApproveDate(dcfNo As String) As String
        Dim dt = New DataTable
        dt = GetRecord("top 1 cast(ProjectDate as date) ProjectDate", "ProjectStatusHist", " where DCFNo='" & dcfNo & "' and ProjectStatus='T' order by ProjectDate desc")
        Try
            getDCFApproveDate = dateToWord(dt.Rows(0)("ProjectDate").ToString())
        Catch ex As Exception
            getDCFApproveDate = ""
        End Try

    End Function


End Class
