Imports System.Data
Imports System.Data.SqlClient

Public Class Login

    ' Replace with your real connection string.
    ' If you already have a connection string somewhere (e.g. cls or My.Settings),
    ' use that instead (eg: Dim connString = cls.ConnectionString)
    Private ReadOnly connString As String = "YOUR_CONNECTION_STRING_HERE"

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If Not checkentry() Then Exit Sub

        frmLoading.Label1.Text = "Please wait while logging in"
        frmLoading.Show()
        Me.Cursor = Cursors.WaitCursor

        Try
            Using conn As New SqlConnection(connString)
                conn.Open()

                ' Parameterized SELECT to avoid SQL injection
                Using cmd As New SqlCommand("SELECT TOP 1 * FROM projectusers WHERE userid = @userid", conn)
                    cmd.Parameters.Add("@userid", SqlDbType.VarChar).Value = Trim(txtUser.Text)

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)

                        If dt.Rows.Count = 0 Then
                            Me.Cursor = Cursors.Default
                            frmLoading.Close()
                            MessageBox.Show("Userid not found", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Exit Sub
                        End If

                        ' Get the stored (encrypted) password value
                        Dim encPass As String = dt.Rows(0)("password").ToString()

                        ' Call the stored proc / function to decrypt (spDePalindrome)
                        ' Using an EXEC with a parameter to keep it parameterized
                        Using cmd2 As New SqlCommand("EXEC spDePalindrome @encPass", conn)
                            cmd2.Parameters.Add("@encPass", SqlDbType.VarChar).Value = encPass

                            ' If the stored proc returns a resultset with the decrypted value in first column
                            Dim dtPass As New DataTable()
                            Using da2 As New SqlDataAdapter(cmd2)
                                da2.Fill(dtPass)
                            End Using

                            Me.Cursor = Cursors.Default
                            frmLoading.Close()

                            If dtPass.Rows.Count = 0 OrElse dtPass.Rows(0)(0).ToString() <> Trim(txtPass.Text) Then
                                MessageBox.Show("Invalid Password", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Exit Sub
                            Else
                                userid = dt.Rows(0)("UserID").ToString()
                                userlevel = dt.Rows(0)("UserLevel").ToString()

                                MDI.Show()
                                Me.Close()
                            End If
                        End Using
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            frmLoading.Close()
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Function checkentry() As Boolean
        checkentry = True

        If Trim(txtUser.Text).Length = 0 Then
            MessageBox.Show("UserID is Required", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            checkentry = False
        End If

        If Trim(txtPass.Text).Length = 0 Then
            MessageBox.Show("Password is Required", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            checkentry = False
        End If
    End Function

    Private Sub txtPass_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call btnLogin_Click(sender, e)
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ' Optionally close the form or clear fields
        Me.Close()
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' IMPORTANT:
        ' The lines below auto-login the user (use only for testing).
        ' Remove or comment them out in production so login runs only when user clicks Login.
        '
        ' userid = "ANESLAGONC"
        ' userlevel = "80"
        ' MDI.Show()
        ' Me.Close()
    End Sub
End Class
