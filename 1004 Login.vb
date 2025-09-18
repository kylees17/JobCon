Public Class Login

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        If checkentry() Then
            frmLoading.Label1.Text = "Please wait while logging in"
            frmLoading.Show()

            Me.Cursor = Cursors.WaitCursor

            'Me.Enabled = False 

            Dim dt As New DataTable
            dt = cls.GetRecord("*", "projectusers", " where userid='" & Trim(txtUser.Text) & "'")

            If dt.Rows.Count = 0 Then
                Me.Cursor = Cursors.Default
                frmLoading.Close()
                'Me.Enabled = True
                MessageBox.Show("Userid not found", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            Else
                Dim dtpass As New DataTable
                dtpass = cls.GetRecordII("exec spDePalindrome '" & dt.Rows(0)("password") & "'")

                Me.Cursor = Cursors.Default
                frmLoading.Close()
                'Me.Enabled = True

                If dtpass.Rows(0)(0).ToString() <> Trim(txtPass.Text) Then
                    MessageBox.Show("Invalid Password", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                Else
                    userid = dt.Rows(0)("UserID")
                    userlevel = dt.Rows(0)("UserLevel")

                    MDI.Show()
                    Me.Close()

                End If
            End If
        End If

    End Sub


    Function checkentry() As Boolean
        If String.IsNullOrWhiteSpace(txtUser.Text) Then
            MessageBox.Show("UserID is Required", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtPass.Text) Then
            MessageBox.Show("Password is Required", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Return True
    End Function



    Private Sub txtPass_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call btnLogin_Click(sender, e)
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

    End Sub

    'Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    userid = "ANESLAGONC"
    '    userlevel = "80"
    '    'usergroup = "ProjectQA"

    '    MDI.Show()
    '    Me.Close()
    'End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
End Class
