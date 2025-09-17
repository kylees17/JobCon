Module MyLibrary

    Public cls As New datacenter

    Public userid As String, userlevel As String, caller As String, usergroup As String


    Public Function enc(parameter As String) As String
        Dim charr As String
        Dim i As Byte
        For i = 1 To Len(parameter)
            charr = Mid(StrReverse(parameter), i, 1)
            enc = enc & Hex(Asc(charr))
        Next i

        End
    End Function

    Public Function decry(parameter As String) As String
        Dim charr As String
        Dim i As Integer
        decry = ""
        For i = 1 To Len(parameter) Step 2
            charr = Mid(parameter, i, 2)
            'decry = decry & Chr(Chr(38) & "H" & charr)
            decry = decry & Chr(Chr(38) & "H" & charr)
        Next i
        decry = StrReverse(decry)
    End Function

    Public Function projectControl(desc As String) As String
        Dim dt As New DataTable
        dt = cls.GetRecordII("exec spProjectDCFNo '" & desc & "'")

        projectControl = dt.Rows(0)("ProjectNo")
    End Function

    Public Function projectSDF(DCFNo As String) As String
        Dim dt As New DataTable
        dt = cls.GetRecordII("exec spProjectSDFNo '" & DCFNo & "'")

        projectSDF = dt.Rows(0)("ProjectNo")
    End Function




End Module
