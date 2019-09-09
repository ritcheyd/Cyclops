Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Deployment
Imports IWshRuntimeLibrary


Public Class Form1
    
    Public stringstatatue As Boolean = False
    Dim connectstring As String = "Data Source=rms-prod-db;Initial Catalog=ossimob;Integrated Security=True"
    Dim helpselect As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        helpselect = False
        CheckBox1.Checked = True
        filterremarks()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim userid As String = My.User.Name
        userid = Mid(userid, 13, 6)
        lablogon.Text = userid
        logfileadd(userid)
        'Dim x As String 'test

        'x = "test" 'test
        'MsgBox(x.PadLeft(15, " ")) 'test
        'Clipboard.SetDataObject(x.PadRight(15, Chr(32))) 'test


        DateTimePicker1.Value = Now.AddDays(-7)
        txtversion.Text = "Version: " & Mid(My.Application.Info.Version.ToString, 1, 2) & Mid(My.Application.Info.Version.ToString, 7, 1)
        'txtversion.Text = System.
    End Sub

    Sub logfileadd(ByVal readline As String)
        Try


            'Dim oFile As System.IO.File
            Dim oWrite As System.IO.StreamWriter
            Dim FILE_NAME As String = "S:\BPD\One_Solution\Login.log"
            ' Dim FILE_NAME As String = "\\rms-prod-rpt\Boulder\Cyclops\Login.log"

            oWrite = System.IO.File.AppendText(FILE_NAME)
            oWrite.WriteLine(Now() & " - " & readline)
            'oWrite.WriteLine(readline)
            oWrite.Flush()
            oWrite.Close()
        Catch ex As Exception

        End Try
    End Sub

    Sub filterremarks()

        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
        'Dim conn As SqlConnection = New SqlConnection("Server=bpd-rms-db;database=ossimob;Trusted_Connection=yes")
        Try
            conn.Open()
            Dim Startdate As Date = DateTimePicker1.Value
            Dim Enddate As Date = DateTimePicker2.Value
            ' MsgBox(Enddate)
            Dim strsql As String = ""
            Dim checkfirst As Boolean = False

            'strsql = "SELECT *  FROM dbo.combined WHERE  (xdate BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            If txtreport.Text <> "" Then
                strsql = "SELECT *  FROM dbo.combined WHERE (xreport = " & txtreport.Text & ")"
            Else
                strsql = "SELECT *  FROM dbo.combined WHERE (xdate >= DATEADD(DAY, -1, CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102)) AND  xdate <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            End If
            'strsql = "SELECT reportstatus.inci_id, reportstatus.status, reportstatus.offense , reportstatus.mobstatus , reportstatus.adduser, reportstatus.date_rept ,  reportstatus.fullname, reportstatus.mobilekey , reportstatus.mobilepkey, location  FROM dbo.reportstatus WHERE  (date_rept BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"


            conn.Close()
            Dim da As New SqlDataAdapter(strsql, conn)
            Dim ds As New DataSet
            da.Fill(ds, "lwmain")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.RowHeadersWidth = 10
            'DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(0).Width = 65
            DataGridView1.Columns(0).HeaderText = "Type"
            DataGridView1.Columns(1).Width = 53
            DataGridView1.Columns(1).HeaderText = "Report"
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(2).HeaderText = "Description"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Status"
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderText = "Report Date"
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 40
            DataGridView1.Columns(5).HeaderText = "Dept"
            DataGridView1.Columns(6).Width = 25
            DataGridView1.Columns(6).HeaderText = "Key"
            DataGridView1.Columns(6).Visible = False
            'DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(7).Width = 55
            DataGridView1.Columns(7).HeaderText = "User"
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(8).HeaderText = "Street#"
            'DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Width = 140
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
        End Try

    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        If stringstatatue = False Then


            Try
                'Dim Systemtype As Object = DataGridView1.Rows(e.RowIndex).Cells(4).Value
                Dim MNItype As Object = DataGridView1.Rows(e.RowIndex).Cells(7).Value
                'MsgBox(Systemtype & " - " & MNItype)
                ' If Systemtype = "HISTRMS" Then
                'HISTRMSRequest(MNItype)
                ' End If
                ' If Systemtype = "PIN" Then
                'PINRequest(MNItype)
                ' End If
                'Form2.Show()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If helpselect = True Then
            Dim Systemtype As Object = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Try

                Dim mnitype As String
                mnitype = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                txttemp.Text = mnitype
                'textmessages = mnitype
                txtmod.Text = Systemtype
                MsgBox(Systemtype & " - " & mnitype)
                'textmessage.Textmessage_note = mnitype
                'textmessage.Show()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else
            If CheckBox1.Checked = True Then


                Dim Systemtype As Object = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                Try

                    Dim mnitype As String
                    mnitype = DataGridView1.Rows(e.RowIndex).Cells(6).Value
                    txttemp.Text = mnitype
                    txtmod.Text = Systemtype
                    ' MsgBox(Systemtype & " - " & mnitype)
                    Form2.Show()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DateTimePicker1.Value = Now.AddDays(-7)
        DateTimePicker2.Value = Now.AddDays(-0)
        txtreport.Text = ""
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        helpselect = False
        CheckBox1.Checked = False
        If Len(txtreport.Text) < 7 Then
            MsgBox("You must enter the report number in first")
        Else

            Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")

            Try
                conn.Open()
                Dim Startdate As Date = DateTimePicker1.Value
                Dim Enddate As Date = DateTimePicker2.Value
                ' MsgBox(Enddate)
                Dim strsql As String = ""
                Dim checkfirst As Boolean = False

                strsql = "SELECT *  FROM dbo.NamesIncident WHERE (inci_id = " & txtreport.Text & ")"

                conn.Close()
                Dim da As New SqlDataAdapter(strsql, conn)
                Dim ds As New DataSet
                da.Fill(ds, "lwmain")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.RowHeadersWidth = 10
                'DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(0).HeaderText = "Report"
                DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(1).Width = 45
                DataGridView1.Columns(1).HeaderText = "Type"
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).Width = 100
                DataGridView1.Columns(2).HeaderText = "First Name"
                DataGridView1.Columns(3).Width = 100
                DataGridView1.Columns(3).HeaderText = "Last Name"
                DataGridView1.Columns(4).Width = 35
                DataGridView1.Columns(4).HeaderText = "Race"
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).Width = 35
                DataGridView1.Columns(5).HeaderText = "Sex"
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(6).Width = 70
                DataGridView1.Columns(6).HeaderText = "DOB"
                DataGridView1.Columns(6).Visible = True
                'DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(7).Width = 35
                DataGridView1.Columns(7).HeaderText = "AGE"
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(8).HeaderText = "Created Date"
                'DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Width = 100
                DataGridView1.Columns(9).HeaderText = "Created By"
            Catch ex As SqlException
                MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
            End Try
        End If
    End Sub

    Private Sub txtreport_TextChanged(sender As Object, e As EventArgs) Handles txtreport.TextChanged
        If txtreport.TextLength = 7 Or txtreport.TextLength = 0 Then
            Button1.Enabled = True
            Button3.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        Else
            Button1.Enabled = False
            Button3.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        End
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        helpselect = False
        CheckBox1.Checked = True
        filterremarks()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        helpselect = False
        CheckBox1.Checked = False
        If Len(txtreport.Text) < 7 Then
            MsgBox("You must enter the report number in first")
        Else


            Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")

            Try
                conn.Open()
                Dim Startdate As Date = DateTimePicker1.Value
                Dim Enddate As Date = DateTimePicker2.Value
                ' MsgBox(Enddate)
                Dim strsql As String = ""
                Dim checkfirst As Boolean = False

                strsql = "SELECT *  FROM dbo.PropertyList WHERE (case_id = " & txtreport.Text & ")"

                conn.Close()
                Dim da As New SqlDataAdapter(strsql, conn)
                Dim ds As New DataSet
                da.Fill(ds, "lwmain")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.RowHeadersWidth = 10
                'DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(0).HeaderText = "Report"
                DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(1).Width = 200
                DataGridView1.Columns(1).HeaderText = "Description"
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).Width = 50
                DataGridView1.Columns(2).HeaderText = "Value"
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).Width = 70
                DataGridView1.Columns(3).HeaderText = "Make"
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(4).HeaderText = "Model"
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(5).HeaderText = "Serial Number"
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(6).Width = 30
                DataGridView1.Columns(6).HeaderText = "EVD"
                DataGridView1.Columns(6).Visible = True
                'DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(7).Width = 100
                DataGridView1.Columns(7).HeaderText = "Created Date"
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(8).HeaderText = "Created By"
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                My.Computer.Clipboard.SetData("specialFormat", ds.Tables(0))
            Catch ex As SqlException
                MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
            End Try
            'My.Computer.Clipboard.SetText("This is a test string.")
            ' copydataview()
            'DataGridView1.SelectAll()
            copytoclipboard()
            'Clipboard.SetDataObject(DataGridView1.GetClipboardContent())
        End If
    End Sub

    Sub copytoclipboard()
        'Dim reportfolder As String = ""
        'Dim currentreportnum As String = ""
        Dim counterfiles As Integer = 0
        Try

            Dim MyConnection As System.Data.SqlClient.SqlConnection
            Dim MyDataAdapter As SqlClient.SqlDataAdapter
            MyConnection = New System.Data.SqlClient.SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
            Dim cmd As New SqlCommand("", MyConnection)
            Dim MyDataTable As New DataTable
            Dim MyDataRow As DataRow
            Dim strSQL As String = "SELECT *  FROM dbo.PropertyList WHERE (case_id = " & txtreport.Text & " and prstatus = 7)"

            MyConnection.Open()
            MyDataAdapter = New SqlClient.SqlDataAdapter(strSQL, MyConnection)
            MyDataAdapter.Fill(MyDataTable)
            Dim filename As String = ""
            Dim description As String = ""


            Dim reportnumber As String = ""

            Dim clipboarddata As String = ""
            Clipboard.Clear()
            Dim descspace As Integer = 41
            Dim makespace As Integer = 16
            Dim modelspace As Integer = 16
            Dim snspace As Integer = 26
            Dim tabdesc As String = ""
            For Each MyDataRow In MyDataTable.Rows
                descspace = 50 - Len(MyDataRow.Item("propdesc"))
                'Select Case Len(MyDataRow.Item("propdesc"))
                '    Case 0 To 10
                '        tabdesc = Chr(9) & Chr(9) & Chr(9) & Chr(9)
                '    Case 11 To 20
                '        tabdesc = Chr(9) & Chr(9)
                '    Case 21 To 50
                '        tabdesc = Chr(9)
                'End Select
                makespace = 16 - Len(MyDataRow.Item("make"))
                modelspace = 16 - Len(MyDataRow.Item("model"))
                snspace = 26 - Len(MyDataRow.Item("serialno"))

                'reportnumber = reportnumber & "Description:" & (MyDataRow.Item("propdesc")) & tabdesc & "Make:" & (MyDataRow.Item("make")) & Chr(9) & "Model:" & (MyDataRow.Item("Model")) & Chr(9) & "Serial#:" & (MyDataRow.Item("serialno")) & Chr(9) & "Value:" & (MyDataRow.Item("Value")) & Environment.NewLine
                reportnumber = reportnumber & "Description:" & (MyDataRow.Item("propdesc")) & Space(descspace) & "Make:" & (MyDataRow.Item("make")) & Space(makespace) & "Model:" & (MyDataRow.Item("Model")) & Space(modelspace) & "Serial#:" & (MyDataRow.Item("serialno")) & Space(snspace) & "Value:" & (MyDataRow.Item("Value")) & Environment.NewLine
                'reportnumber = "Description:" & (MyDataRow.Item("propdesc"))
                'ListBox1.Items.Add(reportnumber)
                counterfiles = counterfiles + 1

                If counterfiles > 10 Then Exit For
            Next
            MyDataAdapter.Dispose()
            MyDataTable.Dispose()
            MyConnection.Dispose()
            Clipboard.SetDataObject(reportnumber)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'Clipboard.SetDataObject("row1_col1" + "\t" + "row1_col2" + Environment.NewLine + "\r" + "row2_col1" + "\t" + "row2_col2")



    End Sub

    Private Sub txtstatute_TextChanged(sender As Object, e As EventArgs) Handles txtstatute.TextChanged
        helpselect = False
        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")

        If Len(txtstatute.Text) > 2 Then
            stringstatatue = True
            Try
                conn.Open()

                ' MsgBox(Enddate)
                Dim strsql As String = ""
                Dim checkfirst As Boolean = False
                Dim caps As String = UCase(txtstatute.Text)
                strsql = "SELECT *  FROM dbo.Statute_List WHERE statutdesc LIKE '%" + caps + "%'"
                'LIKE '%" + strSearchText + "%'"
                conn.Close()
                Dim da As New SqlDataAdapter(strsql, conn)
                Dim ds As New DataSet
                da.Fill(ds, "statutdesc")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.RowHeadersWidth = 10
                'DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(0).Width = 120
                DataGridView1.Columns(0).HeaderText = "Statute"
                DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(1).HeaderText = "Description"
                DataGridView1.Columns(2).Width = 40
                DataGridView1.Columns(2).HeaderText = "NCIC"
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).Width = 40
                DataGridView1.Columns(3).HeaderText = "IBRS"
                DataGridView1.Columns(4).Width = 200
                DataGridView1.Columns(4).HeaderText = "IBRS Description"
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).Width = 20
                DataGridView1.Columns(5).HeaderText = "F/M"
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(6).Width = 60
                DataGridView1.Columns(6).HeaderText = "Repeal Date"
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(7).Width = 100
                DataGridView1.Columns(7).HeaderText = "Other Descr"
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Catch ex As SqlException
                'MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
            End Try
        End If

        stringstatatue = False
    End Sub

    Sub other()
        
    End Sub

    
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        helpselect = False
        CheckBox1.Checked = True
        denied()
    End Sub

    Sub denied()

        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
        'Dim conn As SqlConnection = New SqlConnection("Server=bpd-rms-db;database=ossimob;Trusted_Connection=yes")
        Try
            conn.Open()
            Dim Startdate As Date = DateTimePicker1.Value
            Dim Enddate As Date = DateTimePicker2.Value
            ' MsgBox(Enddate)
            Dim strsql As String = ""
            Dim checkfirst As Boolean = False

            strsql = "SELECT *  FROM dbo.combined WHERE (xstatus = 'DENY')"

            conn.Close()
            Dim da As New SqlDataAdapter(strsql, conn)
            Dim ds As New DataSet
            da.Fill(ds, "lwmain")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.RowHeadersWidth = 10
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderText = "Type"
            DataGridView1.Columns(1).Width = 53
            DataGridView1.Columns(1).HeaderText = "Report"
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(2).HeaderText = "Description"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Status"
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderText = "Report Date"
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 35
            DataGridView1.Columns(5).HeaderText = "Dept"
            DataGridView1.Columns(6).Width = 20
            DataGridView1.Columns(6).HeaderText = "Key"
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Width = 55
            DataGridView1.Columns(7).HeaderText = "User"
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(8).HeaderText = "Street#"
            DataGridView1.Columns(9).Width = 150
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Sub

    

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles BtnAddress.Click
        helpselect = False
        addresssearch()
    End Sub

    Private Sub txtaddreenumber_TextChanged(sender As Object, e As EventArgs) Handles txtaddreenumber.TextChanged
        helpselect = False
        If Len(txtaddreenumber.Text) > 2 Then
            addresssearch()
        End If

    End Sub

    Private Sub txtstreetname_TextChanged(sender As Object, e As EventArgs) Handles txtstreetname.TextChanged
        helpselect = False
        If Len(txtstreetname.Text) > 1 Then
            addresssearch()
        End If
    End Sub

    Sub addresssearch()
        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
        'Dim conn As SqlConnection = New SqlConnection("Server=bpd-rms-db;database=ossimob;Trusted_Connection=yes")
        Try
            conn.Open()
            Dim Startdate As Date = DateTimePicker1.Value
            Dim Enddate As Date = DateTimePicker2.Value
            ' MsgBox(Enddate)
            Dim strsql As String = ""
            Dim checkfirst As Boolean = False

            'strsql = "SELECT *  FROM dbo.combined WHERE  (xdate BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            'If txtaddreenumber.Text <> "" Then
            strsql = "SELECT *  FROM dbo.combined WHERE (streetnbr LIKE '" & txtaddreenumber.Text & "%') and (street LIKE '%" & txtstreetname.Text & "%') and (xdate >= DATEADD(DAY, -1, CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102)) AND  xdate <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            ' Else
            'strsql = "SELECT *  FROM dbo.combined WHERE (xdate >= DATEADD(DAY, -1, CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102)) AND  xdate <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            ' End If


            'strsql = "SELECT reportstatus.inci_id, reportstatus.status, reportstatus.offense , reportstatus.mobstatus , reportstatus.adduser, reportstatus.date_rept ,  reportstatus.fullname, reportstatus.mobilekey , reportstatus.mobilepkey, location  FROM dbo.reportstatus WHERE  (date_rept BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"


            conn.Close()
            Dim da As New SqlDataAdapter(strsql, conn)
            Dim ds As New DataSet
            da.Fill(ds, "lwmain")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.RowHeadersWidth = 10
            'DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderText = "Type"
            DataGridView1.Columns(1).Width = 53
            DataGridView1.Columns(1).HeaderText = "Report"
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(2).HeaderText = "Description"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Status"
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderText = "Report Date"
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 35
            DataGridView1.Columns(5).HeaderText = "Dept"
            DataGridView1.Columns(6).Width = 20
            DataGridView1.Columns(6).HeaderText = "Key"
            DataGridView1.Columns(6).Visible = False
            'DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(7).Width = 55
            DataGridView1.Columns(7).HeaderText = "User"
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(8).HeaderText = "Street#"
            'DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Width = 150
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Sub

  
    Private Sub txtmod_TextChanged(sender As Object, e As EventArgs) Handles txtmod.TextChanged

    End Sub

    Private Sub txthelp_TextChanged(sender As Object, e As EventArgs) Handles txthelp.TextChanged
        helpselect = True
        Dim conn As SqlConnection = New SqlConnection(connectstring)

        If Len(txthelp.Text) > 2 Then
            helpselect = True
            Try
                conn.Open()

                ' MsgBox(Enddate)
                Dim strsql As String = ""
                Dim checkfirst As Boolean = False
                Dim caps As String = LCase(txthelp.Text)
                strsql = "SELECT *  FROM dbo.helpview WHERE Helpkey LIKE '%" + caps + "%'"
                'LIKE '%" + strSearchText + "%'"
                conn.Close()
                Dim da As New SqlDataAdapter(strsql, conn)
                Dim ds As New DataSet
                da.Fill(ds, "statutdesc")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.RowHeadersWidth = 10
                'DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(0).Width = 600
                DataGridView1.Columns(0).HeaderText = "Title"
                'DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(1).HeaderText = "Description"
                DataGridView1.Columns(2).Visible = False
                'DataGridView1.Columns(2).Width = 40
                'DataGridView1.Columns(2).HeaderText = "Key"
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).Width = 70
                DataGridView1.Columns(3).HeaderText = "Added"
                DataGridView1.Columns(4).Visible = False
                'DataGridView1.Columns(4).Width = 200
                'DataGridView1.Columns(4).HeaderText = "Added By"
                'DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(5).HeaderText = "Added By"
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(6).HeaderText = "Repeal Date"
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(7).Width = 100
                DataGridView1.Columns(7).HeaderText = "Other Descr"
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Catch ex As SqlException
                MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
            End Try
        End If

        stringstatatue = False
    End Sub

   
    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        helpselect = False
        CheckBox1.Checked = True
        inprogress()
    End Sub

    Sub inprogress()

        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
        'Dim conn As SqlConnection = New SqlConnection("Server=bpd-rms-db;database=ossimob;Trusted_Connection=yes")
        Try
            conn.Open()
            Dim Startdate As Date = DateTimePicker1.Value
            Dim Enddate As Date = DateTimePicker2.Value
            ' MsgBox(Enddate)
            Dim strsql As String = ""
            Dim checkfirst As Boolean = False

            strsql = "SELECT *  FROM dbo.combined WHERE (xstatus = 'NEW')"

            conn.Close()
            Dim da As New SqlDataAdapter(strsql, conn)
            Dim ds As New DataSet
            da.Fill(ds, "lwmain")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.RowHeadersWidth = 10
            DataGridView1.Columns(0).Width = 75
            DataGridView1.Columns(0).HeaderText = "Type"
            DataGridView1.Columns(1).Width = 53
            DataGridView1.Columns(1).HeaderText = "Report"
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(2).HeaderText = "Description"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Status"
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderText = "Report Date"
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 35
            DataGridView1.Columns(5).HeaderText = "Dept"
            DataGridView1.Columns(6).Width = 20
            DataGridView1.Columns(6).HeaderText = "Key"
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Width = 55
            DataGridView1.Columns(7).HeaderText = "User"
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(8).HeaderText = "Street#"
            DataGridView1.Columns(9).Width = 150
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Sub

    Private Sub TxtNames_TextChanged(sender As Object, e As EventArgs) Handles TxtNames.TextChanged
        If Len(TxtNames.Text) > 2 Then
            ' stringstatatue = True

            Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
            Try
                conn.Open()
                Dim Startdate As Date = DateTimePicker1.Value
                Dim Enddate As Date = DateTimePicker2.Value
                ' MsgBox(Enddate)
                Dim strsql As String = ""
                Dim checkfirst As Boolean = False
                Dim caps As String = UCase(TxtNames.Text)
                'strsql = "SELECT *  FROM dbo.helpview WHERE Helpkey LIKE '%" + caps + "%'"
                strsql = "SELECT *  FROM dbo.CombinedNames WHERE Name LIKE '%" + caps + "%' ORDER BY Name"

                conn.Close()
                Dim da As New SqlDataAdapter(strsql, conn)
                Dim ds As New DataSet
                da.Fill(ds, "lwmain")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.RowHeadersWidth = 10
                'DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(0).Width = 75
                DataGridView1.Columns(0).HeaderText = "Type"
                DataGridView1.Columns(1).Width = 150
                DataGridView1.Columns(1).HeaderText = "Name"
                DataGridView1.Columns(2).Width = 60
                DataGridView1.Columns(2).HeaderText = "Report"
                DataGridView1.Columns(3).Width = 180
                DataGridView1.Columns(3).HeaderText = "Info"
                'DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(4).Width = 40
                DataGridView1.Columns(4).HeaderText = "Race"
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).Width = 40
                DataGridView1.Columns(5).HeaderText = "Sex"
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                'DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 40
                DataGridView1.Columns(6).HeaderText = "Key"
                DataGridView1.Columns(6).Visible = False
                'DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(7).HeaderText = "DOB"
                DataGridView1.Columns(8).Width = 80
                DataGridView1.Columns(8).HeaderText = "Added"
                'DataGridView1.Columns(8).Visible = False
                'DataGridView1.Columns(9).Width = 150
            Catch ex As SqlException
                MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
            End Try
        End If
        ' stringstatatue = False
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim reader1 As StreamReader = New StreamReader("\\rms-prod-rpt\Boulder\Cyclops\shutdown.ini")
        Dim readlinetxt As String
        readlinetxt = (reader1.ReadLine)
        If Microsoft.VisualBasic.Mid(readlinetxt, 1, 1) = "Y" Then
            End
        End If
        reader1.Close()
    End Sub

    Private Sub btnOfficer_Click(sender As Object, e As EventArgs) Handles btnOfficer.Click
        helpselect = False
        CheckBox1.Checked = True
        Dim conn As SqlConnection = New SqlConnection("Server=rms-prod-db;database=ossimob;Trusted_Connection=yes")
        'Dim conn As SqlConnection = New SqlConnection("Server=bpd-rms-db;database=ossimob;Trusted_Connection=yes")
        Try
            conn.Open()
            Dim Startdate As Date = DateTimePicker1.Value
            Dim Enddate As Date = DateTimePicker2.Value
            ' MsgBox(Enddate)
            Dim strsql As String = ""
            Dim checkfirst As Boolean = False
            TbxOfficer.Text = TbxOfficer.Text.ToUpper
            'strsql = "SELECT *  FROM dbo.combined WHERE  (xdate BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            'If txtreport.Text <> "" Then
            'strsql = "SELECT *  FROM dbo.combined WHERE (adduser = '" & TbxOfficer.Text & "')"
            'Else
            strsql = "SELECT *  FROM dbo.combined WHERE (adduser = '" & TbxOfficer.Text & "' AND xdate >= DATEADD(DAY, -1, CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102)) AND  xdate <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"
            'End If
            'strsql = "SELECT reportstatus.inci_id, reportstatus.status, reportstatus.offense , reportstatus.mobstatus , reportstatus.adduser, reportstatus.date_rept ,  reportstatus.fullname, reportstatus.mobilekey , reportstatus.mobilepkey, location  FROM dbo.reportstatus WHERE  (date_rept BETWEEN CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 102) AND CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 102))"


            conn.Close()
            Dim da As New SqlDataAdapter(strsql, conn)
            Dim ds As New DataSet
            da.Fill(ds, "combined")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.RowHeadersWidth = 10
            'DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(0).Width = 65
            DataGridView1.Columns(0).HeaderText = "Type"
            DataGridView1.Columns(1).Width = 53
            DataGridView1.Columns(1).HeaderText = "Report"
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(2).HeaderText = "Description"
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Status"
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(4).HeaderText = "Report Date"
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 40
            DataGridView1.Columns(5).HeaderText = "Dept"
            DataGridView1.Columns(6).Width = 25
            DataGridView1.Columns(6).HeaderText = "Key"
            DataGridView1.Columns(6).Visible = False
            'DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(7).Width = 55
            DataGridView1.Columns(7).HeaderText = "User"
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(8).HeaderText = "Street#"
            'DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Width = 140
        Catch ex As SqlException
            MsgBox(ex.Message, MsgBoxStyle.Critical, "SQL Error")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "General Error")
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Tbxinstall.Click
        CreateShortCut("C:\CAplusApp\CAplusNet\CAPlus.exe", "CAPLUS")

        Tbxinstall.Text = "Installing"
        Tbxinstall.Refresh()
        Dim path As String = "C:\CAplusApp\"
        DeleteDirectory(path)
        My.Computer.FileSystem.CopyDirectory("G:\CAPLUS\CAplusApp", "C:\CAplusApp", True)

        path = "C:\CAplusMaps\"
        DeleteDirectory(path)
        My.Computer.FileSystem.CopyDirectory("G:\CAPLUS\CAplusMaps", "C:\CAplusMaps", True)

        MsgBox("CA Plus is installed")
        Tbxinstall.Text = "Done"
        Tbxinstall.Refresh()
    End Sub

    Private Sub DeleteDirectory(path As String)
        If Directory.Exists(path) Then
            'Delete all files from the Directory
            For Each filepath As String In Directory.GetFiles(path)
                System.IO.File.Delete(filepath)
            Next
            'Delete all child Directories
            For Each dir As String In Directory.GetDirectories(path)
                DeleteDirectory(dir)
            Next
            'Delete a Directory
            Directory.Delete(path)

        End If


    End Sub

    Private Sub CreateShortCut(ByVal FileName As String, ByVal Title As String)

        Dim WshShell As IWshRuntimeLibrary.WshShellClass = New WshShellClass

        Dim MyShortcut As IWshRuntimeLibrary.IWshShortcut

        ' The shortcut will be created on the desktop

        Dim DesktopFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

        MyShortcut = CType(WshShell.CreateShortcut(DesktopFolder & "\CAPLUS.lnk"), IWshRuntimeLibrary.IWshShortcut)

        MyShortcut.TargetPath = "C:\CAplusApp\CAplusNet\CAPlus.exe"  'Specify target file full path

        MyShortcut.Save()


    End Sub
    
End Class
