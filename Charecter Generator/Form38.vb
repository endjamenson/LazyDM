Imports System.Data.OleDb


Public Class Form38
    Dim myDA As OleDbDataAdapter
    Dim myDataSet As DataSet
    ' Dim con As New OleDbConnection
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim da As New OleDbDataAdapter
    Dim cmdUpdate As New OleDbCommand
    Dim inc As Integer


    Dim cmd As New OleDbCommand

    Dim cnnOLEDB As New OleDbConnection

    Dim cmdOLEDB As New OleDbCommand

    Dim cmdInsert As New OleDbCommand

    'Dim cmdUpdate As New OleDbCommand

    Dim cmdDelete As New OleDbCommand

    Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(myDA)

    Dim MaxRows As Integer

    ' Dim con As New OleDb.OleDbConnection

    Dim dbProvider As String

    Dim dbSource As String

    'Dim ds As New DataSet

    ' Dim da As New OleDb.OleDbDataAdapter

    Dim sql As String
    Private Sub Form38_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Are you sure you want to close this?  Have you saved yet?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
        Else
            e.Cancel = True
        End If
    End Sub
    Dim DGVhasChanged As Boolean


    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        DGVhasChanged = True
        CheckBox1.Checked = True

    End Sub

    Private Sub Form38_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        'MsgBox(Form1.TextBox14.Text)

        ' Here
        ' BindingSource1.DataSource = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text)
        ' myDataGridView.DataSource = myBindingSource

        ' cmd = ("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)




        con.Open()

        myDA = New OleDbDataAdapter(cmd)

        'Here one CommandBuilder object is required.

        myDA.UpdateCommand = New OleDbCommand("UPDATE " & TopDisp.TextBox1.Text & " SET Field1 = ?, Field2 = ?")


        'It will automatically generate DeleteCommand,UpdateCommand and InsertCommand for DataAdapter object

        ' Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(myDA) 'Old

        myDataSet = New DataSet()

        myDA.Fill(myDataSet, TopDisp.TextBox1.Text)

        DataGridView1.DataSource = myDataSet.Tables(TopDisp.TextBox1.Text).DefaultView

        con.Close()

        'con = Nothing




    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Sub updategrid()
        DataGridView1.DataSource = Nothing

        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()

        myDA = New OleDbDataAdapter(cmd)

        'Here one CommandBuilder object is required.

        myDA.UpdateCommand = New OleDbCommand("UPDATE " & TopDisp.TextBox1.Text & " SET Field1 = ?, Field2 = ?")


        'It will automatically generate DeleteCommand,UpdateCommand and InsertCommand for DataAdapter object

        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(myDA)

        myDataSet = New DataSet()

        myDA.Fill(myDataSet, TopDisp.TextBox1.Text)

        DataGridView1.DataSource = myDataSet.Tables(TopDisp.TextBox1.Text).DefaultView

        ' con.Close()

    End Sub



    Private Sub senddata()




    End Sub

    Private Sub GetData()


        'The table can be used here to display and edit the data.
        'That will most likely involve data-binding but that is not a data access issue.
    End Sub

    Private Sub SaveData()

    End Sub
    Private Shared Function CreateCustomerAdapter(ByVal connection As OleDbConnection) As OleDbDataAdapter

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim tack As Integer = 1

        Do Until tack = DataGridView1.Rows.Count

            Dim id As Integer = ((DataGridView1.Rows(id).Cells(0).Value)) ' this is the record set ID, need to pull this for every ID that will be updated
            Dim rowsel As Integer = (tack)

            updatejob(id, (rowsel))


            tack = (tack + 1)
        Loop
        MsgBox("Updated.")

    End Sub
    Public Sub uupp()
        wipedata()

        ' Me.YOURDATA.table.clear()

        For Each row As DataGridViewRow In DataGridView1.Rows
            'MsgBox("gets just before")
            If row.Cells(0).Value > 0.5 Then
                ' ...

                'MsgBox("gets just before")
                Try
                    Dim sqlconn As New OleDb.OleDbConnection
                    Dim sqlquery As New OleDb.OleDbCommand
                    Dim connString As String
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Form1.TextBox14.Text
                    sqlconn.ConnectionString = connString
                    sqlquery.Connection = sqlconn
                    sqlconn.Open()
                    ' MsgBox("gets just before")

                    '3 is speachimped, called accent -speed
                    '4 is accent - accent
                    '2 tone - tone
                    '1 pace - pacing
                    Dim name As String = row.Cells(1).Value
                    Dim job As String = row.Cells(5).Value
                    Dim gender As String = row.Cells(3).Value
                    Dim pacing As String = row.Cells(11).Value
                    Dim tone As String = row.Cells(12).Value
                    Dim speed As String = row.Cells(10).Value
                    Dim accent As String = row.Cells(9).Value
                    Dim race As String = row.Cells(4).Value
                    Dim motive As String = row.Cells(6).Value
                    Dim characteristic1 As String = row.Cells(7).Value
                    Dim charic2 As String = row.Cells(8).Value
                    Dim age As Integer = row.Cells(2).Value
                    Dim religion As String = row.Cells(13).Value
                    Dim political As String = row.Cells(14).Value
                    Dim notes As String
                    If row.Cells(15).Value.ToString.Length > 1 Then
                        notes = row.Cells(15).Value
                    Else
                        notes = " "
                    End If


                    sqlquery.CommandText = "INSERT INTO " & TopDisp.TextBox1.Text & "([name], [job], [gender], [pacing], [tone], [SpeechImpediment], [accent], [race], [motive], [characteristic1], [characteristic2], [age], [religion], [political], [Notes])VALUES(@name, @job, @gender, @pacing, @tone, @SpeechImpediment, @accent, @race, @motive, @characteristic1, @characteristic2, @age, @religion, @political, @Notes)"
                    sqlquery.Parameters.AddWithValue("@name", name)
                    sqlquery.Parameters.AddWithValue("@job", job)
                    sqlquery.Parameters.AddWithValue("@gender", gender)
                    sqlquery.Parameters.AddWithValue("@pacing", pacing)
                    sqlquery.Parameters.AddWithValue("@tone", tone)
                    sqlquery.Parameters.AddWithValue("@SpeechImpediment", speed)
                    sqlquery.Parameters.AddWithValue("@accent", accent)
                    sqlquery.Parameters.AddWithValue("@race", race)
                    sqlquery.Parameters.AddWithValue("@motive", motive)
                    sqlquery.Parameters.AddWithValue("@characteristic1", characteristic1)
                    sqlquery.Parameters.AddWithValue("@characteristic2", charic2)
                    sqlquery.Parameters.AddWithValue("@age", age)
                    sqlquery.Parameters.AddWithValue("@religion", religion)
                    sqlquery.Parameters.AddWithValue("@political", political)
                    sqlquery.Parameters.AddWithValue("@Notes", notes)
                    sqlquery.ExecuteNonQuery()
                    sqlconn.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End If
        Next
        MsgBox("Updated")
    End Sub

    Public Sub wipedata()
        ' Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file


        Dim conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file
        Dim com As OleDbCommand

        com = New OleDbCommand("delete * from " & TopDisp.TextBox1.Text, conn)

        conn.Open()
        com.ExecuteNonQuery()
        'MsgBox("Data Deleted")
        conn.Close()

        ' con.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim tack As Integer = 1

        Do Until tack = DataGridView1.Rows.Count

            Dim id As Integer = ((DataGridView1.Rows(id).Cells(0).Value)) ' this is the record set ID, need to pull this for every ID that will be updated
            Dim rowsel As Integer = (tack)

            updatejob(id, (rowsel))


            tack = (tack + 1)
        Loop
        MsgBox("Updated.")
    End Sub
    Public Function updatejob(id As Integer, rowsel As Integer)

        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String
        Dim name As String
        Dim gender As String
        Dim pacing As String
        Dim tone As String
        Dim speed As String
        Dim accent As String
        Dim race As String
        Dim motive As String
        Dim characteristic1 As String
        Dim charic2 As String
        Dim age As String
        Dim religion As String
        Dim political As String
        Dim notes As String


        ' id = (id + 1)

        'BAIL. This shits fucking dumb. 

        ' MsgBox(rowsel)
        If IsDBNull(DataGridView1.Rows(rowsel).Cells(5).Value) Then

            x = " "

        Else
            ' MsgBox("Doesnt See nothing")
            x = DataGridView1.Rows(rowsel).Cells(5).Value
        End If




        'name
        If IsDBNull(DataGridView1.Rows(rowsel).Cells(1).Value) Then

            name = " "

        Else
            ' MsgBox("Doesnt See nothing")
            name = DataGridView1.Rows(rowsel).Cells(1).Value
        End If



        'age
        If IsDBNull(DataGridView1.Rows(rowsel).Cells(2).Value) Then

            age = " "

        Else
            ' MsgBox("Doesnt See nothing")
            age = DataGridView1.Rows(rowsel).Cells(2).Value
        End If


        'gender

        If IsDBNull(DataGridView1.Rows(rowsel).Cells(3).Value) Then

            gender = " "

        Else
            ' MsgBox("Doesnt See nothing")
            gender = DataGridView1.Rows(rowsel).Cells(3).Value
        End If

        'race


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(4).Value) Then

            race = " "

        Else
            ' MsgBox("Doesnt See nothing")
            race = DataGridView1.Rows(rowsel).Cells(4).Value
        End If

        'motive


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(6).Value) Then

            motive = " "

        Else
            ' MsgBox("Doesnt See nothing")
            motive = DataGridView1.Rows(rowsel).Cells(6).Value
        End If

        'characteristic1


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(7).Value) Then

            characteristic1 = " "

        Else
            ' MsgBox("Doesnt See nothing")
            characteristic1 = DataGridView1.Rows(rowsel).Cells(7).Value
        End If


        'charic2

        If IsDBNull(DataGridView1.Rows(rowsel).Cells(8).Value) Then

            charic2 = " "

        Else
            ' MsgBox("Doesnt See nothing")
            charic2 = DataGridView1.Rows(rowsel).Cells(8).Value
        End If

        'accent


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(9).Value) Then

            accent = " "

        Else
            ' MsgBox("Doesnt See nothing")
            accent = DataGridView1.Rows(rowsel).Cells(9).Value
        End If


        'speed (Speech imped)

        If IsDBNull(DataGridView1.Rows(rowsel).Cells(10).Value) Then

            speed = " "

        Else
            ' MsgBox("Doesnt See nothing")
            speed = DataGridView1.Rows(rowsel).Cells(10).Value
        End If

        'pacing


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(11).Value) Then

            pacing = " "

        Else
            ' MsgBox("Doesnt See nothing")
            pacing = DataGridView1.Rows(rowsel).Cells(11).Value
        End If


        'tone

        If IsDBNull(DataGridView1.Rows(rowsel).Cells(12).Value) Then

            tone = " "

        Else
            ' MsgBox("Doesnt See nothing")
            tone = DataGridView1.Rows(rowsel).Cells(12).Value
        End If



        'religion

        If IsDBNull(DataGridView1.Rows(rowsel).Cells(13).Value) Then

            religion = " "

        Else
            ' MsgBox("Doesnt See nothing")
            religion = DataGridView1.Rows(rowsel).Cells(13).Value
        End If


        'political


        If IsDBNull(DataGridView1.Rows(rowsel).Cells(14).Value) Then

            political = " "

        Else
            ' MsgBox("Doesnt See nothing")
            political = DataGridView1.Rows(rowsel).Cells(14).Value
        End If




        If IsDBNull(DataGridView1.Rows(rowsel).Cells(15).Value) Then

            notes = " "

        Else
            ' MsgBox("Doesnt See nothing")
            notes = DataGridView1.Rows(rowsel).Cells(15).Value
        End If









        ' MsgBox(x)
        Try
            Dim sqlconn As New OleDb.OleDbConnection
            Dim sqlquery As New OleDb.OleDbCommand
            Dim connString As String
            connString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & Form1.TextBox14.Text



            sqlconn.ConnectionString = connString
            sqlquery.Connection = sqlconn
            sqlconn.Open()


            '        cmd.CommandText = "UPDATE Records SET FirstName = @firstname, LastName = @lastname, Age = @age, Address = @address, Course = @course WHERE [ID] = @id"
            ' cmd.Parameters.AddWithValue("@firstname", textBox1.Text)
            ' sqlquery.CommandText = "INSERT INTO " & TopDisp.TextBox1.Text & "([ID], [job])VALUES(@ID, @job)"

            sqlquery.CommandText = "UPDATE " & TopDisp.TextBox1.Text & " SET job = @job, name = @name, gender = @gender, pacing = @pacing, tone = @tone, SpeechImpediment = @speed, accent = @accent, race = @race, motive = @motive, characteristic1 = @characteristic1, characteristic2 = @characteristic2, age = @age, religion = @religion, political = @political, notes = @notes WHERE [ID] = @id"
            sqlquery.Parameters.AddWithValue("@job", x)
            sqlquery.Parameters.AddWithValue("@name", name)
            sqlquery.Parameters.AddWithValue("@gender", gender)
            sqlquery.Parameters.AddWithValue("@pacing", pacing)
            sqlquery.Parameters.AddWithValue("@tone", tone)
            sqlquery.Parameters.AddWithValue("@SpeechImpediment", speed)
            sqlquery.Parameters.AddWithValue("@accent", accent)
            sqlquery.Parameters.AddWithValue("@race", race)
            sqlquery.Parameters.AddWithValue("@motive", motive)
            sqlquery.Parameters.AddWithValue("@characteristic1", characteristic1)
            sqlquery.Parameters.AddWithValue("@characteristic2", charic2)
            sqlquery.Parameters.AddWithValue("@age", age)
            sqlquery.Parameters.AddWithValue("@religion", religion)
            sqlquery.Parameters.AddWithValue("@political", political)
            sqlquery.Parameters.AddWithValue("@notes", notes)
            sqlquery.Parameters.AddWithValue("@id", (rowsel + 1))
            sqlquery.ExecuteNonQuery()
            sqlconn.Close()
            'MsgBox("New dec, " & x & ", " & rowsel)
        Catch ex As Exception
            '   MessageBox.Show(ex.Message)
        End Try




        cmdUpdate.Dispose()
    End Function



    Public Function updatename(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String









        con.Close()
        cmdUpdate.Dispose()
    End Function
    Public Function updategender(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String



        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function
    Public Function updaterace(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String



        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function
    Public Function updatemotive(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String



        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function

    Public Function updatecharacteristic1(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function

    Public Function updatecharacteristic2(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function


    Public Function updateaccent(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String

        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function



    Public Function updateSpeechImpediment(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function


    Public Function updatepacing(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String

        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function

    Public Function updatetone(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function

    Public Function updatereligion(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String



        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function




    Public Function updatepolitical(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String

        x = DataGridView1.Rows(rowsel).Cells(14).Value
        'MsgBox(x)
        'Beta2 and Table
        cmdUpdate.CommandText = "UPDATE [" & TopDisp.TextBox1.Text & "] SET [" & "political" & "] = '" & x & "' 
        WHERE ID = " & id & ";"

        'MsgBox(cmdUpdate.CommandText)

        cmdUpdate.CommandType = CommandType.Text

        cmdUpdate.Connection = con

        cmdUpdate.ExecuteNonQuery()

        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function




    Public Function updateNotes(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As String


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function


    Public Function updateage(id As Integer, rowsel As Integer)
        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        ' Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        con.Open()
        '  If txtAu_ID.Text <> "" And txtAuthor.Text <> "" Then
        Dim x As Integer


        ' MsgBox("Record updated.")

        'txtAuthor.Text = ""

        ' Else

        ' MsgBox("Enter the required values:" & vbNewLine & "1. Au_ID" & vbNewLine &
        '        "2.Author")

        ' End If
        con.Close()
        cmdUpdate.Dispose()
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Form1.TextBox14.Text) ' Use relative path to database file

        'Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " & TopDisp.TextBox1.Text, con)

        'con.Open()









        ' Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(myDA)



        'Dim scb As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
        ' Save data back to the original database
        Try
            myDA.Update(myDataSet, TopDisp.TextBox1.Text)
        Catch ex As OleDbException
            MessageBox.Show(ex.Message)
        End Try



    End Sub
End Class