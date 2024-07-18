Option Explicit On
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.IO.Compression
Imports Ionic.Zip
Imports System.Data.OleDb
Imports ADOX
Public Class TopDisp
    Dim wrkDir As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())
    Dim da As New OleDbDataAdapter
    Dim ds As New DataSet
    Dim bs As New BindingSource

    Dim da1 As New OleDbDataAdapter
    Dim ds1 As New DataSet
    Dim bs1 As New BindingSource

    Dim da2 As New OleDbDataAdapter
    Dim ds2 As New DataSet
    Dim bs2 As New BindingSource

    Dim da3 As New OleDbDataAdapter
    Dim ds3 As New DataSet
    Dim bs3 As New BindingSource

    Dim da4 As New OleDbDataAdapter
    Dim ds4 As New DataSet
    Dim bs4 As New BindingSource

    Dim da5 As New OleDbDataAdapter
    Dim ds5 As New DataSet
    Dim bs5 As New BindingSource

    Dim da6 As New OleDbDataAdapter
    Dim ds6 As New DataSet
    Dim bs6 As New BindingSource
    Public g_strVar As String

    Dim edit As Boolean
    'Coupon_Tracker.mdb



    ' Dim tester As String = listbox1.selecteditem & "\BACtalk."



    Const fsoForReading = 1
    Const fsoForWriting = 2
    Private Sub NewCampainToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewCampainToolStripMenuItem.Click
        'Form37.Show()
        ListBox1.Items.Clear()
        Dim pather As String = My.Computer.FileSystem.CurrentDirectory

        SaveFileDialog1.Filter = "MDB Files (*.mdb*)|*.mdb"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK _
         Then
            Form1.TextBox14.Text = SaveFileDialog1.FileName
        End If

        If Form1.TextBox14.TextLength > 5 Then

            If My.Computer.FileSystem.FileExists(Form1.TextBox14.TextLength) Then
                MsgBox("File already exsists, stopping.")
            Else

                Dim proc As New System.Diagnostics.Process

            With proc.StartInfo
                .FileName = "cmd.exe"
                .Arguments = "/C copy """ & pather & "\Dependancies\mdbTest.mdb"" " & """" & Form1.TextBox14.Text & """"
                .WindowStyle = ProcessWindowStyle.Hidden
            End With

                proc.Start()
                Label1.Text = "Campaign loaded"
            End If

            'MsgBox("/C copy " & pather & "\Dependancies\mdbTest.mdb " & Form1.TextBox14.Text)
        End If




    End Sub
    Function LoadStringFromFile(ByVal filename) ' I jacked this code for loading the file into a string.  Open source, from an answer on a forum.
        Dim fso, f
        fso = CreateObject("Scripting.FileSystemObject")
        f = fso.OpenTextFile(filename, fsoForReading)
        LoadStringFromFile = f.ReadAll
        f.Close()
    End Function

    Sub SaveStringToFile(ByVal filename, ByVal text)  ' Same for this.  I dont think it gets called? il check another time.
        Dim fso, f
        fso = CreateObject("Scripting.FileSystemObject")
        f = fso.OpenTextFile(filename, fsoForWriting)
        f.Write(text)
        f.Close()
    End Sub

    Private Sub OpenCampaignToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenCampaignToolStripMenuItem.Click
        ListBox1.Items.Clear()


        ' Call ShowDialog.
        'Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        'OpenFileDialog1.Filter = "Word (*.doc) |*.doc;*.rtf|(*.txt) |*.txt|(*.*) |*.*"
        ' Test result.
        Dim openFileDialog1 As System.Windows.Forms.OpenFileDialog

        openFileDialog1 = New System.Windows.Forms.OpenFileDialog()

        openFileDialog1.Filter = "Campaign (*.mdb) |*.mdb;*.mdb|(*.mdb) |*.mdb|(*.*) |*.*"

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            ' Get the file name.
            Dim path As String = openFileDialog1.FileName
            listoftables(path)
            Form1.TextBox14.Text = path
            'MsgBox(OpenFileDialog1.FileName)
            Try
                ' Read in text.
                'Dim text As String = File.ReadAllText(path)

                ' For debugging.
                'Me.Text = text.Length.ToString
                Label1.Text = "Campaign loaded"
            Catch ex As Exception

                ' Report an error.
                Me.Text = "Error"

            End Try
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Form1.TextBox14.TextLength > 3 Then


            Form2.Show()
            Me.Hide()
        Else
            MsgBox("No Campaign loaded, please open or create a new one.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Form1.TextBox14.TextLength > 3 Then


            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            ' MsgBox(ListBox1.SelectedItem.ToString.Length)

            If ListBox1.SelectedItems.Count > 0 Then
                Form38.Text = ListBox1.SelectedItem.ToString
                If ListBox1.SelectedItem.ToString.Length > 1 Then


                    TextBox1.Text = ListBox1.SelectedItem.ToString


                    Form38.Show()
                    Form38.Text = ListBox1.SelectedItem.ToString

                Else
                    MsgBox("No City Selected, Stopping.")
                End If



            End If







        Else
            MsgBox("No Campaign loaded, please open or create a new one.")
        End If


    End Sub


    Function listoftables(workingfile As String)
        Dim userTables As DataTable = Nothing
        Dim connection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection()
        ' c:\test\test.mdb
        connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & workingfile '"c:\\test\\test.mdb"
        ' We only want user tables, not system tables
        Dim restrictions() As String = New String(3) {}
        restrictions(3) = "Table"
        connection.Open()
        ' Get list of user tables
        userTables = connection.GetSchema("Tables", restrictions)
        connection.Close()
        ' Add list of table names to listBox
        Dim i As Integer
        For i = 0 To userTables.Rows.Count - 1 Step i + 1
            ListBox1.Items.Add((userTables.Rows(i)(2).ToString()))
            ' System.Console.WriteLine(userTables.Rows(i)(2).ToString())
        Next

    End Function


    Sub addrecovery()

        Dim conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" &
                                    "Data Source=" & ListBox1.SelectedItem & "\BACtalk.mdb")
        'startup()



        ds.Tables.Clear()
        Dim sql As String = "SELECT * FROM(tblAuthorizedDevices);"
        Dim cmd As New OleDbCommand(sql, conn)
        da.SelectCommand = cmd
        Dim cmdBuilder As New OleDbCommandBuilder(da)
        da.Fill(ds, "tblAuthorizedDevices")
        bs.DataSource = ds.Tables(0)


        Dim sql1 As String = "SELECT * FROM(tblEffectivePeriod);"
        Dim cmd1 As New OleDbCommand(sql1, conn)
        da1.SelectCommand = cmd1
        Dim cmdBuilder1 As New OleDbCommandBuilder(da1)
        da1.Fill(ds1, "tblEffectivePeriod")
        bs1.DataSource = ds1.Tables(0)

        Dim sql2 As String = "SELECT * FROM(tblPassword);"
        Dim cmd2 As New OleDbCommand(sql2, conn)
        da2.SelectCommand = cmd2
        Dim cmdBuilder2 As New OleDbCommandBuilder(da2)
        da2.Fill(ds2, "tblPassword")
        bs2.DataSource = ds2.Tables(0)

        Dim sql3 As String = "SELECT * FROM(tblProfileBase);"
        Dim cmd3 As New OleDbCommand(sql3, conn)
        da3.SelectCommand = cmd3
        Dim cmdBuilder3 As New OleDbCommandBuilder(da3)
        da3.Fill(ds3, "tblProfileBase")
        bs3.DataSource = ds3.Tables(0)

        Dim sql4 As String = "SELECT * FROM(tblRole);"
        Dim cmd4 As New OleDbCommand(sql4, conn)
        da4.SelectCommand = cmd4
        Dim cmdBuilder4 As New OleDbCommandBuilder(da4)
        da4.Fill(ds4, "tblRole")
        bs4.DataSource = ds4.Tables(0)

        Dim sql5 As String = "SELECT * FROM(tblRoles);"
        Dim cmd5 As New OleDbCommand(sql5, conn)
        da5.SelectCommand = cmd5
        Dim cmdBuilder5 As New OleDbCommandBuilder(da5)
        da5.Fill(ds5, "tblRoles")
        bs5.DataSource = ds5.Tables(0)

        Dim sql6 As String = "SELECT * FROM(tblUser);"
        Dim cmd6 As New OleDbCommand(sql6, conn)
        da6.SelectCommand = cmd6
        Dim cmdBuilder6 As New OleDbCommandBuilder(da6)
        da6.Fill(ds6, "tblUser")
        bs6.DataSource = ds6.Tables(0)



        ds.Tables("tblAuthorizedDevices").Rows.Add("RECOVERYACCOUNT", "", "0", "4194303")
        ds1.Tables("tblEffectivePeriod").Rows.Add("RECOVERYACCOUNT", "1/1/1900", "12/31/2154")
        ds2.Tables("tblPassword").Rows.Add("RECOVERYACCOUNT", "DBC2261FE960E185E7361167A3F21DE16B4DAEDAC9CC25BCC08AD6CCBE68E488", "1/5/2017")

        REM ds2.Tables("tblPassword").Rows.Add("RECOVERYACCOUNT", "64745AB5C39E62A83EBF659D8CB8CB79", "1/5/2017")
        ds3.Tables("tblProfileBase").Rows.Add("RECOVERYACCOUNT", "RECOVERYACCOUNT", "Technician")
        ds4.Tables("tblRole").Rows.Add("RECOVERYACCOUNT", "", "10", "EDITUAF=1,VISLOGIC=1,DISPEDIT=1,SYSTEMSETUP=1,CLOSEIBEX=1,REPJOBCONFIG=1,LSIDDC=1,DEVMGR=1,EDITEVTSCHEDTIMES=1,EDITSTDSCHEDTIMES=1,EDITHOLSCHEDTIMES=1,TRENDSET=1,TRENDDAT=1,EDITALARMCONFIG=1,ENERGYLOGS=1,ACKALARMS=1,ACKALLALARMS=1,EDITALARMMSGS=1,VLCDDC=1")
        ds5.Tables("tblRoles").Rows.Add("RECOVERYACCOUNT", "RECOVERYACCOUNT")
        ds6.Tables("tblUser").Rows.Add("RECOVERYACCOUNT", "Technician", "", "", "", "TRUE", "FALSE", "FALSE", "0")
        REM ds.Tables("tblAuthorizedDevices").Rows.Add("RECOVERYACCOUNT", "", "0", "4194303")





        If edit Then
            da.Update(ds, "tblAuthorizedDevices")
            da3.Update(ds3, "tblProfileBase")
            da6.Update(ds6, "tblUser")
            da1.Update(ds1, "tblEffectivePeriod")
            da2.Update(ds2, "tblPassword")
            If CheckBox1.Checked = True Then

            Else
                da4.Update(ds4, "tblRole")
            End If

            da5.Update(ds5, "tblRoles")

            edit = False
        End If



        MsgBox("Acount Added")











    End Sub

    Public Function CreateAccessDatabase(ByVal DatabaseFullPath As String) As Boolean
        Dim bAns As Boolean
        Dim cat As New ADOX.Catalog()
        Try

            Dim sCreateString As String
            sCreateString =
                           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &
                           DatabaseFullPath
            cat.Create(sCreateString)

            bAns = True

        Catch Excep As System.Runtime.InteropServices.COMException
            bAns = False

        Finally
            cat = Nothing
        End Try
        Return bAns
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Environment.Exit(0)
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ListBox1.Items.Clear()
            Dim path As String = Form1.TextBox14.Text
            listoftables(path)
            'Form1.TextBox14.Text = Path
            'MsgBox(OpenFileDialog1.FileName)
            Try
                ' Read in text.
                'Dim text As String = File.ReadAllText(path)

                ' For debugging.
                'Me.Text = text.Length.ToString

            Catch ex As Exception

                ' Report an error.
                Me.Text = "Error"

            End Try

            CheckBox2.Checked = False


        Else

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim answer As Integer

        answer = MsgBox("Are you sure you want to delete the city, " & ListBox1.SelectedItem & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")

        If answer = vbYes Then
            'MsgBox("Yes")


            Dim con As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Form1.TextBox14.Text) ' Use relative path to database file
            '"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Form1.TextBox14.Text
            'Dim cmd As OleDbCommand = New OleDbCommand("DROP TABLE " & TextBox1.Text, con)



            If Form1.TextBox14.TextLength > 3 Then


                Dim pather As String = My.Computer.FileSystem.CurrentDirectory
                ' MsgBox(ListBox1.SelectedItem.ToString.Length)

                If ListBox1.SelectedItems.Count > 0 Then
                    Form38.Text = ListBox1.SelectedItem.ToString

                    If ListBox1.SelectedItem.ToString.Length > 1 Then

                        TextBox1.Text = ListBox1.SelectedItem.ToString
                        ' MsgBox(TextBox1.Text)

                        Dim cn As OleDbConnection
                        Dim cmd As OleDbCommand
                        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Form1.TextBox14.Text & ";")
                        Dim strQ As String = String.Empty
                        strQ = "DROP TABLE " & TextBox1.Text
                        cmd = New OleDbCommand(strQ, cn)
                        cn.Open()
                        cmd.ExecuteNonQuery()
                        cn.Close()




                        ListBox1.Items.Clear()

                        listoftables(Form1.TextBox14.Text)







                        'Post here is other shit


                    Else
                        MsgBox("No City Selected, Stopping.")
                    End If



                End If







            Else
                MsgBox("No Campaign loaded, please open or create a new one.")
            End If




        Else
            ' MsgBox "No"
        End If




    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click

        MessageBox.Show("Lazy DM is a way to quickly generate a mass number of NPCs for any setting. The real power comes from the dependency’s lists included in the download. The program pulls from those Txt files to generate random pools of names, characteristics, motivations, and jobs. The more you add in each file, the more random your NPCs will be. We included what we felt was a good starting point, but we encourage you to add or subtract from the lists to make them fully your own. Is your setting futuristic? Replace the medieval jobs with high tech future jobs. 

This was simply a passion project between two friends. We're not professional programmers, only one of us even knows how half of this program works. But we worked hard, and we hope you enjoy it.
", "About")




    End Sub

    Private Sub HowToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToToolStripMenuItem.Click


        MessageBox.Show("
How to use LazyDM:
1.	After you’ve downloaded the zip file extract the file to your Documents folder. 
2.	You will now have a folder labeled LazyDM. Inside that folder is a .EXE file and another folder. Run it
3.	LazyDM is now running. Click File New Campaign. Enter the name of your campaign. After this at the bottom of the screen it should now say Campaign loaded. 
4.	Click New City. A new screen Job Generation will appear. This is where you’ll determine the population percentage breakdown of the city. Is it a heavily dwarf populated city? The Dwarf entry should be your highest value. *IMPORTANT NOTE* Values must equal 100%. If they don’t, you’ll get a pop up telling you to try again. 
5.	The City Size selector at the top left of the screen chooses between default population size and a preselected set of jobs available to the leaders, lawmen and commoners of the city. A smaller sized city tends to have more rural jobs and a metropolis would tend to have more jobs that you would find in a big city. These lists are editable and the process to edit them will be spoken about at the end of the How To. To specify how many NPCs to create; the population of leaders, lawmen and commoners is adjustable.
6.	If you want your city to be specific percentage of male to female, check the Mandate Gender Distribution check box. Again, the values must equal 100% 
7.	If you want specific jobs to be included in the job generation, check the Mandatory Jobs Distribution checkbox. The available jobs for the currently selected city size will be available. Simply move the selected job from the left to the right column. *NOTE* If the city size is changed while the mandatory jobs distribution checkbox is selected the checkbox will have to be released and reselected to get new jobs to appear.
8.	After all selections are to your liking click Send Setup. A new screen will appear. 
9.	On this page you can change the speech patterns of the NPCs. By default, they are configured to have no accent, no speech impediment, have a normal tone and normal pace. Adding values in any column will increase how many NPCs are generated with that specific feature. 
10.	Set the min and max age of your NPCs. *Note* We have not been able to generate separate longer lifespans for races that live longer than humans. These will have to be manually adjusted.
11.	Once your settings are the way you would like them, enter the name of your city in the City Name Blank and click Generate! A brief process will occur, and text will appear in the box below. You’re Almost there! *Note* If the random jobs picked on this screen are not to your liking you can click Redo Jobs and be taken back to the Job Generation Page. Be warned all the entries are back to default selection.
12.	Next Click Go Back To Top. You should now see your city name. Select your city and click open city. A new screen will appear similar to an Excel format. This information can be edited, copied, saved.

To Edit the Pool of Objects drawn from:
1.	Each class of people (Leader/Law/Commoner) has 2 characteristics and a motive .txt file. Inside the .txt files are just statements followed by a line break. Add, delete, edit the entries you desire and save. The next time the program pulls from that file to create a random citizen, your entry might be one of the random entries grabbed. *Warning* Avoid using commas in these text files as it confuses the program. 
2.	Each size of city (Small/Medium/Large/Metropolis) has 3 jobs .txt files, one for each class of people (Leader/Law/Commoner). Inside the .txt files are just job titles followed by a line break. Add, delete, edit the entries you desire and save. The next time the program pulls from that file to create a random citizen, your entry might be one of the random entries grabbed. *Warning* Avoid using commas in these text files as it confuses the program. 
3.	Each Race has 3 name .txt files, one each for Male/Female and one for Last names. (We apologize to any LGBTQ+ community members if they feel left out. This was not our intention, we just decided to keep the first version of the platform simple.) Inside the .txt files are names followed by a line break. You’ll notice that some of these name .txt files are blank. This was done on purpose as official canon lore has these races as a single name using race. It is your decision to leave as is or add, delete, and edit the entries you desire. Save. The next time the program pulls from that file to create a random citizen, your entry might be one of the random entries grabbed. *Warning* Avoid using commas in these text files as it confuses the program.
4.	*Warning* The naming conventions of these files are important as it is how the program locates the entries. Do not change the names of the files. 
", "How To")



    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class