Option Explicit On
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.IO.Compression
Imports Ionic.Zip
Imports System.Data.OleDb


Public Class Form1

    'Public Shared lstMain As System.Windows.Forms.ListBox7

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Randomize()

        Dim Name As String = TextBox1.Text

        Dim pacing As String = ListBox1.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
        Dim tone As String = ListBox2.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
        Dim speed As String = ListBox3.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
        Dim accent As String = ListBox4.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)

        'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound

        MsgBox(TextBox1.Text & vbCrLf & pacing & vbCrLf & tone & vbCrLf & speed & vbCrLf & accent)










    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged

    End Sub

    Private Sub ListBox5_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    Function generatecom(race As String, gender As String)
        Dim firstname As String = ""
        Dim lastname As String = ""

        If race = "Aarakocra" Then
            lastname = Form5.ListBox2.Items(CInt(Math.Floor((Form5.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form5.ListBox1.Items(CInt(Math.Floor((Form5.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form5.ListBox3.Items(CInt(Math.Floor((Form5.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Aasimar" Then
            lastname = Form6.ListBox2.Items(CInt(Math.Floor((Form6.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form6.ListBox1.Items(CInt(Math.Floor((Form6.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form6.ListBox3.Items(CInt(Math.Floor((Form6.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Dragonborn" Then
            lastname = Form7.ListBox2.Items(CInt(Math.Floor((Form7.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form7.ListBox1.Items(CInt(Math.Floor((Form7.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form7.ListBox3.Items(CInt(Math.Floor((Form7.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Dwarf" Then
            lastname = Form8.ListBox2.Items(CInt(Math.Floor((Form8.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form8.ListBox1.Items(CInt(Math.Floor((Form8.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form8.ListBox3.Items(CInt(Math.Floor((Form8.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If



        ElseIf race = "Elf" Then
            lastname = Form9.ListBox2.Items(CInt(Math.Floor((Form9.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form9.ListBox1.Items(CInt(Math.Floor((Form9.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form9.ListBox3.Items(CInt(Math.Floor((Form9.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Genasi" Then
            lastname = Form10.ListBox2.Items(CInt(Math.Floor((Form10.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form10.ListBox1.Items(CInt(Math.Floor((Form10.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form10.ListBox3.Items(CInt(Math.Floor((Form10.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Gnome" Then
            lastname = Form11.ListBox2.Items(CInt(Math.Floor((Form11.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form11.ListBox1.Items(CInt(Math.Floor((Form11.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form11.ListBox3.Items(CInt(Math.Floor((Form11.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Goliath" Then
            lastname = Form12.ListBox2.Items(CInt(Math.Floor((Form12.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form12.ListBox1.Items(CInt(Math.Floor((Form12.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form12.ListBox3.Items(CInt(Math.Floor((Form12.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Half-Elf" Then
            lastname = Form13.ListBox2.Items(CInt(Math.Floor((Form13.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form13.ListBox1.Items(CInt(Math.Floor((Form13.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form13.ListBox3.Items(CInt(Math.Floor((Form13.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Human" Then
            lastname = Form14.ListBox2.Items(CInt(Math.Floor((Form14.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form14.ListBox1.Items(CInt(Math.Floor((Form14.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form14.ListBox3.Items(CInt(Math.Floor((Form14.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "LizardFolk" Then
            lastname = Form15.ListBox2.Items(CInt(Math.Floor((Form15.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form15.ListBox1.Items(CInt(Math.Floor((Form15.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form15.ListBox3.Items(CInt(Math.Floor((Form15.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Tabaxi" Then
            lastname = Form16.ListBox2.Items(CInt(Math.Floor((Form16.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form16.ListBox1.Items(CInt(Math.Floor((Form16.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form16.ListBox3.Items(CInt(Math.Floor((Form16.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If



        ElseIf race = "Tiefling" Then
            lastname = Form17.ListBox2.Items(CInt(Math.Floor((Form17.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form17.ListBox1.Items(CInt(Math.Floor((Form17.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form17.ListBox3.Items(CInt(Math.Floor((Form17.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Triton" Then
            lastname = Form18.ListBox2.Items(CInt(Math.Floor((Form18.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form18.ListBox1.Items(CInt(Math.Floor((Form18.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form18.ListBox3.Items(CInt(Math.Floor((Form18.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Warforged" Then
            lastname = Form19.ListBox2.Items(CInt(Math.Floor((Form19.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form19.ListBox1.Items(CInt(Math.Floor((Form19.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form19.ListBox3.Items(CInt(Math.Floor((Form19.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Yuanti-Pureblood" Then
            lastname = Form20.ListBox2.Items(CInt(Math.Floor((Form20.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form20.ListBox1.Items(CInt(Math.Floor((Form20.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form20.ListBox3.Items(CInt(Math.Floor((Form20.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "firbolg" Then
            lastname = Form35.ListBox2.Items(CInt(Math.Floor((Form35.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form35.ListBox1.Items(CInt(Math.Floor((Form35.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form35.ListBox3.Items(CInt(Math.Floor((Form35.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Gith" Then
            lastname = Form34.ListBox2.Items(CInt(Math.Floor((Form34.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form34.ListBox1.Items(CInt(Math.Floor((Form34.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form34.ListBox3.Items(CInt(Math.Floor((Form34.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Goblin" Then
            lastname = Form33.ListBox2.Items(CInt(Math.Floor((Form33.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form33.ListBox1.Items(CInt(Math.Floor((Form33.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form33.ListBox3.Items(CInt(Math.Floor((Form33.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Half-orc" Then
            lastname = Form32.ListBox2.Items(CInt(Math.Floor((Form32.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form32.ListBox1.Items(CInt(Math.Floor((Form32.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form32.ListBox3.Items(CInt(Math.Floor((Form32.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Hobgoblin" Then
            lastname = Form31.ListBox2.Items(CInt(Math.Floor((Form31.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form31.ListBox1.Items(CInt(Math.Floor((Form31.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form31.ListBox3.Items(CInt(Math.Floor((Form31.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Kalashtar" Then
            lastname = Form30.ListBox2.Items(CInt(Math.Floor((Form30.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form30.ListBox1.Items(CInt(Math.Floor((Form30.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form30.ListBox3.Items(CInt(Math.Floor((Form30.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = "Kobold" Then
            lastname = Form29.ListBox2.Items(CInt(Math.Floor((Form29.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form29.ListBox1.Items(CInt(Math.Floor((Form29.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form29.ListBox3.Items(CInt(Math.Floor((Form29.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Loxodon" Then
            lastname = Form28.ListBox2.Items(CInt(Math.Floor((Form28.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form28.ListBox1.Items(CInt(Math.Floor((Form28.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form28.ListBox3.Items(CInt(Math.Floor((Form28.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "ORC" Then
            lastname = Form27.ListBox2.Items(CInt(Math.Floor((Form27.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form27.ListBox1.Items(CInt(Math.Floor((Form27.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form27.ListBox3.Items(CInt(Math.Floor((Form27.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = "Tortle" Then
            lastname = Form26.ListBox2.Items(CInt(Math.Floor((Form26.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form26.ListBox1.Items(CInt(Math.Floor((Form26.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form26.ListBox3.Items(CInt(Math.Floor((Form26.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If




        ElseIf race = TextBox9.Text Then
            ' If TextBox9.TextLength > 2 Then
            lastname = Form21.ListBox2.Items(CInt(Math.Floor((Form21.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form21.ListBox1.Items(CInt(Math.Floor((Form21.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form21.ListBox3.Items(CInt(Math.Floor((Form21.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = TextBox10.Text Then
            ' If TextBox9.TextLength > 2 Then
            lastname = Form22.ListBox2.Items(CInt(Math.Floor((Form22.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form22.ListBox1.Items(CInt(Math.Floor((Form22.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form22.ListBox3.Items(CInt(Math.Floor((Form22.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If



        ElseIf race = TextBox11.Text Then
            ' If TextBox9.TextLength > 2 Then
            lastname = Form23.ListBox2.Items(CInt(Math.Floor((Form23.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form23.ListBox1.Items(CInt(Math.Floor((Form23.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form23.ListBox3.Items(CInt(Math.Floor((Form23.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If


        ElseIf race = TextBox12.Text Then
            ' If TextBox9.TextLength > 2 Then
            lastname = Form24.ListBox2.Items(CInt(Math.Floor((Form24.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form24.ListBox1.Items(CInt(Math.Floor((Form24.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form24.ListBox3.Items(CInt(Math.Floor((Form24.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If

        ElseIf race = TextBox13.Text Then
            ' If TextBox9.TextLength > 2 Then
            lastname = Form25.ListBox2.Items(CInt(Math.Floor((Form25.ListBox2.Items.Count - 0) * Rnd())) + 0)
            If gender = "Female" Then
                firstname = Form25.ListBox1.Items(CInt(Math.Floor((Form25.ListBox1.Items.Count - 0) * Rnd())) + 0)
            Else
                firstname = Form25.ListBox3.Items(CInt(Math.Floor((Form25.ListBox3.Items.Count - 0) * Rnd())) + 0)
            End If





        End If




        Name = firstname & " " & lastname
        If Name.Length < 2 Then
            MsgBox("I wrote something wrong involving the name. I appologize. Please send the setup to me")
        End If
        Return Name

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox8.Items.Clear()
        'Add Speech Speed
        ListBox1.Items.Clear()
        Randomize()

        'Add Super Fast
        If NumericUpDown1.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown1.Value
                ListBox1.Items.Add("Super Fast")
                this = (this + 1)
            Loop
        End If
        'Add Quick
        If NumericUpDown2.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown2.Value
                ListBox1.Items.Add("Quick")
                this = (this + 1)
            Loop
        End If
        'Add Normal
        If NumericUpDown4.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown4.Value
                ListBox1.Items.Add("Normal")
                this = (this + 1)
            Loop
        End If
        'Add Semi Slow
        If NumericUpDown6.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown6.Value
                ListBox1.Items.Add("Semi-Slow")
                this = (this + 1)
            Loop
        End If
        'Add Semi Slow
        If NumericUpDown5.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown5.Value
                ListBox1.Items.Add("Very Slow")
                this = (this + 1)
            Loop
        End If



        'Add Speech Tone
        ListBox2.Items.Clear()


        'Add Very High
        If NumericUpDown10.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown10.Value
                ListBox2.Items.Add("Very High")
                this = (this + 1)
            Loop
        End If
        'Add High
        If NumericUpDown9.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown9.Value
                ListBox2.Items.Add("High")
                this = (this + 1)
            Loop
        End If
        'Add Mid Pitch
        If NumericUpDown8.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown8.Value
                ListBox2.Items.Add("Mid Pitch")
                this = (this + 1)
            Loop
        End If
        'Add Low
        If NumericUpDown7.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown7.Value
                ListBox2.Items.Add("Low")
                this = (this + 1)
            Loop
        End If
        'Add Very Low
        If NumericUpDown3.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown3.Value
                ListBox2.Items.Add("Very Low")
                this = (this + 1)
            Loop
        End If








        'Add Speech Impetatmint
        ListBox3.Items.Clear()


        'Add Stutter
        If NumericUpDown19.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown19.Value
                ListBox3.Items.Add("Stutter")
                this = (this + 1)
            Loop
        End If
        'Add Slurred Speech
        If NumericUpDown18.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown18.Value
                ListBox3.Items.Add("Slurred Speech")
                this = (this + 1)
            Loop
        End If
        'Add Lisping
        If NumericUpDown17.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown17.Value
                ListBox3.Items.Add("Lisping")
                this = (this + 1)
            Loop
        End If
        'Add Cluttering
        If NumericUpDown16.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown16.Value
                ListBox3.Items.Add("Cluttering")
                this = (this + 1)
            Loop
        End If
        'Add No Speech Impediment
        Dim totals As Integer = 0
        totals = (NumericUpDown19.Value + NumericUpDown18.Value + NumericUpDown17.Value + NumericUpDown16.Value)
        Dim rest As Integer = 100
        rest = (rest - totals)
        If rest > 0.5 Then
            Dim this As Integer = 0
            Do Until this = rest
                ListBox3.Items.Add("No Speech Impediment")
                this = (this + 1)
            Loop
        End If





        'Add Speech Accent
        ListBox4.Items.Clear()


        'Add text3
        If NumericUpDown14.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown14.Value
                ListBox4.Items.Add(TextBox3.Text)
                this = (this + 1)
            Loop
        End If
        'Add 13
        If NumericUpDown13.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown13.Value
                ListBox4.Items.Add(TextBox4.Text)
                this = (this + 1)
            Loop
        End If
        'Add 12
        If NumericUpDown12.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown12.Value
                ListBox4.Items.Add(TextBox5.Text)
                this = (this + 1)
            Loop
        End If
        'Add 11
        If NumericUpDown11.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown11.Value
                ListBox4.Items.Add(TextBox6.Text)
                this = (this + 1)
            Loop
        End If
        If NumericUpDown15.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown15.Value
                ListBox4.Items.Add(TextBox7.Text)
                this = (this + 1)
            Loop
        End If

        ' MsgBox(ListBox1.Items.ToString)








        Randomize()
        addnewcity()

        Dim i As Long, R As Long, strSwp As String
        For i = ListBox7.Items.Count - 1 To 0 Step -1
            R = Int(i * Rnd())
            strSwp = ListBox7.Items(i)
            ListBox7.Items(i) = ListBox7.Items(R)
            ListBox7.Items(R) = strSwp
        Next i
        Dim x As Integer = ListBox7.Items.Count
        Dim y As Integer = 0
        Do Until y = x
            Dim gender As String = ListBox24.Items(CInt(Math.Floor((ListBox24.Items.Count - 0) * Rnd())) + 0)
            Dim pacing As String = ListBox1.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
            Dim tone As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
            Dim speed As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
            Dim accent As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
            Dim race As String = ListBox9.Items(CInt(Math.Floor((ListBox9.Items.Count - 0) * Rnd())) + 0)
            Dim age As Integer = (CInt(Math.Floor((NumericUpDown21.Value - NumericUpDown20.Value) * Rnd())) + NumericUpDown20.Value)
            'Dim 
            'Dim age As Integer = CInt(Math.Floor((NumericUpDown21.Value - 0) * Rnd())) + NumericUpDown20.Value)
            Dim job As String = ListBox7.Items(y)
            Dim Name As String = generatecom(race, gender)
            Dim religion As String = ListBox25.Items(CInt(Math.Floor((ListBox25.Items.Count - 0) * Rnd())) + 0)
            Dim political As String = ListBox26.Items(CInt(Math.Floor((ListBox26.Items.Count - 0) * Rnd())) + 0)
            Dim motive As String = Form4.ListBox7.Items(CInt(Math.Floor((Form4.ListBox7.Items.Count - 0) * Rnd())) + 0)
            Dim characteristic1 As String = Form4.ListBox8.Items(CInt(Math.Floor((Form4.ListBox8.Items.Count - 0) * Rnd())) + 0)
            Dim charic2 As String = Form4.ListBox9.Items(CInt(Math.Floor((Form4.ListBox9.Items.Count - 0) * Rnd())) + 0)
            ListBox8.Items.Add(Name & ", " & job & ", " & gender & ", " & pacing & ", " & tone & ", " & speed & ", " & accent & ", " & race &
                                   ", " & motive & ", " & characteristic1 & ",  " & charic2 & ", " & age & ", " & religion & ", " & political & " ")
            Dim con As New OleDb.OleDbConnection
            Dim dbProvider As String
            Dim dbSource As String
            Dim ds As New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            Dim inc As Integer
            Dim MaxRows As Integer

            Try
                Dim sqlconn As New OleDb.OleDbConnection
                Dim sqlquery As New OleDb.OleDbCommand
                Dim connString As String
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TextBox14.Text
                sqlconn.ConnectionString = connString
                sqlquery.Connection = sqlconn
                sqlconn.Open()


                '3 is speachimped, called accent -speed
                '4 is accent - accent
                '2 tone - tone
                '1 pace - pacing



                sqlquery.CommandText = "INSERT INTO " & TextBox8.Text & "([name], [job], [gender], [pacing], [tone], [SpeechImpediment], [accent], [race], [motive], [characteristic1], [characteristic2], [age], [religion], [political])VALUES(@name, @job, @gender, @pacing, @tone, @SpeechImpediment, @accent, @race, @motive, @characteristic1, @characteristic2, @age, @religion, @political)"
                sqlquery.Parameters.AddWithValue("@name", Name)
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
                sqlquery.ExecuteNonQuery()
                sqlconn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            y = (y + 1)
            'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
        Loop
        y = 0
        x = ListBox22.Items.Count
        Do Until y = x

            Dim pacing As String = ListBox1.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
            Dim tone As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
            Dim speed As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
            Dim accent As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
            Dim race As String = ListBox9.Items(CInt(Math.Floor((ListBox9.Items.Count - 0) * Rnd())) + 0)
            Dim age As Integer = (CInt(Math.Floor((NumericUpDown21.Value - NumericUpDown20.Value) * Rnd())) + NumericUpDown20.Value)
            Dim gender As String = ListBox24.Items(CInt(Math.Floor((ListBox24.Items.Count - 0) * Rnd())) + 0)
            Dim job As String = ListBox22.Items(y)
            Dim Name As String = generatecom(race, gender)
            Dim religion As String = ListBox25.Items(CInt(Math.Floor((ListBox25.Items.Count - 0) * Rnd())) + 0)
            Dim political As String = ListBox26.Items(CInt(Math.Floor((ListBox26.Items.Count - 0) * Rnd())) + 0)
            Dim motive As String = Form4.ListBox7.Items(CInt(Math.Floor((Form4.ListBox7.Items.Count - 0) * Rnd())) + 0)
            Dim characteristic1 As String = Form4.ListBox8.Items(CInt(Math.Floor((Form4.ListBox8.Items.Count - 0) * Rnd())) + 0)
            Dim charic2 As String = Form4.ListBox9.Items(CInt(Math.Floor((Form4.ListBox9.Items.Count - 0) * Rnd())) + 0)
            ListBox8.Items.Add(Name & ", " & job & ", " & gender & ", " & pacing & ", " & tone & ", " & speed & ", " & accent & ", " & race _
                               & ", " & motive & ", " & characteristic1 & ",  " & charic2 & ", " & age & ", " & religion & ", " & political & " ")




            Dim con As New OleDb.OleDbConnection
            Dim dbProvider As String
            Dim dbSource As String
            Dim ds As New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            Dim inc As Integer
            Dim MaxRows As Integer

            Try
                Dim sqlconn As New OleDb.OleDbConnection
                Dim sqlquery As New OleDb.OleDbCommand
                Dim connString As String
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TextBox14.Text
                sqlconn.ConnectionString = connString
                sqlquery.Connection = sqlconn
                sqlconn.Open()


                sqlquery.CommandText = "INSERT INTO " & TextBox8.Text & "([name], [job], [gender], [pacing], [tone], [SpeechImpediment], [accent], [race], [motive], [characteristic1], [characteristic2], [age], [religion], [political])VALUES(@name, @job, @gender, @pacing, @tone, @SpeechImpediment, @accent, @race, @motive, @characteristic1, @characteristic2, @age, @religion, @political)"
                sqlquery.Parameters.AddWithValue("@name", Name)
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
                sqlquery.ExecuteNonQuery()
                sqlconn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            y = (y + 1)
        Loop

        y = 0
        x = ListBox23.Items.Count
        Do Until y = x

            Dim pacing As String = ListBox1.Items(CInt(Math.Floor((ListBox1.Items.Count - 0) * Rnd())) + 0)
            Dim tone As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
            Dim speed As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
            Dim accent As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
            Dim race As String = ListBox9.Items(CInt(Math.Floor((ListBox9.Items.Count - 0) * Rnd())) + 0)
            Dim age As Integer = (CInt(Math.Floor((NumericUpDown21.Value - NumericUpDown20.Value) * Rnd())) + NumericUpDown20.Value)
            Dim gender As String = ListBox24.Items(CInt(Math.Floor((ListBox24.Items.Count - 0) * Rnd())) + 0)
            Dim job As String = ListBox23.Items(y)
            Dim Name As String = generatecom(race, gender)
            Dim religion As String = ListBox25.Items(CInt(Math.Floor((ListBox25.Items.Count - 0) * Rnd())) + 0)
            Dim political As String = ListBox26.Items(CInt(Math.Floor((ListBox26.Items.Count - 0) * Rnd())) + 0)
            Dim motive As String = Form4.ListBox7.Items(CInt(Math.Floor((Form4.ListBox7.Items.Count - 0) * Rnd())) + 0)
            Dim characteristic1 As String = Form4.ListBox8.Items(CInt(Math.Floor((Form4.ListBox8.Items.Count - 0) * Rnd())) + 0)
            Dim charic2 As String = Form4.ListBox9.Items(CInt(Math.Floor((Form4.ListBox9.Items.Count - 0) * Rnd())) + 0)
            ListBox8.Items.Add(Name & ", " & job & ", " & gender & ", " & pacing & ", " & tone & ", " & speed & ", " & accent & ", " & race _
                               & ", " & motive & ", " & characteristic1 & ",  " & charic2 & ", " & age & ", " & religion & ", " & political & " ")



            Dim con As New OleDb.OleDbConnection
            Dim dbProvider As String
            Dim dbSource As String
            Dim ds As New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            Dim inc As Integer
            Dim MaxRows As Integer

            Try
                Dim sqlconn As New OleDb.OleDbConnection
                Dim sqlquery As New OleDb.OleDbCommand
                Dim connString As String
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TextBox14.Text
                sqlconn.ConnectionString = connString
                sqlquery.Connection = sqlconn
                sqlconn.Open()



                sqlquery.CommandText = "INSERT INTO " & TextBox8.Text & "([name], [job], [gender], [pacing], [tone], [SpeechImpediment], [accent], [race], [motive], [characteristic1], [characteristic2], [age], [religion], [political])VALUES(@name, @job, @gender, @pacing, @tone, @SpeechImpediment, @accent, @race, @motive, @characteristic1, @characteristic2, @age, @religion, @political)"
                sqlquery.Parameters.AddWithValue("@name", Name)
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
                sqlquery.ExecuteNonQuery()
                sqlconn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            y = (y + 1)
        Loop



        'MsgBox(TextBox1.Text & vbCrLf & pacing & vbCrLf & tone & vbCrLf & speed & vbCrLf & accent)

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListBox7_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = Keys.Delete AndAlso ListBox1.SelectedItem <> Nothing Then
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Form1.Show(False)
        Dim testing As Boolean = depcheck(testing)
        If testing = True Then

            MsgBox("Missing dependancies. Not Loaded, this will fail.")
        Else

            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallCommonerJobs.txt")
            ListBox10.Items.AddRange(lines)

            'Dim pather1 As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines1() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedCommonerJobs.txt")
            ListBox11.Items.AddRange(lines1)


            'Dim pather2 As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines2() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgCommonerJobs.txt")
            ListBox12.Items.AddRange(lines2)

            Dim lines3() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroCommonerJobs.txt")
            ListBox13.Items.AddRange(lines3)



            'Add Law Jobs

            Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgLawJobs.txt")
            ListBox17.Items.AddRange(lines4)
            'MsgBox(pather & "\Dependancies\LrgLawJobs.txt")
            Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedLawJobs.txt")
            ListBox16.Items.AddRange(lines5)

            Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroLawJobs.txt")
            ListBox15.Items.AddRange(lines6)

            Dim lines7() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallLawJobs.txt")
            ListBox14.Items.AddRange(lines7)


            'Add Leade3r Jobs


            Dim lines8() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgLeaderJobs.txt")
            ListBox21.Items.AddRange(lines8)

            Dim lines9() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedLeaderJobs.txt")
            ListBox20.Items.AddRange(lines9)

            Dim lines10() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroLeaderJobs.txt")
            ListBox19.Items.AddRange(lines10)

            Dim lines11() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallLeaderJobs.txt")
            ListBox18.Items.AddRange(lines11)


            'religion deps


            Dim lines12() As String = IO.File.ReadAllLines(pather & "\Dependancies\Religion.txt")
            ListBox25.Items.AddRange(lines12)
            'Political deps

            Dim lines13() As String = IO.File.ReadAllLines(pather & "\Dependancies\Political.txt")
            ListBox26.Items.AddRange(lines13)

        End If
        'Form2.Show()
        'Me.Hide()


        ' File.Close()
    End Sub


    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox8.Text.Length < 1 Then
            MsgBox("Please add a city name before saving.")
        Else

            Dim file As System.IO.StreamWriter
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            file = My.Computer.FileSystem.OpenTextFileWriter((pather & "\" & TextBox8.Text & ".csv"), True)
            Dim x As Integer = ListBox8.Items.Count
            Dim y As Integer = 0
            Do Until y = x
                file.WriteLine(ListBox8.Items(y))
                y = (y + 1)
            Loop

            file.Close()
        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged

    End Sub

    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged

    End Sub

    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub NumericUpDown10_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown10.ValueChanged

    End Sub

    Private Sub NumericUpDown9_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click

    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click

    End Sub

    Private Sub NumericUpDown8_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown8.ValueChanged

    End Sub

    Private Sub NumericUpDown7_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged

    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged

    End Sub

    Private Sub NumericUpDown14_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown14.ValueChanged

    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub

    Private Sub NumericUpDown13_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown13.ValueChanged

    End Sub

    Private Sub NumericUpDown12_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown12.ValueChanged

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click

    End Sub

    Private Sub NumericUpDown11_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown11.ValueChanged

    End Sub

    Private Sub NumericUpDown15_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown15.ValueChanged

    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click

    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub NumericUpDown19_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown19.ValueChanged

    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click

    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click

    End Sub

    Private Sub NumericUpDown18_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown18.ValueChanged

    End Sub

    Private Sub NumericUpDown17_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown17.ValueChanged

    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub NumericUpDown16_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown16.ValueChanged

    End Sub



    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        ListBox5.Items.AddRange(Split(RichTextBox1.Text, vbNewLine))
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ListBox6.Items.AddRange(Split(RichTextBox2.Text, vbNewLine))
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ListBox7.Items.Add(TextBox2.Text)
        TextBox2.Text = ""
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Form2.Show()
    End Sub

    Private Sub ListBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox8.SelectedIndexChanged

    End Sub

    Private Sub ListBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox9.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("This tool is to help you generate entire cities at a time worth of people " &
               vbCrLf & "It is brought to you by sQUEEK and sWEEDISH_F00L" &
               vbCrLf & "If you get an Unhandled exception error, I fucked up Email me at endjamenson@gmail.com and tell me how it happened, with pics if possible" &
               vbCrLf & "If you get an error box, it should tell you how you messed up, however if the FAQ doesnt exsist yet, or you dont understand, do the same.")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()
        ListBox22.Items.Clear()
        ListBox23.Items.Clear()
    End Sub
    Function depcheck(booler As Boolean) As Boolean
        Dim errorcheck As Integer = 0
        Dim pather As String = My.Computer.FileSystem.CurrentDirectory


        'Check for CORE Stuff
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\Political.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\Political.txt")
        End If
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\Religion.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\Religion.txt")
        End If




        'Check for CORE Stuff
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LawCharacteristics1.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LawCharacteristics1.txt")
        End If

        'Check for CORE Stuff
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LawCharacteristics2.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LawCharacteristics2.txt")
        End If

        'Check for CORE Stuff
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LawMotive.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LawMotive.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LeaderCharacteristics1.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LeaderCharacteristics1.txt")
        End If
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LeaderCharacteristics2.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LeaderCharacteristics2.txt")
        End If
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LeaderMotive.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LeaderMotive.txt")
        End If

        'Commoner stuff
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\CommonCharacteristics1.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\CommonCharacteristics1.txt")
        End If
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\CommonCharacteristics2.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\CommonCharacteristics2.txt")
        End If
        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\CommonMotive.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\CommonMotive.txt")
        End If





        'Check for lawman books

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LrgLawJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LrgLawJobs.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\MedLawJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\MedLawJobs.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\MetroLawJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\MetroLawJobs.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\SmallLawJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\SmallLawJobs.txt")
        End If

        'Check for Governer books

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\LrgLeaderJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\LrgLeaderJobs.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\MedLeaderJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\MedLeaderJobs.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\MetroLeaderJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\MetroLeaderJobs.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\SmallLeaderJobs.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\SmallLeaderJobs.txt")
        End If

        'Aarakocra

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aarakocrafemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aarakocrafemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aarakocralast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aarakocralast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aarakocramale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aarakocramale.txt")
        End If

        'Aasimarfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aasimarfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aasimarfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aasimarlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aasimarlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Aasimarmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Aasimarmale.txt")
        End If


        'Dragonbornfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dragonbornfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dragonbornfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dragonbornlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dragonbornlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dragonbornmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dragonbornmale.txt")
        End If

        'Dwarffemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dwarffemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dwarffemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dwarflast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dwarflast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Dwarfmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Dwarfmale.txt")
        End If

        'ELFfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ELFfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ELFfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ELFlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ELFlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ELFmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ELFmale.txt")
        End If

        'Genasifemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Genasifemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Genasifemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Genasilast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Genasilast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Genasimale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Genasimale.txt")
        End If

        'Gnomefemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Gnomefemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Gnomefemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Gnomelast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Gnomelast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Gnomemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Gnomemale.txt")
        End If

        'Goliathfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goliathfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goliathfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goliathlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goliathlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goliathmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goliathmale.txt")
        End If

        'Half-Elffemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-Elffemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-Elffemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-Elflast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-Elflast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-Elfmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-Elfmale.txt")
        End If

        'humanfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\humanfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\humanfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\humanlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\humanlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\humanmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\humanmale.txt")
        End If


        'LizardFolkfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\LizardFolkfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\LizardFolkfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\LizardFolklast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\LizardFolklast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\LizardFolkmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\LizardFolkmale.txt")
        End If


        'Tabaxifemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tabaxifemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tabaxifemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tabaxilast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tabaxilast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tabaximale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tabaximale.txt")
        End If

        'Tieflingfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tieflingfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tieflingfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tieflinglast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tieflinglast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tieflingmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tieflingmale.txt")
        End If


        'Tritonfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tritonfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tritonfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tritonlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tritonlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tritonmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tritonmale.txt")
        End If


        'Warforgedfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Warforgedfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Warforgedfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Warforgedlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Warforgedlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Warforgedmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Warforgedmale.txt")
        End If


        'Yuantifemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Yuantifemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Yuantifemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Yuantilast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Yuantilast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Yuantimale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Yuantimale.txt")
        End If



        'firbolgfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\firbolgfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\firbolgfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\firbolglast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\firbolglast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\firbolgmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\firbolgmale.txt")
        End If






        'Githfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Githfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Githfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Githlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Githlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Githmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Githmale.txt")
        End If





        'Goblinfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goblinfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goblinfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goblinlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goblinlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Goblinmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Goblinmale.txt")
        End If



        'Half-orcfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-orcfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-orcfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-orclast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-orclast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Half-orcmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Half-orcmale.txt")
        End If


        'Hobgoblinfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Hobgoblinfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Hobgoblinfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Hobgoblinlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Hobgoblinlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Hobgoblinmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Hobgoblinmale.txt")
        End If


        'Kalashtarfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Kalashtarfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Kalashtarfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Kalashtarlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Kalashtarlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Kalashtarmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Kalashtarmale.txt")
        End If


        'Koboldfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Koboldfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Koboldfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Koboldlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Koboldlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Koboldmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Koboldmale.txt")
        End If



        'Loxodonfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Loxodonfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Loxodonfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Loxodonlast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Loxodonlast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Loxodonmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Loxodonmale.txt")
        End If





        'ORCfemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ORCfemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ORCfemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ORClast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ORClast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\ORCmale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\ORCmale.txt")
        End If




        'Tortlefemale.txt

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tortlefemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tortlefemale.txt")
        End If


        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tortlelast.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tortlelast.txt")
        End If

        If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\Tortlemale.txt") Then
            ' MsgBox("File found.")
        Else
            errorcheck = (errorcheck + 1)
            MsgBox("Missing Dependancies\names\Tortlemale.txt")
        End If



        If NumericUpDown17.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox1.Text & "female.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox1.Text & "female.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If
        If NumericUpDown17.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox1.Text & "last.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox1.Text & "last.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If
        If NumericUpDown17.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox1.Text & "male.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox1.Text & "male.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown18.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox2.Text & "female.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox2.Text & "female.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown18.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox2.Text & "last.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox2.Text & "last.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown18.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox2.Text & "male.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox2.Text & "male.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If


        If errorcheck > 0.5 Then
            booler = True
        Else
            booler = False

        End If

        Return booler
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TopDisp.ListBox1.Items.Clear()
        TopDisp.Show()
        TopDisp.CheckBox2.Checked = True

        Me.Hide()

    End Sub

    Public Sub addnewcity()

        Dim pather As String = My.Computer.FileSystem.CurrentDirectory


        Dim connectionString As String =
        "Provider=Microsoft.Jet.OLEDB.4.0;" &
        "Data Source=" & TextBox14.Text & ";"
        Using con As New OleDbConnection(connectionString)
            con.Open()
            Using cmd As New OleDbCommand()
                cmd.Connection = con
                '            ListBox8.Items.Add(Name & ", " & job & ", " & gender & ", " & pacing & ", " & tone & ", " & speed & ", " & accent & ", " & race &
                '                  ", " & motive & ", " & characteristic1 & ",  " & charic2 & ", " & age & ", " & religion & ", " & political & " ")

                '3 is speachimped, called accent -spped
                '4 is accent - accent
                '2 tone - tone
                '1 pace - pacing
                cmd.CommandText = "CREATE TABLE " & TextBox8.Text & " ( " &
            "ID Counter, " &
            "name MEMO, " &
            "age Long, " &
            "gender MEMO, " &
            "race MEMO, " &
            "job MEMO, " &
            "motive MEMO, " &
            "characteristic1 MEMO, " &
            "characteristic2 MEMO, " &
            "accent MEMO, " &
            "SpeechImpediment MEMO, " &
            "pacing MEMO, " &
            "tone MEMO, " &
            "religion MEMO, " &
            "political MEMO, " &
            "Notes MEMO )"


                ' "CREATE TABLE " & TextBox8.Text & " (ID COUNTER, Name LONGTEXT job LONGTEXT gender LONGTEXT pacing LONGTEXT )" &
                '      ("tone LONGTEXT speed LONGTEXT, accent LONGTEXT, race LONGTEXT, motive LONGTEXT, characteristic1 LONGTEXT, charic2 LONGTEXT, age LONGTEXT, ") &
                '     ("religion LONGTEXT, political LONGTEXT")
                Try
                    cmd.ExecuteNonQuery()
                    MsgBox("Table created.")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
            con.Close()
        End Using
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub HowToToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AboutToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MessageBox.Show("Lazy DM is a way to quickly generate a mass number of NPCs for any setting. The real power comes from the dependency’s lists included in the download. The program pulls from those Txt files to generate random pools of names, characteristics, motivations, and jobs. The more you add in each file, the more random your NPCs will be. We included what we felt was a good starting point, but we encourage you to add or subtract from the lists to make them fully your own. Is your setting futuristic? Replace the medieval jobs with high tech future jobs. 

This was simply a passion project between two friends. We're not professional programmers, only one of us even knows how half of this program works. But we worked hard, and we hope you enjoy it.
", "About")




    End Sub

    Private Sub HowToToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles HowToToolStripMenuItem.Click
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
End Class
