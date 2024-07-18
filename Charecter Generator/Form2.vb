
Public Class Form2
        Public b As Decimal ' You can access these from Form1 by Form2.b = ?
        Public d As Decimal

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Randomize()

        ' Dim checker As Boolean = depcheck(checker)
        Form1.ListBox7.Items.Clear()
        Form1.ListBox8.Items.Clear()
        Form1.ListBox22.Items.Clear()
        Form1.ListBox23.Items.Clear()

        Dim checker As Boolean = depcheck(checker)
        If checker = False Then
            Form1.Show()
            depload()
            loadtoform4()
            If (NumericUpDown1.Value + NumericUpDown2.Value + NumericUpDown4.Value + NumericUpDown3.Value + NumericUpDown8.Value _
                + NumericUpDown7.Value + NumericUpDown6.Value + NumericUpDown5.Value + NumericUpDown16.Value + NumericUpDown15.Value _
                + NumericUpDown14.Value + NumericUpDown13.Value + NumericUpDown12.Value + NumericUpDown11.Value + NumericUpDown10.Value _
                + NumericUpDown9.Value + NumericUpDown17.Value + NumericUpDown18.Value + NumericUpDown19.Value + NumericUpDown20.Value _
                + NumericUpDown21.Value + NumericUpDown25.Value + NumericUpDown26.Value + NumericUpDown27.Value + NumericUpDown28.Value _
                + NumericUpDown29.Value + NumericUpDown30.Value + NumericUpDown31.Value + NumericUpDown32.Value + NumericUpDown33.Value _
                + NumericUpDown34.Value) = 100 Then
                ' If CheckBox1.Checked = True And (NumericUpDown1.Value + NumericUpDown2.Value) = 100 Or CheckBox1.Checked = False Then
                Dim tester9000 As Boolean
                If (CheckBox1.Checked = False) Or (((CheckBox1.Checked = True) And NumericUpDown35.Value + NumericUpDown36.Value) = 100) Then
                    tester9000 = False
                    'MsgBox((NumericUpDown1.Value + NumericUpDown2.Value))
                Else
                    tester9000 = True
                End If
                'MsgBox(tester9000)



                If tester9000 = False Then


                    If NumericUpDown24.Value > 0.5 Then

                        savemysoulfuckmesideways()

                        'Send the custom shit to form1
                        If NumericUpDown17.Value > 0.5 Then
                            Form1.TextBox9.Text = TextBox1.Text
                        Else
                            Form1.TextBox9.Text = ""
                        End If
                        If NumericUpDown18.Value > 0.5 Then
                            Form1.TextBox10.Text = TextBox2.Text
                        Else
                            Form1.TextBox10.Text = ""
                        End If
                        If NumericUpDown19.Value > 0.5 Then
                            Form1.TextBox11.Text = TextBox3.Text
                        Else
                            Form1.TextBox11.Text = ""
                        End If
                        If NumericUpDown20.Value > 0.5 Then
                            Form1.TextBox12.Text = TextBox4.Text
                        Else
                            Form1.TextBox12.Text = ""
                        End If
                        If NumericUpDown21.Value > 0.5 Then
                            Form1.TextBox13.Text = TextBox5.Text
                        Else
                            Form1.TextBox13.Text = ""
                        End If



                        'Gender add form1

                        If CheckBox1.Checked = True Then
                            Form1.ListBox24.Items.Clear()

                            Dim this As Integer = 0
                            'Dim thisq As Integer = 0
                            Do Until this = NumericUpDown35.Value
                                Form1.ListBox24.Items.Add("Male")
                                this = (this + 1)
                                'MsgBox(this)
                            Loop
                            'MsgBox("Added male")
                            'MsgBox(this)
                            this = 0
                            Do Until this = NumericUpDown36.Value
                                Form1.ListBox24.Items.Add("Female")
                                this = (this + 1)
                            Loop
                        End If



                        'This is the add for jobs, small size


                        If RadioButton1.Checked = True Then

                            Dim this As Integer = 0
                            Do Until this = NumericUpDown24.Value
                                'Form1.ListBox9.Items.Add("Human")

                                If CheckBox2.Checked = True Then
                                    If this < ListBox7.Items.Count Then
                                        Dim small As String = ListBox7.Items(CInt(Math.Floor((ListBox7.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                        'MsgBox(this & "checkingfrom form2")
                                    Else
                                        Dim small As String = ListBox6.Items(CInt(Math.Floor((ListBox6.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                    End If

                                Else
                                    Dim small As String = Form1.ListBox10.Items(CInt(Math.Floor((Form1.ListBox10.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox7.Items.Add(small)
                                End If
                                this = (this + 1)
                            Loop
                            'Do the law first
                            Dim this1 As Integer = 0
                            Do Until this1 = NumericUpDown23.Value
                                'Form1.ListBox9.Items.Add("Human")

                                If CheckBox2.Checked = True Then
                                    If this1 < ListBox5.Items.Count Then
                                        Dim small As String = ListBox5.Items(CInt(Math.Floor((ListBox5.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small)
                                    Else
                                        Dim small2 As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small2)
                                    End If

                                Else
                                    Dim small2 As String = Form1.ListBox14.Items(CInt(Math.Floor((Form1.ListBox14.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox23.Items.Add(small2)
                                End If
                                this1 = (this1 + 1)
                            Loop
                            'Do the ruler second
                            Dim this2 As Integer = 0
                            Do Until this2 = NumericUpDown22.Value
                                'Form1.ListBox9.Items.Add("Human")

                                If CheckBox2.Checked = True Then
                                    If this2 < ListBox3.Items.Count Then

                                        Dim small As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small)
                                    Else
                                        Dim small3 As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small3)

                                    End If
                                Else
                                    Dim small3 As String = Form1.ListBox18.Items(CInt(Math.Floor((Form1.ListBox18.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox22.Items.Add(small3)
                                End If
                                this2 = (this2 + 1)
                            Loop




                            'end small add







                        ElseIf RadioButton2.Checked = True Then 'Med city add

                            Dim this As Integer = 0
                            Do Until this = NumericUpDown24.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this = (this + 1)
                                If CheckBox2.Checked = True Then
                                    If this < ListBox7.Items.Count Then
                                        Dim small As String = ListBox7.Items(CInt(Math.Floor((ListBox7.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                        'MsgBox(this & "checkingfrom form2")
                                    Else
                                        Dim small As String = ListBox6.Items(CInt(Math.Floor((ListBox6.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                    End If

                                Else
                                    Dim small As String = Form1.ListBox11.Items(CInt(Math.Floor((Form1.ListBox11.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox7.Items.Add(small)

                                End If

                            Loop

                            'Do the law first

                            Dim this1 As Integer = 0
                            Do Until this1 = NumericUpDown23.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this1 = (this1 + 1)
                                If CheckBox2.Checked = True Then
                                    If this1 < ListBox5.Items.Count Then
                                        Dim small As String = ListBox5.Items(CInt(Math.Floor((ListBox5.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small)
                                    Else
                                        Dim small2 As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small2)
                                    End If

                                Else
                                    Dim small2 As String = Form1.ListBox15.Items(CInt(Math.Floor((Form1.ListBox15.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox23.Items.Add(small2)

                                End If

                            Loop


                            'Do the ruler second
                            Dim this2 As Integer = 0
                            Do Until this2 = NumericUpDown22.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this2 = (this2 + 1)
                                If CheckBox2.Checked = True Then
                                    If this2 < ListBox3.Items.Count Then

                                        Dim small As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small)
                                    Else
                                        Dim small3 As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small3)

                                    End If
                                Else
                                    Dim small3 As String = Form1.ListBox19.Items(CInt(Math.Floor((Form1.ListBox19.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox22.Items.Add(small3)
                                End If


                            Loop


                        ElseIf RadioButton3.Checked = True Then 'Lrg city add

                            Dim this As Integer = 0
                            Do Until this = NumericUpDown24.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this = (this + 1)
                                If CheckBox2.Checked = True Then
                                    If this < ListBox7.Items.Count Then
                                        Dim small As String = ListBox7.Items(CInt(Math.Floor((ListBox7.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                        'MsgBox(this & "checkingfrom form2")
                                    Else
                                        Dim small As String = ListBox6.Items(CInt(Math.Floor((ListBox6.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                    End If

                                Else
                                    Dim small As String = Form1.ListBox12.Items(CInt(Math.Floor((Form1.ListBox12.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox7.Items.Add(small)
                                End If



                            Loop

                            'Do the law first

                            Dim this1 As Integer = 0
                            Do Until this1 = NumericUpDown23.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this1 = (this1 + 1)
                                If CheckBox2.Checked = True Then
                                    If this1 < ListBox5.Items.Count Then
                                        Dim small As String = ListBox5.Items(CInt(Math.Floor((ListBox5.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small)
                                    Else
                                        Dim small2 As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small2)
                                    End If

                                Else
                                    Dim small2 As String = Form1.ListBox16.Items(CInt(Math.Floor((Form1.ListBox16.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox23.Items.Add(small2)
                                End If


                            Loop


                            'Do the ruler second
                            Dim this2 As Integer = 0
                            Do Until this2 = NumericUpDown22.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this2 = (this2 + 1)
                                If CheckBox2.Checked = True Then
                                    If this2 < ListBox3.Items.Count Then

                                        Dim small As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small)
                                    Else
                                        Dim small3 As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small3)

                                    End If
                                Else
                                    Dim small3 As String = Form1.ListBox20.Items(CInt(Math.Floor((Form1.ListBox20.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox22.Items.Add(small3)
                                End If


                            Loop

                        ElseIf RadioButton4.Checked = True Then 'Med city add

                            Dim this As Integer = 0
                            Do Until this = NumericUpDown24.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this = (this + 1)
                                If CheckBox2.Checked = True Then
                                    If this < ListBox7.Items.Count Then
                                        Dim small As String = ListBox7.Items(CInt(Math.Floor((ListBox7.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                        'MsgBox(this & "checkingfrom form2")
                                    Else
                                        Dim small As String = ListBox6.Items(CInt(Math.Floor((ListBox6.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox7.Items.Add(small)
                                    End If

                                Else
                                    Dim small As String = Form1.ListBox13.Items(CInt(Math.Floor((Form1.ListBox13.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox7.Items.Add(small)
                                End If


                            Loop

                            'Do the law first

                            Dim this1 As Integer = 0
                            Do Until this1 = NumericUpDown23.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this1 = (this1 + 1)
                                If CheckBox2.Checked = True Then
                                    If this1 < ListBox5.Items.Count Then
                                        Dim small As String = ListBox5.Items(CInt(Math.Floor((ListBox5.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small)
                                    Else
                                        Dim small2 As String = ListBox4.Items(CInt(Math.Floor((ListBox4.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox23.Items.Add(small2)
                                    End If

                                Else
                                    Dim small2 As String = Form1.ListBox17.Items(CInt(Math.Floor((Form1.ListBox17.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox23.Items.Add(small2)
                                End If


                            Loop


                            'Do the ruler second
                            Dim this2 As Integer = 0
                            Do Until this2 = NumericUpDown22.Value
                                'Form1.ListBox9.Items.Add("Human")
                                this2 = (this2 + 1)
                                If CheckBox2.Checked = True Then
                                    If this2 < ListBox3.Items.Count Then

                                        Dim small As String = ListBox3.Items(CInt(Math.Floor((ListBox3.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small)
                                    Else
                                        Dim small3 As String = ListBox2.Items(CInt(Math.Floor((ListBox2.Items.Count - 0) * Rnd())) + 0)
                                        Form1.ListBox22.Items.Add(small3)

                                    End If
                                Else
                                    Dim small3 As String = Form1.ListBox21.Items(CInt(Math.Floor((Form1.ListBox21.Items.Count - 0) * Rnd())) + 0)
                                    Form1.ListBox22.Items.Add(small3)

                                End If

                            Loop


                        End If




                        'Form1.ListBox8.Items.Add("item")
                        'Form1.ListBox9.Items.Add("item")
                        'abel1.Text = (d + b).ToString
                        'MsgBox()
                        Me.Close()
                    Else
                        MsgBox("There are no commoners to be generated. Failing.")
                    End If

                    'End If
                Else



                    If CheckBox1.Checked = True Then



                        Dim thiser As Integer = NumericUpDown35.Value + NumericUpDown36.Value


                        If thiser > 100 Then
                            MsgBox("The gender distribution does not equal 100. Please fix. You are over by " & (thiser - 100))
                        Else
                            MsgBox("The gender distribution does not equal 100. Please fix. You are under by " & (100 - thiser))
                        End If

                    End If

                    'MsgBox("The gender distribution does not equal 100. Please fix.")
                    'MsgBox(tester9000)
                    'End If
                End If

            Else
                Dim thiser As Integer = (NumericUpDown1.Value + NumericUpDown2.Value + NumericUpDown4.Value + NumericUpDown3.Value + NumericUpDown8.Value _
                + NumericUpDown7.Value + NumericUpDown6.Value + NumericUpDown5.Value + NumericUpDown16.Value + NumericUpDown15.Value _
                + NumericUpDown14.Value + NumericUpDown13.Value + NumericUpDown12.Value + NumericUpDown11.Value + NumericUpDown10.Value _
                + NumericUpDown9.Value + NumericUpDown17.Value + NumericUpDown18.Value + NumericUpDown19.Value + NumericUpDown20.Value _
                + NumericUpDown21.Value + NumericUpDown25.Value + NumericUpDown26.Value + NumericUpDown27.Value + NumericUpDown28.Value _
                + NumericUpDown29.Value + NumericUpDown30.Value + NumericUpDown31.Value + NumericUpDown32.Value + NumericUpDown33.Value _
                + NumericUpDown34.Value)
                If thiser > 100 Then
                    MsgBox("Race percentages do not add up to exactly 100. Please resolve, you are over by " & (thiser - 100))

                Else
                    MsgBox("Race percentages do not add up to exactly 100. Please resolve, you are under by " & (100 - thiser))
                End If

            End If
        Else
            MsgBox("Missing Dependancies. Aborting.")
        End If

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
    Public Sub savemysoulfuckmesideways()
        'Begin the thousand lines. FML.



        If NumericUpDown1.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown1.Value
                Form1.ListBox9.Items.Add("Human")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown2.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown2.Value
                Form1.ListBox9.Items.Add("Aasimar")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown4.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown4.Value
                Form1.ListBox9.Items.Add("Warforged")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown3.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown3.Value
                Form1.ListBox9.Items.Add("Dwarf")
                this = (this + 1)
            Loop
        End If


        If NumericUpDown8.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown8.Value
                Form1.ListBox9.Items.Add("Yuanti-Pureblood")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown7.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown7.Value
                Form1.ListBox9.Items.Add("Triton")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown6.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown6.Value
                Form1.ListBox9.Items.Add("Goliath")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown5.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown5.Value
                Form1.ListBox9.Items.Add("Tabaxi")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown16.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown16.Value
                Form1.ListBox9.Items.Add("Dragonborn")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown15.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown15.Value
                Form1.ListBox9.Items.Add("Half-Elf")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown14.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown14.Value
                Form1.ListBox9.Items.Add("LizardFolk")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown13.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown13.Value
                Form1.ListBox9.Items.Add("Gnome")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown12.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown12.Value
                Form1.ListBox9.Items.Add("Genasi")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown11.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown11.Value
                Form1.ListBox9.Items.Add("Aarakocra")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown10.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown10.Value
                Form1.ListBox9.Items.Add("Elf")
                this = (this + 1)
            Loop
        End If
        If NumericUpDown9.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown9.Value
                Form1.ListBox9.Items.Add("Tiefling")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown34.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown34.Value
                Form1.ListBox9.Items.Add("firbolg")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown33.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown33.Value
                Form1.ListBox9.Items.Add("Gith")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown32.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown32.Value
                Form1.ListBox9.Items.Add("Goblin")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown31.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown31.Value
                Form1.ListBox9.Items.Add("Half-orc")
                this = (this + 1)
            Loop
        End If


        If NumericUpDown30.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown30.Value
                Form1.ListBox9.Items.Add("Hobgoblin")
                this = (this + 1)
            Loop
        End If


        If NumericUpDown29.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown29.Value
                Form1.ListBox9.Items.Add("Kalashtar")
                this = (this + 1)
            Loop
        End If


        If NumericUpDown28.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown28.Value
                Form1.ListBox9.Items.Add("Kobold")
                this = (this + 1)
            Loop
        End If


        If NumericUpDown27.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown27.Value
                Form1.ListBox9.Items.Add("Loxodon")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown26.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown26.Value
                Form1.ListBox9.Items.Add("ORC")
                this = (this + 1)
            Loop
        End If

        If NumericUpDown25.Value > 0.5 Then
            Dim this As Integer = 0
            Do Until this = NumericUpDown25.Value
                Form1.ListBox9.Items.Add("Tortle")
                this = (this + 1)
            Loop
        End If







        If NumericUpDown17.Value > 0.5 Then
            If TextBox1.TextLength > 1 Then
                Dim this As Integer = 0
                Do Until this = NumericUpDown17.Value
                    Form1.ListBox9.Items.Add(TextBox1.Text)
                    this = (this + 1)
                Loop
            Else
                MsgBox("You fucked up and didnt put in a name for the race. Try again. Not going to add nothing as a race.")
            End If

        End If

        If NumericUpDown18.Value > 0.5 Then
            If TextBox2.TextLength > 1 Then
                Dim this As Integer = 0
                Do Until this = NumericUpDown18.Value
                    Form1.ListBox9.Items.Add(TextBox2.Text)
                    this = (this + 1)
                Loop
            Else
                MsgBox("You fucked up and didnt put in a name for the race. Try again. Not going to add nothing as a race.")
            End If

        End If

        If NumericUpDown19.Value > 0.5 Then
            If TextBox3.TextLength > 1 Then
                Dim this As Integer = 0
                Do Until this = NumericUpDown19.Value
                    Form1.ListBox9.Items.Add(TextBox3.Text)
                    this = (this + 1)
                Loop
            Else
                MsgBox("You fucked up and didnt put in a name for the race. Try again. Not going to add nothing as a race.")
            End If

        End If

        If NumericUpDown20.Value > 0.5 Then
            If TextBox4.TextLength > 1 Then
                Dim this As Integer = 0
                Do Until this = NumericUpDown20.Value
                    Form1.ListBox9.Items.Add(TextBox4.Text)
                    this = (this + 1)
                Loop
            Else
                MsgBox("You fucked up and didnt put in a name for the race. Try again. Not going to add nothing as a race.")
            End If

        End If



        If NumericUpDown21.Value > 0.5 Then
            If TextBox5.TextLength > 1 Then
                Dim this As Integer = 0
                Do Until this = NumericUpDown21.Value
                    Form1.ListBox9.Items.Add(TextBox5.Text)
                    this = (this + 1)
                Loop
            Else
                MsgBox("You fucked up and didnt put in a name for the race. Try again. Not going to add nothing as a race.")
            End If

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim checker As Boolean = depcheck(checker)
        depload()

        'MsgBox(checker)

    End Sub
    Public Sub loadtoform4()



        Dim pather As String = My.Computer.FileSystem.CurrentDirectory
        Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\LawMotive.txt")
        Form4.ListBox1.Items.AddRange(lines4)
        'MsgBox(pather & "\Dependancies\LrgLawJobs.txt")
        Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\LawCharacteristics1.txt")
        Form4.ListBox2.Items.AddRange(lines5)

        Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\LawCharacteristics2.txt")
        Form4.ListBox3.Items.AddRange(lines6)


        Dim lines7() As String = IO.File.ReadAllLines(pather & "\Dependancies\LeaderMotive.txt")
        Form4.ListBox4.Items.AddRange(lines7)
        'MsgBox(pather & "\Dependancies\LrgLawJobs.txt")
        Dim lines8() As String = IO.File.ReadAllLines(pather & "\Dependancies\LeaderCharacteristics1.txt")
        Form4.ListBox5.Items.AddRange(lines8)

        Dim lines9() As String = IO.File.ReadAllLines(pather & "\Dependancies\LeaderCharacteristics2.txt")
        Form4.ListBox6.Items.AddRange(lines9)



        Dim lines10() As String = IO.File.ReadAllLines(pather & "\Dependancies\CommonMotive.txt")
        Form4.ListBox7.Items.AddRange(lines10)
        'MsgBox(pather & "\Dependancies\LrgLawJobs.txt")
        Dim lines11() As String = IO.File.ReadAllLines(pather & "\Dependancies\CommonCharacteristics1.txt")
        Form4.ListBox8.Items.AddRange(lines11)

        Dim lines12() As String = IO.File.ReadAllLines(pather & "\Dependancies\CommonCharacteristics2.txt")
        Form4.ListBox9.Items.AddRange(lines12)



    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            NumericUpDown22.Value = 1
            NumericUpDown23.Value = 1
            NumericUpDown24.Value = 9
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            NumericUpDown22.Value = 1
            NumericUpDown23.Value = 2
            NumericUpDown24.Value = 15
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            NumericUpDown22.Value = 2
            NumericUpDown23.Value = 4
            NumericUpDown24.Value = 40
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            NumericUpDown22.Value = 4
            NumericUpDown23.Value = 8
            NumericUpDown24.Value = 100
        End If
    End Sub
    Function depcheck(booler As Boolean) As Boolean
        Dim errorcheck As Integer = 0
        Dim pather As String = My.Computer.FileSystem.CurrentDirectory

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


        If NumericUpDown19.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox3.Text & "female.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox3.Text & "female.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown19.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox3.Text & "last.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox3.Text & "last.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If
        If NumericUpDown19.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox3.Text & "male.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox3.Text & "male.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown20.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox4.Text & "female.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox4.Text & "female.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown20.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox4.Text & "last.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox4.Text & "last.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If
        If NumericUpDown20.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox4.Text & "male.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox4.Text & "male.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If


        If NumericUpDown21.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox5.Text & "female.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox5.Text & "female.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If NumericUpDown21.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox5.Text & "last.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox5.Text & "last.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If
        If NumericUpDown21.Value > 0.5 Then
            If My.Computer.FileSystem.FileExists(pather & "\Dependancies\names\" & TextBox5.Text & "male.txt") Then
                ' MsgBox("File found.")
            Else
                errorcheck = (errorcheck + 1)
                MsgBox("Missing \Dependancies\names\" & TextBox5.Text & "male.txt  Please add your custom race dependancies in the location shown.")
            End If
        End If

        If errorcheck > 0.5 Then
            booler = True
        Else
            booler = False

        End If

        Return booler
    End Function


    Function depload()

        Dim pather As String = My.Computer.FileSystem.CurrentDirectory
        Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\names\Aarakocrafemale.txt")
        Form5.ListBox1.Items.AddRange(lines4)
        Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\names\Aarakocralast.txt")
        Form5.ListBox2.Items.AddRange(lines5)
        Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\names\Aarakocramale.txt")
        Form5.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Aasimarfemale.txt")
        Form6.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Aasimarlast.txt")
        Form6.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Aasimarmale.txt")
        Form6.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dragonbornfemale.txt")
        Form7.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dragonbornlast.txt")
        Form7.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dragonbornmale.txt")
        Form7.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing




        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dwarffemale.txt")
        Form8.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dwarflast.txt")
        Form8.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Dwarfmale.txt")
        Form8.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\ELFfemale.txt")
        Form9.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\ELFlast.txt")
        Form9.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\ELFmale.txt")
        Form9.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing



        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Genasifemale.txt")
        Form10.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Genasilast.txt")
        Form10.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Genasimale.txt")
        Form10.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Gnomefemale.txt")
        Form11.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Gnomelast.txt")
        Form11.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Gnomemale.txt")
        Form11.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goliathfemale.txt")
        Form12.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goliathlast.txt")
        Form12.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goliathmale.txt")
        Form12.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-Elffemale.txt")
        Form13.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-Elflast.txt")
        Form13.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-Elfmale.txt")
        Form13.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing



        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\humanfemale.txt")
        Form14.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\humanlast.txt")
        Form14.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\humanmale.txt")
        Form14.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\LizardFolkfemale.txt")
        Form15.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\LizardFolklast.txt")
        Form15.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\LizardFolkmale.txt")
        Form15.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing




        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tabaxifemale.txt")
        Form16.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tabaxilast.txt")
        Form16.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tabaximale.txt")
        Form16.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tieflingfemale.txt")
        Form17.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tieflinglast.txt")
        Form17.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tieflingmale.txt")
        Form17.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tritonfemale.txt")
        Form18.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tritonlast.txt")
        Form18.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tritonmale.txt")
        Form18.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing



        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Warforgedfemale.txt")
        Form19.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Warforgedlast.txt")
        Form19.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Warforgedmale.txt")
        Form19.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Yuantifemale.txt")
        Form20.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Yuantilast.txt")
        Form20.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Yuantimale.txt")
        Form20.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\firbolgfemale.txt")
        Form35.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\firbolglast.txt")
        Form35.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\firbolgmale.txt")
        Form35.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Githfemale.txt")
        Form34.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Githlast.txt")
        Form34.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Githmale.txt")
        Form34.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing



        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goblinfemale.txt")
        Form33.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goblinlast.txt")
        Form33.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Goblinmale.txt")
        Form33.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing



        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-orcfemale.txt")
        Form32.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-orclast.txt")
        Form32.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Half-orcmale.txt")
        Form32.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing


        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Hobgoblinfemale.txt")
        Form31.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Hobgoblinlast.txt")
        Form31.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Hobgoblinmale.txt")
        Form31.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Kalashtarfemale.txt")
        Form30.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Kalashtarlast.txt")
        Form30.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Kalashtarmale.txt")
        Form30.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Koboldfemale.txt")
        Form29.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Koboldlast.txt")
        Form29.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Koboldmale.txt")
        Form29.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Loxodonfemale.txt")
        Form28.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Loxodonlast.txt")
        Form28.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Loxodonmale.txt")
        Form28.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\ORCfemale.txt")
        Form27.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\ORClast.txt")
        Form27.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\ORCmale.txt")
        Form27.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing

        lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tortlefemale.txt")
        Form26.ListBox1.Items.AddRange(lines4)
        lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tortlelast.txt")
        Form26.ListBox2.Items.AddRange(lines5)
        lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\Tortlemale.txt")
        Form26.ListBox3.Items.AddRange(lines6)
        lines4 = Nothing
        lines5 = Nothing
        lines6 = Nothing






        'Custom dependacy checks and loads
        '(pather & "\Dependancies\names\" & TextBox1.Text & ".txt")


        If NumericUpDown17.Value > 0.5 Then
            lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox1.Text & "female.txt")
            Form21.ListBox1.Items.AddRange(lines4)
            lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox1.Text & "last.txt")
            Form21.ListBox2.Items.AddRange(lines5)
            lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox1.Text & "male.txt")
            Form21.ListBox3.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If


        If NumericUpDown18.Value > 0.5 Then
            lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox2.Text & "female.txt")
            Form22.ListBox1.Items.AddRange(lines4)
            lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox2.Text & "last.txt")
            Form22.ListBox2.Items.AddRange(lines5)
            lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox2.Text & "male.txt")
            Form22.ListBox3.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If

        If NumericUpDown19.Value > 0.5 Then
            lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox3.Text & "female.txt")
            Form23.ListBox1.Items.AddRange(lines4)
            lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox3.Text & "last.txt")
            Form23.ListBox2.Items.AddRange(lines5)
            lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox3.Text & "male.txt")
            Form23.ListBox3.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If



        If NumericUpDown20.Value > 0.5 Then
            lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox4.Text & "female.txt")
            Form24.ListBox1.Items.AddRange(lines4)
            lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox4.Text & "last.txt")
            Form24.ListBox2.Items.AddRange(lines5)
            lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox4.Text & "male.txt")
            Form24.ListBox3.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If

        If NumericUpDown21.Value > 0.5 Then
            lines4 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox5.Text & "female.txt")
            Form25.ListBox1.Items.AddRange(lines4)
            lines5 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox5.Text & "last.txt")
            Form25.ListBox2.Items.AddRange(lines5)
            lines6 = IO.File.ReadAllLines(pather & "\Dependancies\names\" & TextBox5.Text & "male.txt")
            Form25.ListBox3.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If


    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form4.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            NumericUpDown35.Visible = True
            NumericUpDown36.Visible = True
            Label33.Visible = True
            Label34.Visible = True

        Else

            NumericUpDown35.Visible = False
            NumericUpDown36.Visible = False
            Label33.Visible = False
            Label34.Visible = False
            NumericUpDown35.Value = 50
            NumericUpDown36.Value = 50


        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then

            Dim errorcheck As Integer = 0
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory


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
            If errorcheck > 1 Then
                MsgBox("Missing dependancies. Failed")
            Else

                'loadepstoform2
                loadepstoform2()




                'Ruler
                ListBox2.Visible = True
                    ListBox3.Visible = True
                    Label35.Visible = True
                    Label36.Visible = True

                    'Lawman

                    ListBox4.Visible = True
                    ListBox5.Visible = True
                    Label37.Visible = True
                    Label38.Visible = True

                    'Commoner 

                    ListBox6.Visible = True
                    ListBox7.Visible = True
                    Label39.Visible = True
                    Label40.Visible = True

                    Button5.Visible = True
                    Button6.Visible = True
                    Button7.Visible = True
                    Button8.Visible = True
                    Button9.Visible = True
                    Button10.Visible = True

                End If





                Else

                ListBox2.Visible = False
                ListBox3.Visible = False
                Label35.Visible = False
                Label36.Visible = False
                'NumericUpDown35.Value = 50
                ' NumericUpDown36.Value = 50
                'Lawman

                ListBox4.Visible = False
                ListBox5.Visible = False
                Label37.Visible = False
                Label38.Visible = False

                'Commoner 

                ListBox6.Visible = False
                ListBox7.Visible = False
                Label39.Visible = False
                Label40.Visible = False

                Button5.Visible = False
                Button6.Visible = False
                Button7.Visible = False
                Button8.Visible = False
                Button9.Visible = False
                Button10.Visible = False

                'load up the deps


            End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If CheckBox2.Checked = True Then
            If ListBox2.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox2.SelectedItems.Count - 1

                Do While i >= 0
                    ListBox3.Items.Add(ListBox2.SelectedItems.Item(i))

                    'ListBox2.Items.Remove(ListBox2.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If



    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If CheckBox2.Checked = True Then
            If ListBox3.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox3.SelectedItems.Count - 1

                Do While i >= 0
                    'ListBox2.Items.Add(ListBox3.SelectedItems.Item(i))

                    ListBox3.Items.Remove(ListBox3.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If CheckBox2.Checked = True Then
            If ListBox4.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox4.SelectedItems.Count - 1

                Do While i >= 0
                    ListBox5.Items.Add(ListBox4.SelectedItems.Item(i))

                    ' ListBox4.Items.Remove(ListBox4.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If CheckBox2.Checked = True Then
            If ListBox5.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox5.SelectedItems.Count - 1

                Do While i >= 0
                    'ListBox4.Items.Add(ListBox5.SelectedItems.Item(i))

                    ListBox5.Items.Remove(ListBox5.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If CheckBox2.Checked = True Then
            If ListBox6.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox6.SelectedItems.Count - 1

                Do While i >= 0
                    ListBox7.Items.Add(ListBox6.SelectedItems.Item(i))

                    ' ListBox6.Items.Remove(ListBox6.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If CheckBox2.Checked = True Then
            If ListBox7.SelectedIndices.Count > 0.5 Then


                Dim i As Integer

                i = ListBox7.SelectedItems.Count - 1

                Do While i >= 0
                    'ListBox6.Items.Add(ListBox7.SelectedItems.Item(i))

                    ListBox7.Items.Remove(ListBox7.SelectedItems.Item(i))

                    i = i - 1

                Loop
            Else
                MsgBox("Please select something to move")
            End If
        End If



    End Sub

    Sub loadepstoform2()


        If RadioButton1.Checked = True Then
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallLeaderJobs.txt")
            ListBox2.Items.AddRange(lines4)
            Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallLawJobs.txt")
            ListBox4.Items.AddRange(lines5)
            Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\SmallCommonerJobs.txt")
            ListBox6.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        ElseIf RadioButton2.Checked = True Then
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedLeaderJobs.txt")
            ListBox2.Items.AddRange(lines4)
            Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedLawJobs.txt")
            ListBox4.Items.AddRange(lines5)
            Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\MedCommonerJobs.txt")
            ListBox6.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing




        ElseIf RadioButton3.Checked = True Then
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgLeaderJobs.txt")
            ListBox2.Items.AddRange(lines4)
            Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgLawJobs.txt")
            ListBox4.Items.AddRange(lines5)
            Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\LrgCommonerJobs.txt")
            ListBox6.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        ElseIf RadioButton4.Checked = True Then
            Dim pather As String = My.Computer.FileSystem.CurrentDirectory
            Dim lines4() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroLeaderJobs.txt")
            ListBox2.Items.AddRange(lines4)
            Dim lines5() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroLawJobs.txt")
            ListBox4.Items.AddRange(lines5)
            Dim lines6() As String = IO.File.ReadAllLines(pather & "\Dependancies\MetroCommonerJobs.txt")
            ListBox6.Items.AddRange(lines6)
            lines4 = Nothing
            lines5 = Nothing
            lines6 = Nothing
        End If



    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

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
End Class