Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Math
Imports Microsoft.VisualBasic.FileIO.TextFieldParser
Imports System.Threading

Public Class Form1
    Dim myStreamReaderL1 As System.IO.StreamReader
    Dim myStream As System.IO.StreamWriter
    Dim myStr As String
    '------- Headers as received from Inventor ---------------------------
    Dim header_str() As String = {"ITEM", "QTY", "ARTICLE NO.", "DESCRIPTION", "LENGTH", "STANDARD",
    "MATERIAL", "CERT.", "MASS", "REF.DWG / COMMENT", "TYPE"}
    Public Const aantal_kolommen = 17
    Public builder As New StringBuilder

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim j As Integer
        DataGridView1.ColumnCount = aantal_kolommen
        DataGridView2.ColumnCount = aantal_kolommen
        DataGridView1.AutoSizeColumnsMode = CType(DataGridViewAutoSizeColumnMode.AllCells, DataGridViewAutoSizeColumnsMode)
        DataGridView2.AutoSizeColumnsMode = CType(DataGridViewAutoSizeColumnMode.AllCells, DataGridViewAutoSizeColumnsMode)

        TextBox11.Text = "The proper column headers from Inventor are:" & vbCrLf
        For j = 0 To header_str.Length - 1
            TextBox11.Text &= header_str(j) & vbCrLf
        Next
    End Sub
    'Button Read source file
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBox1.Clear()
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = TextBox1.Text
        OpenFileDialog1.FileName = TextBox1.Text
        OpenFileDialog1.ShowDialog()

        TextBox1.Text = OpenFileDialog1.FileName
        TextBox7.Clear()

        '==============  Read from file into the dataview grid ==============
        Read_from_file()
        Check_header()
    End Sub

    Private Sub Read_from_file()
        Dim coll, row_no As Integer

        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

        If File.Exists(TextBox1.Text) = True Then
            Try
                ProgressBar1.Value = 100
                row_no = -1
                For Each row As String In File.ReadAllLines(TextBox1.Text)
                    '-----------------------------------------------------
                    If ProgressBar1.Value > ProgressBar1.Minimum Then
                        ProgressBar1.Value -= 1
                    Else
                        ProgressBar1.Value = ProgressBar1.Maximum
                    End If
                    '-----------------------------------------------------

                    TextBox7.AppendText(row.ToString)

                    DataGridView1.Rows.Add()
                    row_no += 1
                    coll = 0
                    For Each column As String In row.Split(New String() {";"}, StringSplitOptions.None)
                        DataGridView1.Rows.Item(row_no).Cells(coll).Value = column.ToString
                        coll += 1
                    Next
                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                '---------------- now convert-----------------
            End Try
            Convert()
            TabControl1.SelectedIndex = 1
            TextBox7.Clear()
        Else
            MessageBox.Show("Tough shit, file does not exist ..")
        End If
    End Sub

    'Move data from grid 1 to grid 2 and convert
    Private Sub Convert()
        Dim s1 As String = ""
        Dim length As Double
        Dim Position_nr_counter As Integer = 1

        For row = 0 To DataGridView1.RowCount - 2
            DataGridView2.Rows.Add()
            '-----------------------------------------------------
            If ProgressBar1.Value < ProgressBar1.Maximum Then
                ProgressBar1.Value += 1
            Else
                ProgressBar1.Value = ProgressBar1.Minimum
            End If
            '-------------------------------------------------------
            For column = 0 To aantal_kolommen - 1

                '---- preventing problems------
                If IsNothing(DataGridView1.Rows.Item(row).Cells(column).Value) Then
                    DataGridView1.Rows.Item(row).Cells(column).Value = "0"
                End If

                s1 = CType(DataGridView1.Rows.Item(row).Cells(column).Value, String)   'Remove alle quote sign (inch sign)
                s1 = s1.Replace(ControlChars.Quote, "")

                '------------ ITEM----------------
                If (row = 0) Then   'Only first line
                    '---------------- Replace VTK title with Trimergo titles
                    s1 = s1.Replace("ITEM", "user_position_nr")         '0
                    s1 = s1.Replace("QTY", "amount")                    '1
                    s1 = s1.Replace("ARTICLE NO.", "article_number")    '2
                    s1 = s1.Replace("DESCRIPTION", "art_descr")         '3
                    s1 = s1.Replace("LENGTH", "length")                 '4
                    s1 = s1.Replace("STANDARD", "normalization")        '5
                    s1 = s1.Replace("MATERIAL", "quality")              '6
                    s1 = s1.Replace("CERT.", "userfield1")              '7
                    s1 = s1.Replace("MASS", "userfield2")               '8
                    s1 = s1.Replace("REF.DWG / COMMENT", "dwg_descr")   '9
                    s1 = s1.Replace("TYPE", "material_type")            '10
                End If


                '----------- Article_number------------------
                If (column = 2 And s1.Length() = 0) Then
                    TextBox7.Text &= "Row " & row.ToString & " Article number missing" & vbCrLf
                End If

                '----------- Description----------------------
                If (column = 3) Then
                    s1 = s1.Replace(Chr(34), " ")   'Replace the inch sign with space
                    If s1.Contains("PLATE") And s1.Contains("mm") Then
                        DataGridView2.Rows(row).DefaultCellStyle.BackColor = Color.Red
                    End If
                End If

                '----------- length must be zero or bigger----------------
                '----------- Length ----------------------

                If (column = 4 And row > 0) Then    'Ínteger only
                    s1 = s1.Replace(".0", " ")
                    s1 = s1.Replace(".1", " ")
                    s1 = s1.Replace(".2", " ")
                    s1 = s1.Replace(".3", " ")
                    s1 = s1.Replace(".4", " ")
                    s1 = s1.Replace(".5", " ")
                    s1 = s1.Replace(".6", " ")
                    s1 = s1.Replace(".7", " ")
                    s1 = s1.Replace(".8", " ")
                    s1 = s1.Replace(".9", " ")
                End If

                If (column = 4 And row > 0) Then    'Ínteger only
                    s1 = s1.Replace(",0", " ")
                    s1 = s1.Replace(",1", " ")
                    s1 = s1.Replace(",2", " ")
                    s1 = s1.Replace(",3", " ")
                    s1 = s1.Replace(",4", " ")
                    s1 = s1.Replace(",5", " ")
                    s1 = s1.Replace(",6", " ")
                    s1 = s1.Replace(",7", " ")
                    s1 = s1.Replace(",8", " ")
                    s1 = s1.Replace(",9", " ")
                End If

                If (column = 4 And Not Double.TryParse(s1, length) And row > 0) Then
                    s1 = "0"
                End If

                '----------- Normalized ------------------
                If (column = 5 And s1.Length() = 0) Then
                    s1 = " -- "
                End If

                '----------- Material/Quality ------------------
                If (column = 6 And s1.Length() = 0) Then
                    s1 = " --- "
                End If
                '----------- Certificate ------------------
                If (column = 7 And s1.Length() = 0) Then
                    s1 = " ---- "
                End If

                '----------- Weight-Userfield2 ------------------
                If (column = 8 And s1.Length() = 0) Then
                    s1 = "0"
                Else
                    s1 = s1.Replace("kg", " ")  'Strip kg
                    s1 = s1.Replace(", ", ".")   'swap komma for dot
                End If

                '----------- Drawing number ------------------
                If (column = 9 And row > 0) Then
                    s1 = TextBox4.Text
                End If

                '----------- Material ------------------
                If (column = 10 And row > 0) Then
                    s1 = s1.Replace("ART", "materiaal")
                    s1 = s1.Replace("TXT", "vrij")
                    s1 = s1.Replace("GA", "stuklijst")
                End If
                If (column = 10 And s1.Length() = 0) Then
                    TextBox7.Text &= "Row " & row.ToString & " ART-TXT-GA is missing" & vbCrLf
                End If

                '----------- Engineer -----------
                If (column = 11) Then
                    If (row = 0) Then
                        s1 = "engineer"
                    Else
                        s1 = TextBox2.Text
                    End If
                End If
                '------------ drwg number---------------
                If (column = 12) Then
                    If (row = 0) Then
                        s1 = "dwg_number"
                    Else
                        s1 = TextBox6.Text
                    End If
                End If
                '------------drwg version-----------------
                If (column = 13) Then
                    If (row = 0) Then
                        s1 = "dwg_version"
                    Else
                        s1 = TextBox10.Text
                    End If
                End If

                '----------------- Unit is "stuks" of "M1"-----------
                If (column = 14) Then
                    If (row = 0) Then
                        s1 = "Unit"
                    Else
                        s1 = TextBox8.Text
                        length = 0
                        Double.TryParse(CType(DataGridView2.Rows.Item(row).Cells(4).Value, String), length)
                        If (length > 0) Then
                            s1 = TextBox3.Text
                        End If
                    End If
                End If
                '-----------------Add column ----------------------
                If (column = 15) Then
                    If (row = 0) Then
                        s1 = "position_nr"
                    Else
                        s1 = Position_nr_counter.ToString
                        Position_nr_counter += 1
                    End If
                End If

                '-----------------Request Caroline generated column ----------------------
                '-----------------Must be the last column---------------------------------
                If (column = 16) Then
                    If (row = 0) Then
                        s1 = "generated"
                    Else
                        s1 = "N"
                        If (DataGridView2.Rows.Item(row).Cells(10).Value Is "vrij") Then
                            s1 = "Y"
                        End If

                    End If
                End If

                '----------------------------------------------------
                DataGridView2.Rows.Item(row).Cells(column).Value = s1
            Next
        Next
        ProgressBar1.Value = 0
    End Sub
    'Save grid2 to file
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button4.Click
        Dim dest_file As String = TextBox5.Text
        Dim str As String = ""
        'Dim builder As New StringBuilder

        SaveFileDialog1.Title = "Please Select a File name"
        SaveFileDialog1.InitialDirectory = OpenFileDialog1.InitialDirectory
        SaveFileDialog1.FileName = "VTK-Trimergo_conversion"
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.DefaultExt = "csv"
        SaveFileDialog1.ShowDialog()

        TextBox5.Text = SaveFileDialog1.FileName

        For row = 0 To DataGridView2.RowCount - 2
            If Not (DataGridView2.Rows(row).DefaultCellStyle.BackColor = Color.Red And CheckBox1.Checked) Then
                save_row(row)
            End If
        Next
        str = builder.ToString

        '--------- remove now ascii----------------------
        str = Regex.Replace(str, “[^\u0000-\u007F]”, String.Empty)

        '--------- now write to file-----------------------
        My.Computer.FileSystem.WriteAllText(TextBox5.Text, str, False, System.Text.Encoding.ASCII)
    End Sub
    Private Sub save_row(row As Integer)
        Dim strq As String

        '-------------- first column 12-------------
        strq = CType(DataGridView2.Rows.Item(row).Cells(12).Value, String)
        builder.Append(strq)

        For col = 0 To aantal_kolommen - 1
            '---- preventing problems------
            If IsNothing(DataGridView2.Rows.Item(row).Cells(col).Value) Then
                DataGridView2.Rows.Item(row).Cells(col).Value = " "
            End If

            If col <> 12 Then
                strq = CType(DataGridView2.Rows.Item(row).Cells(col).Value, String)
                builder.Append(";").Append(strq)
            End If
        Next
        builder.AppendLine()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        Size = New Size(360, 530)
    End Sub
    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        Me.Size = New Size(1100, 600)
    End Sub
    Private Sub Check_header()
        Dim col As Integer
        Dim error_flag As Boolean = False
        For col = 0 To header_str.Length - 1
            If CBool(String.Compare(header_str(col), CType(DataGridView1.Rows.Item(0).Cells(col).Value, String))) Then
                error_flag = True
            End If
        Next

        If error_flag = True Then
            MessageBox.Show("Export file from Inventor has a header name problem !!" &
          vbCrLf & vbCrLf & "----------------- See the info Tab --------------------")
        End If
    End Sub
End Class
