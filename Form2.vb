Imports System.Drawing.Printing
Imports System.Data.OleDb

Public Class Form2
    Dim time As Integer
    Dim conString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Database11.accdb"
    Dim pdoc As PdfiumViewer.PdfDocument

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        time = time + 1
        If time = 1 Then
            Label7.ForeColor = Color.Blue
        ElseIf time = 2 Then
            Label7.ForeColor = Color.Chocolate
            time = 0
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Database11DataSet.Sheet1' table. You can move, or remove it, as needed.
        Me.Sheet1TableAdapter.Fill(Me.Database11DataSet.Sheet1)
        LoadComboBox()
        Label11.Text = Date.Now.ToString("dd.MMM.yyyy")
        RadioButton1.Checked = True
        Timer2.Interval = 500
        Timer3.Interval = 200
        Timer2.Enabled = True
        Timer3.Enabled = True
        Timer4.Start()
        DateTimePicker1.CustomFormat = "dd-MM-yyyy"
        DateTimePicker2.CustomFormat = "dd-MM-yyyy"
    End Sub
    Private Sub LoadComboBox()
        LoadComboBoxReference()
        LoadComboBoxBankName()
        LoadComboBoxApplicant()
        LoadComboBoxCategory()
    End Sub

    Private Sub LoadComboBoxReference()
        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select distinct REFERENCE from sheet1 order by REFERENCE"
        Dim cmd As New OleDbCommand(query, con)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        While reader.Read
            ComboBox2.Items.Add(reader.GetString(0))
        End While
        con.Close()
    End Sub

    Private Sub LoadComboBoxBankName()
        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select distinct BANK_NAME from sheet1 order by BANK_NAME"
        Dim cmd As New OleDbCommand(query, con)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        While reader.Read
            ComboBox1.Items.Add(reader.GetString(0))
        End While
        con.Close()
    End Sub

    Private Sub LoadComboBoxApplicant()
        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select distinct APPLICANT from sheet1 order by APPLICANT"
        Dim cmd As New OleDbCommand(query, con)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        While reader.Read
            ComboBox3.Items.Add(reader.GetString(0))
        End While
        con.Close()
    End Sub

    Private Sub LoadComboBoxCategory()
        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select distinct CATEGORY from sheet1 order by CATEGORY"
        Dim cmd As New OleDbCommand(query, con)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        While reader.Read
            ComboBox4.Items.Add(reader.GetString(0))
        End While
        con.Close()
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            Dim FILE_NAME As String = Me.Books.CurrentRow.Cells(6).Value
            If System.IO.File.Exists(FILE_NAME) = True Then
                Process.Start(FILE_NAME)
            Else
                MsgBox("File Does Not Exist")
            End If
        Catch ex As Exception
            MsgBox("File Does Not Exist")
        End Try
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox1.Text = "" Then
            MsgBox("NOT SELECT BANK NAME..")
        Else
            PrintPreviewDialog1.ShowDialog()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage_1(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 20, FontStyle.Bold)
        e.Graphics.DrawString(TextBox1.Text, font, Brushes.Blue, 30, 100)
        e.Graphics.DrawString(TextBox3.Text, font, Brushes.Black, 30, 130)
        e.Graphics.DrawString(TextBox5.Text, font, Brushes.Chocolate, 30, 160)
        e.Graphics.DrawString(TextBox4.Text, font, Brushes.Black, 30, 190)
        e.Graphics.DrawString(TextBox2.Text, font, Brushes.Black, 30, 220)
        'e.Graphics.DrawString(TextBox6.Text, font, Brushes.Black, 30, 250)
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If MsgBox("BACK TO NORMAL OPINION?", MsgBoxStyle.Question Or MsgBoxStyle.ApplicationModal Or MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Me.Close()
        End If
    End Sub

    Private Sub FillByToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Sheet1BindingSource.MoveNext()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Sheet1BindingSource.MoveLast()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Sheet1BindingSource.MoveFirst()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Sheet1BindingSource.MovePrevious()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Form4.Show()
    End Sub

    Private Sub Books_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Books.SelectionChanged
        If Me.Books.CurrentRow Is Nothing Then
            MessageBox.Show("No Data")
            PdfViewer1.Visible = False
        Else
            Dim pdf = Me.Books.CurrentRow.Cells(6).Value
            'AxAcroPDF1.LoadFile(pdf)
            If pdf IsNot DBNull.Value Then
                'PdfViewer1 = New PdfiumViewer.PdfViewer()

                Try
                    pdoc = PdfiumViewer.PdfDocument.Load(pdf)
                    PdfViewer1.Document = pdoc
                    lblMsg.Text = ""
                    PdfViewer1.Visible = True
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                    PdfViewer1.Visible = False
                    lblMsg.Text = ex.Message
                End Try
            End If
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Panel1.Visible = False
        Panel2.Visible = True
    End Sub

    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim REFERENCE As String
        Dim BANK_NAME As String
        Dim APPLICANT As String
        Dim CATEGORY As String
        Dim PROPERTY1 As String
        If ComboBox2.SelectedIndex > -1 Then
            REFERENCE = ComboBox2.SelectedItem.ToString()
        Else
            REFERENCE = "%"
        End If

        If ComboBox1.SelectedIndex > -1 Then
            BANK_NAME = ComboBox1.SelectedItem.ToString()
        Else
            BANK_NAME = "%"
        End If

        If ComboBox3.SelectedIndex > -1 Then
            APPLICANT = ComboBox3.SelectedItem.ToString()
        Else
            APPLICANT = "%"
        End If

        If ComboBox4.SelectedIndex > -1 Then
            CATEGORY = ComboBox4.SelectedItem.ToString()
        Else
            CATEGORY = "%"
        End If

        If TextBox7.Text.Trim.Trim.Length > 0 Then
            PROPERTY1 = "%" + TextBox7.Text + "%"
        Else
            PROPERTY1 = "%"
        End If

        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select BANK_NAME,FORMAT(OPINION, 'DD-MM-YYYY') AS OPINION,REFERENCE,APPLICANT,CATEGORY,PROPERTY,PDF from sheet1 where BANK_NAME LIKE @BANK_NAME and REFERENCE LIKE @REFERENCE and APPLICANT LIKE @APPLICANT and CATEGORY LIKE @CATEGORY and PROPERTY LIKE @PROPERTY order by BANK_NAME"
        Dim cmd As New OleDbCommand(query, con)
        cmd.Parameters.AddWithValue("@BANK_NAME", BANK_NAME)
        cmd.Parameters.AddWithValue("@REFERENCE", REFERENCE)
        cmd.Parameters.AddWithValue("@APPLICANT", APPLICANT)
        cmd.Parameters.AddWithValue("@CATEGORY", CATEGORY)
        cmd.Parameters.AddWithValue("@PROPERTY", PROPERTY1)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Dim dt As DataTable = New DataTable()
        dt.Load(reader)
        Sheet1BindingSource.DataSource = dt
        con.Close()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim fromDate As String
        Dim toDate As String
        fromDate = Format(DateTimePicker1.Value, "MM/dd/yyyy")
        toDate = Format(DateTimePicker2.Value, "MM/dd/yyyy")

        Dim con As New OleDbConnection(conString)
        con.Open()
        Dim query As String = "select BANK_NAME,FORMAT(OPINION, 'DD-MM-YYYY') AS OPINION,REFERENCE,APPLICANT,CATEGORY,PROPERTY,PDF from sheet1 where OPINION between @fromDate and @toDate"
        Dim cmd As New OleDbCommand(query, con)
        cmd.Parameters.AddWithValue("@fromDate", fromDate)
        cmd.Parameters.AddWithValue("@toDate", toDate)
        Dim reader As OleDbDataReader
        reader = cmd.ExecuteReader()
        Dim dt As DataTable = New DataTable()
        dt.Load(reader)
        Sheet1BindingSource.DataSource = dt
        con.Close()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value
    End Sub
End Class
