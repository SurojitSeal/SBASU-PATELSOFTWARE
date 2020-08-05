Imports System.IO
Imports System.Data.OleDb

Public Class Form4
    Dim time As Integer
    Dim conString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Database11.accdb"

    Public Sub ImportData(ByVal sourcePath As String, ByVal destinationPath As String)
        Dim sourceDirectoryInfo As New DirectoryInfo(sourcePath)

        ' If the destination folder don't exist then create it
        If Not Directory.Exists(destinationPath) Then
            Directory.CreateDirectory(destinationPath)
        End If

        Dim fileSystemInfo As FileSystemInfo
        For Each fileSystemInfo In sourceDirectoryInfo.GetFileSystemInfos
            Dim destinationFileName As String =
                Path.Combine(destinationPath, fileSystemInfo.Name)

            ' Now check whether its a file or a folder and take action accordingly
            If TypeOf fileSystemInfo Is FileInfo Then
                Dim PDF As String = "n/a"
                If fileSystemInfo.Extension.ToLower = ".pdf" Then
                    File.Copy(fileSystemInfo.FullName, destinationFileName, True)
                End If
                If fileSystemInfo.Extension.ToLower = ".txt" Then
                    Dim filePath = fileSystemInfo.FullName
                    Using sr As New StreamReader(filePath)
                        Dim line As String
                        Dim arr() As String
                        Dim REFERENCE As String = "n/a"
                        Dim OPINION As String = "n/a"
                        Dim BANK_NAME As String = "n/a"
                        Dim APPLICANT As String = "n/a"
                        Dim CATEGORY As String = "n/a"
                        Dim PROPERTYDESC As String = "n/a"


                        line = sr.ReadLine()
                        While (line <> Nothing)
                            arr = line.Split("|")
                            If arr(0) = "r" Or arr(0) = "R" Then
                                REFERENCE = arr(1)

                            ElseIf arr(0) = "o" Or arr(0) = "O" Then
                                OPINION = arr(1)

                            ElseIf arr(0) = "b" Or arr(0) = "B" Then
                                BANK_NAME = arr(1)

                            ElseIf arr(0) = "a" Or arr(0) = "A" Then
                                APPLICANT = arr(1)

                            ElseIf arr(0) = "c" Or arr(0) = "C" Then
                                CATEGORY = arr(1)

                            ElseIf arr(0) = "p" Or arr(0) = "P" Then
                                PROPERTYDESC = arr(1)
                            End If
                            line = sr.ReadLine()
                        End While
                        arr = destinationFileName.Split(".")
                        PDF = arr(0) + ".pdf"
                        Try
                            Dim con As New OleDbConnection(conString)
                            con.Open()
                            Dim query As String = "insert into Sheet1(BANK_NAME,OPINION,REFERENCE,APPLICANT,CATEGORY,PROPERTY,PDF) values(@BANK_NAME,FORMAT(@OPINION,'DD-MM-YYYY'),@REFERENCE,@APPLICANT,@CATEGORY,@PROPERTY,@PDF)"
                            Dim cmd As New OleDbCommand(query, con)
                            cmd.Parameters.AddWithValue("@BANK_NAME", BANK_NAME)
                            cmd.Parameters.AddWithValue("@OPINION", OPINION)
                            cmd.Parameters.AddWithValue("@REFERENCE", REFERENCE)
                            cmd.Parameters.AddWithValue("@APPLICANT", APPLICANT)
                            cmd.Parameters.AddWithValue("@CATEGORY", CATEGORY)
                            cmd.Parameters.AddWithValue("@PROPERTY", PROPERTYDESC)
                            cmd.Parameters.AddWithValue("@PDF", PDF)
                            cmd.ExecuteNonQuery()
                            con.Close()
                        Catch ex As Exception

                        End Try
                    End Using
                End If
            Else
                ' Recursively call the mothod to copy all the nested folders
                ImportData(fileSystemInfo.FullName, destinationFileName)
            End If
        Next
    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        If TextBox1.Text.Trim <> "" Then
            PictureBox1.Visible = True
            BackgroundWorker1.RunWorkerAsync()
        Else
            MsgBox("PLEASE PORVIDE PATH")
        End If
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
        PictureBox1.Visible = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim srcDir As String = TextBox1.Text
        Dim destDir As String = Application.StartupPath() + "\PDFs\POST OPINION\"
        ImportData(srcDir, destDir)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        time = time + 1
        If time = 1 Then
            Label7.ForeColor = Color.Blue
        ElseIf time = 2 Then
            Label7.ForeColor = Color.Chocolate
            time = 0
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        PictureBox1.ImageLocation = "completed.png"
        MsgBox("IMPORT COMPLETE")
        Application.Restart()
    End Sub
End Class