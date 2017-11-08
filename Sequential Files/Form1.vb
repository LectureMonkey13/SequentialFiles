Option Explicit On

Imports System.IO 'imports classes for file system
Imports System.Drawing.Printing ' Imports classes for printing

Public Class Form1

    'Declare global variables
    Private PrintPageSettings As New PageSettings
    Private StringToPrint As String

    'Handles file . open option from main menu
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click

        Dim AllText As String = "", LineOfText As String = ""
        Dim FileNumber As Integer = FreeFile()

        OpenFileDialog1.Filter = "TextFiles (*.txt) |*.txt"
        OpenFileDialog1.ShowDialog() 'display open dialog box
        If OpenFileDialog1.FileName <> "" Then
            Try 'open file and trap any errors using arror handler
                FileOpen(FileNumber, OpenFileDialog1.FileName, OpenMode.Input)
                Do Until EOF(1) 'read lines from file
                    LineOfText = LineInput(FileNumber)
                    'add each line to the all text variable
                    AllText &= LineOfText & vbCrLf

                Loop
                Lbl_status.Text = "File: " & OpenFileDialog1.FileName 'update label
                TxtInput.Text = AllText 'display file
                TxtInput.Enabled = True 'alllow text cursor
                CloseToolStripMenuItem.Enabled = True 'enable Close Command
                OpenToolStripMenuItem.Enabled = False 'disable Open command
            Catch ex As Exception
                MsgBox("Error opening file", MsgBoxStyle.Critical, "File Error!")
            Finally
                FileClose(FileNumber) 'close file
                FileNumber = FreeFile() ' find a new free file number


            End Try
        End If

    End Sub

    'Handles File > Exit option from main menu
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Are you sure you want to exit?" 'Define message.
        style = MsgBoxStyle.YesNo
        title = "Text Editor" 'Define Title
        'Display message 
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then 'user chose yes
            Application.Exit()

        End If
    End Sub
    'Handles File > Close option from Main men
    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        TxtInput.Text = ""
        Lbl_status.Text = "To load a text file select:File > Open"
        CloseToolStripMenuItem.Enabled = False 'disable close command
        OpenToolStripMenuItem.Enabled = True 'enable open command
    End Sub

    'Handles File > Insert Date option from main menu
    Private Sub InsertDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertDateToolStripMenuItem.Click
        TxtInput.Text = My.Computer.Clock.LocalTime & vbCrLf & TxtInput.Text
        TxtInput.Select(1, 0) 'remove selection
    End Sub
    'Handles File > Save As option from main menu
    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim FileNumber As Integer = FreeFile()
        SaveFileDialog1.Filter = "TextFile (*.txt) |*txt"
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then
            FileOpen(FileNumber, SaveFileDialog1.FileName, OpenMode.Output)
            PrintLine(FileNumber, TxtInput.Text) 'copy text to disk
            FileClose(FileNumber)
            FileNumber = FreeFile()
        End If
    End Sub
    'Handles File > print page setup option from main menu
    Private Sub PageSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageSetupToolStripMenuItem.Click
        Try
            'load page settings and display page setup dialog box
            PageSetupDialog1.PageSettings = PrintPageSettings
            PageSetupDialog1.ShowDialog()
        Catch ex As Exception
            'display error message
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    'Handles File > Print > Print Preview option from main menu
    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Try
            'specify current page settings
            PrintDocument1.DefaultPageSettings = PrintPageSettings
            'specify document for print preview dialog box and show
            StringToPrint = TxtInput.Text
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            'display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'Handles printing of document
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim PrintFont As New Font("Arial", 10)
        Dim NumChars As Integer
        Dim NumLines As Integer
        Dim StringForPage As String
        Dim StrFormat As New StringFormat
        'based on page setup,define drawable rectangle on page
        Dim RectDraw As New RectangleF(e.MarginBounds.Left, e.MarginBounds.Width, e.MarginBounds.Top, e.MarginBounds.Height)
        'define area to determine how much text can fit on the page
        'make height one line shroter to ensure line doesnt clip
        Dim SizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont.GetHeight(e.Graphics))
        'When drawing long strings, break between words
        StrFormat.Trimming = StringTrimming.Word
        'Compute how many characters and lines can fit based on sizeMeasure
        e.Graphics.MeasureString(StringToPrint, PrintFont, SizeMeasure, StrFormat, NumChars, NumLines)
        'Compute string that will fit on page
        StringForPage = StringToPrint.Substring(0, NumChars)
        'Print String on Current Page
        e.Graphics.DrawString(StringForPage, PrintFont, Brushes.Black, RectDraw, StrFormat)
        'if there is more text, indicate there are more pages

        If NumChars < StringToPrint.Length Then
            'subtract text from string that has been printed 
            StringToPrint = StringToPrint.Substring(NumChars)
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            'all text has been printed, so restore string
            StringToPrint = TxtInput.Text
        End If
    End Sub
    'Initialise display when form is loaded
    Private Sub TextEditor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Lbl_status.Text = " To load a text file select:File > Open"
        Me.Text = "File Browser"
        CloseToolStripMenuItem.Enabled = False 'disable Close file command
    End Sub

End Class
