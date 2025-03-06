Imports PuppeteerSharp
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports System.Drawing

Public Class MainForm
    Inherits Form

    ' Declaration of controls
    Private txtUrl As TextBox
    Private numSlides As NumericUpDown
    Private txtLog As TextBox
    Private btnStart As Button

    ' Constructor for the form
    Public Sub New()
        ' Set form properties
        Me.Text = "PDF Generator"
        Me.Size = New System.Drawing.Size(800, 600) ' Use System.Drawing.Size
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Create and configure controls
        txtUrl = New TextBox() With {
            .Location = New System.Drawing.Point(10, 10), ' Use System.Drawing.Point
            .Width = 760
        }
        Me.Controls.Add(txtUrl)

        numSlides = New NumericUpDown() With {
            .Location = New System.Drawing.Point(10, 40), ' Use System.Drawing.Point
            .Width = 100,
            .Minimum = 1,
            .Maximum = 1000
        }
        Me.Controls.Add(numSlides)

        txtLog = New TextBox() With {
            .Location = New System.Drawing.Point(10, 70), ' Use System.Drawing.Point
            .Width = 760,
            .Height = 400,
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .ReadOnly = True
        }
        Me.Controls.Add(txtLog)

        btnStart = New Button() With {
            .Location = New System.Drawing.Point(10, 480), ' Use System.Drawing.Point
            .Width = 760,
            .Height = 40,
            .Text = "Start Processing"
        }
        AddHandler btnStart.Click, AddressOf btnStart_Click
        Me.Controls.Add(btnStart)
    End Sub

    ' Start button code
    Private Async Sub btnStart_Click(sender As Object, e As EventArgs)
        ' Get the link from the TextBox
        Dim url As String = txtUrl.Text
        If String.IsNullOrEmpty(url) Then
            MessageBox.Show("Please, enter a valid URL.")
            Return
        End If

        ' Get the number of slides from NumericUpDown
        Dim totalSlides As Integer = Convert.ToInt32(numSlides.Value)
        If totalSlides <= 0 Then
            MessageBox.Show("Please, select at least 1 slide.")
            Return
        End If

        ' PDF saving folder
        Dim pdfDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "SlidesPDF")

        ' Create the saving folder if it doesn't exist
        If Not Directory.Exists(pdfDirectory) Then
            Directory.CreateDirectory(pdfDirectory)
        End If

        ' Add start log
        LogMessage("Process started...")

        ' Create a BrowserFetcher object to handle Chromium download
        Dim browserFetcher As New BrowserFetcher()
        Await browserFetcher.DownloadAsync()

        ' Launch the headless browser (without GUI)
        Dim browser = Await Puppeteer.LaunchAsync(New LaunchOptions With {
            .Headless = True
        })

        ' Create a new page and navigate to the URL
        Dim page = Await browser.NewPageAsync()
        Await page.GoToAsync(url)

        ' Wait for the page to fully load
        Await page.WaitForSelectorAsync("body")

        ' Path to save the separate PDF files
        Dim slideFiles As New List(Of String)()

        ' Loop through the slides and generate a PDF for each one
        For slideIndex As Integer = 1 To totalSlides ' Use the selected number of slides
            ' Path to save the slide PDF
            Dim slidePdf As String = Path.Combine(pdfDirectory, $"Slide_{slideIndex}.pdf")
            slideFiles.Add(slidePdf)

            ' Generate the current slide
            Await page.PdfAsync(slidePdf, New PdfOptions With {
                .Landscape = True ' Set landscape orientation
            })

            ' Go to the next slide (simulate pressing the right arrow)
            Await page.Keyboard.PressAsync("ArrowRight")
            ' Wait a moment for the slide to load
            Await Task.Delay(1000)

            ' Add to the log
            LogMessage($"Processed slide {slideIndex}/{totalSlides}")
        Next

        ' Check the generated PDFs before combining them
        For Each slideFile As String In slideFiles
            If Not File.Exists(slideFile) Then
                LogMessage($"The file {slideFile} does not exist or is empty!")
                Return
            End If
        Next

        ' Combine the generated PDFs into a single file
        Dim combinedPdf As String = Path.Combine(pdfDirectory, "CombinedSlides.pdf")
        If CombinePdfs(slideFiles, combinedPdf) Then
            LogMessage("Combined PDF of slides generated successfully: " & combinedPdf)
            MessageBox.Show("Combined PDF of slides generated successfully!")
        Else
            LogMessage("Error while combining the PDFs.")
            MessageBox.Show("Error while combining the PDFs.")
        End If

        ' Close the browser
        Await browser.CloseAsync()
    End Sub

    ' Function to add log messages
    Private Sub LogMessage(message As String)
        ' Display messages in the log TextBox (txtLog)
        txtLog.AppendText(message & Environment.NewLine)
    End Sub

    ' Function to combine PDFs into one using PdfSharp
    Private Function CombinePdfs(slideFiles As List(Of String), outputPdf As String) As Boolean
        Try
            ' Create a new empty PDF document
            Dim outputDocument As New PdfDocument()

            ' Add each separate PDF file
            For Each slideFile In slideFiles
                ' Open the separate PDF file
                Using inputDocument As PdfDocument = PdfReader.Open(slideFile, PdfDocumentOpenMode.Import)
                    ' Copy each page from the separate PDF
                    For pageIndex As Integer = 0 To inputDocument.PageCount - 1
                        ' Add the page to the output document
                        outputDocument.AddPage(inputDocument.Pages(pageIndex))
                    Next
                End Using
            Next

            ' Save the combined document
            outputDocument.Save(outputPdf)
            Return True
        Catch ex As Exception
            ' Error handling during combining
            LogMessage("Error during PDF combination: " & ex.Message)
            Return False
        End Try
    End Function
End Class
