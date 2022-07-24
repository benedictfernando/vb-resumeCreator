Imports System.IO
Imports Newtonsoft.Json
Imports Gehtsoft.PDFFlow.Builder
Imports Gehtsoft.PDFFlow.Models.Enumerations
Imports Gehtsoft.PDFFlow.Utils
Imports HorizontalAlignment = Gehtsoft.PDFFlow.Models.Enumerations.HorizontalAlignment

Public Class Form1
    ' declare an empty person
    Dim person As Individual

    Private Sub load_Click(sender As Object, e As EventArgs) Handles load.Click
        ' note: you can change 'json' value accdg. to pathfile of your own .json file
        Dim json As String = "C:\Users\Benedict\Documents\FERNANDO_BENEDICT.json"

        ' load .json file to global 'person' variable @line 10
        LoadJson(json)

        ' enable and disable buttons
        Me.load.Enabled = False
        Me.generate.Enabled = True
    End Sub

    Public Sub LoadJson(ByVal filename As String)
        ' deserialize JSON directly from .JSON file
        Using json As StreamReader = File.OpenText(filename)
            Dim serializer As JsonSerializer = New JsonSerializer()
            person = CType(serializer.Deserialize(json, GetType(Individual)), Individual)
        End Using
    End Sub

    Private Sub generate_Click(sender As Object, e As EventArgs) Handles generate.Click
        ' setup a savefiledialog
        Dim saveFileDialog1 As SaveFileDialog = New SaveFileDialog()
        saveFileDialog1.InitialDirectory = "C:\"
        saveFileDialog1.Title = "Save Resume"
        saveFileDialog1.DefaultExt = "pdf"
        saveFileDialog1.Filter = "Pdf files (*.pdf)|*.pdf|All files (*.*)|*.*"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Me.generate.Enabled = False
            Dim filename As String = saveFileDialog1.FileName
            createResume(filename)

            ' finalize success flow
            MessageBox.Show("Resume successfully saved!")
            Application.[Exit]()
        End If
    End Sub

    Private Sub createResume(ByVal filename As String)
        ' create a document builder
        Dim [resume] As DocumentBuilder = DocumentBuilder.[New]()
        Dim headerSize = 18

        ' customize paper settings
        Dim resumeBuilder = [resume].AddSection().SetMargins(50).SetSize(PaperSize.Letter).SetOrientation(PageOrientation.Portrait).SetStyleFont(Fonts.Courier(11))

        ' For topmost details
        resumeBuilder.AddParagraph($"{person.firstname} {person.lastname}").SetBold().SetAlignment(HorizontalAlignment.Center).SetFontSize(20)
        resumeBuilder.AddParagraph("Full Stack Web Developer").SetAlignment(HorizontalAlignment.Center).SetMarginBottom(10)

        ' For skills
        resumeBuilder.AddParagraph("Skills:").SetBold().SetMarginTop(18).SetMarginBottom(5).SetFontSize(headerSize)
        For Each skill In person.skills
            resumeBuilder.AddParagraph(skill).SetListBulleted(ListBullet.Bullet, 0)
        Next

        ' For techstacks
        resumeBuilder.AddParagraph("Technology Stacks:").SetBold().SetMarginTop(18).SetMarginBottom(5).SetFontSize(headerSize)
        For Each tech In person.techstacks
            resumeBuilder.AddParagraph(tech).SetListBulleted(ListBullet.Bullet, 0)
        Next

        ' For experiences
        resumeBuilder.AddParagraph("Experiences:").SetBold().SetMarginTop(18).SetMarginBottom(5).SetFontSize(headerSize)
        For Each experience In person.experiences
            Dim exp As Experience = JsonConvert.DeserializeObject(Of Experience)(experience.ToString())
            resumeBuilder.AddParagraph(exp.role).SetMarginBottom(2).SetListBulleted(ListBullet.Dash, 0)
            resumeBuilder.AddParagraph($"{exp.company} | {exp.span}")
        Next

        ' For personal
        resumeBuilder.AddParagraph("Details:").SetBold().SetMarginTop(18).SetMarginBottom(5).SetFontSize(headerSize)
        resumeBuilder.AddParagraph($"Location: {person.location}")
        resumeBuilder.AddParagraph($"Email: {person.email}")
        resumeBuilder.AddParagraph($"Number: {person.number}")

        ' build resume
        [resume].Build(filename)
    End Sub

    Public Class Individual
        Public Property firstname As String
        Public Property lastname As String
        Public Property skills As String()
        Public Property techstacks As String()
        Public Property experiences As Object()
        Public Property location As String
        Public Property email As String
        Public Property number As String
    End Class

    Public Class Experience
        Public Property role As String
        Public Property company As String
        Public Property span As String
    End Class
End Class
