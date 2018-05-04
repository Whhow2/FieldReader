Imports System.IO
Imports Word = Microsoft.Office.Interop.Word

Public Class Form1
    Dim SelectedFiles() As String
    Public Sub InitializeOpenFileDialog()
        Dim readRegValue As String
        readRegValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Source", Nothing)
        If String.IsNullOrEmpty(readRegValue) Then
            Me.OpenFileDialog1.InitialDirectory = "H:\Trans\Export"
        Else
            Me.OpenFileDialog1.InitialDirectory = readRegValue
        End If
        Me.OpenFileDialog1.Filter = "All Files *.* | *.*"
        ' Allow the user to select multiple images.
        Me.OpenFileDialog1.Multiselect = True
        Me.OpenFileDialog1.Title = "Document Browser"
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeOpenFileDialog()
        Label3.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Destination", Nothing)
    End Sub

    Private Sub ListBox1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub ListBox1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            For i = 0 To MyFiles.Length - 1
                ListBox1.Items.Add(MyFiles(i))
            Next
            Dim path As String = System.IO.Path.GetDirectoryName(ListBox1.Items.Item(ListBox1.Items.Count - 1))
            My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Translation Tools\Field Reader Document Creator", True)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Source", path)
            If Not String.IsNullOrEmpty(Label3.Text) Then
                Button4.Enabled = True
            End If
            Button2.Enabled = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dr As DialogResult = Me.OpenFileDialog1.ShowDialog()
        If (dr = System.Windows.Forms.DialogResult.OK) Then
            Button2.Enabled = True
            Button6.Enabled = True
            ' Read the files
            Dim filename As String
            For Each file In OpenFileDialog1.FileNames
                filename = System.IO.Path.GetFileName(file)
                ListBox1.Items.Add(file)
            Next file
            Dim path As String = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileNames(OpenFileDialog1.FileNames.Count - 1))
            Dim regValue As String
            regValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Source", Nothing)
            If String.IsNullOrEmpty(regValue) Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Source", path)
            Else
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Translation Tools\Field Reader Document Creator", True)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Source", path)
            End If
            If Not String.IsNullOrEmpty(Label3.Text) Then
                Button4.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' TODO: Add ability to remove multiple items
        ' For i = 0 To ListBox1.SelectedItems.Count - 1
        ' ListBox1.Items.Remove(ListBox1.SelectedItems(i))
        ' Next
        ListBox1.Items.Remove(ListBox1.SelectedItem)
        If ListBox1.Items.Count = 0 Then
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim initialDirectory As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Destination", Nothing)
        If String.IsNullOrEmpty(initialDirectory) Then
        End If
        Using obj As New OpenFileDialog
            obj.Filter = "foldersOnly|*.none"
            obj.CheckFileExists = False
            obj.CheckPathExists = False
            If String.IsNullOrEmpty(initialDirectory) Then
                obj.InitialDirectory = "C:\temp"
            Else
                obj.InitialDirectory = initialDirectory
            End If
            'obj.CustomPlaces.Add("H:\OIS") ' add your custom location, appears upper left
            'obj.CustomPlaces.Add("H:\Permits") ' add your custom location
            obj.Title = "Select folder - click Open to select"
            obj.FileName = "OpenFldrPath"
            If obj.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim dest As String = System.IO.Path.GetDirectoryName(obj.FileName)
                Label3.Text = dest
                My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Translation Tools\Field Reader Document Creator", True)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Translation Tools\Field Reader Document Creator", "Last Destination", Label3.Text)
                Button4.Enabled = True
            End If
        End Using
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = ListBox1.Items.Count * 2
        ProgressBar1.Value = 0
        Dim files() As String = ListBox1.Items.Cast(Of String).ToArray
        Dim MEPSFilesList As ArrayList = New ArrayList()
        Dim destFile As String
        Dim destMPSFile As String
        Dim destWDFile As String
        Dim destPath As String
        destPath = Label3.Text
        Dim appWord As Word.Application
        appWord = New Word.Application
        Dim wordDoc As Word.Document
		wordDoc = New Word.Document
		Dim wordDoc2Save As Word.Document
		wordDoc2Save = New Word.Document
		Dim amkCmd As String
		Dim amkArg As String
		Dim tempMPSFile As String
		amkCmd = """C:\Program Files (x86)\Watchtower\MEPS Prepress 1.8.4\AUTOMARK.exe"""
        'Dim FindObject As Word.Find = Word.Application
        For Each file In files
			destFile = System.IO.Path.Combine(Label3.Text, System.IO.Path.GetFileName(file).ToString)
			System.IO.File.Copy(file, destFile, True)
			ProgressBar1.Value += 1
			If System.IO.Path.GetExtension(Label3.Text) = ".docx" Then
				destMPSFile = Replace(destFile, ".docx", ".mps")
				wordDoc2Save = appWord.Documents.Open(destFile)
				wordDoc2Save.SaveAs2(FileName:=destMPSFile, FileFormat:=100)
				wordDoc2Save.Close()
				System.IO.File.Delete(destFile)
			ElseIf System.IO.Path.GetExtension(Label3.Text) = ".doc" Then
				destMPSFile = Replace(destFile, ".doc", ".mps")
				wordDoc2Save = appWord.Documents.Open(destFile)
				wordDoc2Save.SaveAs2(FileName:=destMPSFile, FileFormat:=100)
				wordDoc2Save.Close()
				System.IO.File.Delete(destFile)
			Else
				tempMPSFile = destFile & "-temp.mps"
				destMPSFile = destFile & ".mps"
				My.Computer.FileSystem.RenameFile(destFile, System.IO.Path.GetFileName(tempMPSFile).ToString)
				wordDoc2Save = appWord.Documents.Open(tempMPSFile)
				wordDoc2Save.SaveAs2(FileName:=destMPSFile, FileFormat:=100)
				wordDoc2Save.Close()
				System.IO.File.Delete(tempMPSFile)
			End If
			MEPSFilesList.Add(destMPSFile)
			'My.Computer.FileSystem.RenameFile(destFile, System.IO.Path.GetFileName(destMPSFile).ToString)
			Label4.Text = Int(ProgressBar1.Value * 100 / ProgressBar1.Maximum) & "%"
		Next file
		amkArg = """" & destPath & """  G:\SYSTEM\AUTOMARK.UP\FieldReader.NV R I"
        Dim p As Process = Process.Start(amkCmd, amkArg)
        p.WaitForInputIdle()
        p.WaitForExit()
        For Each mepsFile In MEPSFilesList
            wordDoc = appWord.Documents.Open(mepsFile)
            appWord.Selection.WholeStory()
            With appWord.Selection.ParagraphFormat
                .SpaceBefore = 6
                .SpaceBeforeAuto = False
                .SpaceAfter = 6
                .SpaceBeforeAuto = False
            End With
            With appWord.Selection
                .LanguageID = Word.WdLanguageID.wdEnglishUS
                .NoProofing = True
            End With
			'appWord.Run("ProcessQuestions")
			'appWord.Run("ProcessOther")

			'ProcessQuestions
			'Selection.HomeKey Unit:=wdStory
			appWord.Selection.HomeKey(Word.WdUnits.wdStory)
			appWord.ActiveDocument.TrackRevisions = False

			'Find source placeholder in document.
			With appWord.Selection.Find
				.ClearFormatting()
				.Text = "[[???]]"
				'.Text = GetPlaceholderText(Source)
				.Replacement.Text = ""
				.Forward = True
				.Wrap = Word.WdFindWrap.wdFindContinue
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchKashida = False
				.MatchDiacritics = False
				.MatchAlefHamza = False
				.MatchControl = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				.Execute()
			End With

			'Begin processing loop if placeholder found.
			Do While appWord.Selection.Find.Found = True
				'Select source placeholder and cut.
				'ActiveDocument.TrackRevisions = False
				appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
				'ActiveDocument.TrackRevisions = True
				appWord.Selection.MoveDown(Word.WdUnits.wdParagraph, 1, Word.WdMovementType.wdExtend)
				appWord.Selection.Cut()
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)
				appWord.Selection.Find.ClearFormatting()
				appWord.Selection.Find.Replacement.ClearFormatting()

				'Find destination placeholder.
				With appWord.Selection.Find
					.Text = "[[@@@]]"
					'.Text = GetPlaceholderText(Destination)
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()

					'Paste into destination placeholder.
					'ActiveDocument.TrackRevisions = False
					appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
					'ActiveDocument.TrackRevisions = True
					appWord.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)

					'Append optional backspace (if specified).
					'If backspace = True Then
					'ActiveDocument.TrackRevisions = False
					'Selection.TypeBackspace
					'Selection.TypeText Text:=" "
					'ActiveDocument.TrackRevisions = True
					'End If
				End With

				'Jump back to the start of the document.
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)

				'Reset find parameters.
				With appWord.Selection.Find
					.ClearFormatting()
					.Text = "[[???]]"
					'.Text = GetPlaceholderText(Source)
					.Replacement.Text = ""
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()
				End With
			Loop

			'ProcessOther
			'Selection.HomeKey Unit:=wdStory
			appWord.Selection.HomeKey(Word.WdUnits.wdStory)
			appWord.ActiveDocument.TrackRevisions = False

			'Find source placeholder in document.
			With appWord.Selection.Find
				.ClearFormatting()
				.Text = "[[%%%]]"
				'.Text = GetPlaceholderText(Source)
				.Replacement.Text = ""
				.Forward = True
				.Wrap = Word.WdFindWrap.wdFindContinue
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchKashida = False
				.MatchDiacritics = False
				.MatchAlefHamza = False
				.MatchControl = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				.Execute()
			End With

			'Begin processing loop if placeholder found.
			Do While appWord.Selection.Find.Found = True
				'Select source placeholder and cut.
				'Selection.Delete Unit:=wdCharacter, Count:=1
				appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
				'ActiveDocument.TrackRevisions = True
				'Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
				appWord.Selection.MoveDown(Word.WdUnits.wdParagraph, 1, Word.WdMovementType.wdExtend)
				appWord.Selection.Cut()
				'Selection.HomeKey Unit:=wdStory
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)
				appWord.Selection.Find.ClearFormatting()
				appWord.Selection.Find.Replacement.ClearFormatting()

				'Find destination placeholder.
				With appWord.Selection.Find
					.Text = "[[$$$]]"
					'.Text = GetPlaceholderText(Destination)
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()

					'Paste into destination placeholder.
					'ActiveDocument.TrackRevisions = False
					'Selection.Delete Unit:=wdCharacter, Count:=1
					appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
					'ActiveDocument.TrackRevisions = True
					appWord.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)

					'Append optional backspace (if specified).
					'If backspace = True Then
					'ActiveDocument.TrackRevisions = False
					'Selection.TypeBackspace
					'Selection.TypeText Text:=" "
					'ActiveDocument.TrackRevisions = True
					'End If
				End With

				'Jump back to the start of the document.
				'Selection.HomeKey Unit:=wdStory
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)

				'Reset find parameters.
				With appWord.Selection.Find
					.ClearFormatting()
					.Text = "[[%%%]]"
					'.Text = GetPlaceholderText(Source)
					.Replacement.Text = ""
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()
				End With
			Loop

			'Begin processing loop if placeholder found.
			Do While appWord.Selection.Find.Found = True
				'Select source placeholder and cut.
				'ActiveDocument.TrackRevisions = False
				'Selection.Delete Unit:=wdCharacter, Count:=1
				appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
				'ActiveDocument.TrackRevisions = True
				'Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
				appWord.Selection.MoveDown(Word.WdUnits.wdParagraph, 1, Word.WdMovementType.wdExtend)
				appWord.Selection.Cut()
				'Selection.HomeKey Unit:=wdStory
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)
				appWord.Selection.Find.ClearFormatting()
				appWord.Selection.Find.Replacement.ClearFormatting()

				'Find destination placeholder.
				With appWord.Selection.Find
					.Text = "[[@@@]]"
					'.Text = GetPlaceholderText(Destination)
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()

					'Paste into destination placeholder.
					'ActiveDocument.TrackRevisions = False
					'Selection.Delete Unit:=wdCharacter, Count:=1
					appWord.Selection.Delete(Word.WdUnits.wdCharacter, 1)
					'ActiveDocument.TrackRevisions = True
					appWord.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)

					'Append optional backspace (if specified).
					'If backspace = True Then
					'ActiveDocument.TrackRevisions = False
					'Selection.TypeBackspace
					'Selection.TypeText Text:=" "
					'ActiveDocument.TrackRevisions = True
					'End If
				End With

				'Jump back to the start of the document.
				'Selection.HomeKey Unit:=wdStory
				appWord.Selection.HomeKey(Word.WdUnits.wdStory)

				'Reset find parameters.
				With appWord.Selection.Find
					.ClearFormatting()
					.Text = "[[???]]"
					'.Text = GetPlaceholderText(Source)
					.Replacement.Text = ""
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchKashida = False
					.MatchDiacritics = False
					.MatchAlefHamza = False
					.MatchControl = False
					.MatchWildcards = False
					.MatchSoundsLike = False
					.MatchAllWordForms = False
					.Execute()
				End With
			Loop

			wordDoc.Sections(1).Footers(1).PageNumbers.Add(1)
            appWord.ActiveDocument.TrackRevisions = True
            appWord.ActiveDocument.ShowRevisions = True
			destWDFile = Replace(mepsFile, ".mps", "")
			appWord.ActiveDocument.SaveAs2(destWDFile)
            wordDoc.Close()
            wordDoc = Nothing
            System.IO.File.Delete(mepsFile)
            ProgressBar1.Value += 1
            Label4.Text = Int(ProgressBar1.Value * 100 / ProgressBar1.Maximum) & "%"
        Next
        appWord.Quit()
        Process.Start(Label3.Text)
        Label4.Text = "Done!"
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ListBox1.Items.Clear()
        Button2.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class
