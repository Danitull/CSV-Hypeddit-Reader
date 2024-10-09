Imports System.Windows.Forms
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Net
Imports System.Net.Mail
Imports Google.Apis.Gmail.v1
Imports Google.Apis.Auth.OAuth2
Public Class Form1
    Dim dataGridView1 As New DataGridView()
    Dim newDataGridView As New DataGridView()
    Dim WidthBool As New Boolean
    Dim comboBox As New ComboBox
    Dim textBoxMyMail As New TextBox
    Dim textBoxMyPass As New TextBox
    Dim textBoxOggetto As New TextBox
    Dim textBoxTesto As New TextBox
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.Text = "CSVReaderMailSender"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox1.Visible = False
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "CSV files (*.csv)|*.csv"
        openFileDialog1.Title = "Select a CSV file"

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            WidthBool = False
            Dim filePath As String = openFileDialog1.FileName
            ' Create a DataTable to store the CSV data
            Dim dt As New DataTable()
            If dataGridView1 IsNot Nothing Then
                Me.Controls.Remove(dataGridView1)
                dataGridView1.Dispose()
            End If
            dataGridView1 = New DataGridView()
            ' Use TextFieldParser to read the CSV file
            Using parser As New TextFieldParser(filePath)
                parser.Delimiters = New String() {","}
                parser.HasFieldsEnclosedInQuotes = True

                ' Get the column names from the first row
                Dim columns() As String = parser.ReadFields()
                For Each column As String In columns
                    dt.Columns.Add(column)
                Next

                ' Read the rest of the rows
                While Not parser.EndOfData
                    Dim row() As String = parser.ReadFields()
                    dt.Rows.Add(row)
                End While
            End Using

            ' Create a DataGridView to display the data

            dataGridView1.DataSource = dt

            ' Add the DataGridView to the form
            Me.Controls.Add(dataGridView1)
            dataGridView1.Location = New Point(0, 60) ' 60px gap from top
            dataGridView1.Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 60) ' occupa tutto lo spazio disponibile
            dataGridView1.Visible = True
            ListBox1.Visible = False
            ' Set the AutoSizeMode for each column
            For Each column As DataGridViewColumn In dataGridView1.Columns
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

            Dim titleHashSet As New HashSet(Of String)
            Dim titleColumnIndex As Integer = dataGridView1.Columns("Title").Index

            For i As Integer = 0 To dataGridView1.Rows.Count - 2
                Dim title As String = dataGridView1.Rows(i).Cells(titleColumnIndex).Value.ToString()
                titleHashSet.Add(title)
            Next

            ComboBox1.DataSource = titleHashSet.ToList()

            Button1.Text = "Change File"
            If dt.Columns.Contains("Email Address") Then
                Button2.Visible = True
                ComboBox1.Visible = True
                Button3.Visible = True
            Else
                Button2.Visible = False
            End If
        End If
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ListBox1.Height = Me.ClientSize.Height - 60
        dataGridView1.Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 60)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "Leggi E-mail singolarmente" Then
            ComboBox1.Visible = False
            Button3.Visible = False
            Button2.Text = "Visualizza tabella precedente"
            Dim emailHashSet As New HashSet(Of String)
            Dim emailColumnIndex As Integer = dataGridView1.Columns("Email Address").Index

            For i As Integer = 0 To dataGridView1.Rows.Count - 2
                Dim email As String = dataGridView1.Rows(i).Cells(emailColumnIndex).Value.ToString().ToLowerInvariant()
                emailHashSet.Add(email)
            Next
            ListBox1.DataSource = emailHashSet.OrderBy(Function(x) x).ToList()
            dataGridView1.Visible = False

            ListBox1.Location = New Point(0, 60)
            If WidthBool = False Then
                ListBox1.Width = ListBox1.GetItemRectangle(0).Width + 100
            End If
            ListBox1.Height = Me.ClientSize.Height - 60
            ListBox1.Visible = True
            WidthBool = True
        Else
            ListBox1.Visible = False
            dataGridView1.Visible = True
            ComboBox1.Visible = True
            Button3.Visible = True
            Button2.Text = "Leggi E-mail singolarmente"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If Button3.Text = "Stampa utenti della traccia selezionata" Then
            ' Disattiva la visualizzazione della GridView1
            Button4.Visible = True
            Button1.Visible = False
            dataGridView1.Visible = False
            Button3.Text = "Visualizza la tabella precedente"
            ComboBox1.Visible = False
            Button2.Visible = False
            newDataGridView.Visible = True
            ' Rimuovi la nuova DataGridView se esiste
            newDataGridView = Me.Controls.OfType(Of DataGridView)().FirstOrDefault(Function(x) x IsNot dataGridView1)
            If newDataGridView IsNot Nothing Then
                Me.Controls.Remove(newDataGridView)
                newDataGridView.Dispose()
            End If

            ' Crea una nuova GridView
            newDataGridView = New DataGridView()
            newDataGridView.Location = New Point(0, 60)
            newDataGridView.Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 60)
            Me.Controls.Add(newDataGridView)

            ' Filtra i dati della GridView1 in base al testo selezionato nella ComboBox1
            Dim filteredRows As New List(Of DataGridViewRow)
            Dim titleColumnIndex As Integer = dataGridView1.Columns("Title").Index

            For i As Integer = 0 To dataGridView1.Rows.Count - 2
                If dataGridView1.Rows(i).Cells(titleColumnIndex).Value.ToString() = ComboBox1.SelectedItem.ToString() Then
                    filteredRows.Add(dataGridView1.Rows(i))
                End If
            Next

            ' Crea un nuovo DataTable per la nuova GridView
            Dim newDataTable As New DataTable()
            For Each column As DataGridViewColumn In dataGridView1.Columns
                newDataTable.Columns.Add(column.HeaderText)
            Next

            ' Aggiunge le righe filtrate al nuovo DataTable
            For Each row As DataGridViewRow In filteredRows
                Dim newRow As DataRow = newDataTable.NewRow()
                For i As Integer = 0 To row.Cells.Count - 1
                    newRow(i) = row.Cells(i).Value
                Next
                newDataTable.Rows.Add(newRow)
            Next

            ' Assegna il nuovo DataTable alla nuova GridView
            newDataGridView.DataSource = newDataTable
            For Each column As DataGridViewColumn In newDataGridView.Columns
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next
        Else
            Button4.Visible = False
            Button2.Visible = True
            dataGridView1.Visible = True
            Button1.Visible = True
            ComboBox1.Visible = True
            newDataGridView.Visible = False
            Button3.Text = "Stampa utenti della traccia selezionata"
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button4.Text = "Invia Mail pubblicitaria" Then
            ' Nascondi tutti gli elementi del form
            For Each control As Control In Me.Controls
                If control IsNot Button4 AndAlso control IsNot Button5 AndAlso control IsNot ComboBox1 Then
                    control.Visible = False
                End If
            Next
            Button5.Visible = True
            Button4.Visible = True
            Button4.Text = "Torna indietro"

            ' Crea una nuova ComboBox con le stesse voci di ComboBox1
            comboBox = New ComboBox()
            comboBox.Location = New Point(10, 10)
            comboBox.Size = New Size(550, 20)
            For Each item As String In ComboBox1.Items
                If item <> ComboBox1.SelectedItem.ToString() Then
                    comboBox.Items.Add(item)
                End If
            Next
            Me.Controls.Add(comboBox)
            If comboBox.Items.Count > 0 Then
                comboBox.SelectedIndex = 0
            End If
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList

            ' Crea una nuova Label con scritto "MyMail"
            Dim labelMyMail As New Label()
            labelMyMail.Location = New Point(10, 40)
            labelMyMail.Size = New Size(50, 20)
            labelMyMail.Text = "E-Mail"
            Me.Controls.Add(labelMyMail)

            ' Crea una nuova TextBox per inserire l'indirizzo email
            textBoxMyMail = New TextBox()
            textBoxMyMail.Location = New Point(70, 40)
            textBoxMyMail.Size = New Size(350, 20)
            Me.Controls.Add(textBoxMyMail)

            ' Crea una nuova Label con scritto "MyMail"
            Dim labelMyPass As New Label()
            labelMyPass.Location = New Point(10, 70)
            labelMyPass.Size = New Size(55, 20)
            labelMyPass.Text = "Password"
            Me.Controls.Add(labelMyPass)

            ' Crea una nuova TextBox per inserire l'indirizzo email
            textBoxMyPass = New TextBox()
            textBoxMyPass.Location = New Point(70, 70)
            textBoxMyPass.Size = New Size(350, 20)
            Me.Controls.Add(textBoxMyPass)

            ' Crea una nuova Label con scritto "Oggetto"
            Dim labelOggetto As New Label()
            labelOggetto.Location = New Point(10, 100)
            labelOggetto.Size = New Size(50, 20)
            labelOggetto.Text = "Oggetto"
            Me.Controls.Add(labelOggetto)

            ' Crea una nuova TextBox per inserire l'oggetto della email
            textBoxOggetto = New TextBox()
            textBoxOggetto.Location = New Point(70, 100)
            textBoxOggetto.Size = New Size(350, 20)
            Me.Controls.Add(textBoxOggetto)

            ' Crea una nuova Label con scritto "Testo"
            Dim labelTesto As New Label()
            labelTesto.Location = New Point(10, 130)
            labelTesto.Size = New Size(50, 20)
            labelTesto.Text = "Testo"
            Me.Controls.Add(labelTesto)

            ' Crea una nuova TextBox per inserire il testo della email
            textBoxTesto = New TextBox()
            textBoxTesto.Location = New Point(70, 130)
            textBoxTesto.Size = New Size(350, 100)
            textBoxTesto.Multiline = True
            Me.Controls.Add(textBoxTesto)
            ComboBox1.Visible = False
        Else
            ' Mostra di nuovo la ComboBox1
            ComboBox1.Visible = True
            ' Nascondi tutti gli elementi creati
            For Each control As Control In Me.Controls
                If control IsNot Button4 AndAlso control IsNot ComboBox1 Then
                    control.Visible = False
                End If
            Next
            newDataGridView.Visible = True
            Button3.Visible = True
            Button4.Text = "Invia Mail pubblicitaria"
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim TrackToSend As String = comboBox.SelectedItem.ToString()
        Dim MyMail As String = textBoxMyMail.Text.ToString()
        Dim MyPass As String = textBoxMyPass.Text.ToString()
        Dim Subject As String = textBoxOggetto.Text.ToString()
        Dim Text As String = textBoxTesto.Text.ToString()

        Dim LinkToSend As String = ""

        Dim titleColumnIndex As Integer = dataGridView1.Columns("Title").Index
        Dim linkColumnIndex As Integer = dataGridView1.Columns("Link").Index

        For Each row As DataGridViewRow In dataGridView1.Rows
            If row.Cells(titleColumnIndex).Value.ToString() = TrackToSend Then
                LinkToSend = row.Cells(linkColumnIndex).Value.ToString()
                Exit For
            End If
        Next

        Dim emailHashSet As New HashSet(Of String)
        Dim emailColumnIndex As Integer = -1
        For i As Integer = 0 To newDataGridView.Columns.Count - 1
            If newDataGridView.Columns(i).HeaderText = "Email Address" Then
                emailColumnIndex = i
                Exit For
            End If
        Next

        If emailColumnIndex <> -1 Then
            For i As Integer = 0 To newDataGridView.Rows.Count - 2
                Dim email As String = newDataGridView.Rows(i).Cells(emailColumnIndex).Value.ToString().ToLower()
                emailHashSet.Add(email)
            Next
        End If

        'If String.IsNullOrWhiteSpace(MyMail) OrElse String.IsNullOrWhiteSpace(Subject) OrElse String.IsNullOrWhiteSpace(Text) OrElse String.IsNullOrWhiteSpace(LinkToSend) OrElse String.IsNullOrWhiteSpace(MyPass) Then
        '    MessageBox.Show("Tutti i campi sono obbligatori", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Return
        'Else
        '    For Each email As String In emailHashSet
        '        Dim gmailService As New GmailService(New Google.Apis.Services.BaseClientService.Initializer() With {
        '        .HttpClientInitializer = New Google.Apis.Auth.OAuth2.GoogleCredential(
        '            "YOURS ID SECRET",
        '            "YOURS DATA",
        '            "urn:ietf:wg:oauth:2.0:oob",
        '            New[] {"https://www.googleapis.com/auth/gmail.send"}
        '        )
        '    })

        '        Dim message As New Message()
        '        message.Subject = Subject
        '        message.Body = Text & vbCrLf & LinkToSend
        '        message.To = email

        '        Dim request As New GmailService.MessagesResource().SendRequest(gmailService, message)
        '    request.Execute()

        '        MessageBox.Show("Mail inviata con successo a " & email)
        '    Next
        'End If
    End Sub
End Class
