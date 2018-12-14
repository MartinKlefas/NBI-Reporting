Imports System.IO
Imports System.Threading


Public Class Form1
    Public runningCount As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'This is the control Thread
        If My.Computer.FileSystem.FileExists(Path.GetTempPath & "NBI.csv") Then
            My.Computer.FileSystem.DeleteFile(Path.GetTempPath & "NBI.csv")
        End If

        Button1.Enabled = False
        Button2.Enabled = False

        BackgroundWorker1.RunWorkerAsync()

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim msg As Outlook.MailItem
        msg = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)
        msg.Attachments.Add(Path.GetTempPath & "NBI.csv")
        msg.Body = "The attached is a report containing information about all Dell Deals in this inbox."
        msg.To = "martin.klefas@insight.com"
        msg.Display()
        Call enableButtons("")
    End Sub

    Private Sub enableButtons(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Button1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf enableButtons)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button2.Text = "Done"
        End If
    End Sub

    Private Sub SetText(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf SetText)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label1.Text = [text]
        End If
    End Sub

    Private Sub SetTextTwo(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf SetTextTwo)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label2.Visible = True
            Me.Label2.Text = [text]
        End If
    End Sub

    Delegate Sub SetTextCallback(ByVal [text] As String)

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim results1 As Outlook.Results, results2 As Outlook.Results
        Dim dealDetails As New List(Of String)
        runningCount = 0
        results1 = findItems("'Inbox'", "Opportunity Approved")
        results2 = findItems("'Inbox'", "Opportunity Declined")

        SetTextTwo("Processing of emails begun.")
        Dim processed As Integer = 0
        For Each result As Outlook.MailItem In results1
            dealDetails = (Globals.ThisAddIn.readDetails(result.HTMLBody))
            writeOut(result.ReceivedTime.ToShortDateString, "Approved", dealDetails)
            processed += 1
            Call SetTextTwo("Processed " & processed & " of " & runningCount & " Emails")
        Next
        For Each result As Outlook.MailItem In results2
            dealDetails = (Globals.ThisAddIn.readDetails(result.HTMLBody))
            writeOut(result.ReceivedTime.ToShortDateString, "Declined", dealDetails)
            processed += 1
            Call SetTextTwo("Processed " & processed & " of " & runningCount & " Emails")
        Next
    End Sub


    Function findItems(folder As String, subject As String) As Outlook.Results

        Dim adsearch As Outlook.Search

        adsearch = Globals.ThisAddIn.Application.AdvancedSearch(folder, Chr(34) & "urn:schemas:httpmail:subject" _
                         & Chr(34) & " like '%" & subject.Trim & "%'", True)

        Dim j As Integer, k As Integer = 0
        j = adsearch.Results.Count
        For i As Integer = 1 To 40
            Call SetText(runningCount + j & " Results found so far")
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1))

            If adsearch.Results.Count = j Then
                If k > 3 Then Exit For
                k += 1
                Dim str As String
                str = runningCount + j & " Results found so far"
                For l = 1 To k
                    str = str & "."
                Next
                Call SetText(str)
            Else
                k = 0
                j = adsearch.Results.Count
            End If
        Next

        runningCount += j

        findItems = adsearch.Results
    End Function

    Sub writeOut(strDate As String, strStatus As String, line As List(Of String))
        If line.Count > 0 Then
            Dim file As System.IO.StreamWriter, tString As String
            tString = Chr(34) & strDate & Chr(34) & "," & Chr(34) & strStatus &
                        Chr(34) & "," & Chr(34) & String.Join(Chr(34) & "," & Chr(34), line) &
                        Chr(34)

            file = My.Computer.FileSystem.OpenTextFileWriter(Path.GetTempPath & "NBI.csv", True)
            file.WriteLine(tString)
            file.Close()
        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()
    End Sub
End Class