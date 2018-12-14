Option Compare Text
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop
Imports xl = Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private writingLock As New Object
    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub


    Private Sub Application_NewMailEx(EntryIDCollection As String) Handles Application.NewMailEx
        For Each itemID In Split(EntryIDCollection, ",")
            Dim item = Application.Session.GetItemFromID(itemID)
            If TypeName(item) = "MailItem" Then
                'is it a Dell Email
                Dim msg As Outlook.MailItem
                msg = item
                If msg.Subject.ToLower.StartsWith("opportunity ") AndAlso Not msg.Subject.ToLower.StartsWith("opportunity submitted") AndAlso msg.SenderEmailAddress.ToLower.Contains("dell.com") Then
                    'extract the info
                    Dim dealdetails As New List(Of String)
                    dealdetails.Add(msg.ReceivedTime.ToShortDateString)
                    If msg.Subject.ToLower.StartsWith("opportunity approved") Then
                        dealdetails.Add("Approved")
                    ElseIf msg.Subject.ToLower.StartsWith("opportunity declined") Then
                        dealdetails.Add("Declined")
                    End If
                    dealdetails.AddRange(readDetails(msg.HTMLBody))

                    'post info online
                    Dim t As Task = Task.Run(Sub() writeOut(dealdetails))
                    't.Start()

                End If
            End If
        Next
    End Sub

    Public Function readDetails(ByVal htmlString As String, Optional sender As String = "") As List(Of String)

        readDetails = New List(Of String)
        Dim spos As Integer, epos As Integer, tableHTML As String
        Dim customer As String, am As String, value As String, dealType As String
        Dim products As String, amArr As String(), dealID As String

        If InStr(htmlString, "strong") = 0 Then Exit Function
        spos = InStr(htmlString, "ID</strong>:")
        If spos = 0 Then

            Exit Function
        Else
            spos += Len("ID</strong>:")
        End If
        epos = InStr(spos, htmlString, "<br>")
        dealID = Mid(htmlString, spos, epos - spos)




        spos = InStr(htmlString, "End User Account Name</strong>:")

        customer = ""

        If spos > 0 Then
            spos += Len("End User Account Name</strong>:")
            epos = InStr(spos, htmlString, "<br>")
            customer = Mid(htmlString, spos, epos - spos)
        End If

        spos = InStr(htmlString, "<strong>Opportunity Name</strong>:") + Len("<strong>Opportunity Name</strong>:")
        epos = InStr(spos, htmlString, "<br>")
        am = Mid(htmlString, spos, epos - spos)

        If InStr(am, "/") = 0 Then Exit Function
        amArr = am.Split("/")
        am = amArr(1)

        If customer = "" Then
            customer = amArr(0)
        End If

        spos = InStr(htmlString, "Total Expected Revenue</strong>:")

        If spos = 0 Then
            spos = InStr(htmlString, "Total Expected  Revenue </strong>:") + Len("Total Expected  Revenue </strong>:")
        Else
            spos += Len("Total Expected Revenue</strong>:")
        End If
        epos = InStr(spos, htmlString, "<br>")
        value = Mid(htmlString, spos, epos - spos)

        If Len(value) > 20 Then
            value = "0"
        End If


        spos = InStr(htmlString, "Deal Type:</strong>")
        If spos = 0 Then
            spos = InStr(htmlString, "Deal Type</strong> :") + Len("Deal Type</strong> :")
        Else
            spos += Len("Deal Type:</strong>")
        End If
        epos = InStr(spos, htmlString, "<br>")
        dealType = Mid(htmlString, spos, epos - spos)
        If Len(dealType) > 100 Then dealType = "Unknown"

        spos = InStr(htmlString, "List of Products Associated to this Opportunity")
        epos = InStr(htmlString, "To view this deal please visit your Partner Portal")
        If epos = 0 Then epos = InStr(htmlString, "For a non-registered quote on this opportunity")
        If epos = 0 Then epos = Len(htmlString)
        tableHTML = Mid(htmlString, spos, epos - spos)

        Dim html As New clsHTMLString(tableHTML)
        Dim tables = html.getTables()

        Try
            products = Join(";", tables(0))
        Catch
            products = ""
        End Try

        'work around for when deals are forwarded to Mat Lawless
        If Environment.UserName.ToLower <> "mlawless" Then
            readDetails.Add(Environment.UserName)
        Else
            readDetails.Add(sender)
        End If

        readDetails.Add(dealID)
        readDetails.Add(customer)
        readDetails.Add(am)
        readDetails.Add(value)
        readDetails.Add(dealType)
        readDetails.Add(products)
        readDetails = Trim(readDetails)
    End Function

    Private Function Trim(myList As List(Of String)) As List(Of String)
        Trim = New List(Of String)

        For Each str As String In myList
            Trim.Add(Strings.Trim(str))
        Next
    End Function
    Private Function Join(delimeter As String, list As List(Of List(Of String))) As String
        Join = ""
        Dim subList As List(Of String)
        For Each subList In list
            Join = Join & delimeter & String.Join(delimeter, subList)
        Next

    End Function
    Private Function Join(delimeter As String, list As List(Of List(Of List(Of String)))) As String
        Join = ""
        Dim subList As List(Of List(Of String))
        For Each subList In list
            Join = Join & delimeter & Join(delimeter, subList)
        Next
    End Function

    'Public Sub writeToExcel(newLine As List(Of String))
    '    SyncLock writingLock ' prevents more than one write taking place at once.
    '        Try
    '            Dim xlApp As New xl.Application
    '            xlApp.Visible = True
    '            xlApp.DisplayAlerts = False
    '            Try


    '                Dim tWb As xl.Workbook
    '                tWb = xlApp.Workbooks.Open("https://insightonlinegbr-my.sharepoint.com/personal/martin_klefas_insight_com/Documents/Dell%20NBI%20Log.xlsx")

    '                Dim nextLine As xl.Range
    '                nextLine = tWb.Sheets("Raw Data").Range("A1").End(xl.XlDirection.xlDown).Offset(1, 0)
    '                For i As Integer = 0 To newLine.Count - 1
    '                    nextLine.Offset(0, i).Value = newLine(i)
    '                Next
    '                tWb.Save()
    '                xlApp.Quit()
    '            Catch
    '                xlApp.Quit()
    '                MsgBox("Could not open the deal log, and so this deal has not been added to the NBI Log")
    '            End Try
    '        Catch
    '            MsgBox("Excel could not be started, and so this deal has not been added to the NBI Log")
    '        End Try

    '    End SyncLock
    'End Sub
    Sub writeOut(line As List(Of String))
        SyncLock writingLock ' prevents more than one write taking place at once.
            If line.Count > 0 Then
                Dim file As System.IO.StreamWriter, tString As String
                tString = Chr(34) & String.Join(Chr(34) & "," & Chr(34), line) & Chr(34)

                file = My.Computer.FileSystem.OpenTextFileWriter("\\insight.com\root\Shared\Sales\Public sector\Martin Klefas\NBI Reporting\NBI.csv", True)
                file.WriteLine(tString)
                file.Close()
            End If
        End SyncLock
    End Sub
End Class
