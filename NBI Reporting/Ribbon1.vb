Option Compare Text

Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports xl = Microsoft.Office.Interop.Excel


Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim myForm As New Form1
        myForm.Show()


    End Sub

    Private Sub Button2_Click_1(sender As Object, e As RibbonControlEventArgs) 
        Try
            Dim xlApp As New xl.Application
            xlApp.Visible = True
            xlApp.DisplayAlerts = False
            Try


                Dim tWb As xl.Workbook
                tWb = xlApp.Workbooks.Open("https://insightonlinegbr-my.sharepoint.com/personal/martin_klefas_insight_com/Documents/Dell%20NBI%20Log.xlsx")

                Dim nextLine As xl.Range
                nextLine = tWb.Sheets("Raw Data").Range("A1").End(xl.XlDirection.xlDown).Offset(1, 0)


            Catch
                xlApp.Quit()
                MsgBox("Could not open the deal log, and so this deal has not been added to the NBI Log")
            End Try
        Catch
            MsgBox("Excel could not be started, and so this deal has not been added to the NBI Log")
        End Try


        '            


        '            


    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) 
        Try
            Dim xlApp As New xl.Application
            xlApp.Visible = True
            xlApp.DisplayAlerts = False
            Try


                Dim tWb As xl.Workbook
                tWb = xlApp.Workbooks.Open("https://insightonlinegbr-my.sharepoint.com/:x:/g/personal/martin_klefas_insight_com/ERYnTa25G7JFgUxwxkCvC2EBxCqsEcUUENx6M3MQBzmJFw?e=byDpGJ")

                Dim nextLine As xl.Range
                nextLine = tWb.Sheets("Raw Data").Range("A1").End(xl.XlDirection.xlDown).Offset(1, 0)


            Catch
                xlApp.Quit()
                MsgBox("Could not open the deal log, and so this deal has not been added to the NBI Log")
            End Try
        Catch
            MsgBox("Excel could not be started, and so this deal has not been added to the NBI Log")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) 
        Try
            Dim xlApp As New xl.Application
            xlApp.Visible = True
            xlApp.DisplayAlerts = False
            Try


                Dim tWb As xl.Workbook
                tWb = xlApp.Workbooks.Open("\\insight.com\root\Shared\Sales\Public sector\Martin Klefas\NBI Reporting\Dell NBI Log.xlsx")

                Dim nextLine As xl.Range
                nextLine = tWb.Sheets("Raw Data").Range("A1").End(xl.XlDirection.xlDown).Offset(1, 0)


            Catch
                xlApp.Quit()
                MsgBox("Could not open the deal log, and so this deal has not been added to the NBI Log")
            End Try
        Catch
            MsgBox("Excel could not be started, and so this deal has not been added to the NBI Log")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        Dim olCurrExplorer As Outlook.Explorer
        Dim olCurrSelection As Outlook.Selection
        Dim i As Integer = 0

        olCurrExplorer = Globals.ThisAddIn.Application.ActiveExplorer
        olCurrSelection = olCurrExplorer.Selection

        For Each item In olCurrSelection

            If TypeName(item) = "MailItem" Then
                'is it a Dell Email
                Dim msg As Outlook.MailItem
                msg = item
                If msg.Subject.ToLower.Contains("opportunity ") AndAlso Not msg.Subject.ToLower.Contains("opportunity submitted") Then
                    'extract the info
                    Dim dealdetails As New List(Of String)
                    dealdetails.Add(msg.ReceivedTime.ToShortDateString)
                    If msg.Subject.ToLower.Contains("opportunity approved") Then
                        dealdetails.Add("Approved")
                    ElseIf msg.Subject.ToLower.Contains("opportunity declined") Then
                        dealdetails.Add("Declined")
                    End If
                    dealdetails.AddRange(Globals.ThisAddIn.readDetails(msg.HTMLBody, msg.SenderEmailAddress))

                    'post info online
                    Call Globals.ThisAddIn.writeOut(dealdetails)
                    't.Start()
                    i += 1
                End If
            End If
        Next

        MsgBox(i & " items added to the log.")

    End Sub
End Class
