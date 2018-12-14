Imports System.Diagnostics
Imports System.Net
Imports System.Text.RegularExpressions

Public Class clsHTMLString
    Private rawString As String
    Public Sub New(ByVal html As String)
        rawString = html
    End Sub
    Function getTables() As List(Of List(Of List(Of String)))
        'List of tables(list of lines)
        getTables = New List(Of List(Of List(Of String)))
        Dim pos As Integer, endpos As Integer
        Dim tableEnd As Integer, rowEnd As Integer

        Dim rawElement As String
        Dim elements As New List(Of String)
        Dim rows As New List(Of List(Of String))

        pos = 1

        While pos > 0

            pos = InStr(pos, rawString, "<table")
            If pos > 0 Then
                tableEnd = InStr(pos, rawString, "</table")

                While pos < tableEnd And pos > 0

                    If pos > 0 And pos < tableEnd Then
                        pos = InStr(pos, rawString, "<tr")

                        If pos > 0 And pos < tableEnd Then
                            rowEnd = InStr(pos, rawString, "</tr")
                            While pos < rowEnd And pos > 0 And pos < tableEnd
                                pos = InStr(pos, rawString, "<td")
                                If pos > 0 And pos < rowEnd Then
                                    pos = InStr(pos, rawString, ">") + 1
                                    endpos = InStr(pos, rawString, "</td")

                                    rawElement = Mid(rawString, pos, endpos - pos)
                                    rawElement = stripTags(rawElement)

                                    elements.Add(TrimMore(WebUtility.HtmlDecode(rawElement)))
                                Else
                                    pos = rowEnd
                                End If
                            End While
                            rows.Add(elements.ToList) 'https://stackoverflow.com/questions/52478152/is-there-a-method-to-add-to-a-list-of-listof-string-byval/52478564#52478564
                            elements.Clear()
                        End If


                    End If

                End While
                getTables.Add(rows.ToList)
                rows.Clear()
                pos = tableEnd
            End If



        End While

    End Function

    Function stripTags(ByVal html As String) As String

        While InStr(html, "<script>") > 0
            Dim newstr As String = ""
            newstr = Left(html, InStr(html, "<script>"))
            newstr = newstr & Mid(html, InStr(html, "</script>"))
            html = newstr
        End While


        Dim tChar As Char, ignore As Boolean
        stripTags = ""
        ignore = False

        For Each tChar In html
            If tChar = "<" Then ' start ignoring characters starting with the open tag
                ignore = True
            End If

            If Not ignore Then
                stripTags = stripTags & tChar
            End If

            If tChar = ">" Then ' stop ignoring characters after a close tag
                ignore = False
            End If
        Next

    End Function
    Function TrimMore(ByVal tStr) As String
        TrimMore = Regex.Replace(tStr, "[^A-Za-z0-9\-/ ]", "")
        TrimMore = Trim(TrimMore)
    End Function

End Class
