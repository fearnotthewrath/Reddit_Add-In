Attribute VB_Name = "Reddit_AddIn"
Option Explicit
Sub Convert_Selection_To_Reddit_Table()
'iterators
Dim i As Integer
Dim j As Integer
Dim k As Integer

'Object for clipboard
Dim DataObj As Object
'Selection Range
Dim MatrixArray As Range

'strings used for formatting and output
Dim formatString As String
Dim revFormatStr As String
Dim tempString As String
Dim FinalString As String
Dim cleanString As String

'helper measures
Dim tableRows As Integer
Dim tableCols As Integer

'add characters here that need to be escaped (have a backslash added so they display in Reddit).
'The backslash MUST be the first character, else it will double up all of the slashes
cleanString = "\^*~`"

Set DataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
Set MatrixArray = Selection
tableRows = MatrixArray.Rows.Count
tableCols = MatrixArray.Columns.Count

'Check size, must be at least 2x2
If tableRows < 2 Or tableCols < 2 Then
        MsgBox "Selection Too Small, must be at least 2x2"
        Exit Sub
End If

'For each row
For i = 1 To tableRows
        If i = 2 Then 'Set the alignment and table formatting for the table. Based on the alignment of the second row of the selected table
                For j = 1 To tableCols
                        Select Case MatrixArray(2, j).HorizontalAlignment
                                Case xlGeneral: FinalString = FinalString & "-:|" ' General
                                Case xlLeft: FinalString = FinalString & ":-|" ' Left
                                Case xlCenter: FinalString = FinalString & ":-:|" ' Center
                                Case xlRight: FinalString = FinalString & "-:| " ' Right
                                Case Else
                        End Select
                Next
                FinalString = FinalString & Chr(10)
        End If
        'For each column
        For j = 1 To tableCols
            'Using .Text here so that the formatted Excel display is used, instead of the underlying .Value
            tempString = MatrixArray(i, j).Text
            For k = 1 To Len(cleanString) 'escape characters are escaped. add characters in variable definition above
                tempString = Replace(tempString, Mid(cleanString, k, 1), "\" & Mid(cleanString, k, 1))
            Next k
                'Reddit formatting
                If MatrixArray(i, j).Font.Strikethrough Then
                    formatString = formatString & "~~" 'StrikeThrough
                    revFormatStr = "~~" & revFormatStr
                End If
                If MatrixArray(i, j).Font.Bold Then
                    formatString = formatString & "**" 'Bold
                    revFormatStr = "**" & revFormatStr
                End If
                If MatrixArray(i, j).Font.Italic Then
                    formatString = formatString & "*" 'Italic
                    revFormatStr = "*" & revFormatStr
                End If
                If MatrixArray(i, j).Font.Superscript Then
                    formatString = formatString & "^" 'SuperScript
                End If
                
                'Build the cell contents
                FinalString = FinalString & formatString & tempString & revFormatStr & "|"
                formatString = vbNullString 'Clear format
                revFormatStr = vbNullString
        Next
        FinalString = FinalString & Chr(10) 'line break
Next

        'Max chars in Reddit comments is 10k. Hope you only wanted the table!
        If Len(FinalString) > 10000 Then
            MsgBox ("There are too many characters for Reddit comment! 10,000 characters copied.")
            FinalString = Left(FinalString, 9999)
        End If


DataObj.SetText FinalString
DataObj.PutInClipboard

'cleanup
Set MatrixArray = Nothing
Set DataObj = Nothing

MsgBox "Data copied to clipboard!", vbOKOnly, "Written by: /u/norsk"

End Sub
