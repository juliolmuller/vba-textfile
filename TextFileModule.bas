Attribute VB_Name = "TextFileModule"

' =================================================================
' Text File Handler Module for Visual Basic for Application
' =================================================================
'
' Author:       Julio L. Muller
' Version:      1.0.0
' Repository:   https://github.com/juliolmuller/VBA-Module-TextFile
'
' =================================================================

Option Private Module
Option Explicit

'Create a new text file or overwrite an old one
Public Sub CreateTextFile(fullFileName As String, Optional content As String)

    'Dimension local varable
    Dim fileIndex As Integer

    'Determine the next file index available for use and open it
    fileIndex = FreeFile
    Open fullFileName For Output As fileIndex

    'Write text content in the file
    Print #fileIndex, content

    'Save and close text file
    Close fileIndex

End Sub

'Returns the content of a text file
Public Function GetTextFileContent(fullFileName As String) As String

    'Dimension local varable
    Dim fileIndex As Integer

    'Determine the next file number available for use by the FileOpen function and open it
    fileIndex = FreeFile
    Open fullFileName For Input As fileIndex

    'Capture file content
    GetTextFileContent = Input(LOF(fileIndex), fileIndex)

    'Close text fFile
    Close fileIndex

End Function

'Replace text fragments at a text file.
Public Function ReplaceAtTextFile(fullFileName As String, oldText As String, newText As String, Optional newFullFileName As String) As String

    'Dimension local varable
    Dim content As String

    'Define the destination file
    If (newFullFileName = vbNullString) Then
        newFullFileName = fullFileName
    End If

    'Store file content into a variable
    content = GetTextFileContent(fullFileName)

    'Repalce selected text and apply to a new or an existing text file
    content = Replace(content, oldText, newText)
    Call CreateTextFile(newFullFileName, content)

    'Expose results as return
    ReplaceAtTextFile = content

End Function

'Append text to a text file.
Public Function AppendToTextFile(fullFileName As String, newContent As String) As String

    'Dimension local varable
    Dim fileIndex As Integer

    'Determine the next file number available for use by the FileOpen function and open it
    fileIndex = FreeFile
    Open fullFileName For Append As fileIndex

    'Add the text into the selected file and close it
    Print #fileIndex, newContent

    'Save and close text file
    Close fileIndex

    'Expose results as return
    AppendToTextFile = GetTextFileContent(fullFileName)

End Function

'Capture the delimited content from a text file and load a 2-D array (columns must have all the same size).
Public Function ConvertTextFileToArray(fullFileName As String, horizontalDelimiter As String) As String()

    'Dimension local varables
    Dim fileContent As String
    Dim fullArray() As String
    Dim tempArray() As String
    Dim arrayRows() As String
    Dim colNum As Long
    Dim rowNum As Long
    Dim i As Long
    Dim j As Long
  
    'Store file content inside a variable
    fileContent = GetTextFileContent(fullFileName)
  
    'Separate Out lines of data
    arrayRows() = Split(fileContent, vbCrLf)

    'Read Data into an Array Variable
    For i = LBound(arrayRows) To UBound(arrayRows)
        If Len(Trim(arrayRows(i))) <> 0 Then

            'Split up fields within a line of text abd resize array boundaries
            tempArray = Split(arrayRows(i), horizontalDelimiter)
            colNum = UBound(tempArray)
            ReDim Preserve fullArray(colNum, rowNum)
      
            'Load line of data into Array variable
            For j = LBound(tempArray) To UBound(tempArray)
                fullArray(j, rowNum) = tempArray(j)
            Next j
        End If
        rowNum = rowNum + 1
    Next i

    'Return function result
    ConvertTextFileToArray = fullArray

End Function
