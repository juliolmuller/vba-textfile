
# Text File Handler Module for Visual Basic for Application

- **Developed by:** Julio L. Muller
- **Released on:** Jan 24, 2019
- **Updated on:** Jan 24, 2019
- **Latest version:** 1.0.0
- **License:** *FREE*

## Installation

Imprort the `*.bas` file into your Visual Basic project by following the steps:

1. With the Excel workbook open. start the VBE window (`Alt + F11`);
2. In the menu, click on *File* > *Import File...* (`Ctrl + M`);
3. Through the file explorer, select the **TextFileModule.bas** file;
4. An item called *TextFileModule* will show up on your *Modules* list;
5. Enjoy!

Alternatively, copy and paste the plain text from the `*.txt` file into an existing module in your project.

**It is important to mention** that VBA does not handle **relative paths** very well, like other programming languages, so always prefar to use the **absolute paths** instead. To *emulate* relative paths, use the object `ThisWorkbook` and its method `Path` and then concatenate the folder/file paths starting on the current directory of the workbook you are working on. The use of this resource will be shown in the examples that follow. Also, when using `ThisWorkbook.Path`, always use the back slash (`\`) to navigate through the directories.

## Content Summary

| Type         | Name                                                         | Return Type |
|:------------:|:-------------------------------------------------------------|:-----------:|
| **Sub**      | [CreateTextFile](#create-text-file)                          | -           |
| **Function** | [GetTextFileContent](#get-text-file-content)                 | *String*    |
| **Function** | [ReplaceAtTextFile](#replace-text-file-content)              | *String*    |
| **Function** | [AppendToTextFile](#append-content-to-text-file)             | *String*    |
| **Function** | [ConvertTextFileToArray](#turn-text-file-content-into-array) | *Array*     |

## Resources Documentation

### Create Text File

This subroutine creates a text file and adds content to it if necessary. ***Warning:*** if the file informed already exists in the informed directory, the function will overwrite the old file and create a brand new one.

#### Structure

```vbnet
Sub CreateTextFile(fullFileName As String, Optional content As String)
```

- **fullFileName** - Target file name with its directory path. If the path informed contains folders that do not exist, the compiler will throw a **run-time error 76**.
- **content** (optional) - String which should be written in the created file. If this parameter is not passed, the file will be created blank.

#### Example

```vbnet
'Creates a text file called "my-file.txt" in the same directory of the workbook

    Dim file As String
    dim text As String

    file = ThisWorkbook.Path & "\my-file.txt"
    text = "Hello, Excel!"
    Call CreateTextFile(file, text)
```

```vbnet
'Creates a blank file called "my-file.txt" in the level-up folder "data"

    Dim path As String

    file = ThisWorkbook.Path & "\..\data\my-file.txt"
    Call CreateTextFile(file)
```

### Get Text File Content

Retrieves the content of a text file as *String*.

#### Structure

```vbnet
Function GetTextFileContent(fullFileName As String) As String
```

- **fullFileName** - Target file name with its directory path. If the file or the folder informed does not exist, the compiler will throw **run-time error 53** and **76** respectively.
- ***return*** - The function returns the exact text content of the file passed as parameter.

#### Example

```vbnet
'Copy the text file content into the variable "myContent"

    Dim file As String
    Dim myContent As String

    file = "C:/Users/Julio/Downloads/my-data.txt"
    myContent = GetTextFileContent(file)
```

### Replace Text File Content

Function to access a text file content, replace certain term by other and, alternatively. save the result into a new text file. ***Warning:*** if the new file path informed already exists in the informed directory, the function will overwrite the old file without any prompt.

#### Structure

```vbnet
Function ReplaceAtTextFile(fullFileName As String, oldText As String, newText As String, Optional newFullFileName As String) As String
```

- **fullFileName** - Target file name with its directory path. If the file or the folder informed does not exist, the compiler will throw **run-time error 53** and **76** respectively.
- **oldText** - The string term you want to find inside the text file.
- **newText** - The string term you want instead of the `oldText`. If you want to remove the `oldText`, simply pass `""` or `Empty` in this parameter.
- **newFullFileName** (optional) - If you want to keep the target file as is and have the result in another file instead, use this parameter to inform the file with the new content.
- ***return*** - The function returns the updated content of the target file. If you want to capture this return, attribute the function to a variable.

#### Example

```vbnet
'Capture the content "Hello, Excel!" in file "target.txt" and replace "Hello" by "Good-bye"

    Dim file As String
    Dim result As String

    file = ThisWorkbook.Path & "\target.txt"
    result = ReplaceAtTextFile(file, "Hello", "Good-bye")

    MsgBox result   'displays "Good-bye, Excel!"
```

### Append Content to Text File

Place a given string at the end of an existing text file. ***Warning:*** if the new file path informed already exists in the informed directory, the function will overwrite the old file without any prompt.

#### Structure

```vbnet
Function AppendToTextFile(fullFileName As String, newContent As String) As String
```

- **fullFileName** - Target file name with its directory path. If the file or the folder informed does not exist, the compiler will throw **run-time error 53** and **76** respectively.
- **newContent** - String which should be written at the end of the passed file.
- - ***return*** - The function returns the updated content of the target file. 

#### Example

```vbnet
'Open the file "products.txt" and add a record at the end of the file

    Dim file As String
    Dim newProd As String

    file = "products.txt"
    newProd = "1812,Gamer Mouse,20,$19.99"
    Call AppendToTextFile(file, newProd)
```

### Turn Text File Content into Array

Excellent function to capture tables stored in text files, manipulating the data as a bidimentional array. ***Warning:*** this function works fine only with verywell dimensioned data set. You **MUST** have your table header on the first line and immediately followed by the data, or you can ommit the header and start the file with the data directly. **Blamk rows or rows with different sizes than your header are not supported.**

#### Structure

```vbnet
Function ConvertTextFileToArray(fullFileName As String, horizontalDelimiter As String) As String()
```

- **fullFileName** - Target file name with its directory path. If the file or the folder informed does not exist, the compiler will throw **run-time error 53** and **76** respectively.
- **horizontalDelimiter** - Inform the string which represents the jump from a column to another (the delimiter between two data cells). For example, the most commonly used delimiters are: `Tab`, `","` (comma), `";"` (semicolon) and `"|"` (pipeline). Just for instance, the vertical delimiter is the *line break* (`Enter`). This character (or string) is used only as metadata, and it is not imported into the array.
- - ***return*** - The function returns a bidimentioned array of strings, so its return should be attributed to variable of type `String()` or `Variant`. The indexes of the array are designed as *Columns & Rows*, meaning the firs index indicates the column, and the second one indicates the row of the content obtained from the file. Both indexes start in '0' (zero) and it does not change whether you have header line or not.

#### Example

```vbnet
'gets a table 4C x 5000R

    Dim myFile As String
    Dim myRange As Range
    Dim myTable() As String
    Dim i As Long, j As Long

    myFile = ThisWorkbook.Path & "\dabase\customers.txt"
    Set myRange = ActiveSheet.Range("A:J")
    myTable = ConvertTextFileToArray(myFile, vbTab)   'columns delimited by TAB

    For i = 0 To 50
        For j = 0 To UBound(myTable, 1)
            myRange.Cells(i + 1, j + 1).Value = myTable(j, i)
        Next j
    Next i
```

### Compatibility

The scripts were tested **ONLY** in MS Excel 2013 and 2016. MS Access and other MS Office applications were not tested.

Please, report any issues (or even success on running it in other applications or Excel versions) through the commentary session.

### Other Contents

- [File Handler Class for VBA](https://github.com/juliolmuller/VBA-Class-FileHandler)
- [MS Outlook Module for VBA](https://github.com/juliolmuller/VBA-Module-Outlook)
- [Native File Handler API in VBA](http://www.homeandlearn.org/open_a_text_file_in_vba.html)
