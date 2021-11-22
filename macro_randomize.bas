REM  *****  BASIC  *****

Sub RandomizeContent()
	On Error GoTo Except

	' Grabs config file in directory of spreadsheet file
	configFile = GetFilePath("config.txt")
	numTag = FreeFile()
	Open configFile for input as #numTag
	Input #numTag, fileToUse
	Close #numTag

	' Grabs file specified from config in directory of spreadsheet file
	fileName = fileToUse
	textFile = GetFilePath(fileName)

	' Reads file specified by config
	Open textFile for input as #numTag
	totalLines = 0
	Do While Not eof(numTag)
		ReDim Preserve fileLines(totalLines)
		Input #numTag, fileLines(totalLines)
		totalLines = totalLines + 1
	Loop
	Close #numTag

	' Stores elements in array
	totalElements = ArrayLen(fileLines) - 1
	For iEle = 0 To totalElements Step 1
		ReDim Preserve elements(iEle)
		If Not IsEmpty(fileLines(iEle)) Then
			elements(iEle) = fileLines(iEle)
		End If
	Next

	' Assigns elements to cell contents
	oService = createUnoService("com.sun.star.sheet.addin.Analysis")
	currSheet = ThisComponent.CurrentController.ActiveSheet
	activeRange = ThisComponent.CurrentSelection
	rangeCoor = activeRange.RangeAddress
	For iCol = rangeCoor.StartColumn To rangeCoor.EndColumn
		For iRow = rangeCoor.StartRow To rangeCoor.EndRow
			Cel = currSheet.getCellByPosition(iCol, iRow)
			RandBetween = oService.getRandbetween(0, totalElements)
			Cel.String = elements(RandBetween)
		Next
	Next

	Except:
	If Err = 57 Then
		If Not IsEmpty(fileName) Then
			msgbox "Missing file. Create a new '" & fileName & "'."
		Else
			msgbox "Missing file. Create a new 'config.txt'."
		End If
	ElseIf Error = "Reading exceeds EOF." Then
		msgbox "Invalid content in '" & fileName & "'. Please remove any empty lines then try again."
	ElseIf Err = 91 Then
		msgbox "'" & fileName & "' is empty. Please add content then try again."
	ElseIf Err <> 0 Then
		msgbox Error
	End If
End Sub

Function ArrayLen(arr as Variant) as Integer
	ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function GetFilePath(query as String) as String
	pathURL = ThisComponent.getURL()
	pathOS = ConvertFromURL(pathURL)
	fileName = Dir(pathOS)
	pathLen = Len(pathOS)
	nameLen = Len(fileName)
	path = Left(pathOS, pathLen - nameLen)
	GetFilePath = path & query
End Function