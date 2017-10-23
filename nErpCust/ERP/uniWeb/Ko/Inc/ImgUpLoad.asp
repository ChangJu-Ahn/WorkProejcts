<%
' Author Philippe Collignon
' Email PhCollignon@email.com

Sub BuildUploadRequest(RequestBin)
    Dim PosBeg
    Dim PosEnd
    Dim boundary
    Dim BoundaryPos
    Dim Pos
    Dim Name
    Dim PosFile
    Dim PosBound
    Dim FileName
    Dim ContentType
    Dim Value
    
	'Get the boundary
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	'Get all data inside the boundaries
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		'Members variable of objects are put in a dictionary object
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		'Get an object name
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		'Test if object is of file type
		If  PosFile<>0 AND (PosFile<PosBound) Then
			'Get Filename, content-type and content of file
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'Add filename to dictionary object
			UploadControl.Add "FileName", FileName
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			PosBeg = Pos+14
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			'Add content-type to dictionary object
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "ContentType",ContentType
			'Get content of object
			PosBeg = PosEnd+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
			'Get content of object
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		End If
		'Add content to dictionary object
	UploadControl.Add "Value" , Value	
		'Add dictionary object to main dictionary
	UploadRequest.Add name, UploadControl	
		'Loop to next object
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop

End Sub

'String to byte string conversion
Function getByteString(StringStr)
   Dim i
   Dim Char
   For i = 1 to Len(StringStr)
       char = Mid(StringStr,i,1)
       getByteString = getByteString & chrB(AscB(char))
   Next
End Function

'Byte string to string conversion
Function getString(StringBin)
    Dim intCount
    Dim iTemp
    
    getString =""
    For intCount = 1 to LenB(StringBin)
        If AscB(MidB(StringBin,intCount,1)) >= 128 And gLang = "KO" Then
            iTemp = AscB(MidB(StringBin,intCount,1)) * 256 + AscB(MidB(StringBin,intCount+1,1))
            iTemp = iTemp - 65536
            getString = getString & ChrW(iTemp) 
            intCount = intCount + 1
        Else
            getString = getString & ChrW(AscB(MidB(StringBin,intCount,1))) 
        End If
    Next
End Function
%>