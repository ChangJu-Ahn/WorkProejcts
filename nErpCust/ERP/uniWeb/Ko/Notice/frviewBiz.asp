<%@ LANGUAGE=VBSCript %>

<!-- #Include file="../inc/IncServer.asp" -->

<%

Dim iStrFileInfo
Dim arrTemp, arrTemp2
Dim iDx, iPos, iFilePath, iFileId, iFileTempPath
Dim iExtensionPos, iStrExtension

Const C_FileName = 1
Const C_FilePath = 2
Const C_FileId = 3	
Const C_FileSize = 8
Const C_FileCDate = 9

iStrFileInfo = Request("FileInfo")
iFileMode    = Request("FileMode")

arrTemp  = Split(iStrFileInfo, chr(31))
arrTemp2 = Split(arrTemp(1), chr(29))

	'파일 복사 
	iPos      = InStrRev(arrTemp2(C_FileId), "/", -1)		
	iFileTempPath   = Mid(arrTemp2(C_FileId), 1,iPos)
	iFileTempPath = Replace(iFileTempPath,"/", "\")
	iFileId   = right(arrTemp2(C_FileId),len(arrTemp2(C_FileId)) - iPos)	

	iPos      = InStrRev(arrTemp2(C_FilePath), "\", -1)		

	iFilePath = left(arrTemp2(C_FilePath),iPos)

	Call FileTransfer(iFilePath & iFileId, iFileTempPath,iFileId)

'	Call Writelog("iFilePath:" & iFilePath & vbcrlf)
'	Call Writelog("iFileTempPath:" & iFileTempPath & vbcrlf)
'	Call Writelog("iFileId:" & iFileId & vbcrlf)

    iExtensionPos = InStrRev(iFileId, ".", -1)		
	iStrExtension = right(iFileId,len(iFileId) - iExtensionPos)	

    Response.Write "<Script Language=VBS>"   & vbCr
    Response.Write " If Trim(UCASE(""" & iFileMode & """)) = ""W"" Then " & vbCr 
    Response.Write "         Call Parent.vbViewFile(""W"",""" & iStrFileInfo & """" & ")  "            & vbCr    '파일저장    
    Response.Write " Else "  & vbCr            '파일보기 
'    Response.Write "      If Trim(Parent.vbCheckFileAssociation(""" & iStrExtension & """)) <> """" Then "  & vbCr
    Response.Write "         Call Parent.vbViewFile(""F"",""" & iStrFileInfo & """" & ")  "            & vbCr	'파일보기 
'    Response.Write "      Else "  & vbCr       '만약 파일을 Open할 프로그램이 존재하지않으면 
'    Response.Write "         Call Parent.vbViewFile(""W"",""" & iStrFileInfo & """" & ")  "            & vbCr    '파일저장 
'    Response.Write "      End If "  & vbCr
    Response.Write " End If "  & vbCr                            
    Response.Write "</Script>"                    & vbCr
	
Response.End



Function FileTransfer(SourceFilePath,TargetPath, TargetFileName)
	
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.FileCopy(SourceFilePath,TargetPath, TargetFileName)
   
   Set pfile = Nothing
   
End Function

Function WriteLog(LogData)
        Dim objFileForLog
		Dim objFSO		
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")                     'For testing
        Set objFileForLog = objFSO.OpenTextFile("C:\Notice.log", 8, True, -2) 'For testing
        objFileForLog.WriteLine LogData
        objFileForLog.Close
        Set objFileForLog = Nothing
End Function

%>