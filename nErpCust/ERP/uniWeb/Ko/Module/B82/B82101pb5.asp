<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : 
'*  5. Program Desc         : �������ϰ���(���Ϻ��������)
'*  6. Comproxy List        :
'*  7. Modified date(First) :
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncServer.asp" -->

<%    
on Error Resume Next


Dim iStrFileInfo, iFileMode
Dim strSourceFilePath, strTargetPath, strTargetFileName
iStrFileInfo = Request("txtViewRet")
iFileMode    = Request("txtFileMod")
		
strSourceFilePath = Request("txtSourceFilePath")
strTargetPath     = Request("txtTargetPath")

strTargetPath     = Replace(strTargetPath,"/", "\")

strTargetFileName = Request("txtTargetFileName")


Call FileTransfer(strSourceFilePath & strTargetFileName , strTargetPath, strTargetFileName)
Response.Write "<Script Language=VBS>"   & vbCr
Response.Write " If Trim(UCASE(""" & iFileMode & """)) = ""W"" Then " & vbCr 
Response.Write "    Call Parent.ViewFile(""W"",""" & iStrFileInfo & """" & ")  "            & vbCr  '��������    
Response.Write " Else "  & vbCr 
Response.Write "    Call Parent.ViewFile(""F"",""" & iStrFileInfo & """" & ")  "            & vbCr	'���Ϻ��� 
Response.Write " End If "  & vbCr                            
Response.Write "</Script>"                    & vbCr
	
Response.End

Function FileTransfer(SourceFilePath,TargetPath, TargetFileName)
	
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.FileCopy(SourceFilePath,TargetPath, TargetFileName)
   
   Set pfile = Nothing
   
End Function


%>