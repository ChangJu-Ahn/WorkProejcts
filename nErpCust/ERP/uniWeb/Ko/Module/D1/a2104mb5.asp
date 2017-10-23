<%
'**********************************************************************************************
'*  1. Module Name          : 파일로 존재하는 업무메뉴와 DB에서 읽은 유저메뉴를 되돌려준다.
'*  2. Function Name        :
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) :
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  :
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<Script Language=vbscript src="../../inc/incUni2KTV.vbs"></Script>
<Script LANGUAGE=VBScript>
Sub Document_onReadyStateChange()
	parent.frm1.uniTree1.MousePointer = 0
End Sub
</Script>
<%		

Call HideStatusWnd		

On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 파일 열기에 필요한 상수 
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

Const C_SEP = "::"
Const C_MNU_ID = 0
Const C_MNU_UPPER = 1
Const C_MNU_LVL = 2
Const C_MNU_TYPE = 3
'Const C_MNU_SEQ = 4
Const C_MNU_NM = 4
Const C_MNU_AUTH = 5
'Const C_MNU_PGM = 6

Const C_Open  = "Open"
Const C_Folder  = "Folder"
Const C_URL  = "URL"
Const C_None = "None"
Const C_Const = "Const"
Const C_UNDERBAR = "_"

'Dim pAb0018, TextStream
Dim arrMenu, i, arrLine, strImg

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim intColCnt																'Popup이 Display해야 할 columns수 
Dim LngRow
Dim GroupCount    
Dim strCmd
Dim strKey
Dim IntLvl
Dim intSeq

Dim pAb0018


'Response.End 
Set pAb0018 = Server.CreateObject("Ab0018.AListAcctSvr")
strCmd =  Request("strCmd")

IF Request("NextCd") = "" then
	strKey = ""	
else
	strKey =  Request("NextCd")
end if

IF Request("NextCd1") = "" then
	IntLvl = 0
else
	IntLvl = Cint(Request("NextCd1"))
end if

IF Request("NextCd2") = "" then
	intSeq = 0	
else
	intSeq = Cint(Request("NextCd2"))
end if

Select Case strCmd
    Case "LISTTOP"		
		pAb0018.CommandSent = "LISTTOP"         
        pAb0018.ImportAAcctGpGpCd = Trim(strKey)
    Case "LISTGP"		
        pAb0018.CommandSent = "LISTGP"
%>		               
        IF parent.lgQueryFlag = "1" Then
<%		       
			pAb0018.ImportAAcctGpGpCd = ""
			pAb0018.ImportAAcctGpGpLvl = 0
			pAb0018.ImportAAcctGpGpSeq = 0
%>					
		Else
<%				
			pAb0018.ImportAAcctGpGpCd = Trim(strKey)
			pAb0018.ImportAAcctGpGpLvl = IntLvl
			pAb0018.ImportAAcctGpGpSeq = IntSeq
%>					
		End IF 	
<%				
    Case "LISTACCT"		
		pAb0018.CommandSent = "LISTACCT"
%>		
		IF parent.lgQueryFlag = "1" Then		
<%		
			pAb0018.ImportAAcctGpGpCd = ""
			pAb0018.ImportAAcctAcctSeq = 0
%>			
		Else
<%		
			pAb0018.ImportAAcctGpGpCd = Trim(strKey)
			pAb0018.ImportAAcctAcctSeq = IntSeq
%>			
		End IF			
<%		
End Select

pAb0018.ServerLocation = ggServerIP
pAb0018.Comcfg = gConnectionString
pAb0018.Execute

'-----------------------
'Com action result check area(DB,internal)
'-----------------------
If Not (pAb0018.OperationStatusMessage = MSG_OK_STR) Then
	Call DisplayMsgBox(pAb0018.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)

	Set pAb0018 = Nothing
	Response.End
End If

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '⊙:
   Set pAb0018 = Nothing																	'☜: ComProxy UnLoad
   Response.End																				'☜: Process End
End If

IF 	strCmd <> "LISTACCT" Then
	GroupCount = pAb0018.ExportGroupGpCount
Else
	GroupCount = pAb0018.ExportGroupAcctCount
end if	
'Set pAb0018 = Nothing
%>
<Script Language=vbscript>
	Dim NodX
	With parent.frm1
<%
For LngRow = 1 To GroupCount	
	Select Case strCmd
		Case "LISTTOP"
%>						
			Set NodX = .uniTree1.Nodes.Add (parent.C_USER_MENU_KEY, tvwChild, "G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>", "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpNm(LngRow))%>", "<%=C_Folder%>" )			
			.uniTree1.Nodes("G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>").Tag = cstr("<%=pAb0018.ExportGpAAcctGpGpLvl(LngRow)%>") & "|" & cstr("<%=pAb0018.ExportGpAAcctGpGpSeq(LngRow)%>")
<%	
		Case "LISTGP"
%>	
			Set NodX = .uniTree1.Nodes.Add ("G" & "<%=ConvSPChars(pAb0018.ExportParentGpAAcctGpGpCd(LngRow))%>", tvwChild, "G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>", "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpNm(LngRow))%>", "<%=C_Folder%>")
			.uniTree1.Nodes("G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>").Tag = cstr("<%=pAb0018.ExportGpAAcctGpGpLvl(LngRow)%>") & "|" & cstr("<%=pAb0018.ExportGpAAcctGpGpSeq(LngRow)%>")
<%			
		Case "LISTACCT"
%>	
			Set NodX = .uniTree1.Nodes.Add ("G" & "<%=ConvSPChars(pAb0018.ExportAAcctGpGpCd(LngRow))%>", tvwChild, "A" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(LngRow))%>", "<%=ConvSPChars(pAb0018.ExportAAcctAcctNm(LngRow))%>", "<%=C_URL%>" )			
			.uniTree1.Nodes("A" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(LngRow))%>").Tag =  cstr("<%=pAb0018.ExportAAcctAcctSeq(LngRow)%>")
<%			
	END SELECT			
	'Response.Flush	
Next
Select Case strCmd
	Case "LISTTOP"
%>	
		If "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(GroupCount))%>" = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>" Then
			parent.lgQueryFlag = "1"
			parent.lgStrPrevKey = ""
			parent.lgStrPrevKey1 = 0	
			parent.lgStrPrevKey2 = 0
			parent.AddNodes(parent.C_CMD_GP_LEVEL)
		Else
			parent.lgQueryFlag = "0"			
			parent.lgStrPrevKey = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>"
			parent.AddNodes(parent.C_CMD_TOP_LEVEL)
		End If		
<%		
	Case "LISTGP"
%>		
		If "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(GroupCount))%>" = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>" AND _
			"<%=pAb0018.ExportGpAAcctGpGpLvl(GroupCount)%>" = "<%=pAb0018.ExportNextAAcctGpGpLvl%>" AND _
			"<%=pAb0018.ExportGpAAcctGpGpSeq(GroupCount)%>" = "<%=pAb0018.ExportNextAAcctGpGpSeq%>" THEN
			parent.lgQueryFlag = "1"			
			parent.lgStrPrevKey = ""
			parent.lgStrPrevKey1 = 0	
			parent.lgStrPrevKey2 = 0
			parent.AddNodes(parent.C_CMD_ACCT_LEVEL)			
		Else
			parent.lgQueryFlag = "0"			
			parent.lgStrPrevKey = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>"
			parent.lgStrPrevKey1 = "<%=pAb0018.ExportNextAAcctGpGpLvl%>"	
			parent.lgStrPrevKey2 = "<%=pAb0018.ExportNextAAcctGpGpSeq%>"			
			parent.AddNodes(parent.C_CMD_GP_LEVEL)
		End If		
<%		
	Case "LISTACCT"
%>	
		If "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(GroupCount))%>" = "<%=ConvSPChars(pAb0018.ExportNextAAcctAcctCd)%>" AND _			
			"<%=pAb0018.ExportAAcctAcctSeq(GroupCount)%>" = "<%=pAb0018.ExportNextAAcctAcctSeq%>" Then
			parent.lgQueryFlag = "1"			
			parent.lgStrPrevKey = ""	
			parent.lgStrPrevKey2 = ""
		Else
			parent.lgQueryFlag = "0"			
			parent.lgStrPrevKey = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>"	
			parent.lgStrPrevKey2 = "<%=pAb0018.ExportNextAAcctAcctSeq%>"
			parent.AddNodes(parent.C_CMD_ACCT_LEVEL)
		End If						
<%		
END SELECT		
%>
	.uniTree1.MousePointer = 0
End With
</Script>
<%
    Set pAb0018 = nothing    
%>
