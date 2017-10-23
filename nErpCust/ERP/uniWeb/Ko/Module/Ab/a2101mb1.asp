
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
pAb0018.ComCfg = gConnectionString
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
End if 	

%>
<Script Language=vbscript>
	Dim NodX
	With parent.frm1
<%
For LngRow = 1 To GroupCount	
	Select Case strCmd
		Case "LISTTOP"
%>			
			Set NodX = .uniTree1.Nodes.Add (parent.C_USER_MENU_KEY, tvwChild, "G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>", "[" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>" & "]" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpNm(LngRow))%>", C_Folder, C_Open )			
			.uniTree1.Nodes("G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>").Tag = cstr("<%=pAb0018.ExportGpAAcctGpGpLvl(LngRow)%>") & "|" & cstr("<%=pAb0018.ExportGpAAcctGpGpSeq(LngRow)%>")
<%	
		Case "LISTGP"
%>	
			Set NodX = .uniTree1.Nodes.Add ("G" & "<%=ConvSPChars(pAb0018.ExportParentGpAAcctGpGpCd(LngRow))%>", tvwChild, "G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>","[" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>" & "]" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpNm(LngRow))%>",  C_Folder, C_Open )		
			.uniTree1.Nodes("G" & "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(LngRow))%>").Tag = cstr("<%=pAb0018.ExportGpAAcctGpGpLvl(LngRow)%>") & "|" & cstr("<%=pAb0018.ExportGpAAcctGpGpSeq(LngRow)%>")
<%			
		Case "LISTACCT"
%>	
			Set NodX = .uniTree1.Nodes.Add ("G" & "<%=ConvSPChars(pAb0018.ExportAAcctGpGpCd(LngRow))%>", tvwChild, "A" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(LngRow))%>", "[" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(LngRow))%>" & "]" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctNm(LngRow))%>",  C_URL, C_URL )		
			.uniTree1.Nodes("A" & "<%=ConvSPChars(pAb0018.ExportAAcctAcctCd(LngRow))%>").Tag =  cstr("<%=pAb0018.ExportAAcctAcctSeq(LngRow)%>")
<%			
	END SELECT		
	Response.Flush	
Next
Select Case strCmd
	Case "LISTTOP"
	
%>	
		If "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(GroupCount))%>" = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>" Then
			parent.lgQueryFlag = "1"
			parent.lgStrPrevKey = ""
			parent.lgStrPrevKey1 = ""
			parent.lgStrPrevKey2 = ""
			parent.AddNodes(parent.C_CMD_GP_LEVEL)
		Else
			parent.lgQueryFlag = "0"
			'Response.Write "lgQueryFlag : " & parent.lgQueryFlag & "<br>"
			parent.lgStrPrevKey = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>"
			parent.AddNodes(parent.C_CMD_TOP_LEVEL)
		End If		
		' 트리의 각 노드를 확장 
		For intColCnt = 1 To .uniTree1.Nodes.count
			  .uniTree1.Nodes(intColCnt).Expanded = True
		Next		

<%		
	Case "LISTGP"
%>		
		If "<%=ConvSPChars(pAb0018.ExportGpAAcctGpGpCd(GroupCount))%>" = "<%=ConvSPChars(pAb0018.ExportNextAAcctGpGpCd)%>" AND _
			"<%=pAb0018.ExportGpAAcctGpGpLvl(GroupCount)%>" = "<%=pAb0018.ExportNextAAcctGpGpLvl%>" AND _
			"<%=pAb0018.ExportGpAAcctGpGpSeq(GroupCount)%>" = "<%=pAb0018.ExportNextAAcctGpGpSeq%>" THEN
			parent.lgQueryFlag = "1"			
			parent.lgStrPrevKey = ""
			parent.lgStrPrevKey1 = ""	
			parent.lgStrPrevKey2 = ""
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

			.uniTree1.MousePointer = 0
			parent.DisplayAcctOK

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
    Set oAp0018 = nothing    
%>

