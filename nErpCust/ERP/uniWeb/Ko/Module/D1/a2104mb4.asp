<%
'**********************************************************************************************
'*  1. Module Name          : ���Ϸ� �����ϴ� �����޴��� DB���� ���� �����޴��� �ǵ����ش�.
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

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<Script Language=vbscript src="../../inc/incUni2KTV.vbs"></Script>
<Script LANGUAGE=VBScript>
Sub Document_onReadyStateChange()	
	parent.frm1.uniTree2.MousePointer = 0
End Sub
</Script>
<%
		
Call HideStatusWnd		

On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' ���� ���⿡ �ʿ��� ��� 
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
'Const C_UNDERBAR = "_"

Dim arrMenu, i, arrLine, strImg

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim intColCnt																'Popup�� Display�ؾ� �� columns�� 
Dim LngRow
Dim GroupCount    
Dim strCmd
Dim strKey
Dim IntLvl			
Dim intSeq
Dim intacct

Dim pAb0028

Set pAb0028 = Server.CreateObject("Ab0028.AListAcctClassSvr")

strCmd =  Request("strCmd")

IF Request("NextCd") = "" then
	strKey = ""	
else
	strKey =  Request("NextCd")
end if

if strCmd = "ACCTDIST" then
	IF Request("NextCd1") = "" then
		intacct = ""
	else
		intacct = Request("NextCd1")
	end if
else
	IF Request("NextCd1") = "" then
		IntLvl = 0
	else
		IntLvl = Cint(Request("NextCd1"))
	end if
end if

IF Request("NextCd2") = "" then
	intSeq = 0	
else
	intSeq = Cint(Request("NextCd2"))
end if

Select Case strCmd
    Case "LISTTOP"				
		pAb0028.CommandSent = "LISTTOP"         
		pAb0028.ImportAAcctClassTypeClassType = strKey
    Case "LIST"		
        pAb0028.CommandSent = "LIST"
%>        
        IF parent.lgQueryFlag = "1" Then
<%        
			pAb0028.ImportAAcctClassTypeClassType = strKey	
			pAb0028.ImportNextAAcctClassClassLvl = 0
			pAb0028.ImportNextAAcctClassClassSeq = 0
%>			
		Else
<%		
			pAb0028.ImportAAcctClassTypeClassType = strKey
			pAb0028.ImportNextAAcctClassClassLvl = IntLvl
			pAb0028.ImportNextAAcctClassClassSeq = IntSeq
%>			
		End IF 	    		
<%		
	Case "ACCTDIST"		
        pAb0028.CommandSent = "ACCTDIST"
%>        
        IF parent.lgQueryFlag = "1" Then
<%        
			pAb0028.ImportAAcctClassTypeClassType = strKey	
			pAb0028.ImportNextAAcctClassClassLvl = 0
			pAb0028.ImportNextAAcctClassClassSeq = 0
			pAb0028.ImportNextAAcctAcctCd = ""
%>			
		Else
<%		
			pAb0028.ImportAAcctClassTypeClassType = strKey
			pAb0028.ImportNextAAcctClassClassLvl = 0
			pAb0028.ImportNextAAcctClassClassSeq = IntSeq
			pAb0028.ImportNextAAcctAcctCd		 = intacct
%>			
		End IF 	    		
<%		
End Select


If Err.Number <> 0 Then	
   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
   Set pAb0028 = Nothing																	'��: ComProxy UnLoad
   Response.End																				'��: Process End
End If

pAb0028.ServerLocation = ggServerIP
pAb0028.Comcfg = gConnectionString
pAb0028.Execute

'-----------------------
'Com action result check area(DB,internal)
'-----------------------
If Not (pAb0028.OperationStatusMessage = MSG_OK_STR) Then
	Call DisplayMsgBox(pAb0028.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)

	Set pAb0028 = Nothing
	Response.End
End If

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then	
   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
   Set pAb0028 = Nothing																	'��: ComProxy UnLoad
   Response.End																				'��: Process End
End If

GroupCount = pAb0028.ExportGroupCount
%>
<Script Language=vbscript>
	Dim NodX	
	Dim StrTrim
	Dim StrPar
	Dim StrCon
	dIM LngMaxRow
		
	With parent.frm1
	LngMaxRow = .vspdData.MaxRows
<%
For LngRow = 1 To GroupCount	
	Select Case strCmd
		Case "LISTTOP"			
			StrTrim	= Trim(pAb0028.ExportItemAAcctClassClassNm(LngRow))
%>		
			
			Set NodX = .uniTree2.Nodes.Add (parent.C_USER_MENU_KEY, tvwChild, "K" & "<%=ConvSPChars(pAb0028.ExportItemAAcctClassClassCd(LngRow))%>" , "<%=StrTrim%>", "<%=C_Folder%>" )			
			.uniTree2.Nodes("K" & "<%=ConvSPChars(pAb0028.ExportItemAAcctClassClassCd(LngRow))%>").Tag = cstr("<%=pAb0028.ExportItemAAcctClassClassLvl(LngRow)%>") & "|" & cstr("<%=pAb0028.ExportItemAAcctClassClassSeq(LngRow)%>")

<%	
		Case "LIST"
			StrPar = "K" & pAb0028.ExportParentItemAAcctClassClassCd(LngRow)			
			StrTrim	= Trim(pAb0028.ExportItemAAcctClassClassNm(LngRow))
%>	
			Set NodX = .uniTree2.Nodes.Add ("<%=StrPar%>", tvwChild, "K" & "<%=ConvSPChars(pAb0028.ExportItemAAcctClassClassCd(LngRow))%>", "<%=StrTrim%>", "<%=C_Folder%>")
			.uniTree2.Nodes("K" & "<%=ConvSPChars(pAb0028.ExportItemAAcctClassClassCd(LngRow))%>").Tag = cstr("<%=pAb0028.ExportItemAAcctClassClassLvl(LngRow)%>") & "|" & cstr("<%=pAb0028.ExportItemAAcctClassClassSeq(LngRow)%>")
<%			
		Case "ACCTDIST"			
			IF Trim(pAb0028.ExportItemAAcctAcctCd(LngRow)) <> "" THEN				
				StrCon = "K" & pAb0028.ExportItemAAcctClassClassCd(LngRow) & "#" & pAb0028.ExportItemAAcctAcctCd(LngRow) 
				StrTrim	= Trim(pAb0028.ExportItemAAcctAcctNm(LngRow))			
%>		
				Set NodX = .uniTree2.Nodes.Add ("K" & "<%=ConvSPChars(pAb0028.ExportItemAAcctClassClassCd(LngRow))%>" , tvwChild, "<%=StrCon%>" , "<%=StrTrim%>", "<%=C_URL%>" )			
<%			END IF

	END SELECT	
%>	
	
<%			
	'Response.Flush	
Next
Select Case strCmd
	Case "LISTTOP"
%>	
		parent.frm1.txtClassType.value		= "<%=ConvSPChars(pAb0028.ExportAAcctClassTypeClassType)%>"
		parent.frm1.txtClassTypeNm.value	= "<%=ConvSPChars(pAb0028.ExportAAcctClassTypeClassTypeNm)%>"
		
		if  strData <> "" THEN
			parent.ggoSpread.Source = .vspdData 
			parent.ggoSpread.SSShowData strData
		END IF	
	
		parent.lgQueryFlag = "1"				
		parent.lgStrPrevKey1 = "0"	
		parent.lgStrPrevKey2 = "0"	
		parent.AddClassNodes(parent.C_CMD_LIST_LEVEL)					
<%	
	Case "LIST"
%>		
		if  strData <> "" THEN	
			parent.ggoSpread.Source = .vspdData 
			parent.ggoSpread.SSShowData strData
		END IF
						
<%		
		IF GroupCount = 0 Then
%>
			parent.lgQueryFlag = "1"			
			parent.lgStrPrevKey = ""
			parent.lgStrPrevKey1 = "0"	
			parent.lgStrPrevKey2 = "0"						
			parent.AddClassNodes(parent.C_CMD_LIST_DIST)								
<%		
		Else
%>
			If "<%=pAb0028.ExportItemAAcctClassClassLvl(GroupCount)%>" = "<%=pAb0028.ExportNextAAcctClassClassLvl%>" AND _ 
				"<%=pAb0028.ExportItemAAcctClassClassSeq(GroupCount)%>" = "<%=pAb0028.ExportNextAAcctClassClassSeq%>" THEN
				parent.lgQueryFlag = "1"			
				parent.lgStrPrevKey = ""
				parent.lgStrPrevKey1 = "0"	
				parent.lgStrPrevKey2 = "0"						
				parent.AddClassNodes(parent.C_CMD_LIST_DIST)								
			Else
				parent.lgQueryFlag = "0"			
				parent.lgStrPrevKey = strkey
				parent.lgStrPrevKey1 = "<%=pAb0028.ExportNextAAcctClassClassLvl%>"	
				parent.lgStrPrevKey2 = "<%=pAb0028.ExportNextAAcctClassClassSeq%>"			
				parent.AddClassNodes(parent.C_CMD_LIST_LEVEL)
			End If				
<%
		End IF	
	
	Case "ACCTDIST"
%>		
		if  strData <> "" THEN	
			parent.ggoSpread.Source = .vspdData 
			parent.ggoSpread.SSShowData strData
		END IF			
<%		
		IF GroupCount <> 0 Then		
%>
			If "<%=ConvSPChars(pAb0028.ExportItemAAcctAcctCd(GroupCount))%>" = "<%=ConvSPChars(pAb0028.ExportNextAAcctAcctCd)%>" AND _ 
				"<%=pAb0028.ExportItemAAcctClassClassSeq(GroupCount)%>" = "<%=pAb0028.ExportNextAAcctClassClassSeq%>" THEN
			
				parent.lgQueryFlag = "1"			
				parent.lgStrPrevKey = ""
				parent.lgStrPrevKey1 = "0"	
				parent.lgStrPrevKey2 = "0"						
				parent.GridQuery()				
			Else
				parent.lgQueryFlag = "0"			
				parent.lgStrPrevKey = strkey
				parent.lgStrPrevKey1 = "<%=ConvSPChars(pAb0028.ExportNextAAcctAcctCd)%>"	
				parent.lgStrPrevKey2 = "<%=pAb0028.ExportNextAAcctClassClassSeq%>"			
				parent.AddClassNodes(parent.C_CMD_LIST_DIST)
			End If						
<%		
		Else
%>	
			parent.GridQuery()					
<%							
		END IF	
END SELECT		
%>
	'.uniTree2.MousePointer = 0
End With

</Script>
<%
    Set pAb0028 = nothing    
%>
<Script Language=vbscript RUNAT=server>

'==============================================================================
' Function : TreeClear
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function TreeClear()
	
	
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.uniTree2.Nodes.Clear " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	
End Function
</Script>
