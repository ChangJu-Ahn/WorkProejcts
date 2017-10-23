<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1114MB2
'*  4. Program Name         : 검사분류별 불량률단위등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG080
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd 
	
Dim PQBG070																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strPlant
Dim txtSpread
Dim iErrorPosition
	
	If Trim(Request("hPlantCd")) <> "" Then 
		strPlant = Request("hPlantCd")	
	Else
		strPlant = Request("txtPlantCd")	
	End If
	
	txtSpread = Request("txtSpread")	

	Set PQBG070 = Server.CreateObject("PQBG070.cQMtDefRtInsClsSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	
	Call PQBG070.Q_MT_DFCT_RT_BY_INS_CLS_SVR (gStrGlobalCollection, _
											  strPlant, _
											  txtSpread, _
											  iErrorPosition)

'##############################################################################
	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "125000"	'공장이 존재하지 않습니다.
			If CheckSYSTEMError(Err,True) = True Then
				%>
				<Script Language=vbscript>			
					Parent.frm1.txtPlantNm.Value = ""
					Parent.frm1.txtPlantCd.Focus
				</Script>
				<%
				Set PQBG070 = Nothing
				Response.End
			End If
		Case Else
			If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
				If iErrorPosition <> "" Then	
					Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
					Set PQBG070 = Nothing
					Response.End
				End If
			End If	
	End Select
		
	Set PQBG070 = Nothing
'##############################################################################
		
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	Dim strHTML
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>
