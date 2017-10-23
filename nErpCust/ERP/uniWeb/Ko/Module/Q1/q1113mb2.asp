<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1113MB2
'*  4. Program Name         : 불량률단위등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG050
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/09
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd
	
Dim PQBG050																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim txtSpread
Dim iErrorPosition
	
	txtSpread = Request("txtSpread")	
		
	Set PQBG050 = Server.CreateObject("PQBG050.cQMtDfctRatUnitSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
		
	Call PQBG050.Q_MAINT_DFCT_RATIO_UNIT_SVR (gStrGlobalCollection, _
											  txtSpread, _
											  iErrorPosition)
		
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		If iErrorPosition <> "" Then	
			Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
			Set PQBG050 = Nothing
			Response.End
		End If
	End If	
		                                      '☜: Unload Comproxy
	Set PQBG050 = Nothing	      
                                
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