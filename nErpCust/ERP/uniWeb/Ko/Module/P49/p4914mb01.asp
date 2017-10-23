<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4914mb01.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd

Dim pPP4G903

Dim lgIntFlgMode
Dim LngMaxRow

Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 

Dim iErrorPosition
Dim LngRow
Dim txtSpread
Dim txtSpread1
Dim txtSpread2
Dim txtSpread3
Dim txtSpread4
Dim txtSpread5

Dim strPlantCd
Dim strReportDt
Dim strWcCd

strPlantCd = UCase(Trim(Request("txtPlantCd")))
strReportDt = Trim(Request("txtprodDt"))
strWcCd = UCase(Trim(Request("txtWcCd")))

txtSpread = Request("txtSpread")		' Create, Update
txtSpread1 = Request("txtSpread1")		' Create, Update
txtSpread2 = Request("txtSpread2")		' Create, Update
txtSpread3 = Request("txtSpread3")		' Create, Update
txtSpread4 = Request("txtSpread4")		' Create, Update
txtSpread5 = Request("txtSpread5")		' Create, Update

'Response.Write "strPlantCd   : " & strPlantCd & "<P>"
'Response.Write "strReportDt   : " & strReportDt & "<P>"
'Response.Write "strWcCd   : " & strWcCd & "<P>"
'
'Response.Write "txtSpread  : " & txtSpread & "<P>"
'Response.Write "txtSpread1  : " & txtSpread1 & "<P>"
'Response.Write "txtSpread2  : " & txtSpread2 & "<P>"
'Response.Write "txtSpread3  : " & txtSpread3 & "<P>"
'Response.Write "txtSpread4  : " & txtSpread4 & "<P>"
'Response.Write "txtSpread5  : " & txtSpread5 & "<P>"


'-------------------------------------------------------------------------------
'	COM+ Action
'-------------------------------------------------------------------------------

Set pPP4G903 = Server.CreateObject("PP4G903.cPMngPrdtDailyReport")
                                            
If CheckSYSTEMError(Err,True) = True Then
   Response.End
End If

Call pPP4G903.P_MANAGE_PRDT_DAILY_REPORT(gStrGlobalCollection, strPlantCd, strReportDt, strWcCd, txtSpread, txtSpread1, txtSpread2, txtSpread3, txtSpread4, txtSpread5, iErrorPosition)


'Response.Write iErrorPosition & "<P>"

If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
	If iErrorPosition <> "" Then
'		Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
		Set pPP4G903 = Nothing
		Response.End
	End If
End If

'If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
'	Set pPP1G303 = Nothing															'☜: Unload Component
'	Response.End
'End If

Set pPP4G903 = Nothing
'Response.End

%>
<Script Language=vbscript>

Select Case parent.gSelframeFlg
	Case 1	'TAB1
		parent.DbSaveOk																		'☜: 화면 처리 ASP 를 지칭함 
	Case 2	'TAB2
		parent.DbSaveFormOk																	'☜: 화면 처리 ASP 를 지칭함 
End Select

</Script>
<%
' Server Side 로직은 여기서 끝남 
'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
'Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
'	Dim strHTML
'	If iLoc = I_INSCRIPT Then
'		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
'		Response.Write strHTML
'	ElseIf iLoc = I_MKSCRIPT Then
'		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
'		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
'		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
'		Response.Write strHTML
'	End If
'End Function
</Script>