<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1502mb2.asp
'*  4. Program Name         : ManageResourceGroup 저장 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/07
'*  8. Modified date(Last)  : 2003/01/28
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd 

Dim oPP1G604

Dim lgIntFlgMode
Dim LngMaxRow
		
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 

Dim iErrorPosition
Dim txtSpread
Dim LngRow
Dim strPlantCd


txtSpread = Request("txtSpread")
strPlantCd = UCase(Trim(Request("txtPlantCd")))

Set oPP1G604 = Server.CreateObject("PP1G604.cPMngRsrcGrp")
    
If CheckSYSTEMError(Err,True) = True Then
   Response.End 
End If
    
Call oPP1G604.P_MANAGE_RESOURCE_GROUP(gStrGlobalCollection, _
									  strPlantCd, _
									  txtSpread, _
									  iErrorPosition)

If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
	If iErrorPosition <> "" Then	
		Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
		Set oPP1G604 = Nothing
		Response.End
	End If
End If
    
Set oPP1G604 = Nothing	
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
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
