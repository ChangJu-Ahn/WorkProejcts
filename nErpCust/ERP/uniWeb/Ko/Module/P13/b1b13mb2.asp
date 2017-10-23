<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b13mb2.asp
'*  4. Program Name         : 대체품목등록 
'*  5. Program Desc         :
'*  6. Component List       : +PP1G303.cBMngAltItem.B_MANAGE_ALTERNATIVE_ITEM_Svr 
'*  7. Modified date(First) : 2000/11/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

'Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
'Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

On Error Resume Next														'☜: 
Err.Clear																'☜: Protect system from crashing

Dim I1_b_item_cd
Dim I2_b_plant_cd
Dim IG0_import_group
Dim pPP1G303
Dim iErrorPosition

Dim intFlgMode

Call LoadBasisGlobalInf() 
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
 		
intFlgMode = CInt(Request("txtFlgMode"))
    
If intFlgMode = OPMD_CMODE Then
	I2_b_plant_cd = Trim(UCase(Request("txtPlantCd")))
	I1_b_item_cd= Trim(UCase(Request("txtItemCd")))
Else
	I2_b_plant_cd = Trim(UCase(Request("hPlantCd")))
	I1_b_item_cd= Trim(UCase(Request("hItemCd")))
End If
    
    
If I2_b_plant_cd = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)				'⊙:
	Response.End
End If
If I1_b_item_cd = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)					'⊙:
	Response.End
End If		
    
IG0_import_group = Request("txtSpread")
    
Set pPP1G303 = Server.CreateObject("PP1G303.cBMngAltItem")    
    
If CheckSYSTEMError(Err, True) = True Then
    
	Response.End 
End if

Call pPP1G303.B_MANAGE_ALTERNATIVE_ITEM_Svr  (gStrGlobalCollection, I1_b_item_cd, _
				I2_b_plant_cd, IG0_import_group)
    
If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
	Set pPP1G303 = Nothing															'☜: Unload Component
	Response.End
End If
	
Set pPP1G303 = Nothing
            
Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
%>
' Server Side 로직은 여기서 끝남 

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
<Script Language=vbscript RUNAT=server>

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
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