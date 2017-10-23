<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1503mb2.asp
'*  4. Program Name         : 자원별 Shift Save
'*  5. Program Desc         :
'*  6. Comproxy List        : +P15031ManageResourceOnShift
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2000/09/15
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf


Dim oPP1G610																	'☆ : 저장용 ComProxy Dll 사용 변수 

Dim StrNextKey											'⊙: 다음 값 
Dim lgStrPrevKey										'⊙: 이전 값 
Dim strMode												'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngMaxRow											'⊙: 현재 그리드의 최대Row
Dim intFlgMode
Dim LngRow

Dim IG1_Import_Group
Dim I1_Plant_Cd
Dim I2_Resource_Cd 

Dim arrCols, arrRows									'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus											'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt												'☜: Group Count
Dim strCode												'⊙: Lookup 용 리턴 변수 

Const C_Select_Char = 0
Const C_Shift_Cd = 1

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 
	intFlgMode = CInt(Request("txtFlgMode"))	
	LngMaxRow = CInt(Request("txtMaxRows"))									'☜: 최대 업데이트된 갯수 
	
	If intFlgMode = OPMD_CMODE Then
		I1_Plant_Cd = UCase(Trim(Request("txtPlantCd")))
		I2_Resource_Cd= UCase(Trim(Request("txtResourceCd")))
    Else
		I1_Plant_Cd = UCase(Trim(Request("hPlantCd")))
		I2_Resource_Cd= UCase(Trim(Request("hResourceCd")))
    End If
    
	arrRows = Split(Request("txtSpread"), gRowSep)							'☆: Spread Sheet 내용을 담고 있는 Element명 
	ReDim IG1_Import_Group(UBound(arrRows,1),1)

	For LngRow = 0 To LngMaxRow - 1 
		arrCols = Split(arrRows(LngRow), gColSep)

 		IG1_Import_Group(LngRow, C_Shift_Cd) = UCase(arrCols(2))
		IG1_Import_Group(LngRow, C_Select_Char) = UCase(Trim(arrCols(0)))

        If LngRow >= 99 Or LngRow = LngMaxRow - 1 Then						'⊙: 5개를 Group으로, 나머지 일때 
            Exit For
		End If
	Next

	Set oPP1G610 = Server.CreateObject("PP1G610.cPMngRsrcOnShift")    

	If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

	Call oPP1G610.P_MANAGE_RESOURCE_ON_SHIFT(gStrGlobalCollection, _
										IG1_Import_Group, _
										I2_Resource_Cd, _
										I1_Plant_Cd)

	If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G610 = Nothing		                                        '☜: Unload Comproxy DLL
       Response.End		
    End If

	Set oPP1G610 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr					
	
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

%>