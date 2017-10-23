<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4116mb2
'*  4. Program Name         : 
'*  5. Program Desc         : Insert, Delete, Update Production Order
'*  6. Comproxy List        : +PP4G124.cPconvProdOrdToPudSvr
'*  7. Modified date(First) : 2002/04/02
'*  8. Modified date(Last)  : 2002/08/19
'*  9. Modifier (First)     : Park, Bum Soo
'* 10. Modifier (Last)      : RYU, SUNG WON
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")																					'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

Dim oPP4G124
Dim strMode																			'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey
Dim lgStrPrevKey
Dim LngMaxRow
Dim LngRow          

Dim I1_B_Plant_Plant_Cd
Dim IG1_Import_Group
	Const C_IG1_I1_count = 0
	Const C_IG1_I2_prodt_order_no = 1
	Const C_IG1_I2_remark = 2
	Const C_IG1_I3_pr_no = 3
	Const C_IG1_I3_req_qty = 4
	Const C_IG1_I3_req_unit = 5
	Const C_IG1_I3_req_dt = 6
	Const C_IG1_I3_req_dept = 7
	Const C_IG1_I3_req_prsn = 8
	Const C_IG1_I3_dlvy_dt = 9
	Const C_IG1_I3_sl_cd = 10
	Const C_IG1_I3_base_req_qty = 11
	Const C_IG1_I3_base_req_unit = 12
	Const C_IG1_I3_pur_grp = 13
	Const C_IG1_I3_tracking_no = 14
	Const C_IG1_I3_pur_org = 15

Dim R1_Ief_Supplied_Count	'Return Value
Dim R2_P_Prodt_Order_No		'Return Value

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

On Error Resume Next

Err.Clear																		'☜: Protect system from crashing

	strMode = Request("txtMode")														'☜ : 현재 상태를 받음 
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
	itxtSpread = ""
             
	iCUCount = Request.Form("txtCUSpread").Count
	
	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount)
	
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next

	itxtSpread = Join(itxtSpreadArr,"")
	
	Dim arrCols, arrRows																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																		'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim lGrpCnt																			'☜: Group Count
	
	ReDim IG1_Import_Group(LngMaxRow - 1, C_IG1_I3_pur_org)
	
	arrRows = Split(itxtSpread, gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 

    For LngRow = 0 To LngMaxRow - 1

		arrCols = Split(arrRows(LngRow), gColSep)
		
		IG1_Import_Group(LngRow, C_IG1_I1_count)		  = Trim(arrCols(0))						'Row Number
		IG1_Import_Group(LngRow, C_IG1_I2_prodt_order_no) = UCase(Trim(arrCols(1)))
		IG1_Import_Group(LngRow, C_IG1_I2_remark)		  = Trim(arrCols(2))
		IG1_Import_Group(LngRow, C_IG1_I3_pr_no)		  = UCase(Trim(arrCols(3)))	'구매요청번호 
		IG1_Import_Group(LngRow, C_IG1_I3_req_qty)		  = UNIConvNum(arrCols(4),0)
		IG1_Import_Group(LngRow, C_IG1_I3_req_unit)		  = UCase(Trim(arrCols(5)))
		IG1_Import_Group(LngRow, C_IG1_I3_req_dt)		  = UNIConvDate(arrCols(6))
		IG1_Import_Group(LngRow, C_IG1_I3_req_dept)		  = UCase(Trim(arrCols(7)))
		IG1_Import_Group(LngRow, C_IG1_I3_req_prsn)		  = UCase(Trim(arrCols(8)))
		IG1_Import_Group(LngRow, C_IG1_I3_dlvy_dt)		  = UNIConvDate(arrCols(9))
		IG1_Import_Group(LngRow, C_IG1_I3_sl_cd)		  = UCase(Trim(arrCols(10)))
		IG1_Import_Group(LngRow, C_IG1_I3_tracking_no)	  = UCase(Trim(arrCols(14)))
		IG1_Import_Group(LngRow, C_IG1_I3_pur_org)	      = UCase(Trim(arrCols(15)))
		
		If LngRow >= 99 Or LngRow = LngMaxRow - 1 Then								'⊙: 100개를 Group으로, 나머지 일때 
			I1_B_Plant_Plant_Cd = UCase(Request("txtPlantCd"))
			Exit For
		End If
    Next
	
	Set oPP4G124 = Server.CreateObject("PP4G124.cPconvProdOrdToPudSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>	" & vbCr															
		Response.Write "	Call Parent.RemovedivTextArea" & vbCr
		Response.Write "</Script>					" & vbCr
        Response.End 
    End If
	
	Call oPP4G124.P_CONV_PROD_ORDER_TO_PUR(gStrGlobalCollection, _
										I1_B_Plant_Plant_Cd, _
										IG1_Import_Group, _
										R1_Ief_Supplied_Count, _
										R2_P_Prodt_Order_No)
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>	" & vbCr															
		Response.Write "	Call Parent.RemovedivTextArea" & vbCr
		Response.Write "</Script>					" & vbCr
		
		
		Call SheetFocus(R1_Ief_Supplied_Count,1,I_MKSCRIPT)
		Set oPP4G124 = Nothing       
		Response.End		
    End If

	Set oPP4G124 = Nothing
	
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