<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4111pb1.asp
'*  4. Program Name         : List Production Order Header (Query)
'*  5. Program Desc         :
'*  6. Comproxy List        : +P32118ListProdOrderHeader
'*  7. Modified date(First) : 2000/03/28
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3							'DBAgent Parameter 선언 
Dim strQryMode
Dim i

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

Dim strStartDt
Dim strEndDt
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strOrderType
Dim strOrderStatus
Dim strOrderStatus1, strOrderStatus2, strOrderStatus3
Dim strItemGroupCd
Dim strFlag

strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(Ucase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.txtItemNm.value = ""
	</Script>	
	<%    	

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
	Else
		rs1.Close
		Set rs1 = Nothing
	End If

	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
		
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		End If
	End IF
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.txtItemGroupNm.value = """ & ConvSPChars(rs3("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4111pb1"
	
	IF Request("txtFromDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF

	IF Request("txtToDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
	End IF
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtProdOrderNo") = "" Then
				strProdOrderNo = "|"
			Else
				strProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
			End If	
		Case CStr(OPMD_UMODE) 
			strProdOrderNo = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	End Select 

	IF Request("cboOrderType") = "" Then
		strOrderType = "|"
	Else
		strOrderType = " " & FilterVar(UCase(Request("cboOrderType")), "''", "S") & ""
	End IF

	IF Request("txtFromStstus") = "" Then
		strOrderStatus1 = "|"
	Else
		strOrderStatus1 = " " & FilterVar(UCase(Request("txtFromStstus")), "''", "S") & ""
	End IF

	IF Request("txtToStstus") = "" Then
		strOrderStatus2 = "|"
	Else
		strOrderStatus2 = " " & FilterVar(UCase(Request("txtToStstus")), "''", "S") & ""
	End IF
	
	IF Request("txtThirdStstus") = "" Then
		strOrderStatus3 = "|"
	Else
		strOrderStatus3 = " " & FilterVar(UCase(Request("txtThirdStstus")), "''", "S") & ""
	End IF

	IF strOrderStatus1 = "|" and strOrderStatus2 = "|" and strOrderStatus3 = "|" Then
		strOrderStatus = "|"
	Else
		strOrderStatus = ""
		strOrderStatus = "a.order_status in ( " & "" & strOrderStatus1 & "" & "," & "" & strOrderStatus2 & "" & "," & "" & strOrderStatus3 & "" & " ) "
	End If
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	,"''", "S") & " ))"
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strStartDt
	UNIValue(0, 3) = strEndDt
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo		
	UNIValue(0, 7) = strOrderType
	UNIValue(0, 8) = strOrderStatus
	UNIValue(0, 9) = strItemGroupCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

%>
<Script Language=vbscript>      
    Dim LngMaxRow                 
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr
	
    With parent												'☜: 화면 처리 ASP 를 지칭함 
		
 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
					
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
	
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Prodt_Order_Qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Prod_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Good_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Rcpt_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_Unit"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Order_Qty_In_Base_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Prod_Qty_In_Base_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Good_Qty_In_Base_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Rcpt_Qty_In_Base_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"

				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Rout_No"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"

				If "<%=rs0("Re_Work_Flg")%>" = "N" Then
					strData = strData & Chr(11) & "작업"
				Else
					strData = strData & Chr(11) & "재작업"
				End If

				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Plan_Order_No"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
			End If
		Next
%>
	
	iTotalStr = Join(TmpBuffer, "")
	
	.ggoSpread.Source = .vspdData
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey = "<%=Trim(rs0("Prodt_Order_No"))%>"
	
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	
	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then	 ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
		.InitData(LngMaxRow)
		.DbQuery
	Else
		.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.hProdFromDt.value	= "<%=Request("txtFromDt")%>"
		.hProdToDt.value	= "<%=Request("txtToDt")%>"
		.hOrderType.value	= "<%=ConvSPChars(Request("cboOrderType"))%>"
		.hFromStatus.value	= "<%=ConvSPChars(Request("txtFromStstus"))%>"
		.hToStatus.value	= "<%=ConvSPChars(Request("txtToStstus"))%>"
		.hTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.hitemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.DbQueryOk(LngMaxRow)
	End If  
    
    End With
</Script>	
