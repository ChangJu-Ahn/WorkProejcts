<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4600mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2002/01/02
'*  7. Modified date(Last)  : 2002/02/21
'*  8. Modifier (First)     : Park, BumSoo 
'*  9. Modifier (Last)      : Park, BumSoo 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim	rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim strConsumedDtFrom, strConsumedDtTo, strItemCd, strWcCd, strProdtOrderNo, strResourceCd, strResourceGroupCd
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i
Dim lgStrPrevKey	

Dim strTrackingNo
Dim strSoldToParty
Dim strSalesGrp
Dim strSoDtFrom
Dim strSoDtTo
Dim strDlvryDtFrom
Dim strDlvryDtTo
Dim strSoType

On Error Resume Next
Err.Clear																'☜: Protect system from crashing

	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000sab"	' LookUp Item
	UNISqlId(1) = "p4600pb11"	' LookUp Biz Partner
	UNISqlId(2) = "p4600pb12"	' LookUp Sales Group
	UNISqlId(3) = "p4600pb13"	' LookUp Sales Order Type
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtSoldToParty")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtSalesGrp")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtSoType")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtItemNm.value = ""
			parent.txtItemNm.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemNm.value = ""
		</Script>	
		<%
	End IF
	rs1.Close
	Set rs1 = Nothing

	' 거래처명 Display
	IF Request("txtSoldToParty") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtSoldToPartyNm.value = ""
			parent.txtSoldToPartyNm.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.txtSoldToPartyNm.value = "<%=ConvSPChars(rs2("bp_nm"))%>"
			</Script>	
			<%
		End If
	Else
		%>
		<Script Language=vbscript>
			parent.txtSoldToPartyNm.value = ""
		</Script>	
		<%
	End IF
	rs2.Close
	Set rs2 = Nothing
	
	' 영업그룹명 Display
	IF Request("txtSalesGrp") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("125400", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtSalesGrpNm.value = ""
			parent.txtSalesGrpNm.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.txtSalesGrpNm.value = "<%=ConvSPChars(rs3("sales_grp_nm"))%>"
			</Script>	
			<%
		End IF
	Else
		%>
		<Script Language=vbscript>
			parent.txtSalesGrpNm.value = ""
		</Script>	
		<%
	End IF
	rs3.Close
	Set rs3 = Nothing

	' 수주형태명 Display
	IF Request("txtSoType") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			Call DisplayMsgBox("201600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtSoTypeNm.value = ""
			parent.txtSoTypeNm.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.txtSoTypeNm.value = "<%=ConvSPChars(rs4("so_type_nm"))%>"
			</Script>	
			<%
		End IF
	Else
		%>
		<Script Language=vbscript>
			parent.txtSoTypeNm.value = ""
		</Script>	
		<%
	End IF
	rs4.Close
	Set rs4 = Nothing
	Set ADF = Nothing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 13)

	UNISqlId(0) = "p4600pb1"

	lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	strQryMode = Request("lgIntFlgMode")
	
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF		
		
	IF Trim(Request("txtSoldToParty")) = "" Then
	   strSoldToParty = "|"
	ELSE
	   strSoldToParty = FilterVar(UCase(Request("txtSoldToParty")), "''", "S")
	END IF

	IF Trim(Request("txtSalesGrp")) = "" Then
	   strSalesGrp = "|"
	ELSE
	   strSalesGrp = FilterVar(UCase(Request("txtSalesGrp")), "''", "S")
	END IF	

	IF Trim(Request("txtSoDtFrom")) = "" Then
	   strSoDtFrom = "|"
	ELSE
	   strSoDtFrom = " " & FilterVar(UNIConvDate(Request("txtSoDtFrom")), "''", "S") & ""
	END IF

	IF Trim(Request("txtSoDtTo")) = "" Then
	   strSoDtTo = "|"
	ELSE
	   strSoDtTo = " " & FilterVar(UNIConvDate(Request("txtSoDtTo")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtDlvryDtFrom")) = "" Then
	   strDlvryDtFrom = "|"
	ELSE
	   strDlvryDtFrom = " " & FilterVar(UNIConvDate(Request("txtDlvryDtFrom")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtDlvryDtTo")) = "" Then
	   strDlvryDtTo = "|"
	ELSE
	   strDlvryDtTo = " " & FilterVar(UNIConvDate(Request("txtDlvryDtTo")), "''", "S") & ""
	END IF

	IF Trim(Request("txtSoType")) = "" Then
	   strSoType = "|"
	ELSE
	   strSoType = " " & FilterVar(UCase(Request("txtSoType")), "''", "S") & ""
	END IF	

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNIValue(0, 2) = strTrackingNo
	UNIValue(0, 3) = strItemCd
	UNIValue(0, 4) = strSoldToParty
	UNIValue(0, 5) = strSalesGrp
	UNIValue(0, 6) = strSoDtFrom
	UNIValue(0, 7) = strSoDtTo
	UNIValue(0, 8) = strDlvryDtFrom	
	UNIValue(0, 9) = strDlvryDtTo
	UNIValue(0,10) = strSoType
	
	If Trim(Request("txtrdoflag")) = "O" Then
		UNIValue(0,11) = "" & FilterVar("Y", "''", "S") & " "
		UNIValue(0,12) = "|"
	ElseIf Trim(Request("txtrdoflag")) = "C" Then
		UNIValue(0,11) = "|"
		UNIValue(0,12) = "" & FilterVar("Y", "''", "S") & " "
	Else
		UNIValue(0,11) = "|"
		UNIValue(0,12) = "|"
	End If
	

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0,13) = "|"
		Case CStr(OPMD_UMODE) 
			UNIValue(0,13) =  lgStrPrevKey 
	End Select

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .vspdData.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS_D Then
%>			
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("so_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("so_type"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("so_type_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sold_to_party"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("so_dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dlvy_dt"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("so_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_in_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sales_grp"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sales_grp_nm"))%>"
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

	.lgStrPrevKey = "<%=Trim(rs0("tracking_no"))%>"		

	.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.hSoldToParty.value		= "<%=ConvSPChars(Request("txtSoldToParty"))%>"
	.hSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
	.hSoDtFrom.value		= "<%=UNIDateClientFormat(Request("txtSoDtFrom"))%>"
	.hSoDtTo.value			= "<%=UNIDateClientFormat(Request("txtSoDtTo"))%>"
	.hDlvryDtFrom.value		= "<%=UNIDateClientFormat(Request("txtDlvryDtFrom"))%>"
	.hDlvryDtTo.value		= "<%=UNIDateClientFormat(Request("txtDlvryDtTo"))%>"
	.hSoType.value			= "<%=Request("txtSoType")%>"
	.hrdoflag.value			= "<%=Request("txtrdoflag")%>"
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
