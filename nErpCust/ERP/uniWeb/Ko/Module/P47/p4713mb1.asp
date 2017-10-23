<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4713mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :	2001-12-10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     :  Park, Bumsoo
'*  9. Modifier (Last)      :  Kang, Seong Moon
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" --> 
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> 
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim	rs0, rs1, rs2, rs3, rs4, rs5
Dim strPlantCd, strItemCd, strWcCd, strProdtOrderNo, strSlCd, strTrackingNo

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Call HideStatusWnd


	Redim UNISqlId(5)
	Redim UNIValue(5, 7)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"
	UNISqlId(2) = "180000sab"
	UNISqlId(3) = "180000sac"
	UNISqlid(4) = "p4419mb1h" ' Order Data Check
	UNISqlId(5) = "p4713mb1"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(4, 1) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	
	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtConsumedDtFrom")) = "" Then
	   strConsumedDtFrom = "|"
	ELSE
	   strConsumedDtFrom = " " & FilterVar(UNIConvDate(Request("txtConsumedDtFrom")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtConsumedDtTo")) = "" Then
	   strConsumedDtTo = "|"
	ELSE
	   strConsumedDtTo = " " & FilterVar(UNIConvDate(Request("txtConsumedDtTo")), "''", "S") & ""
	END IF
	
	UNIValue(5, 0) = "^"
	UNIValue(5, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(5, 2) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	UNIValue(5, 3) = strProdtOrderNo
	UNIValue(5, 4) = strWcCd
	UNIValue(5, 5) = strItemCd
	UNIValue(5, 6) = strConsumedDtFrom
	UNIValue(5, 7) = strConsumedDtTo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs0)
	Set ADF = Nothing

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbcr
		Response.Write "parent.frm1.txtPlantCd.focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	rs1.Close
	Set rs1 = Nothing
	
	' 자원명 Display
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		Call DisplayMsgBox("181600", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """"" & vbcr
		Response.Write "parent.frm1.txtResourceCd.focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """ & ConvSPChars(rs2("description")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	rs2.Close
	Set rs2 = Nothing
	
	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbcr
			Response.Write "parent.frm1.txtItemCd.focus()" & vbcr
			Response.Write "</Script>" & vbcr
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs3("ITEM_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	End IF
	rs3.Close
	Set rs3 = Nothing

	' 작업장명 Display
	IF Request("txtWcCd") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbcr
			Response.Write "parent.frm1.txtWCCD.focus()" & vbcr
			Response.Write "</Script>" & vbcr
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs4("WC_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtWCNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	End IF
	rs4.Close
	Set rs4 = Nothing

	' Prodt Order No Check
	IF Request("txtProdtOrderNo") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			rs5.Close
			Set rs5 = Nothing
			Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtProdtOrderNo.focus()" & vbcr
			Response.Write "</Script>" & vbcr
			Response.End
		Else
			rs5.Close
			Set rs5 = Nothing
		End If
	End IF

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.DbQueryNotOk" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("consumed_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("consumed_time"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodt_order_qty"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("good_qty_in_order_unit"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("bad_qty_in_order_unit"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("release_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("real_start_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("order_status"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("order_status"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_type"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_type"))%>"			
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData1
	.ggoSpread.SSShowDataByClip iTotalStr

	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hResourceCd.value		= "<%=ConvSPChars(Request("txtResourceCd"))%>"
	.frm1.hProdtOrderNo.value	= "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hConsumedDtFrom.value	= "<%=ConvSPChars(Request("txtConsumedDtFrom"))%>"
	.frm1.hConsumedDtTo.value	= "<%=ConvSPChars(Request("txtConsumedDtTo"))%>"
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk LngMaxRow
End With	
</Script>	

<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</script>
