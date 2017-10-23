<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4714mb1.asp
'*  4. Program Name         : 자원소비실적조회(오더별)
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/12/04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeon, JaeHyun 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4								'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS_D = 50

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim strTemp
Dim i

'@Var_Declare

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strFromDt
Dim strToDt
Dim StrProdOrderNo
Dim StrWcCd
Dim StrItemCd
Dim StrTrackingNo


lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
lgStrPrevKey3 = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
lgStrPrevKey4 = " " & FilterVar(UNIConvDate(Request("lgStrPrevKey4")), "''", "S") & ""

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sam"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	Set ADF = Nothing
	
	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If

	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtItemCd.Focus()" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	' 작업장명 Display
	IF Request("txtWcCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtWCCD.focus()" & vbcr
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs3("WC_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	' Tracking No Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtTrackingNo.Focus()" & vbCrLf
			Response.Write "</Script>	" & vbCrLf
			Response.End
		Else
			rs4.Close
			Set rs4 = Nothing
		End If
	End IF
		        
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 8)

	UNISqlId(0) = "189754SAA"
	
	IF Request("txtFromDt") = "" Then
		strFromDt = "|"
	Else
		strFromDt = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF
	
	IF Request("txtToDt") = "" Then
		strToDt = "|"
	Else
		strToDt = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
	End IF

	Select Case strQryMode
		Case CStr(OPMD_CMODE)		
			IF Request("txtProdOrderNo") = "" Then
				StrProdOrderNo = "|"
			Else
				StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
			End IF
			
		Case CStr(OPMD_UMODE) 
			StrProdOrderNo = "|"
	End Select

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strFromDt
	UNIValue(0, 7) = strToDt
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 8) = "|"
		Case CStr(OPMD_UMODE) 
		
			 strTemp = ""
			 strTemp = "(a.prodt_order_no > " & lgStrPrevKey
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey  'second condition  for group view
			 strTemp = strTemp  & " and a.opr_no > " & lgStrPrevKey2 & ") "  'second condition  for group view
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey	'third condition  for group view
			 strTemp = strTemp  & " and a.opr_no = " & lgStrPrevKey2 		'third condition  for group view
			 strTemp = strTemp  & " and a.resource_cd > " & lgStrPrevKey3 & ") "  'third condition  for group view
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey  'Forth condition  for group view
			 strTemp = strTemp  & " and a.opr_no = " & lgStrPrevKey2			'Forth condition  for group view
			 strTemp = strTemp  & " and a.resource_cd = " & lgStrPrevKey3  'Forth condition  for group view
			 strTemp = strTemp  & " and a.consumed_dt >= " & lgStrPrevKey4 & ")) " 'Forth condition  for group view
			UNIValue(0, 8) = strTemp
	End Select
		
	UNILock = DISCONNREAD :	UNIFlag = "1"	 
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing
    
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																						'☜: 화면 처리 ASP 를 지칭함 

	LngMaxRow = .frm1.vspdData.MaxRows															'Save previous Maxrow
		
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
			If i < C_SHEETMAXROWS_D THEN
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"	        '제조오더 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"					'공정코드 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"			'자원코드 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"			'자원코드명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM_RESOURCE_TYPE"))%>"	'자원구분 
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("CONSUMED_DT"))%>"	'자원소비일 
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("CONSUMED_TIME"))%>"		'자원소비시간 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>" '실적수량 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"	'단위							
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'양품수량 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("BAD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'불량수량 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>"		'자원그룹코드 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_NM"))%>"			'자원그룹명			
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"				'품목 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"				'품목명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"				'품목명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ROUT_NO"))%>"				'라우팅 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"					'작업장 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"					'작업장명			
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"			'Tracking No.
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
				
			END IF
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=ConvSPChars(Trim(rs0("PRODT_ORDER_NO")))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(Trim(rs0("OPR_NO")))%>"
		.lgStrPrevKey3 = "<%=ConvSPChars(Trim(rs0("RESOURCE_CD")))%>"
		.lgStrPrevKey4 = "<%=UniDateClientFormat(rs0("CONSUMED_DT"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hWcCd.value = "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hFromDt.value = "<%=ConvSPChars(Request("txtFromDt"))%>"
		.frm1.hToDt.value = "<%=ConvSPChars(Request("txtToDt"))%>"
		     
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbQueryOk

End With

</Script>	

<script Language = vbscript RUNAT = server>
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
