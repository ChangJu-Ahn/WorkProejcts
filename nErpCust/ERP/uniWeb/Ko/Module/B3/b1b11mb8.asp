<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b11mb8.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/11/13
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter ���� 
Dim rs0, rs1, rs2
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Dim strPlantCd
Dim strItemCd
Dim strItemAccunt
Dim strProcType
Dim strAvailableItem
Dim strItemGroupCd

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = FilterVar(Request("lgStrPrevKey1") , "''", "S")
iLngMaxRows = Request("txtMaxRows")

'======================================================================================================
'	ǰ���̸� ó�����ִ� �κ� 
'======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 0)

UNISqlId(0) = "122700sab"	
UNISqlId(1) = "122600sac"
UNISqlId(2) = "127400saa"	

strPlantCd = FilterVar(UCase(Request("txtPlantCd")) , "''", "S") 

IF Trim(Request("txtItemCd")) = "" Then
    strItemCd = "|"
ELSE
	strItemCd = FilterVar(UCase(Request("txtItemCd")) , "''", "S")	   
END IF

IF Trim(Request("txtItemGroupCd")) = "" Then
    strItemGroupCd = "|"
ELSE
	strItemGroupCd = FilterVar(UCase(Request("txtItemGroupCd"))  , "''", "S") 
END IF

UNIValue(0, 0) = strPlantCd	
UNIValue(1, 0) = strItemCd
UNIValue(2, 0) = strItemGroupCd

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing
	rs2.Close
	Set rs2 = Nothing
	Set ADF = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf					'��: ȭ�� ó�� ASP �� ��Ī�� 
		Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("plant_nm")) & """" & vbCrLf '��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "</Script>" & vbCrLf
End If

rs0.Close
Set rs0 = Nothing
      
If rs1.EOF And rs1.BOF Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf							'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs1("item_nm")) & """" & vbCrLf	'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "</Script>" & vbCrLf
End If

rs1.Close
Set rs1 = Nothing
	
If Not(rs2.EOF AND rs2.BOF) Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs2("item_group_nm")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	IF Trim(Request("txtItemGroupCd")) <> "" Then
		Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End 
	End If
End If

rs2.Close		
Set rs2 = Nothing
Set ADF = Nothing

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 8)
	
UNISqlId(0) = "B1B11MB8"
IF Trim(Request("txtItemCd")) = "" Then
   strItemCd = "|"
ELSE
   strItemCd = FilterVar(Request("txtItemCd")  , "''", "S")
END IF
		
IF Trim(Request("cboAccount")) = "" Then
   strItemAccunt = "|"
ELSE
   strItemAccunt = FilterVar(Request("cboAccount")  , "''", "S")
END IF
	
IF Trim(Request("cboProcType")) = "" Then
   strProcType = "|"
ELSE
   strProcType = FilterVar(Request("cboProcType")  , "''", "S")
END IF
	
IF Trim(Request("rdoAvailableItem")) = "A" Then
   strAvailableItem = "|"
ELSE
   strAvailableItem = FilterVar(Request("rdoAvailableItem")  , "''", "S")
END IF	
	
IF Trim(Request("txtItemGroupCd")) = "" Then
   strItemGroupCd = "|"
Else
   strItemGroupCd =   FilterVar(Request("txtItemGroupCd") , "''", "S")
END IF	

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	
Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 2) = strItemCd
	Case CStr(OPMD_UMODE) 
		UNIValue(0, 2) = iStrPrevKey
End Select
	
UNIValue(0, 3) = strItemAccunt
UNIValue(0, 4) = strProcType
UNIValue(0, 5) = strAvailableItem
UNIValue(0, 6) = FilterVar("P1001", "''", "S")
UNIValue(0, 7) = FilterVar("P1003", "''", "S")
IF Trim(Request("txtItemGroupCd")) = "" Then
	UNIValue(0,8) = "|"
Else
	UNIValue(0,8) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Request("txtItemGroupCd"), "''", "S") & " ))"
End IF
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing					
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	If Not(rs0.EOF And rs0.BOF) Then
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		Else
			ReDim TmpBuffer(rs0.RecordCount - 1)
		End If
	    For iIntCnt = 0 To rs0.RecordCount - 1 
			If iIntCnt < C_SHEETMAXROWS_D Then
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_CD"))			'1 ǰ�� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_NM"))			'2 ǰ��� 
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))			'3 �԰� 
				strData = strData & Chr(11) & ConvSPChars(rs0("BASIC_UNIT"))			'3 ���� 
				strData = strData & Chr(11) & rs0("MINOR_NM_ITEM_ACCT")				'4 ǰ����� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_GROUP_CD"))		'5 ǰ��׷� 
				strData = strData & Chr(11) & rs0("MINOR_NM_PROC_TYPE")				'6 ���ޱ���			
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("SS_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)			'7 ������� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("MAX_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'8 �ִ�������� 
				strData = strData & Chr(11) & rs0("DAMPER_MAX")		'36 DAMPER �ִ���� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("MIN_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'9 �ּҿ������� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("FIXED_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)	'10 ������������ 
				strData = strData & Chr(11) & rs0("LINE_NO")			'40 ���μ� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("ROUND_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'11 �ø��� 
				strData = strData & Chr(11) & rs0("REQ_ROUND_FLG")	'13 �ҿ䷮�ø����� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("SCRAP_RATE_MFG"), 0)	'14 ����ǰ��ҷ��� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("SCRAP_RATE_PUR"), 0)	'15 ����ǰ��ҷ��� 
				strData = strData & Chr(11) & rs0("INSPEC_LT_MFG")	'16	�����˻�L/T
				strData = strData & Chr(11) & rs0("INSPEC_LT_PUR")	'17 ���Ű˻�L/T
				strData = strData & Chr(11) & rs0("INV_CHECK_FLG")	'18 �������üũ 
				strData = strData & Chr(11) & rs0("INV_MGR")			'21 ������� 
				strData = strData & Chr(11) & rs0("INV_MGR")			'21 ������� 
				strData = strData & Chr(11) & rs0("MRP_MGR")			'22 MRP����� 
				strData = strData & Chr(11) & rs0("MRP_MGR")			'22 MRP����� 
				strData = strData & Chr(11) & rs0("PROD_MGR")		'23 �������� 
				strData = strData & Chr(11) & rs0("PROD_MGR")		'23 �������� 
				strData = strData & Chr(11) & rs0("MPS_MGR")			'24 MPS����� 
				strData = strData & Chr(11) & rs0("MPS_MGR")			'24 MPS����� 
				strData = strData & Chr(11) & rs0("INSPEC_MGR")		'25 �˻����� 
				strData = strData & Chr(11) & rs0("INSPEC_MGR")		'25 �˻����� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("STD_TIME"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'26 ǥ�� ST
				strData = strData & Chr(11) & rs0("VAR_LT")			'27 ����L/T
				strData = strData & Chr(11) & rs0("LOT_FLG")			'29 LOT �������� 
				strData = strData & Chr(11) & ConvSPChars(rs0("CAL_TYPE"))		'38 Į����Ÿ�� 
				strData = strData & Chr(11)
				
				strData = strData & Chr(11) & rs0("VALID_FLG")		'28 ��ȿǰ�� 
				
				strData = strData & Chr(11) & rs0("ATP_LT")			'35 ATP L/T
				strData = strData & Chr(11) & rs0("ETC_FLG1")		'41 ETC FLAG1
				strData = strData & Chr(11) & rs0("ETC_FLG2")		'42 ETC FLAG2
				
				strData = strData & Chr(11) & rs0("OVER_RCPT_FLG")	'19 ���԰���뿩�� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("OVER_RCPT_RATE"), 0)	'20 ���԰������ 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("DAMPER_MIN"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'37 DAMPER �ּҼ��� 
				strData = strData & Chr(11) & rs0("DAMPER_FLG")		'37 DAMPER ���� 
				strData = strData & Chr(11) & ConvSPChars(rs0("LOCATION"))		'38 location
				
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) = strData
				
				rs0.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("ITEM_CD") = Null Then
			Response.Write ".lgStrPrevKey1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey1 = """ & Trim(rs0("ITEM_CD")) & """" & vbCrLf
		End If
	End If

	rs0.Close
	Set rs0 = Nothing

	Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemAccunt.value = """ & Request("cboAccount") & """" & vbCrLf
	Response.Write ".frm1.hProcType.value = """ & Request("cboProcType") & """" & vbCrLf
	Response.Write ".frm1.hItemGroupCd.value = """ & ConvSPChars(Request("txtItemGroupCd")) & """" & vbCrLf
	Response.Write ".frm1.hAvailableItem.value = """ & Request("rdoAvailableItem") & """" & vbCrLf

	Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
