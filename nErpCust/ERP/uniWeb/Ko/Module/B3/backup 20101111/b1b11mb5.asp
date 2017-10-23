<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b11mb5.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 2002/11/14
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'��: 

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Dim pPB6S101
Dim strPlantCd
Dim strItemCd
Dim strItemAccunt
Dim strProcType
Dim strStartDt
Dim strEndDt
Dim strAvailableItem

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = FilterVar(Request("lgStrPrevKey1") , "''", "S")
iLngMaxRows = Request("txtMaxRows")

'======================================================================================================
'	ǰ���̸� ó�����ִ� �κ� 
'======================================================================================================
Redim UNISqlId(1)
Redim UNIValue(1, 0)
	
UNISqlId(0) = "122600sac"
UNISqlId(1) = "122700sab"

strItemCd = FilterVar(Request("txtItemCd")  , "''", "S")
strPlantCd= FilterVar(UCase(Request("txtPlantCd")), "''", "S")   

	
UNIValue(0, 0) = strItemCd
UNIValue(1, 0) = strPlantCd
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
      
If rs0.EOF And rs0.BOF Then
	Response.Write "<Script Language = VBScript>" &vbCrLF
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language = VBScript>" &vbCrLF
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If
	
If rs1.EOF And rs1.BOF Then
	Response.Write "<Script Language = VBScript>" &vbCrLF
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language = VBScript>" &vbCrLF
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If

rs0.Close
rs1.Close
		
Set rs0 = Nothing
Set rs1 = Nothing

Set ADF = Nothing
	
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 7)
	
UNISqlId(0) = "122700saa"
IF Request("txtItemCd") = "" Then
   strItemCd = "|"
ELSE
   strItemCd = FilterVar(Request("txtItemCd")  , "''", "S")
END IF
		
IF Request("cboAccount") = "" Then
   strItemAccunt = "|"
ELSE
   strItemAccunt = FilterVar(Request("cboAccount")  , "''", "S")
END IF
	
IF Request("cboProcType") = "" Then
   strProcType = "|"
ELSE
   strProcType = FilterVar(Request("cboProcType")  , "''", "S")
END IF
	
IF Request("txtStartDt") = "" Then
   strStartDt = "|"
ELSE
   strStartDt = FilterVar(UniConvDate(Request("txtStartDt")) , "''", "S")
END IF
	
IF Request("txtEndDt") = "" Then
   strEndDt = "|"
ELSE
   strEndDt = FilterVar(UniConvDate(Request("txtEndDt")) , "''", "S")
END IF
	
IF Request("rdoAvailableItem") = "A" Then
   strAvailableItem = "|"
ELSE
   strAvailableItem = FilterVar(Request("rdoAvailableItem") , "''", "S")
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
UNIValue(0, 5) = strStartDt
UNIValue(0, 6) = strEndDt
UNIValue(0, 7) = strAvailableItem	
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
	rs0.Close
	Set rs0 = Nothing					
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
				strData = strData & Chr(11) & ConvSPChars(rs0("BASIC_UNIT"))			'3 ���� 
				strData = strData & Chr(11) & rs0("NM_ITEM_ACCT")	'4 item_acct		
				strData = strData & Chr(11) & rs0("NM_PROC_TYPE")	'5 ���ޱ���			
				strData = strData & Chr(11) & rs0("NM_PROD_ENV")		'6 ��������			
				strData = strData & Chr(11) & rs0("NM_ITEM_CLASS")	'7 �����ǰ��Ŭ���� 
				strData = strData & Chr(11) & rs0("PHANTOM_FLG")		'8 Phontom���� 
				strData = strData & Chr(11) & rs0("MPS_FLG")			'9 MPSǰ�� 
				strData = strData & Chr(11) & rs0("TRACKING_FLG")	'10 Tracking���� 
				strData = strData & Chr(11) & rs0("SINGLE_ROUT_FLG")	'11 �ܰ������� 
				strData = strData & Chr(11) & ConvSPChars(rs0("WORK_CENTER")) '12�۾��� 
				strData = strData & Chr(11) & rs0("VALID_FLG")		'13 ��ȿǰ�� 
				strData = strData & Chr(11) & rs0("NM_MPS_MGR")		'14 MPS����� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("STD_TIME"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'15 ǥ�� ST
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_FROM_DT"))	'16 ������ 
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_TO_DT"))	'17 ������ 
				strData = strData & Chr(11) & rs0("ORDER_TYPE") 		'18 MRP�������� 
				strData = strData & Chr(11) & rs0("NM_ORDER_RULE") 	'19 LotSizing
				strData = strData & Chr(11) & rs0("VAR_LT")			'20 ����L/T
				strData = strData & Chr(11) & rs0("ROUND_PERD")		'21 �ø��Ⱓ 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("FIXED_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)	'22 ������������ 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("MIN_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'23 �ּҿ������� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("MAX_MRP_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'24 �ִ�������� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("ROUND_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'25 �ø��� 
				strData = strData & Chr(11) & rs0("REQ_ROUND_FLG")	'26 �ҿ䷮�ø����� 
				strData = strData & Chr(11) & rs0("NM_MRP_MGR")		'27 MRP����� 
				strData = strData & Chr(11) & rs0("DAMPER_FLG")		'28 Damper���� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("DAMPER_MAX"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'29 Damper�ִ���� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("DAMPER_MIN"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'30 Damper�ּҼ��� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ORDER_UNIT_MFG"))	'31 ������������ 
				strData = strData & Chr(11) & rs0("ORDER_LT_MFG")	'32 ��������L/T
				strData = strData & Chr(11) & rs0("NM_PROD_MGR")		'33 �������� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("SCRAP_RATE_MFG"), 0)	'34 ����ǰ��ҷ��� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ORDER_UNIT_PUR"))	'35 ���ſ������� 
				strData = strData & Chr(11) & rs0("ORDER_LT_PUR")	'36 ���ſ���L/T
				strData = strData & Chr(11) & ConvSPChars(rs0("PUR_ORG"))		'37 �������� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("SCRAP_RATE_PUR"), 0)	'38 ����ǰ��ҷ��� 
				strData = strData & Chr(11) & ConvSPChars(rs0("MAJOR_SL_CD"))	'39 �԰�â�� 
				strData = strData & Chr(11) & ConvSPChars(rs0("NM_ISSUE_MTHD"))	'40 ����� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ISSUED_SL_CD"))	'41 ���â�� 
				strData = strData & Chr(11) & ConvSPChars(rs0("ISSUED_UNIT"))	'42 ������ 
				strData = strData & Chr(11) & rs0("LOT_FLG")			'43 Lot���� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("SS_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)			'44 ������� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("REORDER_PNT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'45 ������ 
				strData = strData & Chr(11) & rs0("INV_CHECK_FLG")	'46 �������üũ 
				strData = strData & Chr(11) & rs0("OVER_RCPT_FLG")	'47 ���԰���뿩�� 
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("OVER_RCPT_RATE"), 0)	'48 ���԰������ 
				strData = strData & Chr(11) & rs0("CYCLE_CNT_PERD")	'49 ���ǻ��ֱ� 
				strData = strData & Chr(11) & rs0("ABC_FLG")			'50 ǰ��ABC���� 
				strData = strData & Chr(11) & rs0("NM_INV_MGR")		'51 ������� 
				strData = strData & Chr(11) & rs0("RECV_INSPEC_FLG") '52 ���԰˻翩�� 
				strData = strData & Chr(11) & rs0("PROD_INSPEC_FLG") '53 �����˻翩�� 
				strData = strData & Chr(11) & rs0("FINAL_INSPEC_FLG")'54	�����˻翩�� 
				strData = strData & Chr(11) & rs0("SHIP_INSPEC_FLG") '55 ���ϰ˻翩�� 
				strData = strData & Chr(11) & rs0("INSPEC_LT_MFG")	'56	�����˻�L/T
				strData = strData & Chr(11) & rs0("INSPEC_LT_PUR")	'57 ���Ű˻�L/T
				strData = strData & Chr(11) & rs0("NM_INSPEC_MGR")	'58 �˻����� 
				strData = strData & Chr(11) & rs0("NM_PRC_CTRL_INCTR")	'59 �ܰ����� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("STD_PRC"), ggUnitCost.DecPoint,  ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)			'60 ǥ�شܰ� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("MOVING_AVG_PRC"), ggUnitCost.DecPoint,  ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)  '61 �̵���մܰ� 
				strData = strData & Chr(11) & rs0("LINE_NO")			'62 ���μ� 
				strData = strData & Chr(11) & UCase(rs0("NM_ORDER_FROM"))		'63 ������������ 
				strData = strData & Chr(11) & rs0("ATP_LT")			'ATP ����Ÿ�� 
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("CAL_TYPE")))	'Į����Ÿ�� 
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))	'ǰ��԰� 
				strData = strData & Chr(11) & ConvSPChars(rs0("Tracking_NO"))	'67 Tracking_NO
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
	Response.Write ".frm1.hStartDt.value = """ & UNIDateClientFormat(strStartDt) & """" & vbCrLf
	Response.Write ".frm1.hEndDt.value = """ & UNIDateClientFormat(strEndDt) & """" & vbCrLf
	Response.Write ".frm1.hAvailableItem.value = """ & Request("rdoAvailableItem") & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>