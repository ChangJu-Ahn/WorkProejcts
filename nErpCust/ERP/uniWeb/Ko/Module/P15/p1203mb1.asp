<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb1.asp
'*  4. Program Name         : Query Routing Detail
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2005/12/20
'*  9. Modifier (First)     : Im, HyunSoo
'* 10. Modifier (Last)      : Chen, Jae Hyun 
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData, strTemp
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
Redim UNISqlId(4)
Redim UNIValue(4, 1)

UNISqlId(0) = "180000saa"
UNISqlId(1) = "180000sab"
UNISqlId(2) = "180000saf"
UNISqlId(3) = "B1254MA804"
UNISqlId(4) = "180000sat"

UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(2, 0) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(2, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(3, 0) = FilterVar(Request("txtCostCd"), "''", "S")
UNIValue(4, 0) = FilterVar(Request("txtPlantCd"), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
Response.Write "</Script>" & vbCrLf

' Plant �� Display      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("PLANT_NM")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs0.Close
	Set rs0 = Nothing
End If


' GET OPR. COST FLAG    
If (rs4.EOF And rs4.BOF) Then
	Call DisplayMsgBox("180600", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs4.Close
	Set rs4 = Nothing
	Set ADF = Nothing
	Response.End
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.hOprCostFlag.value = """ & ConvSPChars(rs4("OPR_COST_FLAG")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs4.Close
	Set rs4 = Nothing
End If

' ǰ��� Display
If (rs1.EOF And rs1.BOF) Then
	Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtItemCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs1.Close
	Set rs1 = Nothing
	Set ADF = Nothing
	Response.End
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs1("ITEM_NM")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs1.Close
	Set rs1 = Nothing
End If
		
' ���庰ǰ�� Display
If (rs2.EOF And rs2.BOF) Then
	Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtItemCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs2.Close
	Set rs2 = Nothing
	Set ADF = Nothing
	Response.End
End If
rs2.Close
Set rs2 = Nothing


' Routing Header Display
If strQryMode = CStr(OPMD_CMODE) Then

	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "p1203mb1h"
	
	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(0, 1) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(0, 2) = FilterVar(Request("txtRoutingNo"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs3)
			  
	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("181300", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutingNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtRoutingNo.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent" & vbCrLf
			Response.Write ".frm1.txtRoutingNm.value = """ & ConvSPChars(rs3("DESCRIPTION")) & """" & vbCrLf
			Response.Write ".frm1.txtItemCd1.value = """ & ConvSPChars(rs3("ITEM_CD")) & """" & vbCrLf
			Response.Write ".frm1.txtItemNm1.value = """ & ConvSPChars(rs3("ITEM_NM")) & """" & vbCrLf
			Response.Write ".frm1.txtRoutingNo1.value = """ & ConvSPChars(rs3("ROUT_NO")) & """" & vbCrLf
			Response.Write ".frm1.txtRoutingNm1.value = """ & ConvSPChars(rs3("DESCRIPTION")) & """" & vbCrLf
			Response.Write ".frm1.txtBomNo.value = """ & ConvSPChars(rs3("BOM_NO")) & """" & vbCrLf
			Response.Write ".frm1.txtValidFromDt.Text = """ & UNIDateClientFormat(rs3("VALID_FROM_DT")) & """" & vbCrLf
			Response.Write ".frm1.txtValidToDt.Text = """ & UNIDateClientFormat(rs3("VALID_TO_DT")) & """" & vbCrLf
			Response.Write ".frm1.txtCostCd.value = """ & ConvSPChars(rs3("COST_CD")) & """" & vbCrLf
			Response.Write ".frm1.txtCostNm.value = """ & ConvSPChars(rs3("COST_NM")) & """" & vbCrLf
			If Len(rs3("ALT_RT_VALUE")) <> 0 Then
				Response.Write ".frm1.txtALTRTVALUE.value = """ & ConvSPChars(rs3("ALT_RT_VALUE")) & """" & vbCrLf
			Else
				Response.Write ".frm1.txtALTRTVALUE.value = """ & 0 & """" & vbCrLf
			End If
			
			If Trim(rs3("MAJOR_FLG")) = "Y" Then
				Response.Write ".frm1.rdoMajorRouting(0).checked = True" & vbCrLf
				Response.Write ".lgRdoOldVal = 1" & vbCrLf
			Else
				Response.Write ".frm1.rdoMajorRouting(1).checked = True" & vbCrLf
				Response.Write ".lgRdoOldVal = 2" & vbCrLf
			End IF
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
	End If

End If
	
Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "p1203mb1d"
	
UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(0, 3) = FilterVar(Request("txtRoutingNo"), "''", "S")
	
If iStrPrevKey <> "" Then
	UNIValue(0, 4) = FilterVar(iStrPrevKey, "''", "S")
Else
	UNIValue(0, 4) = "|"
End If

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs4)

If (rs4.EOF And rs4.BOF) Then
	Call DisplayMsgBox("181200", vbOKOnly, "", "", I_MKSCRIPT)
	rs4.Close
	Set rs4 = Nothing
	Set ADF = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.DbQueryOk(0)" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	If Not(rs4.EOF And rs4.BOF) Then
	
		'If C_SHEETMAXROWS_D < rs4.RecordCount Then 
		'	ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		'Else
			ReDim TmpBuffer(rs4.RecordCount - 1)
		'End If

		For iIntCnt = 0 To rs4.RecordCount - 1
			'If iIntCnt < C_SHEETMAXROWS_D Then
				strData = "" 
				strData = strData & Chr(11) & ConvSPChars(rs4("OPR_NO"))									'C_OprNo
		        strData = strData & Chr(11) & ConvSPChars(rs4("WC_CD"))										'C_WcCd
		        strData = strData & Chr(11) & ""															'C_WcPopup
		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("JOB_CD")))								'C_JobCd
				strData = strData & Chr(11) & ""															'C_JobNm
		        strTemp = UCase(rs4("INSIDE_FLG"))															'C_InsideFlg	
		        
		        If strTemp = "Y" Then
					strData = strData & Chr(11) & "�系"
				ElseIf strTemp = "N" Then
					strData = strData & Chr(11) & "����"
				Else
					strData = strData & Chr(11) & ""
				End If

				strData = strData & Chr(11) & rs4("MFG_LT")
				strData = strData & Chr(11) & ConvToTimeFormat(rs4("QUEUE_TIME"))							'C_QueueTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("SETUP_TIME"))							'C_SetupTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("WAIT_TIME"))							'C_WaitTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("Fix_Run_Time"))							'C_Fix_Run_Time
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("RUN_TIME"))								'C_RunTime
		        strData = strData & Chr(11) & UniConvNumberDBToCompany(rs4("RUN_TIME_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'C_RunTimeQty
		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("RUN_TIME_UNIT")))						'C_RunTimeUnit
				strData = strData & Chr(11) & ""															'C_UnitPopup  
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("MOVE_TIME"))							'C_MoveTime
				strData = strData & Chr(11) & ConvSPChars(rs4("OVERLAP_OPR"))								'C_OverlapOpr
		        strData = strData & Chr(11) & rs4("OVERLAP_LT")												'C_OverlapLt
		        strData = strData & Chr(11) & rs4("BP_CD")													'C_BpCd

		        strData = strData & Chr(11) & ""															'C_BpPopup
		        strData = strData & Chr(11) & ConvSPChars(rs4("BP_NM"))										'C_BpNm

		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("CUR_CD")))								'C_CurCd
		        strData = strData & Chr(11) & ""															'C_CurPopup
		        
		        strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs4("SUBCONTRACT_PRC"), 0)

		        strData = strData & Chr(11) & ConvSPChars(rs4("TAX_TYPE"))									'C_TaxType
		        strData = strData & Chr(11) & ""															'C_TaxPopup
		        strData = strData & Chr(11) & UCase(rs4("MILESTONE_FLG"))									'C_MileStoneFlg
		        strData = strData & Chr(11) & UCase(rs4("INSP_FLG"))											'C_InspFlg
		        strData = strData & Chr(11) & ConvSPChars(rs4("ROUT_ORDER_DESC"))							'C_RoutOrder
		        strData = strData & Chr(11) & UNIDateClientFormat(rs4("VALID_FROM_DT"))						'C_ValidFromDt
		        strData = strData & Chr(11) & UNIDateClientFormat(rs4("VALID_TO_DT"))						'C_ValidToDt
				strData = strData & Chr(11) & UCase(rs4("INSIDE_FLG"))										'C_HiddenInsideFlg	
				strData = strData & Chr(11) & ConvSPChars(UCase(rs4("ROUT_ORDER")))							'C_HiddenRoutOrder	

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				rs4.MoveNext
				
				TmpBuffer(iIntCnt) = strData
			'End If
		Next
		
		iTotalStr =  Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """,""F""" & vbCrLf
		Response.Write "Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRows + 1 & "," & iLngMaxRows + iIntCnt + 1 & ",.C_CurCd,.C_SubconPrc, ""C"" ,""I"",""X"",""X"")" & vbCrLf
		If rs4("OPR_NO") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs4("OPR_NO")) & """" & vbCrLf
		End If
	End If

	rs4.Close
	Set rs4 = Nothing
	
		Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
		Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
		Response.Write ".frm1.hRoutingNo.value = """ & ConvSPChars(Request("txtRoutingNo")) & """" & vbCrLf
		Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf

	'Response.Write "If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then" & vbCrLf
	'	Response.Write ".initData(" & iLngMaxRows & " + 1)" & vbCrLf
	'	Response.Write ".DbQuery" & vbCrLf
	'Response.Write "Else" & vbCrLf
	'	Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
	'	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	'	Response.Write ".frm1.hRoutingNo.value = """ & ConvSPChars(Request("txtRoutingNo")) & """" & vbCrLf
	'	Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf
	'Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>

<Script Language = VBScript RUNAT = Server>
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
on error resume next
err.Clear

	If iVal = 0 or Len(Trim(iVal)) = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(iVal Mod 3600)
		iTime = Fix(iVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	End If
End Function
</Script>