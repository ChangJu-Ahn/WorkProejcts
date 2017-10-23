<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1208mb1.asp
'*  4. Program Name         : List Routing Detail
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/03/22
'*  7. Modified date(Last)  : 2002/11/20
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0	, rs1 , rs2 , rs3, rs4				'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData

Dim TmpBuffer
Dim iTotalStr

Dim strWCCd
Dim strOprNo
Dim strFlag

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
Redim UNISqlId(3)
Redim UNIValue(3, 2)

UNISqlId(0) = "180000saa"
UNISqlId(1) = "180000sab"
UNISqlId(2) = "180000sac"
UNISqlId(3) = "p1205mb1h"

UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(2, 0) = FilterVar(Request("txtWcCd"), "''", "S")
UNIValue(3, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(3, 1) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(3, 2) = FilterVar(Request("txtRoutNo"), "''", "S")
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
	Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
	Response.Write "parent.frm1.txtRoutingNm.value = """"" & vbCrLf
Response.Write "</Script>" & vbCrLf

' �۾���� Display
IF Request("txtWcCd") <> "" Then
	If (rs3.EOF And rs3.BOF) Then
		rs3.Close
		Set rs3 = Nothing
		strFlag = "ERROR_WCCD"
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtWcNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtWcNm.value = """ & ConvSPChars(rs3("WC_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
	End If
End IF

' ����� �� Display      
If (rs4.EOF And rs4.BOF) Then
	rs4.Close
	Set rs4 = Nothing
	strFlag = "ERROR_ROUT"
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutingNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutingNm.value = """ & ConvSPChars(rs4("DESCRIPTION")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs4.Close
	Set rs4 = Nothing
End If

' ǰ��� Display
IF Request("txtItemCd") <> "" Then
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		strFlag = "ERROR_ITEM"
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs2.Close
		Set rs2 = Nothing
	End If
End IF

' Plant �� Display
If (rs1.EOF And rs1.BOF) Then
	rs1.Close
	Set rs1 = Nothing
	strFlag = "ERROR_PLANT"
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs1.Close
	Set rs1 = Nothing
End If

If strFlag <> "" Then
	If strFlag = "ERROR_PLANT" Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	ElseIf strFlag = "ERROR_ITEM" Then
		Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	ElseIf strFlag = "ERROR_WCCD" Then		
		Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtWcNm.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	ElseIf strFlag = "ERROR_ROUT" Then		
		Call DisplayMsgBox("181300", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutingNm.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	End If
End IF

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 7)

UNISqlId(0) = "P1208MB1"
	
IF Request("txtWCCd") = "" Then
			strWCCd = "|"
Else
	StrWCCd = FilterVar(UCase(Request("txtWCCd")), "''", "S")
End IF

Select Case strQryMode
	Case CStr(OPMD_CMODE)
		IF Request("txtOprNo") = "" Then
			strOprNo = "|"
		Else
			StrOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
		End IF	
	Case CStr(OPMD_UMODE) 
		StrOprNo = FilterVar(UCase(iStrPrevKey), "''", "S")
End Select 
		
UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(0, 3) = FilterVar(Request("txtRoutNo"), "''", "S")
UNIValue(0, 4) = strOprNo
UNIValue(0, 5) = strWCCd
UNIValue(0, 6) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
UNIValue(0, 7) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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
				strData = strData & Chr(11) & ConvSPChars(rs0("OPR_NO"))			'���� 
				strData = strData & Chr(11) & ConvSPChars(rs0("WC_CD"))				'�۾��� 
				strData = strData & Chr(11) & ConvSPChars(rs0("WC_NM"))				'�۾���� 
				strData = strData & Chr(11) & ConvSPChars(rs0("JOB_CD"))			'�����۾� 
				strData = strData & Chr(11) & ConvSPChars(rs0("JOB_NM"))			'�����۾��� 
				If ConvSPChars(UCase(rs0("INSIDE_FLG"))) = "Y" Then
					strData = strData & Chr(11) & "�系"
				ElseIf ConvSPChars(UCase(rs0("INSIDE_FLG"))) = "N" Then
					strData = strData & Chr(11) & "����"
				Else
					strData = strData & Chr(11) & ""
				End If
				strData = strData & Chr(11) & ConvSPChars(rs0("MFG_LT"))					'����LT
				strData = strData & Chr(11) & ConvToTimeFormat(rs0("QUEUE_TIME"))			'Queue �ð� 
				strData = strData & Chr(11) & ConvToTimeFormat(rs0("SETUP_TIME"))			'Setup �ð� 
				strData = strData & Chr(11) & ConvToTimeFormat(rs0("WAIT_TIME"))			'Wait �ð� 
				If Trim(CStr(rs0("FIX_RUN_TIME"))) = "" Then
					strData = strData & Chr(11) & ConvToTimeFormat(0)						'C_FixRunTime
				Else
					strData = strData & Chr(11) & ConvToTimeFormat(rs0("FIX_RUN_TIME"))		'C_FixRunTime
				End If
				strData = strData & Chr(11) & ConvToTimeFormat(rs0("RUN_TIME"))				'run �ð� 
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("RUN_TIME_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'C_RunTimeQty
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("RUN_TIME_UNIT")))		'���ش��� 
				strData = strData & Chr(11) & ConvToTimeFormat(rs0("MOVE_TIME"))			'move time
				strData = strData & Chr(11) & ConvSPChars(rs0("OVERLAP_OPR"))				'overlap opr
				strData = strData & Chr(11) & ConvSPChars(rs0("OVERLAP_LT"))				'overlap leadtime
				strData = strData & Chr(11) & ConvSPChars(rs0("BP_CD"))						'business partner
				strData = strData & Chr(11) & ConvSPChars(rs0("BP_NM"))						'business partner name		'��ȿ������			
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("CUR_CD")))		
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("SUBCONTRACT_PRC"), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("MILESTONE_FLG")))		'milestone flag
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("INSP_FLG")))			'�˻����� 
				strData = strData & Chr(11) & ConvSPChars(UCase(rs0("ROUT_ORDER")))			'�����ܰ� 
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_FROM_DT"))		'��ȿ������ 
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_TO_DT"))		'��ȭ 
				strData = strData & Chr(11) & ConvSPChars(rs0("TAX_TYPE"))					'tax type
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				rs0.MoveNext
				TmpBuffer(iIntCnt) = strData
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """,""F""" & vbCrLf
		Response.Write "Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData1," & iLngMaxRows + 1 & "," & iLngMaxRows + iIntCnt & ",.C_CurCd,.C_SubcontractPrc, ""C"" ,""I"",""X"",""X"")" & vbCrLf		
		If rs0("OPR_NO") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs0("OPR_NO")) & """" & vbCrLf
		End If
	End If

	rs0.Close
	Set rs0 = Nothing

	Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hRoutNo.value = """ & ConvSPChars(Request("txtRoutNo")) & """" & vbCrLf
	Response.Write ".frm1.hOprNo.value = """ & ConvSPChars(Request("txtoprNo")) & """" & vbCrLf
	Response.Write ".frm1.hWcCd.value = """ & ConvSPChars(Request("txtWCCd")) & """" & vbCrLf
	Response.Write ".frm1.hStdDt.value = """ & Request("txtStdDt") & """" & vbCrLf
			
	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf	
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>

<Script Language = VBScript RUNAT = Server>
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
