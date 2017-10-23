<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb1.asp
'*  4. Program Name         : Routing Detail Query
'*  5. Program Desc         :
'*  6. Component List       : DB Agent
'*  7. Modified date(First) : 2002/06/24
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Park, BumSoo 
'* 10. Modifier (Last)      : Hong Chang Ho 
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData, strTemp

Const C_SHEETMAXROWS = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(2)
Redim UNIValue(2, 1)

UNISqlId(0) = "180000saa"
UNISqlId(1) = "180000sab"
UNISqlId(2) = "180000saf"

UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(2, 0) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(2, 1) = FilterVar(Request("txtPlantCd"), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
Response.Write "</Script>" & vbCrLf

' Plant 명 Display      
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

' 품목명 Display
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
		
' 공장별품목 Display
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

	UNISqlId(0) = "p1205mb1h"
	
	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(0, 1) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(0, 2) = FilterVar(Request("txtRoutNo"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs3)
			  
	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("181300", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtRoutNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtRoutNo.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent" & vbCrLf
			Response.Write ".frm1.txtRoutNm.value = """ & ConvSPChars(rs3("DESCRIPTION")) & """" & vbCrLf
			Response.Write ".frm1.txtBomNo.value = """ & ConvSPChars(rs3("BOM_NO")) & """" & vbCrLf
			Response.Write ".frm1.txtValidFromDt.Text = """ & UNIDateClientFormat(rs3("VALID_FROM_DT")) & """" & vbCrLf
			Response.Write ".frm1.txtValidToDt.Text = """ & UNIDateClientFormat(rs3("VALID_TO_DT")) & """" & vbCrLf

'			If Len(rs3("ALT_RT_VALUE")) <> 0 Then
'				Response.Write ".frm1.txtALTRTVALUE.value = """ & ConvSPChars(rs3("ALT_RT_VALUE")) & """" & vbCrLf
'			Else
'				Response.Write ".frm1.txtALTRTVALUE.value = """ & 0 & """" & vbCrLf
'			End If
		
			If Trim(rs3("MAJOR_FLG")) = "Y" Then
				Response.Write ".frm1.rdoMajorRouting(0).checked = True" & vbCrLf
			Else
				Response.Write ".frm1.rdoMajorRouting(1).checked = True" & vbCrLf
			End IF
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
	End If

End If
	
Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "p1205mb1d"
	
UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
UNIValue(0, 3) = FilterVar(Request("txtRoutNo"), "''", "S")
	
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
	Response.End
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	strData = ""
	If Not(rs4.EOF And rs4.BOF) Then
		For iIntCnt = 0 To rs4.RecordCount - 1
			If iIntCnt < C_SHEETMAXROWS Then
				strData = strData & Chr(11) & ConvSPChars(rs4("OPR_NO"))									'C_OprNo
		        strData = strData & Chr(11) & ConvSPChars(rs4("WC_CD"))										'C_WcCd
		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("JOB_CD")))								'C_JobCd
				strData = strData & Chr(11) & ""															'C_JobNm
		        strTemp = UCase(rs4("INSIDE_FLG"))															'C_InsideFlg	
		        
		        If strTemp = "Y" Then
					strData = strData & Chr(11) & "사내"
				ElseIf strTemp = "N" Then
					strData = strData & Chr(11) & "외주"
				Else
					strData = strData & Chr(11) & ""
				End If

				strData = strData & Chr(11) & rs4("MFG_LT")
				strData = strData & Chr(11) & ConvToTimeFormat(rs4("QUEUE_TIME"))							'C_QueueTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("SETUP_TIME"))							'C_SetupTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("WAIT_TIME"))							'C_WaitTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("FIX_RUN_TIME"))								'C_RunTime
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("RUN_TIME"))								'C_RunTime
		        strData = strData & Chr(11) & UniConvNumberDBToCompany(rs4("RUN_TIME_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)		'C_RunTimeQty
		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("RUN_TIME_UNIT")))						'C_RunTimeUnit
		        strData = strData & Chr(11) & ConvToTimeFormat(rs4("MOVE_TIME"))							'C_MoveTime
				strData = strData & Chr(11) & ConvSPChars(rs4("OVERLAP_OPR"))								'C_OverlapOpr
		        strData = strData & Chr(11) & rs4("OVERLAP_LT")												'C_OverlapLt
		        strData = strData & Chr(11) & rs4("BP_CD")													'C_BpCd
		        strData = strData & Chr(11) & ConvSPChars(UCase(rs4("CUR_CD")))								'C_CurCd
		        'strData = strData & Chr(11) & UniConvNumberDBToCompany(rs4("SUBCONTRACT_PRC"), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
		        strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs4("SUBCONTRACT_PRC"), 0)
		        strData = strData & Chr(11) & ConvSPChars(rs4("TAX_TYPE"))									'C_TaxType
		        strData = strData & Chr(11) & UCase(rs4("MILESTONE_FLG"))									'C_MileStoneFlg
		        strData = strData & Chr(11) & ConvSPChars(rs4("ROUT_ORDER_DESC"))							'C_RoutOrder
		        strData = strData & Chr(11) & UNIDateClientFormat(rs4("VALID_FROM_DT"))						'C_ValidFromDt
		        strData = strData & Chr(11) & UNIDateClientFormat(rs4("VALID_TO_DT"))						'C_ValidToDt

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				rs4.MoveNext
			End If
		Next

		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & strData & """,""F""" & vbCrLf
		Response.Write "Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData1," & iLngMaxRows + 1 & "," & iLngMaxRows + iIntCnt + 1 & ",.C_CurCd,.C_UnitPriceOfOprSubcon, ""C"" ,""I"",""X"",""X"")" & vbCrLf		
		
		If rs4("OPR_NO") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs4("OPR_NO")) & """" & vbCrLf
		End If
	End If

	rs4.Close
	Set rs4 = Nothing

	Response.Write "If .frm1.vspdData1.MaxRows < .VisibleRowCnt(.frm1.vspdData1, 0) And .lgStrPrevKey <> """" Then" & vbCrLf
		Response.Write ".initData(" & iLngMaxRows & " + 1)" & vbCrLf
		Response.Write ".DbQuery" & vbCrLf
	Response.Write "Else" & vbCrLf
		Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
		Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
		Response.Write ".frm1.hRoutNo.value = """ & ConvSPChars(Request("txtRoutNo")) & """" & vbCrLf

		Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf
	Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

<Script Language = VBScript RUNAT = Server>
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

On Error Resume Next
Err.Clear
				
	If IVal = 0 Or Len(Trim(IVal)) = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</Script>


