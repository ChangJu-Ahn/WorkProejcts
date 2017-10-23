<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1209mb1_ko441.asp
'*  4. Program Name         : Routing Detail Query
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im, HyunSoo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData, strTemp
Dim TmpBuffer
Dim iTotalStr

ReDim TmpBuffer(0)

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 1)

UNISqlId(0) = "180000saa"

UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
Response.Write "</Script>" & vbCrLf

' Plant 명 Display      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)

	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End
Else
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("PLANT_NM")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs0.Close
	Set rs0 = Nothing
End If

Redim UNISqlId(0)
Redim UNIValue(0, 3)

UNISqlId(0) = "p1209mb1_ko441"
	
UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 1) = FilterVar(Request("txtRoutNo"), "''", "S")
	
UNIValue(0, 2) = FilterVar(Request("lgStrPrevKey"), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

If (rs1.EOF And rs1.BOF) Then
	Call DisplayMsgBox("181400", vbOKOnly, "", "", I_MKSCRIPT)
	rs1.Close
	Set rs1 = Nothing
	Set ADF = Nothing
	Response.End
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	Response.Write ".frm1.txtRoutNo.value = """ & ConvSPChars(rs1("ITEM_GROUP_CD")) & """" & vbCrLf
	Response.Write ".frm1.txtRoutingNm.value = """ & ConvSPChars(rs1("DESCRIPTION")) & """" & vbCrLf
	Response.Write ".frm1.txtRoutingNo.value = """ & ConvSPChars(rs1("ITEM_GROUP_CD")) & """" & vbCrLf
	Response.Write ".frm1.txtRoutingNm1.value = """ & ConvSPChars(rs1("DESCRIPTION")) & """" & vbCrLf
	Response.Write ".frm1.txtValidFromDt.text = """ & UNIDateClientFormat(rs1("VALID_FROM_DT")) & """" & vbCrLf	       		
	Response.Write ".frm1.txtValidToDt.text = """ & UNIDateClientFormat(rs1("VALID_TO_DT")) & """" & vbCrLf

	Response.Write ".frm1.txtCostCd.value = """ & ConvSPChars(rs1("COST_CD")) & """" & vbCrLf
	Response.Write ".frm1.txtCostNm.value = """ & ConvSPChars(rs1("COST_NM")) & """" & vbCrLf
	Response.Write ".frm1.txtALTRTVALUE.Text = """ & ConvSPChars(rs1("ALT_RT_VALUE")) & """" & vbCrLf

	if ConvSPChars(rs1("MAJOR_FLG")) = "Y" Then
		Response.Write ".frm1.rdoMajorRouting1.Checked = True " & vbCrLf
	 Else
		Response.Write ".frm1.rdoMajorRouting2.Checked = True " & vbCrLf
	End if

	If Not(rs1.EOF And rs1.BOF) Then
		
		If C_SHEETMAXROWS_D < rs1.RecordCount Then 
			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		Else
			ReDim TmpBuffer(rs1.RecordCount - 1)
		End If
		
		For iIntCnt = 0 To rs1.RecordCount - 1
			If iIntCnt < C_SHEETMAXROWS_D Then
			    strData = "" 
			    strData = strData & Chr(11) & ConvSPChars(rs1("OPR_NO"))			'C_OprNo
			    strData = strData & Chr(11) & ConvSPChars(rs1("WC_CD"))			'C_WcCd
			    strData = strData & Chr(11) & ""						'C_WcCdPopUp
			    strData = strData & Chr(11) & ConvSPChars(rs1("WC_NM"))			'C_WcNm
			    strData = strData & Chr(11) & ConvSPChars(UCase(rs1("JOB_CD")))		'C_JobCd
			    strData = strData & Chr(11) & ""						'C_JobNm

			    strData = strData & Chr(11) & UCase(rs1("INSIDE_FLG"))			'C_InsideFlg			
			    strTemp = Trim(UCase(rs1("Inside_Flg")))					'C_InsideFlg

			    If strTemp = "Y" Then
					strData = strData & Chr(11) & "사내"
				ElseIf strTemp = "N" Then
					strData = strData & Chr(11) & "외주"
				Else
					strData = strData & Chr(11) & ""
			    End If

			    strData = strData & Chr(11) & ConvSPChars(rs1("BP_CD"))			'C_BpCd
			    strData = strData & Chr(11) & ""						'C_BpPopup
			    strData = strData & Chr(11) & ConvSPChars(rs1("BP_NM"))			'C_BpNm
			    strData = strData & Chr(11) & ConvSPChars(rs1("CUR_CD"))			'C_CurCd
			    strData = strData & Chr(11) & ""						'C_CurPopup
			    strData = strData & Chr(11) & ConvSPChars(rs1("SUBCONTRACT_PRC"))		'C_SubconPrc
			    strData = strData & Chr(11) & ConvSPChars(rs1("TAX_TYPE"))			'C_TaxType
			    strData = strData & Chr(11) & ""						'C_TaxPopup
			    strData = strData & Chr(11) & ConvSPChars(rs1("MILESTONE_FLG"))		'C_MilestoneFlg

			    strData = strData & Chr(11) & UCase(rs1("ROUT_ORDER"))
			    strData = strData & Chr(11) & ""
					
		            strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
			    strData = strData & Chr(11) & Chr(12)
				
			    TmpBuffer(iIntCnt) = strData
				
			    rs1.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs1("OPR_NO") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs1("OPR_NO")) & """" & vbCrLf
		End If
	End If

	rs1.Close
	Set rs1 = Nothing

	Response.Write "If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then" & vbCrLf
		Response.Write ".initData(" & iLngMaxRows & " + 1)" & vbCrLf
		Response.Write ".DbQuery" & vbCrLf
	Response.Write "Else" & vbCrLf
		Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
		Response.Write ".frm1.hRoutNo.value = """ & ConvSPChars(Request("txtRoutNo")) & """" & vbCrLf

		Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf
	Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
