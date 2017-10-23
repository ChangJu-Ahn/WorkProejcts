
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
'*  3. Program ID           : p4815mb1.asp
'*  4. Program Name         : Called By p4815ma1
'*  5. Program Desc         : Daily Mfg. Plan & Prod. results
'*  6. Modified date(First) : 2003/03/10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim strFlag
Dim TmpBuffer
Dim iTotalStr
Dim strItemCd
Dim strItemGroupCd

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey1")
iLngMaxRows = Request("txtMaxRows")

	'=======================================================================================================
	'	Handle Description
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sas"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
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
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs3("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
		
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
		Set ADF = Nothing
		Response.End
	End IF
	Set ADF = Nothing
	'=======================================================================================================
	'	Main Query
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)

	UNISqlId(0) = "p4815mb1a"
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF
		
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
		
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("cboYear")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("cboMonth")), "''", "S")
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 4) = strItemCd
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 4) = FilterVar(iStrPrevKey, "''", "S")
	End Select
	
	UNIValue(0, 5) = strItemGroupCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT) 	
		rs0.Close
		Set rs0 = Nothing					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
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
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))
				strData = strData & Chr(11) & ConvSPChars(rs0("BASIC_UNIT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("item_group_cd"))
				strData = strData & Chr(11) & ConvSPChars(rs0("item_group_nm"))
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_1"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_1"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_2"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_2"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_3"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_3"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_4"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_4"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_5"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_5"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_6"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_6"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_7"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_7"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_8"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_8"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_9"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_9"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_10"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_10"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_11"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_11"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_12"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_12"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_13"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_13"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_14"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_14"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_15"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_15"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_16"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_16"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_17"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_17"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_18"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_18"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_19"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_19"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_20"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_20"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_21"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_21"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_22"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_22"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_23"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_23"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_24"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_24"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_25"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_25"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_26"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_26"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_27"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_27"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_28"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_28"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_29"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_29"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_30"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_30"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("MFG_31"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs0("PROD_31"),ggQty.DecPoint,0)		

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
	Response.Write ".frm1.hcboYear.value = """ & ConvSPChars(Request("cboYear")) & """" & vbCrLf
	Response.Write ".frm1.hcboMonth.value = """ & ConvSPChars(Request("cboMonth")) & """" & vbCrLf
	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemGroupCd.value = """ & ConvSPChars(Request("txtItemGroupCd")) & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

