<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3111MB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : PI3G111
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/05/25
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : KiHong Han
'* 11. Comment
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change" 
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")   
Call HideStatusWnd

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4

Const C_SHEETMAXROWS_D = 100

Dim strPlantCd
Dim strPrevNextFlag
Dim strKeyVal

Err.Clear

	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "I3111MB1"
	
	strPlantCd = UCase(Request("txtPlantCd"))
	strPrevNextFlag = Request("PrevNextFlg")
	
	Select Case strPrevNextFlag
		Case "P"
			strKeyVal = " WHERE A.PLANT_CD < " & FilterVar(strPlantCd,"''","S") _
					  & " ORDER BY A.PLANT_CD "
		Case "N"
			strKeyVal = " WHERE A.PLANT_CD > " & FilterVar(strPlantCd,"''","S") _
					  & " ORDER BY A.PLANT_CD "
		Case Else
			strKeyVal = " WHERE A.PLANT_CD = " & FilterVar(strPlantCd,"''","S")
	End Select

	UNIValue(0, 0) = strKeyVal
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>

<Script Language=vbscript>
Dim strData

With Parent.frm1
	.txtPlantCd1.value					= "<%=ConvSPChars(Trim(rs0("PLANT_CD")))%>"
	.txtPlantNm1.value					= "<%=ConvSPChars(Trim(rs0("PLANT_NM")))%>"
	.txtPlantCd2.Value					= "<%=ConvSPChars(Trim(rs0("PLANT_CD")))%>"
	.txtPlantNm2.Value					= "<%=ConvSPChars(Trim(rs0("PLANT_NM")))%>"
	.txtLongtermStockCalPeriod.Value	= "<%=ConvSPChars(Trim(rs0("LONGTERM_STOCK_CAL_PERIOD")))%>"
	.txtPerniciousStockCalPeriod.Value	= "<%=ConvSPChars(Trim(rs0("PERNICIOUS_STOCK_CAL_PERIOD")))%>"

	'If "<%=Trim(rs0("PLAN_STOCK_CAL_OPTION"))%>" = "Y" then
	'	.cboplanflag(0).Checked = True
	'Else
	'	.cboplanflag(1).Checked = True
	'End If
		
	'.txtplanStockCalPeriod.value = "<%=Trim(rs0(PLAN_STOCK_CAL_PERIOD))%>"
	
End With

parent.DbQueryOk
</Script>

