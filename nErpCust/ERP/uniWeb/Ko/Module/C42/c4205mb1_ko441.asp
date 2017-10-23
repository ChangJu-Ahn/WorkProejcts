<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		: 공정별원가 
'*  3. Program ID			: C4205MA1.asp
'*  4. Program Name			:공정별수불장 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4205Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/09/16
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "C", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3,	rs4, rs5
Dim strQryMode

Dim StrNextKey, lgStrPrevKey
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strGubun, strYYYYMM
Dim strOrderNo
Dim strCcCd
Dim strPlantCd
Dim strSpId
Dim strFlag

Dim strWhere
Err.Clear

strGubun		= request("txtGubun")
strYYYYMM		= replace( request( "txtYYYYMM"),"-","")
strCcCd			= UCase(Request("txtCCCd"))
strPlantCd		= UCase(Request("txtPlantCD"))
strOrderNo		= UCase(Request("txtProdOrderNo"))
lgStrPrevKey	= Request("lgStrPrevKey")	' choe0tae 2007-03-29



If strGubun = "A"	Then				'cost
	'  Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)


	UNISqlId(0) = "c4205ma101ko441"
	UNIValue(0,0) =FilterVar (strYYYYMM ,"''","S")
	
	strWhere =""
	IF Request("txtPlantCD") <> "" Then
		strWhere = strWhere & " AND A.PLANT_CD =" & FilterVar(strPlantCd, "''", "S")
	End IF	
	IF Request("txtCCCd") <> "" Then
		strWhere = strWhere & " AND A.COST_CD =" &  FilterVar(strCcCd, "''", "S")
	End IF
	IF Request("txtProdOrderNo") <> "" Then
		strWhere = strWhere & " AND A.ORDER_NO =" &  FilterVar(strOrderNo, "''", "S")
	End IF

	UNIValue(0, 1) = strWhere	
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

Else				'batch
	

	if strPlantCd = ""	then strPlantCd = "%"
	if strCcCd = ""		then strCcCd = "%"


    Call SubCreateCommandObject(lgObjComm)	 
	'get sp_id
	With lgObjComm
		.CommandTimeout = 0				

		.CommandText = "dbo.USP_C_PRNT_MVMT_BY_OPR_BATCH_S" 	
	    .CommandType = adCmdStoredProc

		.Parameters.Append	.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)	
		.Parameters.Append	.CreateParameter("@YYYYMM",		adVarXChar,	adParamInput, 6,Replace(strYYYYMM, "'", "''"))
		.Parameters.Append	.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(strPlantCd, "'", "''"))
		.Parameters.Append	.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(strCcCd, "'", "''"))		
		.Parameters.Append	.CreateParameter("@SP_ID",		adSmallInt, adParamOutput, 100)	
		.Parameters.Append	.CreateParameter("@MSG_CD",		adVarXChar,	adParamOutput, 6)		
       	.Execute ,, adExecuteNoRecords

   		If err.number <> 0 Then
   			strSpId = ""
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
			Response.End			
		Else
			strSpId = .Parameters("@SP_ID").Value
		End If 
    End With


	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "c4205ma102"
	UNIValue(0,0) =FilterVar (strYYYYMM ,"''","S")
	UNIValue(0,1) =FilterVar (strSpId ,"''","S")
		
	strWhere =""
		
	IF Request("txtPlantCD") <> "" Then
		strWhere = strWhere & " AND A.PLANT_CD =" & FilterVar(strPlantCd, "''", "S")
	End IF	
	IF Request("txtCCCd") <> "" Then
		strWhere = strWhere & " AND A.COST_CD =" &  FilterVar(strCcCd, "''", "S")
	End IF
	IF Request("txtProdOrderNo") <> "" Then
		strWhere = strWhere & " AND A.ORDER_NO =" &  FilterVar(strOrderNo, "''", "S")
	End IF
		

	UNIValue(0, 2) = strWhere	

				
	UNILock = DISCONNREAD :	UNIFlag = "1"
			
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
		      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
End If
%>

<Script Language=vbscript>

Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		
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
		
		' -- choe0tae 2007-03-29  조회환 레코드셋에서 페이지만큼 점프하는 기능 
		If ("" & lgStrPrevKey) <> "" Then
			For i = 0 To CDbl(lgStrPrevKey)-1
				rs0.MoveNext
			Next 
		Else 
			lgStrPrevKey = 0
		End If
		i = 0
		
		Do Until rs0.EOF
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("plant_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cost_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cost_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("order_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"				
							
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("close_flag"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("inside_flg"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("milestone_flg"))%>"
				
				
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("bas_wip_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prior_opr_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("next_opr_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("last_wip_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("before_bad_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("this_bad_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("before_bad_rework_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("bal_bad_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		

				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_rate"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
				
				' -- choe0tae 2007-03-29  조회환 레코드셋에서 페이지만큼 점프하는 기능 
				i = i + 1
		
				If i >= C_SHEETMAXROWS_D Then 
					StrNextKey = CStr(CDbl(lgStrPrevKey) + i)
					Exit Do
				End If
		Loop
		
		
%>
		iTotalStr = Join(TmpBuffer,"")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=StrNextKey %>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"	
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrdNo"))%>"
	.frm1.hCCCd.value			= "<%=ConvSPChars(Request("txtCCCd"))%>"
	.frm1.hYYYYMM.value			= "<%=ConvSPChars(Request("txtYYYYMM"))%>"
	.frm1.hSpId.value			= "<%=strSpId%>"
			
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
