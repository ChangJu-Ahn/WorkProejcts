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
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "C", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3,	rs4, rs5
Dim strQryMode

Dim StrNextKey
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strGubun
Dim strOrderNo
Dim strOprNo,strYYYYMM

Dim strFlag

Dim strWhere
Err.Clear

strGubun= request("txtGubun")
strYYYYMM=replace( request( "txtYYYYMM"),"-","")
strOprNo= UCase(Request("txtOprNo"))
strOrderNo = UCase(Request("txtProdOrderNo"))

If strGubun = "A"	Then				'cost
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "c4205ma1011"
	UNIValue(0,0) =FilterVar (strYYYYMM ,"''","S")
	UNIValue(0,1) =FilterVar ( strOrderNo,"''","S")
	UNIValue(0,2) =FilterVar (strOprNo ,"''","S")
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	      
	If (rs0.EOF And rs0.BOF) Then
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

Else				'batch
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "c4205ma1021"
	UNIValue(0,0) =FilterVar (strYYYYMM ,"''","S")
	UNIValue(0,1) =FilterVar ( strOrderNo,"''","S")
	UNIValue(0,2) =FilterVar (strOprNo ,"''","S")
	UNIValue(0,3) =FilterVar (request("txtSpId") ,"''","S")
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	      
	If (rs0.EOF And rs0.BOF) Then
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		
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
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("child_item_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"				
							
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
				
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("this_wip_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("this_wip_amt"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("unit_price"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer,"")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
		'.lgStrPrevKey2 = "<%=i%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
		
	.DbDtlQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
