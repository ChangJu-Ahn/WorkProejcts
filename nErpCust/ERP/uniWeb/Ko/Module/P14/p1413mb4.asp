<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1413mb4.asp
'*  4. Program Name         : Look Up Item Info
'*  5. Program Desc         :
'*  6. Comproxy List        :  
'*  7. Modified date(First) : 2003/03/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*************************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strItemNm
Dim strSpec
Dim strItemAcctNm
Dim strProcurTypeNm
Dim Row 

Call HideStatusWnd

On Error Resume Next
Err.Clear
	
	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
	Row = Request("Row")	

	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "180000saq"
	UNIValue(0, 0) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF and rs0.BOF)Then
		Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
		strItemNm	= ""
		strSpec		= ""
		strItemAcctNm = ""
		strProcurTypeNm = ""
		rs0.Close
		Set rs0 = Nothing
	Else
		strItemNm	= rs0("ITEM_NM")
		strSpec		= rs0("SPEC")
		strItemAcctNm = UCase(Trim(rs0("ITEM_ACCT_NM")))
		strProcurTypeNm = UCase(Trim(rs0("PROCUR_TYPE_NM")))
		rs0.Close
		Set rs0 = Nothing
	End If

%>
<Script Language=vbscript>
	Call parent.LookUpItemInfoOk("<%=strItemNm%>","<%=strSpec%>","<%=strItemAcctNm%>","<%=strProcurTypeNm%>","<%=Row%>")
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
