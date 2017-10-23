<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4111rb1.asp
'*  4. Program Name			: List Production Order Detail (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/06/27
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment				:
'********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1								'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim i

Dim strWhere

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear																	'��: Protect system from crashing
	
	If Trim(Request("lgStrPrevKey1")) = "" Then
		strWhere = ""
	Else
		strWhere = " AND ( A.OPR_NO > " &  FilterVar(Request("lgStrPrevKey1"), "''", "S")  _
					& " OR ( A.OPR_NO = " & FilterVar(Request("lgStrPrevKey1"), "''", "S") _
					& " AND A.CHANGE_SEQ >= " & FilterVar(Request("lgStrPrevKey2"), "''", "S") & "))"
	End If
	
	'// Order Detail Display
	Redim UNISqlId(1)
	Redim UNIValue(1, 4)

	UNISqlId(0) = "p4111mb1"
	UNISqlId(1) = "p4114rb1"
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	UNIValue(0, 3) = "|"
	UNIValue(0, 4) = "|"
	UNIValue(1, 0) = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	UNIValue(1, 1) = strWhere
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs1, rs0)
    
    If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemCd.value = "<%=ConvSPChars(rs1("ITEM_CD"))%>"
			parent.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
		</Script>	
		<%
		Set rs1 = Nothing
	End If
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If

%>
<Script Language=vbscript>
    Dim LngMaxRow
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr
	
    With parent												'��: ȭ�� ó�� ASP �� ��Ī�� 
		
 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
	
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Opr_No"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHANGE_SEQ"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ACTION_FLG"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("APPLY_DATE"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRE_WC_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("USR_NM"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr

		.lgStrPrevKey1 = "<%=Trim(rs0("Opr_No"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then
		.InitData(LngMaxRow)
		.DbQuery
	Else
		.DbQueryOk(LngMaxRow)
	End If

    End With
</Script>	
<%    
    Set ADF = Nothing
%>
