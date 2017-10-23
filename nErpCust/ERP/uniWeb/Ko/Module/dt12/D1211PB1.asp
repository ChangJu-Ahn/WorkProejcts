<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: DT
'*  2. Function Name		: 
'*  3. Program ID			: d1211PB1.asp
'*  4. Program Name			: Digital Tax (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2009/12/20
'*  8. Modified date(Last)	: 2009/12/22
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M","NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs1										'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim i
Dim strFlag
Dim strInvNo 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strInvNo = Request("txtInvNo")

On Error Resume Next
Err.Clear
																	'��: Protect system from crashing
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "D1211PA1"
	
	UNIValue(0, 0) = FilterVar(strInvNo, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
	'// QUERY REWORK ORDER HISTORY
	
	
%>
<Script Language=vbscript>
	Dim TmpBuffer1
    Dim iTotalStr
    Dim LngMaxRow
    Dim strData
	
    With parent												'��: ȭ�� ó�� ASP �� ��Ī�� 
		
	 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
		
<%  
		If Not(rs1.EOF And rs1.BOF) Then
%>	
		
			
			Redim TmpBuffer1(<%=rs1.RecordCount-1%>)
<%		
			For i=0 to rs1.RecordCount-1
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("inv_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("wrk_dtm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_flag"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_flag_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_usr_id"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("proc_usr_name"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("attr02"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs1("sup_tot_amt"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs1("sur_tax"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(<%=i%>) = strData
<%		
				rs1.MoveNext
				
			Next
%>
			
		iTotalStr = Join(TmpBuffer1,"") 
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr

<%	
		End If
		

		rs1.close

		Set rs1 = Nothing

%>	
		
		.DbQueryOk(LngMaxRow)
		
    End With
</Script>	
<%    
    Set ADF = Nothing
%>
