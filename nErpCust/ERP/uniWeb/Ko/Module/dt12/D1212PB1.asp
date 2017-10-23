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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs1,  rs2								'DBAgent Parameter ���� 
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
	Set ADF = Nothing
	
	'// QUERY REWORK ORDER HISTORY
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)
	
	UNISqlId(0) = "D1212PA11"
	UNISqlId(1) = "D1212PA12"
	
	UNIValue(0, 0) = FilterVar(strInvNo, "''", "S")
	UNIValue(1, 0) = FilterVar(strInvNo, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)
    
    If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	  
	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs1 = Nothing
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>
<Script Language=vbscript>
	Dim TmpBuffer1
    Dim iTotalStr
    Dim LngMaxRow
    Dim strData
	
    With parent												'��: ȭ�� ó�� ASP �� ��Ī�� 
    
	 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
		
<%  
		If Not(rs0.EOF And rs0.BOF) Then
%>	
			
			.dtCreateDate.text = "<%=UNIDateClientFormat(rs1("cre_date"))%>"
			.numSumAmt.value = "<%=UniNumClientFormat(rs1("sum_amt"),ggAmtOfMoney.DecPoint,0)%>"  
			.numNetAmt.value = "<%=UniNumClientFormat(rs1("sup_tot_amt"),ggAmtOfMoney.DecPoint,0)%>"
			.numVatAmt.value = "<%=UniNumClientFormat(rs1("sur_tax"),ggAmtOfMoney.DecPoint,0)%>"  
			
			.txtRegNoS.value =  "<%=ConvSPChars(rs1("sup_reg_num"))%>"
			.txtRegNoB.value =  "<%=ConvSPChars(rs1("dem_reg_num"))%>"
			.txtSubRegnoS.value =  "<%=ConvSPChars(rs1("sup_reg_id"))%>"
			.txtSubRegnoB.value =  "<%=ConvSPChars(rs1("dem_reg_id"))%>"
			.txtBizAreaS.value =  "<%=ConvSPChars(rs1("sup_cmp_name"))%>"
			.txtBizAreaB.value =  "<%=ConvSPChars(rs1("dem_cmp_name"))%>"
			.txtOwnerS.value =  "<%=ConvSPChars(rs1("sup_owner"))%>"
			.txtOwnerB.value =  "<%=ConvSPChars(rs1("dem_owner"))%>"
			.txtAddressS.value =  "<%=ConvSPChars(rs1("sup_cmp_addr"))%>"
			.txtAddressB.value =  "<%=ConvSPChars(rs1("dem_cmp_addr"))%>"
			.txtBizTypeS.value =  "<%=ConvSPChars(rs1("sup_biz_type"))%>"
			.txtBizTypeB.value =  "<%=ConvSPChars(rs1("dem_biz_type"))%>"
			.txtBizKindS.value =  "<%=ConvSPChars(rs1("sup_biz_kind"))%>"
			.txtBizKindB.value =  "<%=ConvSPChars(rs1("dem_biz_kind"))%>"
			
			Redim TmpBuffer1(<%=rs2.RecordCount-1%>)
<%		
			For i=0 to rs2.RecordCount-1
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("sale_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("ln_ord"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("sup_date"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("item"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("item_std1"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("item_unit"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("item_qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("item_prc"),ggUnitCost.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("item_amt"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("item_tax"),ggAmtOfMoney.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("item_memo"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("code_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("ser_no"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(<%=i%>) = strData
<%		
				rs2.MoveNext
				
			Next
%>
			

		iTotalStr = Join(TmpBuffer1,"") 
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr

<%	
		End If
		
		rs1.close
		rs2.close

		Set rs1 = Nothing
		Set rs2 = Nothing
%>	
		
		.DbQueryOk(LngMaxRow)
		
    End With
</Script>	
<%    
    Set ADF = Nothing
%>
