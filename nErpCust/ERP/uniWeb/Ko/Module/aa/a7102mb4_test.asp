<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

fod<% 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7102mb4
'*  4. Program Name         : �����ڻ���泻�����
'*  5. Program Desc         : ����ȣ�� �ڻ�Master ��ȸ(6th)
'*  6. Comproxy List        : +As0029LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : ������
'* 10. Modifier (Last)      : ������
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. ���Ǻ�
'##########################################################################################################
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
''On Error Resume Next														'��: 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
strMode = Request("txtMode")		
										'�� : ���� ���¸� ����
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
'#########################################################################################################
'												2.1 ���� üũ
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then											'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)		'��: ��ȸ �����ε� �ٸ� ���·� ��û�� ���� ���, �ʿ������ ���� ��, �޼����� ID������ ����ؾ� ��
	Response.End 
ElseIf Trim(Request("txtAcqNo")) = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)						'��:
	Response.End 
End If

'#########################################################################################################
'												2. ���� ó�� �����
'##########################################################################################################
'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim strAs0029																	'�� : ��ȸ�� ComProxy Dll ��� ����
Dim IntRows
Dim IntCols
Dim sList
Dim strData_i, strData_m
Dim vbIntRet
Dim intCount,IntCount1
Dim IntCurSeq
Dim LngMaxRow
Dim StrNextKey_m,plgStrPrevKey_m

'#########################################################################################################
'												2.2. ��û ���� ó��
'##########################################################################################################
	plgStrPrevKey_m = Request("lgStrPrevKey_m")

	Set strAs0029 = Server.CreateObject("As0029.As0029AcqLookupSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set strAs0029 = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������
	End If
	
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	If plgStrPrevKey_m = "" Then
		strAs0029.ImportNextAAssetMasterAsstNo = ""
	Else
		strAs0029.ImportNextAAssetMasterAsstNo = plgStrPrevKey_m
	End If
	
	strAs0029.ImportAAssetAcqAcqNo	= Trim(Request("txtAcqNo"))
	strAs0029.ImportFgIefSuppliedSelectChar = "B"      ' �ι�° Tab�� ���� ��ȸ ǥ��
	strAs0029.ServerLocation		= ggServerIP

	strAs0029.ComCfg = gConnectionString
	strAs0029.Execute
   
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.number <> 0 Then
		Set strAs0029 = Nothing
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End 								'��: �����Ͻ� ���� ó���� ������
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (strAs0029.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(strAs0029.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		Set strAs0029 = Nothing												'��: ComProxy Unload
		Response.End
	End If

	'#########################################################################################################
	'												2.4. HTML ��� ������
	'##########################################################################################################

%>
	<Script Language=vbscript>
	Dim LngMaxRow
	Dim varMakeFg
	Dim varApAmt

		'**************************************************************************************
		'        The Query Part Of Multi for Asset Master
		'**************************************************************************************	
<%
		LngMaxRow = Request("txtMaxRows_m")											'Save previous Maxrow                                                
		intCount = strAs0029.ExpGroupCount

		If strAs0029.ExportItemAAssetMasterAsstNo(intCount) = strAs0029.ExportNextAAssetMasterAsstNo Then
			StrNextKey_m = ""
		Else
			StrNextKey_m = strAs0029.ExportNextAAssetMasterAsstNo
		End If
    
%>


	With parent
	
		LngMaxRow = .frm1.vspdData.MaxRows
		.frm1.vspdData.MaxRows = LngMaxRow

<%    
		For IntRows = 1 To intCount
%>


<% 
	Dim lgCurrency
	lgCurrency = ConvSPChars(strAs0029.ExportItemAAssetMasterdoccur(IntRows))
%>	
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemBAcctDeptDeptCd(IntRows))%>"
			strData_m = strData_m & Chr(11) & ""	    
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemBAcctDeptDeptNm(IntRows))%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAcctAcctCd(IntRows))%>"
			strData_m = strData_m & Chr(11) & ""
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAcctAcctNm(IntRows))%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAsstNo(IntRows))%>"                		
        	strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAsstNm(IntRows))%>"        
    
			'strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqAmt(IntRows),    ggAmtOfMoney.DecPoint, 0)%>"
			'strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqLocAmt(IntRows), ggAmtOfMoney.DecPoint, 0)%>"

			strData_m = strData_m & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(strAs0029.ExportItemAAssetMasterAcqAmt(IntRows),lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
			strData_m = strData_m & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(strAs0029.ExportItemAAssetMasterAcqLocAmt(IntRows),gCurrency,ggAmtOfMoneyNo, gLocRndPolocyNo, "X")%>"
			
			strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqQty(IntRows),    ggQty.DecPoint, 0)%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterRefNo(IntRows))%>"   	    	    	    	    
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAssetDesc(IntRows))%>"

			strData_m = strData_m & Chr(11) & LngMaxRow + <%=IntRows%>
			strData_m = strData_m & Chr(11) & Chr(12)

<%
		Next
%>    
	    .ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData_m
		
		.lgStrPrevKey_m = "<%=StrNextKey_m%>"		
		
		If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS_m And .lgStrPrevKey_m <> "" Then	
		 	.DbQuery2
		Else    
			.DbQueryOk2
		End If
	
	End With	

<%
	Set strAs0029 = Nothing
%>

</Script>
