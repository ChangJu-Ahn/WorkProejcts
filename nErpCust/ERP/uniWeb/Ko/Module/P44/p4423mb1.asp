<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4423mb1.asp
'*  4. Program Name         : ���ְ����񳻿� ��ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.11.28
'*  7. Modified date(Last)  : 2003/06/30
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter ���� 
Dim strQryMode, strFlag

Dim strBpCd
Dim strFromDt
Dim strToDt
Dim StrPlantCd
Dim StrWcCd
Dim strTemp

Const C_SHEETMAXROWS_D = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim i

Call HideStatusWnd

On Error Resume Next								'��: 

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
lgStrPrevKey3 = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
lgStrPrevKey4 = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 0)

	UNISqlId(0) = "180000sak"
	UNISqlId(1) = "180000saa"
	UNISqlId(2) = "180000sac"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtBpCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")		

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	' ����ó�� Display
	IF Request("txtBpCd") <> "" Then
		If (rs1.EOF And rs1.BOF) Then
			rs1.Close
			Set rs1 = Nothing
			strFlag = "ERROR_BPCD"
			%>
			<Script Language=vbscript>
				parent.frm1.txtBpNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtBpNm.value = "<%=ConvSPChars(rs1("Bp_NM"))%>"
			</Script>	
			<%
			rs1.Close
			Set rs1 = Nothing
		End If
	End IF
	
	' Plant �� Display
	IF Request("txtPlantCd") <> "" Then	    
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_PLANT"
			%>
			<Script Language=vbscript>
				parent.frm1.txtPlantNm.value = ""
			</Script>	
			<%    	
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs2("PLANT_NM"))%>"
			</Script>	
			<%    	
			rs2.Close
			Set rs2 = Nothing
		End If
	End If		
		
	' �۾���� Display
	IF Request("txtWcCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_WCCD"
			%>
			<Script Language=vbscript>
				parent.frm1.txtWCNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtWCNm.value = "<%=ConvSPChars(rs3("WC_NM"))%>"
			</Script>	
			<%
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	If strFlag <> "" Then
		If strFlag = "ERROR_BPCD" Then
			Call DisplayMsgBox("189629", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtBPCD.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtWCCD.Focus()
			</Script>	
			<%
			Response.End	
		End If
	End IF
		        
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 6)

	UNISqlId(0) = "p4423mb1a"
	
	IF Request("txtBpCd") = "" Then
		strBpCd = "|"
	Else
		strBpCd = FilterVar(UCase(Request("txtBpCd")), "''", "S")
	End IF
	
	IF UNIConvDate(Request("txtFromDt")) = UNIConvDate("") Then
		strFromDt = "|"
	Else
		strFromDt = " " & FilterVar(UniConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF
	
	IF UNIConvDate(Request("txtToDt")) = UNIConvDate("") Then
		strToDt = "|"
	Else
		strToDt = " " & FilterVar(UniConvDate(Request("txtToDt")), "''", "S") & ""
	End IF

	IF Request("txtPlantCd") = "" Then
		strPlantCd = "|"
	Else
		strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF
		
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strBpCd
	UNIValue(0, 2) = strFromDt
	UNIValue(0, 3) = strToDt
	UNIValue(0, 4) = strPlantCd
	UNIValue(0, 5) = strWcCd
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)		
			UNIValue(0, 6) = "|" 
		Case CStr(OPMD_UMODE)
			 strTemp = ""
			 strTemp = "(B.BP_CD > " & lgStrPrevKey 
			 strTemp = strTemp  & " or (B.BP_CD = " & lgStrPrevKey		  'second condition  for group view
			 strTemp = strTemp  & " and A.PLANT_CD > " & lgStrPrevKey2 & ") "	  'second condition  for group view
			 strTemp = strTemp  & " or (B.BP_CD = " & lgStrPrevKey		  'third condition  for group view
			 strTemp = strTemp  & " and A.PLANT_CD = " & lgStrPrevKey2 		  'third condition  for group view
			 strTemp = strTemp  & " and B.CUR_CD > " & lgStrPrevKey3 & ") "	  'third condition  for group view 
			 strTemp = strTemp  & " or (B.BP_CD = " & lgStrPrevKey		  'fourth condition  for group view
			 strTemp = strTemp  & " and A.PLANT_CD = " & lgStrPrevKey2 		  'fourth condition  for group view
			 strTemp = strTemp  & " and B.CUR_CD = " & lgStrPrevKey3		  'fourth condition  for group view
			 strTemp = strTemp  & " and B.TAX_TYPE >= " & lgStrPrevKey4 & ")) "  'fourth condition  for group view  
			UNIValue(0, 6) = strTemp
	End Select	
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																						'��: ȭ�� ó�� ASP �� ��Ī�� 

	LngMaxRow = .frm1.vspdData1.MaxRows															'Save previous Maxrow
		
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
			If i < C_SHEETMAXROWS_D THEN
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"									'����ó 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"									'����ó�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CUR_CD"))%>"										'ȭ����� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("SUBCONTRACT_AMT"), 0)%>"	'���ֱݾ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TAX_TYPE"))%>"								'VAT���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TAX_TYPE_NM"))%>"							'VAT���� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("TAX_AMT"), 0)%>"		'VAT�ݾ� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("TOTAL_COST"), 0)%>"	'�ѱݾ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"								'���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_NM"))%>"								'����� 
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			END IF
		Next
		
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source  = .frm1.vspdData1
		Call .ggoSpread.SSShowDataByClip(iTotalStr, "F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData1, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd,.C_SubcontractAmt, "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData1, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd,.C_TaxAmt, "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData1, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd,.C_TotalCost, "A" ,"I","X","X")
		
		.lgStrPrevKey =	"<%=ConvSPChars(rs0("BP_CD"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("PLANT_CD"))%>"
		.lgStrPrevKey3 = "<%=ConvSPChars(rs0("CUR_CD"))%>"
		.lgStrPrevKey4 = "<%=ConvSPChars(rs0("TAX_TYPE"))%>"
		
		.frm1.hBPCd.value = "<%=ConvSPChars(Request("txtBPCd"))%>"
		.frm1.hFromDt.value = "<%=ConvSPChars(Request("txtFromDt"))%>"
		.frm1.hToDt.value = "<%=ConvSPChars(Request("txtToDt"))%>"
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hWcCd.value = "<%=ConvSPChars(Request("txtWcCd"))%>"
		     
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbQueryOk												'��: ��ȸ ������ ������� 

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
