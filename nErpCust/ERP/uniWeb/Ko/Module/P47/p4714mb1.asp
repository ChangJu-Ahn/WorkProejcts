<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4714mb1.asp
'*  4. Program Name         : �ڿ��Һ������ȸ(������)
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/12/04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeon, JaeHyun 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4								'DBAgent Parameter ���� 
Dim strQryMode

Const C_SHEETMAXROWS_D = 50

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim strTemp
Dim i

'@Var_Declare

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strFromDt
Dim strToDt
Dim StrProdOrderNo
Dim StrWcCd
Dim StrItemCd
Dim StrTrackingNo


lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
lgStrPrevKey3 = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
lgStrPrevKey4 = " " & FilterVar(UNIConvDate(Request("lgStrPrevKey4")), "''", "S") & ""

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sam"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	Set ADF = Nothing
	
	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If

	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtItemCd.Focus()" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	' �۾���� Display
	IF Request("txtWcCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtWCCD.focus()" & vbcr
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs3("WC_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	' Tracking No Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtTrackingNo.Focus()" & vbCrLf
			Response.Write "</Script>	" & vbCrLf
			Response.End
		Else
			rs4.Close
			Set rs4 = Nothing
		End If
	End IF
		        
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 8)

	UNISqlId(0) = "189754SAA"
	
	IF Request("txtFromDt") = "" Then
		strFromDt = "|"
	Else
		strFromDt = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF
	
	IF Request("txtToDt") = "" Then
		strToDt = "|"
	Else
		strToDt = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
	End IF

	Select Case strQryMode
		Case CStr(OPMD_CMODE)		
			IF Request("txtProdOrderNo") = "" Then
				StrProdOrderNo = "|"
			Else
				StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
			End IF
			
		Case CStr(OPMD_UMODE) 
			StrProdOrderNo = "|"
	End Select

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strFromDt
	UNIValue(0, 7) = strToDt
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 8) = "|"
		Case CStr(OPMD_UMODE) 
		
			 strTemp = ""
			 strTemp = "(a.prodt_order_no > " & lgStrPrevKey
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey  'second condition  for group view
			 strTemp = strTemp  & " and a.opr_no > " & lgStrPrevKey2 & ") "  'second condition  for group view
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey	'third condition  for group view
			 strTemp = strTemp  & " and a.opr_no = " & lgStrPrevKey2 		'third condition  for group view
			 strTemp = strTemp  & " and a.resource_cd > " & lgStrPrevKey3 & ") "  'third condition  for group view
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey  'Forth condition  for group view
			 strTemp = strTemp  & " and a.opr_no = " & lgStrPrevKey2			'Forth condition  for group view
			 strTemp = strTemp  & " and a.resource_cd = " & lgStrPrevKey3  'Forth condition  for group view
			 strTemp = strTemp  & " and a.consumed_dt >= " & lgStrPrevKey4 & ")) " 'Forth condition  for group view
			UNIValue(0, 8) = strTemp
	End Select
		
	UNILock = DISCONNREAD :	UNIFlag = "1"	 
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing
    
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

	LngMaxRow = .frm1.vspdData.MaxRows															'Save previous Maxrow
		
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"	        '�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"					'�����ڵ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"			'�ڿ��ڵ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"			'�ڿ��ڵ�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM_RESOURCE_TYPE"))%>"	'�ڿ����� 
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("CONSUMED_DT"))%>"	'�ڿ��Һ��� 
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("CONSUMED_TIME"))%>"		'�ڿ��Һ�ð� 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>" '�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"	'����							
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'��ǰ���� 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("BAD_QTY"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'�ҷ����� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>"		'�ڿ��׷��ڵ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_NM"))%>"			'�ڿ��׷��			
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"				'ǰ�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"				'ǰ��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"				'ǰ��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ROUT_NO"))%>"				'����� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"					'�۾��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"					'�۾����			
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"			'Tracking No.
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
				
			END IF
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=ConvSPChars(Trim(rs0("PRODT_ORDER_NO")))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(Trim(rs0("OPR_NO")))%>"
		.lgStrPrevKey3 = "<%=ConvSPChars(Trim(rs0("RESOURCE_CD")))%>"
		.lgStrPrevKey4 = "<%=UniDateClientFormat(rs0("CONSUMED_DT"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hWcCd.value = "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hFromDt.value = "<%=ConvSPChars(Request("txtFromDt"))%>"
		.frm1.hToDt.value = "<%=ConvSPChars(Request("txtToDt"))%>"
		     
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbQueryOk

End With

</Script>	

<script Language = vbscript RUNAT = server>
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
			
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</script>
