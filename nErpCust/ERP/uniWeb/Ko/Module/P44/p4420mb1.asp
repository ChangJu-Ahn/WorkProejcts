<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4420mb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2003/05/26
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
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

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter ���� 
Dim strQryMode

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")


On Error Resume Next

Dim strItemCd
Dim strItemAcct
Dim strWcCd
Dim strShiftCd
Dim strItemGroupCd
Dim strFlag

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	
	
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemGroupNm.value = ""
		parent.frm1.txtWCNm.value = ""
	</Script>	
	<%
	
   	' Plant �� Display      
	
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs4.EOF AND rs4.BOF Then
			strFlag = "ERROR_GROUP"
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs4("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	End If

	' ǰ��� Display
	IF Request("txtItemCd") <>      "" Then
		If (rs2.EOF And rs2.BOF) Then
			
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
		End If
   	End IF

	' �۾���� Display
	IF Request("txtWcCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_WCCD"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtWCNm.value = "<%=ConvSPChars(rs3("WC_NM"))%>"
			</Script>	
			<%
		End If
	End IF
	
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	rs3.Close	:	Set rs3 = Nothing
	rs4.Close	:	Set rs4 = Nothing
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantNm.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End	
		End If
		
    	If strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		
	    End If

		If strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtWcNm.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		End If
		
		If strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Set ADF = Nothing
			Response.End
		End If
	End IF

	
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)
	
	UNISqlId(0) = "189611saa"
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF
		
	IF Request("cboItemAcct") = "" Then
	   strItemAcct = "|"
	ELSE
	   strItemAcct = " " & FilterVar(Request("cboItemAcct"), "''", "S") & ""
	END IF
	
	IF Request("txtWcCd") = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Request("txtShliftCd") = "" Then
	   strShiftCd = "|"
	ELSE
	   strShiftCd = FilterVar(UCase(Request("txtShliftCd")), "''", "S")
	END IF	
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "d.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = " " & FilterVar(Request("txtWorkDt"), "''", "S") & "" 
	UNIValue(0, 3) = " " & FilterVar(UniDateAdd("M",1,Request("txtWorkDt"),gServerDateFormat), "''", "S") & ""  	
	UNIValue(0, 4) = strShiftCd
	UNIValue(0, 5) = strItemCd
	UNIValue(0, 6) = strItemAcct
	UNIValue(0, 7) = strWcCd
	UNIValue(0, 8) = "" & FilterVar("P1001", "''", "S") & "" 'minor code for item account
	UNIValue(0, 9) = strItemGroupCd

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
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
    	
With parent																'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows									'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  		
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEMCD"))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEMNM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEMACCT"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOODQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BADQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPTQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("UNIT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData
	.ggoSpread.SSShowDataByClip iTotalStr
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
