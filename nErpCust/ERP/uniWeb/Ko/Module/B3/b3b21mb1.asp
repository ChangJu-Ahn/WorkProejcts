<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b21mb1.asp
'*  4. Program Name         : ����׸� ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3						'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim lgStrPrevKey

Dim strCharCd
Dim strCharValueCd
Dim i
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = Request("lgStrPrevKey")

strCharCd = " " & FilterVar(UCase(Request("txtCharCd")), "''", "S") & " "

If Trim(lgStrPrevKey) = "" Then
	If Trim(Request("txtCharValueCd")) <> "" Then
		strCharValueCd = " " & FilterVar(UCase(Request("txtCharValueCd")), "''", "S") & " "
	Else
		strCharValueCd = "|"
	End If
Else
	strCharValueCd = " " & FilterVar(UCase(lgStrPrevKey), "''", "S") & " "
End If
	
On Error Resume Next
Err.Clear

	Redim UNISqlId(2)
	Redim UNIValue(2,1)

	UNISqlId(0) = "b3b21mb1a"
	UNISqlId(1) = "b3b21mb1b"
	UNISqlId(2) = "b3b21mb1c"

	UNIValue(0,0) = strCharCd
	UNIValue(1,0) = strCharCd
	UNIValue(1,1) = strCharValueCd
	UNIValue(2,0) = strCharCd
	UNIValue(2,1) = strCharCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs3)

	%>
	<Script Language=vbscript>
		parent.frm1.txtCharNm.value = ""
		parent.frm1.txtCharCd1.value = ""
		parent.frm1.txtCharNm1.value = ""
	</Script>	
	<%    	
	' Char�� Display      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtCharCd.Focus()
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
		parent.frm1.txtCharNm.value = "<%=ConvSPChars(rs0("CHAR_NM"))%>"
		parent.frm1.txtCharCd1.value = "<%=ConvSPChars(rs0("CHAR_CD"))%>"
		parent.frm1.txtCharNm1.value = "<%=ConvSPChars(rs0("CHAR_NM"))%>"
		parent.frm1.txtCharValueDigit.value = "<%=ConvSPChars(rs0("CHAR_VALUE_DIGIT"))%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If

	' char_value_digit Protected
	If (rs3.EOF And rs3.BOF) Then
		%>
		<Script Language=vbscript>
		Call parent.ggoOper.SetReqAttr(parent.frm1.txtCharValueDigit, "N")
		</Script>	
		<%
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
	Else
		%>
		<Script Language=vbscript>
		Call parent.ggoOper.SetReqAttr(parent.frm1.txtCharValueDigit, "Q")
		</Script>	
		<%
		rs3.Close
		Set rs3 = Nothing
	End If
	%>
	<Script Language=vbscript>
		parent.frm1.txtCharValueNm.value = ""
	</Script>	
	<%
	' CharValue�� Display
	If strCharValueCd <> "|" Then
		If (rs1.EOF And rs1.BOF) Then	
			%>
			<Script Language=vbscript>
			parent.frm1.txtCharValueCd.Focus()
			</Script>	
			<%
			rs1.Close
			Set rs1 = Nothing
		Else
			%>
			<Script Language=vbscript>
			parent.frm1.txtCharValueNm.value = "<%=ConvSPChars(rs1("CHAR_VALUE_NM"))%>"
			</Script>	
			<%
			rs1.Close
			Set rs1 = Nothing
		End If
	End If
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "b3b21mb1"
	
	UNIValue(0, 0) = strCharCd
	UNIValue(0, 1) = strCharValueCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)

	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.DbQueryOk(false)
		</Script>
		<%
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End
	End If

%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent
	LngMaxRow = .frm1.vspdData.MaxRows

<%  
	If Not(rs2.EOF And rs2.BOF) Then
		If C_SHEETMAXROWS_D < rs2.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs2.RecordCount - 1%>)
<%
		End If

		For i=0 to rs2.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("CHAR_VALUE_CD"))%>"			'1
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("CHAR_VALUE_NM"))%>"		'3
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%
			rs2.MoveNext
			End If
		Next
%>

		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs2("CHAR_VALUE_CD"))%>"
<%	
	End If

	rs2.Close
	Set rs2 = Nothing
%>	
		If .frm1.vspdData.MaxRows < parent.parent.VisibleRowCnt(.frm1.vspdData, 0) and .lgStrPrevKey <> "" Then
			.DbQuery
		Else
			.frm1.hCharCd.value	= "<%=ConvSPChars(Request("txtCharCd"))%>"
			.frm1.hCharValueCd.value = "<%=ConvSPChars(Request("txtCharValueCd"))%>" 
			.DbQueryOk(true)
		End If

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
