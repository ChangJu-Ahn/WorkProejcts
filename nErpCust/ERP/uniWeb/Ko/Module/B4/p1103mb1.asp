<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1103mb1.asp
'*  4. Program Name         : Mfg Calendar Type Query
'*  5. Program Desc         :
'*  6. Component List       :  DB Agent (p1103mb1)
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2000/06/24
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1								'DBAgent Parameter ���� 
Dim lgStrPrevKey
Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd
Call LoadBasisGlobalInf() 

lgStrPrevKey = Request("lgStrPrevKey")

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================

	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saj"
	UNISqlId(1) = "p1103mb1"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtClnrType")), "''", "S")

	If lgStrPrevKey = "" Then
		UNIValue(1, 0) =FilterVar(UCase(Request("txtClnrType")), "''", "S")
	Else
		UNIValue(1, 0) = FilterVar(lgStrPrevKey, "''", "S")	
	End If
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	' Plant �� Display      
	If (rs0.EOF And rs0.BOF) Then
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtClnrTypeNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtClnrTypeNm.value = """ & ConvSPChars(rs0("CAL_TYPE_NM")) & """" & vbCrLf	'''''
		Response.Write "</Script>" & vbCrLf
	End If
	rs0.Close
	Set rs0 = Nothing

	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("180300", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
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
	If Not(rs1.EOF And rs1.BOF) Then
		If C_SHEETMAXROWS_D < rs1.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs1.RecordCount - 1%>)
<%
		End If

		For i=0 to rs1.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("cal_type"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs1("cal_type_nm"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData

<%		
			rs1.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		.lgStrPrevKey = "<%=ConvSPChars(rs1("cal_type"))%>"
		
<%	
	End If

	rs1.Close
	Set rs1 = Nothing

%>	
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hClnrType.value = "<%=ConvSPChars(Request("txtClnrType"))%>"

		.DbQueryOk
	End If

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
