<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1504mb1.asp
'*  4. Program Name         : Shift Exception Query
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/06/24
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Park, BumSoo 
'* 10. Modifier (Last)      : Ryu Sung Won 
'* 11. Comment              : 
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3						'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim lgStrPrevKey
Dim TmpBuffer
Dim iTotalStr
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

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	Redim UNISqlId(3)
	Redim UNIValue(3, 3)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"
	UNISqlId(2) = "180000sap"
	UNISqlId(3) = "180000sao"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(1, 1) = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "
	
	UNIValue(2, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(2, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(2, 2) = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "
	UNIValue(2, 3) = " " & FilterVar(UCase(Request("txtShiftCd")), "''", "S") & " "
	
	UNIValue(3, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(3, 1) = " " & FilterVar(UCase(Request("txtShiftCd")), "''", "S") & " "

	UNILock = DISCONNREAD :	UNIFlag = "1"
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtResourceNm.value = ""
		parent.frm1.txtShiftNm.value = ""
	</Script>	
	<%    	

	' Plant �� Display      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM")) & ConvSPChars(rs3("Description"))%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If

	' �ڿ��� Display
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("181604", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtResourceCd.Focus()
		</Script>	
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtResourceNm.value = "<%=ConvSPChars(rs1("Description"))%>"
		</Script>	
		<%
		rs1.Close
		Set rs1 = Nothing
	End If

	' Shift Description Display
	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("181800", vbOKOnly, "", "", I_MKSCRIPT)	'#####
		%>
		<Script Language=vbscript>
		parent.frm1.txtShiftCd.Focus()
		</Script>	
		<%
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtShiftNm.value = "<%=ConvSPChars(rs3("Description"))%>"
		</Script>	
		<%
		rs2.Close
		Set rs2 = Nothing
	End If
	rs3.Close
	Set rs3 = Nothing

	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "p1504mb1"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtShiftCd")), "''", "S") & " "
	UNIValue(0, 2) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(0, 3) = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "
	UNIValue(0, 4) = " " & FilterVar(lgStrPrevKey, "''", "S") & "" 	

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs3)

	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("181900", vbOKOnly, "", "", I_MKSCRIPT)
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim LngRow

With parent
	LngMaxRow = .frm1.vspdData.MaxRows

<%  
	If Not(rs3.EOF And rs3.BOF) Then
		If C_SHEETMAXROWS_D < rs3.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs3.RecordCount - 1%>)
<%
	End If

		For i=0 to rs3.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs3("Shift_Exception_Cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs3("Description"))%>"
			strData = strData & Chr(11) & ""
        
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs3("Start_Dt"))%>"

			strHour = Hour("<%=rs3("Start_Dt")%>")
			strMin  = Minute("<%=rs3("Start_Dt")%>")
			strSec  = Second("<%=rs3("Start_Dt")%>")

			If Len(strHour) < 2 Then
				If Len(strHour) = 1 Then
					strHour = "0" & strHour
				Else
					strHour = "00"
				End If
			End If
			If Len(strMin) < 2 Then
				If Len(strMin) = 1 Then
					strMin = "0" & strMin
				Else
					strMin = "00"
				End If
			End If
			If Len(strSec) < 2 Then
				If Len(strSec) = 1 Then
					strSec = "0" & strSec
				Else
					strSec = "00"
				End If
			End If

			strTime = strHour & ":" & strMin & ":" & strSec
			strData = strData & Chr(11) & strTime
        
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs3("End_Dt"))%>"
        
			strHour = Hour("<%=rs3("End_Dt")%>")
			strMin  = Minute("<%=rs3("End_Dt")%>")
			strSec  = Second("<%=rs3("End_Dt")%>")

			If Len(strHour) < 2 Then
				If Len(strHour) = 1 Then
					strHour = "0" & strHour
				Else
					strHour = "00"
				End If
			End If
			If Len(strMin) < 2 Then
				If Len(strMin) = 1 Then
					strMin = "0" & strMin
				Else
					strMin = "00"
				End If
			End If
			If Len(strSec) < 2 Then
				If Len(strSec) = 1 Then
					strSec = "0" & strSec
				Else
					strSec = "00"
				End If
			End If

			strTime = strHour & ":" & strMin & ":" & strSec
			strData = strData & Chr(11) & strTime
  
			strData = strData & Chr(11) & "<%=ConvSPChars(rs3("Exception_Type"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs3("Work_flg"))%>"
        
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs3.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs3("Shift_Cd"))%>"
		
<%	
	End If

	rs3.Close
	Set rs3 = Nothing

%>	
		If .frm1.vspdData.MaxRows < parent.parent.VisibleRowCnt(.frm1.vspdData, 0) and .lgStrPrevKey <> "" Then	<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
			.DbQuery
		Else
			.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
			.frm1.hResourceCd.value = "<%=ConvSPChars(Request("txtResourceCd"))%>"
			.frm1.hShiftCd.value	= "<%=ConvSPChars(Request("txtShiftCd"))%>"
			Call .DbQueryOk(LngMaxRow + 1)
		End If

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
