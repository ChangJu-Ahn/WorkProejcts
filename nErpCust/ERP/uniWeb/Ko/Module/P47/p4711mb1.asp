<%@ Language=vbscript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4711mb1.asp
'*  4. Program Name			: List Shift (Query)
'*  5. Program Desc			: List Shift (Called By Confirm By Operation and Confirm By Order)
'*  6. Comproxy List		: +
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/06/26
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment		:
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()

On Error Resume Next
Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1										'DBAgent Parameter 선언 
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Call HideStatusWnd

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "p4400mb1"
	UNISqlId(1) = "184000saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "	parent.frm1.txtPlantCd.Focus()"
		Response.Write "</Script>" & vbCrLf			
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("180400", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If


%>

<Script Language=vbscript>
    	
With parent
	.frm1.txtPlantNm.value		= "<%= ConvSPChars(rs1("plant_nm")) %>"

<%	rs1.Close
	Set rs1 = Nothing
	  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1
		
%>
			Call .SetCombo(.frm1.cboShiftCdFrom,"<%=ConvSPChars(rs0("Shift_Cd"))%>","<%=ConvSPChars(rs0("Shift_Cd"))%>")
			Call .SetCombo(.frm1.cboShiftCdTo,"<%=ConvSPChars(rs0("Shift_Cd"))%>","<%=ConvSPChars(rs0("Shift_Cd"))%>")
<%		
			rs0.MoveNext
		Next
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.lgShiftCnt = "<%=i%>"

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
