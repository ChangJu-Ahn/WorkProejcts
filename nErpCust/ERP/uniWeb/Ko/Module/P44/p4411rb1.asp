<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4411rb1.asp
'*  4. Program Name			: List Production Results (Query)
'*  5. Program Desc			: Called By Confirm By Operation and Confirm By Order
'*  6. Comproxy List		: DB Agent (p4411rb1)
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/12/12
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Ryu Sung Won
'* 11. Comment				: COOL:GEN -> DB Agent
'**********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strQryMode
Dim lgStrPrevKey
Dim i
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = Request("lgStrPrevKey")

On Error Resume Next
Err.Clear
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 4)

	UNISqlId(0) = "p4111mb1"
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("lgPlantCD")), "''", "S") & ""
	UNIValue(0, 2) = " " & FilterVar(UCase(Request("lgProdOrdNo")), "''", "S") & ""
	UNIValue(0, 3) = "|"
	UNIValue(0, 4) = "|"
	
	If Request("lgOprCd") <> "" Then
	
		' Operation Results Display
		UNISqlId(1) = "p4411rb1d"
	
		UNIValue(1, 0) = " " & FilterVar(UCase(Request("lgProdOrdNo")), "''", "S") & ""
		UNIValue(1, 1) = " " & FilterVar(UCase(Request("lgOprCd")), "''", "S") & "" 	
	
		If Request("lgStrPrevKey") <> "" Then
			UNIValue(1, 2) = " " & FilterVar(UCase(Request("lgStrPrevKey")), "''", "S") & ""
		Else
			UNIValue(1, 2) = "0"
		End If
	
	Else
	
		' Order Results Display
		UNISqlId(1) = "p4411rb1h"
	
		UNIValue(1, 0) = " " & FilterVar(UCase(Request("lgProdOrdNo")), "''", "S") & ""

		If Request("lgStrPrevKey") <> "" Then
			UNIValue(1, 1) = " " & FilterVar(UCase(Request("lgStrPrevKey")), "''", "S") & ""
		Else
			UNIValue(1, 1) = "0"
		End If
	
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs0)

	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.txtProdOrdNo.value = "<%=ConvSPChars(UCase(Trim(Request("lgProdOrdNo"))))%>"
			parent.txtItemCd.value = "<%=ConvSPChars(rs1("ITEM_CD"))%>"
			parent.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
		</Script>	
		<%
		Set rs1 = Nothing
	End If

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("189600", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
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
	LngMaxRow = .vspdData.MaxRows

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
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Report_Dt"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Shift_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Report_Type"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Prod_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_Unit"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Prod_Qty_In_Base_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Reason_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Remark"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>		
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs0("Seq"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	

	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then
		.initData(LngMaxRow+1)
		.DbQuery
	Else
		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
