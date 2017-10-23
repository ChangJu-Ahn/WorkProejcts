<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4111rb1.asp
'*  4. Program Name			: List Production Order Detail (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/12/10
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Ryu Sung Won
'* 11. Comment				:
'**********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF										'ActiveX Data Factory
Dim strRetMsg								'Record Set Return Message
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter
Dim rs0										'DBAgent Parameter
Dim strMode
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

    ' Order Header Check
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p4713mb4"
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If

	rs0.Close
	Set rs0 = Nothing

    
	' Order Detail Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p4112pb1"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				If "<%=rs0("Inside_Flg")%>" = "Y" Then
					strData = strData & Chr(11) & "사내"
				Else
					strData = strData & Chr(11) & "외주"
				End If
				strData = strData & Chr(11) & (LngMaxRow + <%=i%>)
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
	
	.lgStrPrevKey = "<%=Trim(rs0("Opr_No"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then	<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
		.InitData LngMaxRow + 1, 1
		.DbQuery
	Else
		.DbQueryOk(LngMaxRow + 1)
	End If
    
    End With
</Script>	
<%    
    Set ADF = Nothing
%>
