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

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1								'DBAgent Parameter 선언 
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
Call LoadBasisGlobalInf() 

lgStrPrevKey = Request("lgStrPrevKey")

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
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

	' Plant 명 Display      
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
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
