<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb1.asp
'*  4. Program Name         : Routing Detail Query
'*  5. Program Desc         :
'*  6. Comproxy List        : +P12038ListRoutingDetail
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im, HyunSoo
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

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
Err.Clear																	'☜: Protect system from crashing
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(1, 1) = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtResourceNm.value = ""
	</Script>	
	<%    	

	' Plant 명 Display      
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
		parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If

	' 자원명 Display
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

	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "p1503mb1"
	
	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtResourceCd")), "''", "S") & " "
	UNIValue(0, 2) = " " & FilterVar(lgStrPrevKey, "''", "S") & "" 

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)

	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("181800", vbOKOnly, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim LngRow
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Shift_Cd"))%>"			'1
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Description"))%>"		'3
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
		
		.lgStrPrevKey = "<%=Trim(rs2("Shift_Cd"))%>"
		
<%	
	End If

	rs2.Close
	Set rs2 = Nothing

%>	
		If .frm1.vspdData.MaxRows < parent.parent.VisibleRowCnt(.frm1.vspdData, 0) and .lgStrPrevKey <> "" Then
			.DbQuery
		Else
			.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
			.frm1.hResourceCd.value = "<%=ConvSPChars(Request("txtResourceCd"))%>"    
			.DbQueryOk
		End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
