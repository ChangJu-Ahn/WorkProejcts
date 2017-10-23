<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4116rb1.asp
'*  4. Program Name         : List Conversion History
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002-04-25
'*  7. Modified date(Last)  : 2002/12/20
'*  8. Modifier (First)     : Park , Bumsoo
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1								'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

Dim StrProdOrderNo
Dim StrPRNo
Dim strFlag

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

	%>
	<Script Language=vbscript>
		parent.txtPlantNm.value = ""
	</Script>	
	<%

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		%>
		<Script Language=vbscript>
			parent.txtPlantNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If


	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtPlantNm.Focus()
			</Script>	
			<%
			Response.End
		End If
	End IF

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	If Trim(Request("txtProdtOrderNo")) = "" Then
				StrProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	End If

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			StrPRNo = "|"
		Case CStr(OPMD_UMODE)
			If Trim(Request("lgStrPrevKey")) = "" Then
				StrPRNo = "|"
			Else
				StrPRNo = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
			End If
	End Select		

	UNISqlId(0) = "P4116RB1"
	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo
	UNIValue(0, 3) = StrPRNo
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .vspdData.MaxRows										'Save previous Maxrow
		
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PR_NO"))%>"								'PR
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REQ_DT"))%>"						'요청일 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DLVY_DT"))%>"					'필요일	
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"	'요청수량 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_UNIT"))%>"							'요청단위 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>"							'구매조직 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"								'입고창고 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_PRSN"))%>"							'요청자 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_DEPT"))%>"							'요청부서 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"						'제조지시상태 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("INSRT_DT"))%>"					'변환일 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"								'비고 
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
		
		.lgStrPrevKey = "<%=ConvSPChars(rs0("PR_NO"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.hProdOrderNo.value		= "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
	.DbQueryOk()

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
