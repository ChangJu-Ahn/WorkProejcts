<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4110mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/01/23
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Chen Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2	'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i
Dim j

Call HideStatusWnd

On Error Resume Next

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(2)
	Redim UNIValue(2, 2)
	
	UNISqlId(0) = "189702saa"
	UNISqlId(1) = "189702sab"
	UNISqlId(2) = "189702sae"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")

	UNIValue(2, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIvalue(2, 1) = FilterVar("C", "''", "S") 
	UNIValue(2, 2) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
      
	If (rs0.EOF And rs0.BOF) and (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing			
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>

Dim LngMaxRow1
Dim LngMaxRow2
Dim strData1
Dim strData2
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
    IF .parent.CompareDateByFormat("<%=UNIDateClientFormat(rs2("end_dt"))%>",.frm1.txtDlvyDt.text,"","","970025",gDateFormat,gComDateType,False) = False And "<%=rs2("push_flg")%>" = "Y" Then  
		.frm1.txtDlvyDt.text = "<%= UNIDateClientFormat(rs2("end_dt")) %>"
	End If

	LngMaxRow1 = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	LngMaxRow2 = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>		
		ReDim TmpBuffer(<%=rs0.RecordCount-1%>)
<%		
		For i=0 to rs0.RecordCount-1
%>	
			strData1 = ""		
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"			
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"			
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("start_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("due_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData1 = strData1 & Chr(11) & LngMaxRow + <%=i%>
			strData1 = strData1 & Chr(11) & Chr(12)

			TmpBuffer(<%=i%>) = strData1
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
<%	
	End If
%>	
		
<%  
	If Not(rs1.EOF And rs1.BOF) Then
%>
		ReDim TmpBuffer(<%=rs1.RecordCount-1%>)
<%	
		For j=0 to rs1.RecordCount-1
%>			
			strData2 = ""
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_cd"))%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_nm"))%>"			
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("spec"))%>"			
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("start_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("due_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UniConvNumberDBToCompany(rs1("plan_qty"),  ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("tracking_no"))%>"
			strData2 = strData2 & Chr(11) & LngMaxRow + <%=j%>
			strData2 = strData2 & Chr(11) & Chr(12)
			
			TmpBuffer(<%=j%>) = strData2
<%		
			rs1.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		

<%
	End If

		rs0.Close
		Set rs0 = Nothing

		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
