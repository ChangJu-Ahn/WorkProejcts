<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4120mb1_ko119.asp
'*  4. Program Name         : List Production Results
'*  5. Program Desc         : 
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2								'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strOprNo
Dim StrNextKey1
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount
Dim i, ii
Dim MaxCount
Dim GridColCount
Dim JobLine1
Dim JobLine2
Dim HiddenCol
Dim JobPlanTime, ProdtOrderNo1, ProdtOrderNo2
Dim Check
Dim FValue
Dim strStartDt
Dim strFlag
Dim ItemCd1, ItemCd2
Dim cboLine
Dim StrItemCd
Dim StrProdOrderNo
Dim strQryMode

Call HideStatusWnd

strMode = Request("txtMode")
strQryMode = Request("lgIntFlgMode")												'☜ : 현재 상태를 받음 
HiddenCol = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "", "SNM")
lgStrPrevKey1 = FilterVar(UCase(Request("lgStrPrevKey1")), "", "SNM")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "", "SNM")

On Error Resume Next

'Dim StrProdOrderNo

Err.Clear

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================

	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
'	UNISqlId(2) = "180000sac"
'	UNISqlId(2) = "180000sam"    
'	UNISqlId(4) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
'	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
'	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
'	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
	</Script>	
	<%

	'Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	
	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	End IF
	
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	
	If strFlag <> "" Then
		%>
		<Script Language=vbscript>
			Call parent.SetFieldColor(False)
		</Script>	
		<%
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
'		ElseIf strFlag = "ERROR_GROUP" Then
'			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
'			Response.Write "<Script Language=VBScript>" & vbCrLf
'				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
'			Response.Write "</Script>" & vbCrLf
'			Set ADF = Nothing
'			Response.End	
		End If
	End IF
																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	' Production Results Display
	Redim UNISqlId(1)
	Redim UNIValue(1, 3)

'	UNISqlId(0) = "P4412MB2"
	UNISqlId(0) = "p4120ma101ko119"
	UNISqlId(1) = "p4120ma102ko119"
	
'	IF Request("txtProdOrderNo") = "" Then
'		strProdOrderNo = "|"
'	Else
'		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
'	End IF
	IF Request("txtProdFromDt") = "" Then
		strStartDt = "" & FilterVar("1900-01-01", "", "SNM") & ""
	Else
		strStartDt = "" & FilterVar(UniConvDate(Request("txtProdFromDt")), "", "SNM") & ""
	End IF
	
	IF Request("txtItemCd") = "" Then
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "", "SNM")
	End IF	

    If Request("txtMaxCount") = "" then
	Else
	   MaxCount = int(Request("txtMaxCount"))
    End if
    
    If Request("GridColCount") = "" then
    Else
		GridColCount = int(Request("GridColCount"))
    End if
    
    If Request("cboLine") = "" then
    Else
		cboLine = Request("cboLine")
    End if
    
    IF Request("txtProdOrderNo") <> "" Then
	    StrProdOrderNo = FilterVar(Request("txtProdOrderNo"),"","SNM") 
	End IF 
	
	
'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "", "SNM")
	UNIValue(0, 1) = strStartDt
	UNIValue(0, 2) = ""
	

	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "", "SNM")
	UNIValue(1, 1) = strStartDt
	UNIValue(1, 2) = ""
	
	If cboLine <> "" then
	UNIValue(0, 2) = UNIValue(0, 2) & " and a.job_line = '" & cboLine & "'"
	UNIValue(1, 2) = UNIValue(1, 2) & " and a.job_line = '" & cboLine & "'"
	End if
	
	if StrProdOrderNo <> "" then
	UNIValue(0, 2) = UNIValue(0, 2) & " and a.prodt_order_no = '" & StrProdOrderNo & "'"
	UNIValue(1, 2) = UNIValue(1, 2) & " and a.prodt_order_no = '" & StrProdOrderNo & "'"
	end if
	
	if StrItemCd <> "" then
	UNIValue(0, 2) = UNIValue(0, 2) & " and a.item_cd = '" & StrItemCd & "'"
	UNIValue(1, 2) = UNIValue(1, 2) & " and a.item_cd = '" & StrItemCd & "'"
	end if
	
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
'			UNIValue(0, 2) = UNIValue(0, 2) & " and a.prodt_order_no >= '" & strProdOrderNo	& "'"
		Case CStr(OPMD_UMODE) 			
			UNIValue(0, 2) = UNIValue(0, 2) & " and (a.job_line > '" & lgStrPrevKey & "'"  
			UNIValue(0, 2) = UNIValue(0, 2) & " or (a.job_line >= '" & lgStrPrevKey & "'"
			UNIValue(0, 2) = UNIValue(0, 2) & " and (a.prodt_order_no > '" & lgStrPrevKey1 & "'"
			UNIValue(0, 2) = UNIValue(0, 2) & " or (a.prodt_order_no >= '" & lgStrPrevKey1 & "'"
			UNIValue(0, 2) = UNIValue(0, 2) & " and (a.item_cd >= '" & lgStrPrevKey2 & "' and a.item_cd >= '" & lgStrPrevKey2 & "'))))) "
			
			UNIValue(1, 2) = UNIValue(1, 2) & " and (a.job_line > '" & lgStrPrevKey & "'"  
			UNIValue(1, 2) = UNIValue(1, 2) & " or (a.job_line >= '" & lgStrPrevKey & "'"
			UNIValue(1, 2) = UNIValue(1, 2) & " and (a.prodt_order_no > '" & lgStrPrevKey1 & "'"
			UNIValue(1, 2) = UNIValue(1, 2) & " or (a.prodt_order_no >= '" & lgStrPrevKey1 & "'"
			UNIValue(1, 2) = UNIValue(1, 2) & " and (a.item_cd >= '" & lgStrPrevKey2 & "' and a.item_cd >= '" & lgStrPrevKey2 & "'))))) "
	End Select
	
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
			parent.DbQueryNotOk
		</Script>	
		<%		
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow, LngMaxRows
Dim strTemp
Dim strData, strData1
Dim TmpBuffer1, TmpBuffer2
Dim iTotalStr1, iTotalStr2
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
<%  
	If Not(rs0.EOF And rs0.BOF) Then
	
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer1(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
<%
		End If
		JobLine1 = ""
		JObLIne2 = ""
		JobPlanTime = ""
		ProdtOrderNo1 = ""
		ProdtOrderNo2 = ""
		ItemCd1 = ""
		ItemCd2 = ""
		check = 0
			
'	For i=0 to MaxCount - 1
    For i=0 to rs0.RecordCount-1
		If i < C_SHEETMAXROWS_D Then	
		JobLine1 = rs0("job_line") 
		ProdtOrderNo1 = rs0("prodt_order_no")
		ItemCd1 = rs0("item_cd")
%>  
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_line"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_line"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prodt_Order_SumQty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			
			
<%    ii = 0
    Do Until ii = GridColCount  
     If Not(rs1.EOF And rs1.BOF) Then 
'        if FValue <> i then
'	         rs1.MoveFirst
'        End if 
          JobPlanTime = ""
          if isempty(rs1("job_line")) then
          else   
			JobLine2 = rs1("job_line")
          end if
          
          if isempty(rs1("prodt_order_no")) then
          else   
			ProdtOrderNo2 = rs1("prodt_order_no")
          end if
          
		  if isempty(rs1("job_plan_time")) then
		  else
		    JobPlanTime = rs1("job_plan_time")
		  end if  
		  
		  if isempty(rs1("item_cd")) then
		  else
		    ItemCd2 = rs1("item_cd")
		  end if  
			
	    if JobLine2 = JobLine1 and ItemCd1 = ItemCd2 and ProdtOrderNo1 = ProdtOrderNo2 and HiddenCol(ii) = JobPlanTime then 
	     Check = 1
%>	
			strData = strData & Chr(11) & "<%=ConvSPChars(HiddenCol(ii))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("job_order_no"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs1("job_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("job_seq"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("checkbox"))%>"
<%      Else 
		Check = 0
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(HiddenCol(ii))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "0"
			strData = strData & Chr(11) & "0"
<%      End if
%>						

<%     
          ii = ii + 1
	      if Check = 0 then
	      else
             rs1.MoveNext
	      end if 
	     if ii >  GridColCount  then
	       rs1.MoveFirst
	     end if
      else
%>            
		strData = strData & Chr(11) & "<%=ConvSPChars(HiddenCol(ii))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "0"
		strData = strData & Chr(11) & "0"
<%    
      end if        
    Loop
%>	
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData

<% 			rs0.MoveNext
			FValue = i
	 end if		
	Next
%>
		iTotalStr1 = Join(TmpBuffer1, "")
		
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr1
		
		.lgStrPrevKey = Trim("<%=ConvSPChars(rs0("job_line"))%>")
		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("item_cd"))%>"
		

<%	
	End If

	rs0.Close
	Set rs0 = Nothing
	
	rs1.Close
	Set rs1 = Nothing
	
%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.htxtItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hcboLine.value		= "<%=ConvSPChars(Request("cboLine"))%>"
	.frm1.hProdFromDt.value		= "<%=Request("txtProdFromDt")%>"
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
