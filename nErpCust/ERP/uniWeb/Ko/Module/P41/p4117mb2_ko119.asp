<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4117mb2_ko119.asp
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
Dim rs0, rs1								'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strOprNo
Dim StrNextKey1

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
Dim JobLine
Dim JobLine2
Dim HiddenCol
Dim JobPlanTime
Dim Check
Dim FValue

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
HiddenCol = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)

On Error Resume Next

Dim StrProdOrderNo

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	' Production Results Display
	Redim UNISqlId(1)
	Redim UNIValue(1, 3)

'	UNISqlId(0) = "P4412MB2"
	UNISqlId(0) = "p4117ma102ko119"
	UNISqlId(1) = "p4117ma103ko119"
	
	IF Request("txtProdOrderNo") = "" Then
'		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End IF

'	IF Request("txtOprNo") = "" Then
'		strOprNo = "|"
'	Else
'		StrOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
'	End IF

    If Request("txtMaxCount") = "" then
	Else
	   MaxCount = int(Request("txtMaxCount"))
    End if
    
    If Request("GridColCount") = "" then
    Else
		GridColCount = int(Request("GridColCount"))
    End if

'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = StrProdOrderNo

	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = StrProdOrderNo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
	If (rs1.EOF And rs1.BOF) Then
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
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
<%
	If Not(rs0.EOF And rs0.BOF) Then
%>	
		ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
<%	    
		JobLine = ""
		JObLIne2 = ""
		JobPlanTime = ""
		check = 0
			
	For i=0 to MaxCount - 1
		JobLine = rs0("job_line") 
%>  
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_line"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_line"))%>"
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
          
		  if isempty(rs1("job_plan_time")) then
		  else
		    JobPlanTime = rs1("job_plan_time")
		  end if  
			
	    if JobLine2 = JobLine and HiddenCol(ii) = JobPlanTime then 
	     Check = 1
%>	
			strData = strData & Chr(11) & "<%=ConvSPChars(HiddenCol(ii))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("job_order_no"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs1("job_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
<%      Else 
		Check = 0
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(HiddenCol(ii))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			
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
<%    
      end if        
    Loop
%>	
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData

<% 			rs0.MoveNext
			FValue = i
	Next
%>
		iTotalStr1 = Join(TmpBuffer1, "")
		
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr1

<%	
	End If

	rs0.Close
	Set rs0 = Nothing
	
	rs1.Close
	Set rs1 = Nothing
	
%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.DbDtlQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
