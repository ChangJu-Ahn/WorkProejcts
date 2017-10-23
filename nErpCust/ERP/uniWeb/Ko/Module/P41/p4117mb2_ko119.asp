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
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
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

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1								'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strOprNo
Dim StrNextKey1

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
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

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
HiddenCol = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)

On Error Resume Next

Dim StrProdOrderNo

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
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
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow, LngMaxRows
Dim strTemp
Dim strData, strData1
Dim TmpBuffer1, TmpBuffer2
Dim iTotalStr1, iTotalStr2
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
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
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
