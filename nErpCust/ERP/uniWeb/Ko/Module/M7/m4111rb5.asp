<%
'**********************************************************************************************
'*  1. Module Name			: P
'*  2. Function Name		: 
'*  3. Program ID			: m4111rb5.asp
'*  4. Program Name			: BackFlush Simulation
'*  5. Program Desc			: 
'*  6. Comproxy List		: +PM7CSBF.cMBackFlushSimulation
'*  7. Modified date(First)	: 2003/06/18
'*  8. Modified date(Last) 	: 2005/10/27
'*  9. Modifier (First)		: KIM, JIHYUN
'* 10. Modifier (Last)		: KIM DUKHYUN
'* 11. Comment				:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE","PB")
														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd
 
On Error Resume Next

Dim OBJ_PM7CSBF											'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strTxtSpread, strSupplierCd
Dim iLngGrpCnt, iLngGrpCnt1
Dim iLngRow, iLngRow1
Dim strData, strData1
Dim EG1_back_simulation_a
Dim EG2_back_simulation_m


    ' Export View for Auto Issue
	Const M462_E1_ParCnt = 0
	Const M462_E1_PoNo = 1
	Const M462_E1_ParItemCd = 2
	Const M462_E1_ParItemNm = 3
	Const M462_E1_Issuemthd = 4
	Const M462_E1_ChildItemCd = 5
	Const M462_E1_ChildItemNm = 6
	Const M462_E1_ChildItemSpec = 7
	Const M462_E1_BaseUnit = 8
	Const M462_E1_ReqmtQty = 9
	Const M462_E1_IOnHandQty = 10
	Const M462_E1_OOnHandQty = 11
	Const M462_E1_IssueQty = 12
	Const M462_E1_SpplType = 13
        
    ' Export View for Manual Issue
	Const M462_E2_PoSeqNo = 0
	Const M462_E2_PoNo = 1
	Const M462_E2_ParItemCd = 2
	Const M462_E2_ParItemNm = 3
	Const M462_E2_Issuemthd = 4
	Const M462_E2_ChildItemCd = 5
	Const M462_E2_ChildItemNm = 6
	Const M462_E2_ChildItemSpec = 7
	Const M462_E2_BaseUnit = 8
	Const M462_E2_ReqmtQty = 9
	Const M462_E2_IOnHandQty = 10
	Const M462_E2_OOnHandQty = 11
	Const M462_E2_IssueQty = 12
	Const M462_E2_SpplType = 13


  Err.Clear												'��: Protect system from crashing

    strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
 
	LngMaxRow = CInt(Request("txtMaxRows"))				'��: �ִ� ������Ʈ�� ���� 
   
    Set OBJ_PM7CSBF = CreateObject("PM7CSBF.cMBackFlushSimulation")
    
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If
	

	strSupplierCd = Request("txtSupplierCd")
	strTxtSpread = Request("txtSpread")

	
	Call OBJ_PM7CSBF.M_BACKFLUSH_SIMULATION(gStrGlobalCollection, _
											strSupplierCd, _
											strTxtSpread, _
											EG1_back_simulation_a, _
											EG2_back_simulation_m)
	
    If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM7CSBF = Nothing
		Response.End
	End If

	If Not (OBJ_PM7CSBF is nothing)  Then
		Set OBJ_PM7CSBF = Nothing
	End If

	If Not IsNull(EG1_back_simulation_a) Then
		iLngGrpCnt = UBound(EG1_back_simulation_a, 1)
		
		For iLngRow = 0 To iLngGrpCnt
			If Cdbl(EG1_back_simulation_a(iLngRow, M462_E1_ParCnt)) > 1 Then
				strData = strData & Chr(11) & "����"
			Else
				strData = strData & Chr(11) & "����"
			End If		
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_ParItemCd))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_ParItemNm))
			If Trim(UCase(ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_IssueMthd)))) = "A" Then
				strData = strData & Chr(11) & "�ڵ�"
			Elseif Trim(UCase(ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_IssueMthd)))) = "M" Then
				strData = strData & Chr(11) & "����"
			End If
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_ChildItemCd))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_ChildItemNm))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_ChildItemSpec))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, M462_E1_BaseUnit))						
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, M462_E1_ReqmtQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, M462_E1_IOnhandQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, M462_E1_OOnHandQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, M462_E1_IssueQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)			
			strData = strData & Chr(11) & iLngMaxRow + iLngRow
			strData = strData & Chr(11) & Chr(12)
		Next
	End If


	If Not IsNull(EG2_back_simulation_m) Then
		iLngGrpCnt1 = UBound(EG2_back_simulation_m, 1)
		    
		For iLngRow1 = 0 To iLngGrpCnt1 
	   		strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_PoSeqNo))		
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_ParItemCd))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_ParItemNm))
			
'			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_Issuemthd))
			If Trim(UCase(ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_Issuemthd)))) = "A" Then
				strData1 = strData1 & Chr(11) & "�ڵ�"
			ElseIf Trim(UCase(ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_Issuemthd)))) = "M" Then
				strData1 = strData1 & Chr(11) & "����"
			End If
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_ChildItemCd))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_ChildItemNm))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_ChildItemSpec))
	   		strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_BaseUnit))
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow1, M462_E2_ReqmtQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow1, M462_E2_IOnHandQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow1, M462_E2_OOnHandQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow1, M462_E2_IssueQty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
'	=== 2005.07.04 ��ޱ��� �߰� =====================================================================================
'			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_SpplType))
			If Trim(UCase(ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_SpplType)))) = "F" Then
				strData1 = strData1 & Chr(11) & "����"
			ElseIf Trim(UCase(ConvSPChars(EG2_back_simulation_m(iLngRow1, M462_E2_SpplType)))) = "C" Then
				strData1 = strData1 & Chr(11) & "����"
			End If			
'	=== 2005.07.04 ��ޱ��� �߰� =====================================================================================			
			strData1 = strData1 & Chr(11) & iLngMaxRow + iLngRow1
			strData1 = strData1 & Chr(11) & Chr(12)
		Next
	End If


	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf										'��: ȭ�� ó�� ASP �� ��Ī�� 

	If IsEmpty(EG1_back_simulation_a) = False Then
		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowData """ & strData & """" & vbCrLf
	End If

	If IsEmpty(EG2_back_simulation_m) = False Then
		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
		Response.Write ".ggoSpread.SSShowData """ & strData1 & """" & vbCrLf
	End If	
	
	Response.Write ".DbQueryOk()" & vbCrLf

	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	Response.End

%>
