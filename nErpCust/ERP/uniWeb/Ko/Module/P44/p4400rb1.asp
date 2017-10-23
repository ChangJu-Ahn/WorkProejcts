<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4400rb1.asp
'*  4. Program Name			: BackFlush Simulation
'*  5. Program Desc			: 
'*  6. Comproxy List		: +PP4G461.cPBackFlushSimulation
'*  7. Modified date(First)	: 2003/06/18
'*  8. Modified date(Last) 	: 2003/06/18
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Park, Bum-Soo
'* 11. Comment		:
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
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")
														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

On Error Resume Next

Dim oPP4G461											'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strTxtSpread
Dim iLngGrpCnt
Dim strData, strData1
Dim EG1_back_simulation_a
Dim EG2_back_simulation_m

Const P462_E1_item_cd = 0
Const P462_E1_item_nm = 1
Const P462_E1_spec = 2
Const P462_E1_tracking_no = 3
Const P462_E1_sl_cd = 4
Const P462_E1_sl_nm = 5
Const P462_E1_to_be_issued_qty = 6
Const P462_E1_base_unit = 7
Const P462_E1_available_qty = 8
Const P462_E1_good_on_hand_qty = 9
Const P462_E1_tot_stk_qty = 10

Const P462_E2_order_no = 0
Const P462_E2_item_cd = 1
Const P462_E2_item_nm = 2
Const P462_E2_spec = 3
Const P462_E2_tracking_no = 4
Const P462_E2_sl_cd = 5
Const P462_E2_sl_nm = 6
Const P462_E2_to_be_issued_qty = 7
Const P462_E2_base_unit = 8
Const P462_E2_issued_qty = 9
Const P462_E2_consumed_qty = 10
Const P462_E2_available_qty = 11
Const P462_E2_good_on_hand_qty = 12
Const P462_E2_tot_stk_qty = 13

    Err.Clear											'��: Protect system from crashing

    strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

    LngMaxRow = CInt(Request("txtMaxRows"))				'��: �ִ� ������Ʈ�� ���� 
    
    Set oPP4G461 = CreateObject("PP4G461.cPBackFlushSimulation")
    
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	strTxtSpread = Request("txtSpread")
	
	Call oPP4G461.P_BACKFLUSH_Simulation(	gStrGlobalCollection, _
											strTxtSpread, _
											EG1_back_simulation_a, _
											EG2_back_simulation_m)
	
    If CheckSYSTEMError(Err,True) = True Then
		Set oPP4G461 = Nothing
		Response.End
	End If

	If Not (oPP4G461 is nothing)  Then
		Set oPP4G461 = Nothing
	End If

	If Not IsNull(EG1_back_simulation_a) and Not IsEmpty(EG1_back_simulation_a)Then
		iLngGrpCnt = UBound(EG1_back_simulation_a, 1)
		    
		For iLngRow = 0 To iLngGrpCnt
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_item_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_item_nm))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_spec))
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, P462_E1_to_be_issued_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_base_unit))
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, P462_E1_good_on_hand_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, P462_E1_available_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & UniConvNumberDBToCompany(EG1_back_simulation_a(iLngRow, P462_E1_tot_stk_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_sl_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_sl_nm))
			strData = strData & Chr(11) & ConvSPChars(EG1_back_simulation_a(iLngRow, P462_E1_tracking_no))
			strData = strData & Chr(11) & iLngMaxRow + iLngRow
			strData = strData & Chr(11) & Chr(12)
		Next
	End If

	If Not IsNull(EG2_back_simulation_m) and Not IsEmpty(EG2_back_simulation_m)Then
		iLngGrpCnt = UBound(EG2_back_simulation_m, 1)
		    
		For iLngRow = 0 To iLngGrpCnt
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_item_cd))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_item_nm))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_spec))
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow, P462_E2_to_be_issued_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_base_unit))
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow, P462_E2_issued_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow, P462_E2_consumed_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow, P462_E2_available_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & UniConvNumberDBToCompany(EG2_back_simulation_m(iLngRow, P462_E2_tot_stk_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_order_no))
			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_back_simulation_m(iLngRow, P462_E2_tracking_no))
			strData1 = strData1 & Chr(11) & iLngMaxRow + iLngRow
			strData1 = strData1 & Chr(11) & Chr(12)
		Next
	End If

	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf										'��: ȭ�� ó�� ASP �� ��Ī�� 

	If IsEmpty(EG1_back_simulation_a) = False Then
		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & strData & """" & vbCrLf
	End If
	If IsEmpty(EG2_back_simulation_m) = False Then
		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & strData1 & """" & vbCrLf
	End If

	Response.Write ".DbQueryOk()" & vbCrLf

	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	Response.End

%>
