<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4110mb2.asp
'*  4. Program Name			: Execute Order Explosion in batch
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP2G101.cPExecMrpSvr
'*  7. Modified date(First)	: 
'*  8. Modified date(Last) 	: 2002/08/20
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Call HideStatusWnd

On Error Resume Next									'��: 

'--------------------------------------------------------------------------------------------------------------------
Dim ADF													'ActiveX Data Factory ���� �������� 
Dim strRetMsg											'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0			'DBAgent Parameter ���� 


Dim strMode												'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim strStatus
Dim prErrorPosition

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

	lgStrPrevKey = Request("lgStrPrevKey")
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	UNISqlId(0) = "189702sae"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(0, 2) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	If NOT(rs0.EOF And rs0.BOF) Then
		strStatus = rs0("confirm_flg")
	  IF strStatus = "Y" THEN
		%>
		<Script Language=vbscript>
			With parent
<%				IF rs0("inv_flg") = "Y" THEN
%>
					.frm1.chkInvStock.checked = True
<%				ELSE
%>
					.frm1.chkInvStock.checked = False
<%				END IF	
%>
				
<%				IF rs0("ss_flg") = "Y" THEN
%>
					.frm1.chkSFStock.checked = True
<%				ELSE
%>
					.frm1.chkSFStock.checked = False
<%				END IF	
%>
				
<%				IF rs0("push_flg") = "Y" THEN
%>
					.frm1.chkForward.checked = True
<%				ELSE
%>
					.frm1.chkForward.checked = False
<%				END IF												
%>
			End With
		</script>	
<%	
	  END IF
	    If strStatus = "Y" Then	'��: ������ ���� ���� ���Դ��� üũ 
			Call DisplayMsgBox("187743", vbOKOnly, "", "", I_MKSCRIPT)
			rs0.Close
			Set rs0 = Nothing
			Set ADF = Nothing         
			Response.End
	    End If
	END IF

	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing

'--------------------------------------------------------------------------------------------------------------------

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim strcurrentdate
Dim pPP2G105												'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim dtf, ptf, PlanDt
Dim mpsscope
Dim I1_mrp_parameter


Err.Clear																				'��: Protect system from crashing
	
	Const P202_I1_plant_cd = 0    '[CONVERSION INFORMATION]  View Name : import mrp_parameter
	Const P202_I1_current_date = 1
	Const P202_I1_plan_date = 2
	Const P202_I1_open_date = 3
	Const P202_I1_flag = 4
	Const P202_I1_safe_flg = 5
	Const P202_I1_inv_flg = 6
	Const P202_I1_idep_flg = 7
	Const P202_I1_option_flg = 8
	Const P202_I1_item_cd = 9
	Const P202_I1_warning_flg = 10
	Const P202_I1_order_no = 11
	Const P202_I1_codr_flg = 12
	Const P202_I1_net_flg = 13
	Const P202_I1_pack_flg = 14
	Const P202_I1_scrap = 15
	Const P202_I1_forward = 16
	Const P202_I1_mpsscope = 17
	    
	Redim I1_mrp_parameter(P202_I1_mpsscope)
	
	Err.Clear																				'��: Protect system from crashing

    '-----------------------
    'Data manipulate area
    '-----------------------
    strcurrentdate = UniConvDate(GetSvrDate)
   
    I1_mrp_parameter(P202_I1_plant_cd) = UCase(Trim(Request("txtPlantCd")))										'��: Plant Code    

    I1_mrp_parameter(P202_I1_current_date) = UniConvDateToYYYYMMDD(Request("txtExecFromDt"),gDateFormat,"")
                                      
    I1_mrp_parameter(P202_I1_plan_date) =    Year(UniConvDate(Request("txtEndDt"))) & _
                                            Right("0" & Month(UniConvDate(Request("txtEndDt"))),2) & _
                                            Right("0" & Day(UniConvDate(Request("txtEndDt"))),2)
                                            
    I1_mrp_parameter(P202_I1_open_date) =    Year(UniConvDate(Request("txtEndDt"))) & _
                                            Right("0" & Month(UniConvDate(Request("txtEndDt"))),2) & _
                                            Right("0" & Day(UniConvDate(Request("txtEndDt"))),2)
                                            
    I1_mrp_parameter(P202_I1_flag) = ""

    If Request("chkSFStock") = "True" Then
         I1_mrp_parameter(P202_I1_safe_flg)  = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_safe_flg)  = "N"
    End If

    If Request("chkInvStock") = "True" Then
         I1_mrp_parameter(P202_I1_inv_flg)  = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_inv_flg)  = "N"    	 
    End If

    I1_mrp_parameter(P202_I1_idep_flg) = "Y"
    I1_mrp_parameter(P202_I1_option_flg) = "N"
    I1_mrp_parameter(P202_I1_item_cd) = UCase(Trim(Request("txtItemCd")))
'    pP23132.ImportMrpParameterIsrtUserId = Request("txtUserId")
    I1_mrp_parameter(P202_I1_warning_flg) = "N"
    I1_mrp_parameter(P202_I1_order_no) = UCase(Request("txtPlanOrderNO"))
    I1_mrp_parameter(P202_I1_codr_flg) = "N"
    If Request("chkInvStock") = "True" Then
         I1_mrp_parameter(P202_I1_net_flg)  = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_net_flg)  = "N"    	 
    End If

    I1_mrp_parameter(P202_I1_pack_flg) = "N"
    I1_mrp_parameter(P202_I1_scrap) = ""

   
    If Request("chkForWard") = "True" Then
         I1_mrp_parameter(P202_I1_forward) = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_forward) = "N"
    End If

    I1_mrp_parameter(P202_I1_mpsscope) = ""

    '-----------------------
    'Com Action Area
    '-----------------------
    Set pPP2G105 = Server.CreateObject("PP2G105.cPOrderExecMrpSvr")
	    
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G105 = Nothing		
		Response.End
	End If
	
	Call pPP2G105.P_ORDER_EXEC_MRP_SVR(gStrGlobalCollection, I1_mrp_parameter,prErrorPosition)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G105 = Nothing															'��: Unload Component
		Response.End
	Else
	   If prErrorPosition = 0 then
	      Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
	      
	   Else
	      Call DisplayMsgBox("209002", vbInformation,"������" &  (prErrorPosition) , "�����޽����˾�", I_MKSCRIPT)
	      
	   End If
	      	
	End If

	Set pPP2G105 = Nothing   
	
	
    
%>
<Script Language=vbscript>
	parent.FncQuery
</script>	
