<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : MES 
'*  2. Function Name        : ��ǰ�����Ȳ 
'*  3. Program ID           : xi315ma1_KO119
'*  4. Program Name         : ��ǰ�����Ȳ(�Ѽ�)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'*                            -2002/04/11 : ADO ��ȯ 
'**********************************************************************************************

'								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next
'err.clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q","*", "NOCOOKIE", "QB") 

'============================================  2002-04-10 ����  =============================================
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9, rs10, rs11
															'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim strQryMode

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
DIM strPlantNm		'�����̸�
DIM strSecItemNm	'SECǰ���
Dim MsgDisplayFlag
Dim lgADF
Dim lgstrRetMsg

Dim lgtotMQty          '����԰�
Dim lgtotSampleQty     'Sample�԰� 
Dim lgtotMOutQty       '������ 
Dim lgtotSampleOutQty  'Sample��� 
Dim lgtotVOutQty       '�������  

Dim lgtotMHoldQty      '���Hold��� 
Dim lgtotSHoldQty      'Sample Hold��� 
Dim lgtotMUseQty       '��갡����� 
Dim lgtotSUseQty       'Sample������� 

															'�� : ���� ���ڵ���� ������ŭ �迭 ũ�� ����			
	MsgDisplayFlag = False															
'--------------- ������ coding part(��������,End)----------------------------------------------------------
   	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100						                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"
	strQryMode		 = Request("lgIntFlgMode")

    Call  FixUNISQLData1()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call  FixUNISQLData()
    call  QueryData()                                                               '�� : DB-Agent�� ���� ADO query

'/////////////////////////////////////////////////////////////////////////////////////////////
'Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
'Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
'Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
'Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
'Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
'Dim lgStrPrevKey                                            '�� : ���� �� 
'Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
'Dim lgTailList
'Dim lgSelectList
'Dim lgSelectListDT
'============================================  2002-04-10 ��  ===============================================
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData1()

	Dim strVal
	Dim strVal2
	Dim strVal3
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------	
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"    ' �����ڵ�
	UNISqlId(1) = "commonqry"    'SECǰ���ڵ�		
				
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
		          		   
	'�����ڵ�
    UNIValue(0,0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")					'�����ڵ�
    'SECǰ���
    UNIValue(1,0)  = "SELECT FR2_CD PAR_ITEM_NM from J_CODE_MAPPING (nolock) where major_Cd = 'J0010' and minor_cd = '0000' and FR1_CD = " & FilterVar(Request("txtSecItemCd"), "''", "S") 	

	'--------------- ������ coding part(�������,End)----------------------------------------------------
'	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
		   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)
    
   	If SetConditionData = False Then Exit Sub

End Sub



' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If        
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
	
'	On Error Resume Next
 
	SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPlantNm =  rs1(1)
        rs1.Close	:	Set rs1 = Nothing 
    Else
		rs1.Close	:	Set rs1 = Nothing
		If Len(Request("txtPlantCd")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                parent.frm1.txtPlantCd.Focus()   
                </Script>
            <% 
            Set lgADF = Nothing       		    	
		End If
	End If 
	
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strSecItemNm =  rs2(0)
        rs2.Close	:	Set rs2 = Nothing 
    Else
		rs2.Close	:	Set rs2 = Nothing
		If Len(Request("txtSecItemCd")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("XX1137", vbOKOnly, "", "", I_MKSCRIPT)
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                parent.frm1.txtSecItemCd.Focus()   
                </Script>
            <% 
            Set lgADF = Nothing       		    	
		End If
	End If 
	
	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Dim strVal2
	Dim strVal3
	Dim strVal4
	Dim strVal5
	Redim UNISqlId(9)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------	
  if Trim(Request("txtRadio")) = "MINV" or Trim(Request("txtRadio")) = "SINV" or Trim(Request("txtRadio")) = "MOUT" or Trim(Request("txtRadio")) = "SOUT" then
	Redim UNIValue(9, 3)
  else
    Redim UNIValue(9, 2)
  end if	

  if Trim(Request("txtRadio")) = "VOUT" then   '�������
	UNISqlId(0) = "xi315ma2_KO119"
  elseif Trim(Request("txtRadio")) = "MU" or Trim(Request("txtRadio")) = "SU" then  '��갡��/Sample����
    UNISqlId(0) = "xi315ma3_KO119"
  elseif Trim(Request("txtRadio")) = "MH" or Trim(Request("txtRadio")) = "SH" then  '���Hold���/Sample Hold���
    UNISqlId(0) = "xi315ma1_KO119"
  else
    UNISqlId(0) = "xi315ma4_KO119"      
  end if	
	UNISqlId(1) = "commonqry"    '����԰�		
	UNISqlId(2) = "commonqry"    'Sample�԰�	
	UNISqlId(3) = "commonqry"    '������	
	UNISqlId(4) = "commonqry"    'Sample���
	UNISqlId(5) = "commonqry"    '�������
	UNISqlId(6) = "commonqry"    '���Hold���	
	UNISqlId(7) = "commonqry"    'Sample Hold���		 
	UNISqlId(8) = "commonqry"    '��갡�����
	UNISqlId(9) = "commonqry"    'Sample�������
				
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = " "
	strVal2 = " "
	strVal3 = " "
	strVal4 = " "
	strVal5 = " "
		
    '---�����ڵ�(�ʼ�)
    If Len(Trim(Request("txtPlantCd"))) Then
    	strVal = strval & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " " 
    	strVal2 = strval2 & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "    			
    	strVal3 = strval3 & "  a.plant_cd =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "    			
    End If
	
	'---��������(�ʼ�) 
    If Len(Trim(Request("txtPlanStartDt"))) Then
    	strVal = strval & " AND a.production_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.production_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    	strVal4 = strval4 & " product_date >=  " & FilterVar(uniConvDate(Trim(Request("txtPlanStartDt"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtPlanEndDt"))) Then
    	strVal = strval & " AND a.production_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.production_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    	strVal4 = strval4 & " AND product_date <=  " & FilterVar(uniConvDate(Trim(Request("txtPlanEndDt"))), "''", "S") & ""
    End If
    
    '---MES�۽űⰣ 
    If Len(Trim(Request("txtSendStartDt"))) Then
    	strVal = strval & " AND a.send_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtSendStartDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.send_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtSendStartDt"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtSendEndDt"))) Then
    	strVal = strval & " AND a.send_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtSendEndDt"))), "''", "S") & ""
    	strVal3 = strval3 & " AND a.send_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtSendEndDt"))), "''", "S") & ""
    End If
	
	'---SECǰ���ڵ� 
	If Len(Trim(Request("txtSecItemCd"))) Then
    	strVal5 = strval5 & " AND b.sec_item_cd =  " & FilterVar(Request("txtSecItemCd"), "''", "S") & " "
    End If    
	
	
	'---���� 
	If Trim(Request("txtRadio")) = "MU"  Then  '��갡�����
'		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119)"
		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y'"   	
	elseif Trim(Request("txtRadio")) = "SU" then  'Sample�������
'		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119)"
		strVal = strval & " AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y'"   	
	elseif Trim(Request("txtRadio")) = "MH" then  '���Hold���
	    strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG = 'Y' "
	elseif Trim(Request("txtRadio")) = "SH" then  'Sample Hold���
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG = 'Y' "  			
	elseif Trim(Request("txtRadio")) = "MINV" then  '����԰�
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS'"
	elseif Trim(Request("txtRadio")) = "SINV" then  'Sample �԰�
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS'"
	elseif Trim(Request("txtRadio")) = "MOUT" then  '������
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
	elseif Trim(Request("txtRadio")) = "SOUT" then  'Sample���
		strVal = strval & " AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
	elseif Trim(Request("txtRadio")) = "VOUT" then  '�������
	    	
    End If   	    
     
   if Trim(Request("txtRadio")) = "MINV" or Trim(Request("txtRadio")) = "SINV" or Trim(Request("txtRadio")) = "MOUT" or Trim(Request("txtRadio")) = "SOUT" then
      UNIValue(0,1) = strVal
      UNIValue(0,2) = strVal4
	  UNIValue(0,3) = strVal5 
   else 		   
		UNIValue(0,1) = strVal
		UNIValue(0,2) = strVal5
   end if 
    
	'����԰�
	UNIValue(1,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS'" & strVal5 & ""
	'Sample�԰�
	UNIValue(2,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS'"	& strVal5 & ""
	'������
	UNIValue(3,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' " & strVal5 & ""
	'Sample���
	UNIValue(4,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND ISNULL(A.SEC_INVOICE_NO,'') <> '' " & strVal5 & ""
    '�������
    UNIValue(5,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd inner join  S_Prev_Integrate_Lbl_Ko119  C on  A.PALLET_NO  =  C.PALLET_NO Where " & strVal3 & strVal5 & ""	
    '���Hold���
    UNIValue(6,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG = 'Y' " & strVal5 & ""
	'Sample Hold ���
	UNIValue(7,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG = 'Y' " & strVal5 & ""
    '��갡�����
    UNIValue(8,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd "
	UNIValue(8,0)  = UNIValue(8,0) & "where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119 (nolock)) " & strVal5 & ""
	'���ð������
	UNIValue(9,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  b_item_mapping_ko119 b (nolock) on a.item_cd = b.item_cd "
	UNIValue(9,0)  = UNIValue(9,0) & "where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' and a.pallet_no not in (select pallet_no from S_Prev_Integrate_Lbl_Ko119 (nolock)) " & strVal5 & ""
			

'	UNIValue(3,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' "
'	UNIValue(4,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) = 'PS' "
'	UNIValue(5,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & ""
'	UNIValue(6,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND ISNULL(A.SEC_INVOICE_NO,'') <> '' "
'	UNIValue(7,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal2 & "AND ISNULL(A.SEC_INVOICE_NO,'') = '' AND UPPER(LEFT(A.PRODT_ORDER_NO,2)) <> 'PS' AND A.DELIVERY_HOLD_FG <> 'Y' "
'	UNIValue(8,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) where " & strVal3 & "AND A.DELIVERY_HOLD_FG = 'Y' "
'	UNIValue(9,0)  = "SELECT sum(isnull(A.tray_item_qty,0)) from T_IF_RCV_TRAY_INFO_KO119 A (nolock) inner join  S_Prev_Integrate_Lbl_Ko119  C on  A.PALLET_NO  =  C.PALLET_NO Where " & strVal3 & ""

	'--------------- ������ coding part(�������,End)----------------------------------------------------
'	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
	Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim iStr
	    
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs3, rs4, rs5, rs6, rs7, rs8, rs9, rs10, rs11)
	    
	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)
   
   
	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
		

   	IF NOT (rs3.EOF or rs3.BOF) then          '����԰�
		lgtotMQty     = rs3(0)				
	ELSE
		lgtotMQty     = ""
	End if
    rs3.Close
    Set rs3 = Nothing 

    IF NOT (rs4.EOF or rs4.BOF) then			'Sample�԰�
		lgtotSampleQty = rs4(0)			
	ELSE
		lgtotSampleQty = ""
	End if
    rs4.Close
    Set rs4 = Nothing 
    
    IF NOT (rs5.EOF or rs5.BOF) then		'������
		lgtotMOutQty = rs5(0)			
	ELSE
		lgtotMOutQty = ""
	End if
    rs5.Close
    Set rs5 = Nothing
        
    IF NOT (rs6.EOF or rs6.BOF) then		'Sample���
		lgtotSampleOutQty = rs6(0)			
	ELSE
		lgtotSampleOutQty = ""
	End if
    rs6.Close
    Set rs6 = Nothing  
    
        
    IF NOT (rs7.EOF or rs7.BOF) then		'�������
		lgtotVOutQty = rs7(0)				
	ELSE
		lgtotVOutQty = ""
	End if
    rs7.Close
    Set rs7 = Nothing      


    IF NOT (rs8.EOF or rs8.BOF) then		'���Hold���
		lgtotMHoldQty = rs8(0)				
	ELSE
		lgtotMHoldQty = ""
	End if
    rs8.Close
    Set rs8 = Nothing   
    
    
    IF NOT (rs9.EOF or rs9.BOF) then		'Sample Hold���
		lgtotSHoldQty = rs9(0)				
	ELSE
		lgtotSHoldQty = ""
	End if
    rs9.Close
    Set rs9 = Nothing   
    
    IF NOT (rs10.EOF or rs10.BOF) then		'��갡�����
		lgtotMUseQty = rs10(0)				
	ELSE
		lgtotMUseQty = ""
	End if
   rs10.Close
    Set rs10 = Nothing   
    
    IF NOT (rs11.EOF or rs11.BOF) then		'Sample�������
		lgtotSUseQty = rs11(0)				
	ELSE
		lgtotSUseQty = ""
	End if
	   rs11.Close
    Set rs11 = Nothing     

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
		MsgDisplayFlag = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtPlantCd.focus    
            </Script>
        <%
	Else
		Call  MakeSpreadSheetData()
	End If

'	Call  SetConditionData()
    
End Sub

%>
<Script Language=vbscript>

With Parent 	

	.frm1.txtPlantNm.value			= "<%=ConvSPChars(strPlantNm)%>"
	.frm1.txtSecItemNm.value			= "<%=ConvSPChars(strSecItemNm)%>"
	
	.frm1.txtMassSumQty.text     = "<%=UNINumClientFormat(lgtotMQty,ggQty.DecPoint, 0)%>"      '����԰� 
    .frm1.txtSampleSumQty.text		= "<%=UNINumClientFormat(lgtotSampleQty,ggQty.DecPoint, 0)%>"      'Sample�԰� 
    .frm1.txtMOutSumQty.text    = "<%=UNINumClientFormat(lgtotMOutQty,ggQty.DecPoint, 0)%>"      '������ 
    .frm1.txtSampleOutSumQty.text	= "<%=UNINumClientFormat(lgtotSampleOutQty,ggQty.DecPoint, 0)%>"          'Sample��� 
    .frm1.txtVOutSumQty.text		= "<%=UNINumClientFormat(lgtotVOutQty,ggQty.DecPoint, 0)%>"			  '������� 
    .frm1.txtMHoldSumQty.text = "<%=UNINumClientFormat(lgtotMHoldQty,ggQty.DecPoint, 0)%>"      '���Hold��� 
    .frm1.txtSampleHoldSumQty.text = "<%=UNINumClientFormat(lgtotSHoldQty,ggQty.DecPoint, 0)%>"      'Sample Hold��� 
	.frm1.txtMUseSumQty.text = "<%=UNINumClientFormat(lgtotMUseQty,ggQty.DecPoint, 0)%>"      '��갡����� 
	.frm1.txtSampleUseSumQty.text = "<%=UNINumClientFormat(lgtotSUseQty,ggQty.DecPoint, 0)%>"      'Sample������� 

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area				
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txthPlantCd.value			= "<%=ConvSPChars(Request("txtPlantCd"))%>"		
			.frm1.txthSecItemCd.value		= "<%=ConvSPChars(Request("txtSecItemCd"))%>"									
			.frm1.txthPlanStartDt.value		= "<%=ConvSPChars(Request("txtPlanStartDt"))%>"
			.frm1.txthPlanEndDt.value		= "<%=ConvSPChars(Request("txtPlanEndDt"))%>"
			.frm1.txthSendStartDt.value		= "<%=ConvSPChars(Request("txtSendStartDt"))%>"
			.frm1.txthSendEndDt.value		= "<%=ConvSPChars(Request("txthSendEndDt"))%>"			
		End If      
		
		With parent
			.ggoSpread.Source = .frm1.vspdData 
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>"									'��: Display data 
			.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"								'��: set next data tag
			.frm1.vspdData.Redraw = True
			.DbQueryOk
		End with 
	else
	    parent.dbquerynotok	    	      
    End If  
    
End With     
</Script>
