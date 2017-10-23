<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3112MA1
'*  4. Program Name         : ���ֳ������ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/11/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho in kuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%													

On Error Resume Next
														
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

Call HideStatusWnd

Dim I1_s_so_hdr	
Dim I2_s_so_hdr	
Const l2_s_so_no = 0
Const l2_s_cfm_flag = 1	
Redim I2_s_so_hdr(1)

Dim iLngRow	
Dim iLngMaxRow
Dim istrData
Dim iStrPrevKey
Dim iStrNextKey
Dim I1_command
Dim istrMode
Dim iPS3G152
Const C_SHEETMAXROWS_D  = 100	
Dim C_s_so_no
Dim C_s_so_seq

'--------------
' ���ֳ������� 
'--------------
Dim EG1_exp_grp													
Const EG1_so_seq         = 0   '---���ּ��� 
Const EG1_hs_no          = 1   'HS��ȣ 
Const EG1_so_price       = 2   '�ܰ� 
Const EG1_net_amt        = 3   '�ݾ�(���ּ��ݾ�[�ŷ�ȭ��])
Const EG1_so_qty         = 4   '����(���ַ�)
Const EG1_bonus_qty      = 5   '������(��������[��]) 
Const EG1_req_qty        = 6   '����û�� 
Const EG1_req_bonus_qty  = 7
Const EG1_bill_qty       = 8   '---������� 
Const EG1_so_unit        = 9   '���� 
Const EG1_lc_qty         = 10
Const EG1_tol_more_rate  = 11  '�����������(+)
Const EG1_tol_less_rate  = 12  '�����������(-)
Const EG1_close_flag     = 13
Const EG1_so_status      = 14  '---����������� 
Const EG1_remark         = 15  '��� 
Const EG1_cust_item_cd   = 16
Const EG1_gi_qty         = 17
Const EG1_gi_bonus_qty   = 18
Const EG1_pre_doc_seq    = 19   '---�������ּ��� 
Const EG1_vat_amt        = 20   'VAT�ݾ� 
Const EG1_dlvy_dt        = 21   '������ 
Const EG1_tracking_no    = 22   'Tracking No
Const EG1_so_base_qty    = 23   '---������ 
Const EG1_bonus_base_qty = 24   '---�������� 
Const EG1_dn_seq         = 25   '�������ϼ��� 
Const EG1_cust_po_seq    = 26
Const EG1_bom_num        = 27
Const EG1_price_flag     = 28   '�ܰ�����(���ܰ�:N, ���ܰ�:Y)
Const EG1_ctp_times      = 29   '---CTP Time
Const EG1_pur_qty        = 30
Const EG1_atp_flag       = 31
Const EG1_pre_doc_no     = 32   '---�������ֹ�ȣ 
Const EG1_ext1_qty       = 33
Const EG1_ext2_qty       = 34
Const EG1_ext3_qty       = 35
Const EG1_ext1_amt       = 36
Const EG1_ext2_amt       = 37
Const EG1_ext3_amt       = 38
Const EG1_ext1_cd        = 39
Const EG1_ext2_cd        = 40
Const EG1_ext3_cd        = 41
Const EG1_net_amt_loc    = 42
Const EG1_vat_amt_loc    = 43
Const EG1_dn_no          = 44   '---�������Ϲ�ȣ 
Const EG1_lot_seq        = 45   'Lot seq
Const EG1_lot_no         = 46   'Lot no
Const EG1_ret_type       = 47   '��ǰ����(��ǰó������)
Const EG1_vat_type       = 48   'VAT���� 
Const EG1_vat_rate       = 49   'VAT�� 
Const EG1_vat_inc_flag   = 50   '---VAT���Ա���(1:����, 2:����)
Const EG1_sl_cd          = 51   'â�� 
Const EG1_sl_nm          = 52   '---â��� 
Const EG1_plant_cd       = 53   '���� 
Const EG1_plant_nm       = 54   '---����� 
Const EG1_item_cd        = 55   'ǰ�� 
Const EG1_item_nm        = 56   'ǰ���  
Const EG1_spec           = 57   '�԰� 
Const EG1_bp_cd          = 58   '��ǰó 
Const EG1_bp_nm          = 59   '---��ǰó�� 

Const EG1_promise_dt      = 60   '���Ͽ�û��(���Ͽ�������)
Const EG1_vat_type_nm     = 61   'VAT������ 
Const EG1_ret_type_nm     = 62   '��ǰ������(��ǰó�����и�)
Const EG1_vat_inc_flag_nm = 63   'VAT���Ա��и�(1:����, 2:����)
Const EG1_aps_host        = 64   '---APSHost 
Const EG1_aps_port        = 65   '---APSPort
Const EG1_flag            = 66   '---CTPCheckFlag 
'---- v3.0 Tracking No ����ä�� 
Const EG1_tracking_flag   = 67 
Const EG1_OldNet_amt      = 68
Const EG1_OriginalNet_amt = 69


'--------------
' ����������� 
'--------------
Dim EG2_exp_grp 
Const EG2_so_dt           = 0
Const EG2_req_dlvy_dt     = 1
Const EG2_cfm_flag        = 2
Const EG2_price_flag      = 3
Const EG2_cur             = 4
Const EG2_net_amt         = 5
Const EG2_cust_po_no      = 6
Const EG2_deal_type       = 7
Const EG2_pay_meth        = 8
Const EG2_vat_inc_flag    = 9
Const EG2_vat_type        = 10
Const EG2_vat_rate        = 11
Const EG2_vat_amt         = 12
Const EG2_pre_doc_no      = 13
Const EG2_ret_item_flag   = 14
Const EG2_export_flag     = 15
Const EG2_so_sts          = 16
Const EG2_maint_no        = 17
Const EG2_auto_dn_flag    = 18
Const EG2_so_type         = 19
Const EG2_bp_cd2          = 20
Const EG2_bp_cd3          = 21
Const EG2_bp_nm3          = 22
Const EG2_vat_inc_flag_nm = 23
Const EG2_ci_flag	      = 24
const EG2_dn_req_flag	  = 25

Const lsConfirm = "CONFIRM"													' Ȯ��ó���� 
istrMode = Request("txtMode")												' ���� ���¸� ���� 

Select Case istrMode

Case CStr(UID_M0001)														' ���� ��ȸ/Prev/Next ��û�� ���� 

    Err.Clear     
    
    '--------------------------------------------------------------------------------------------------------
    ' ���� HDR�� DTL�� �о�´�.
    '--------------------------------------------------------------------------------------------------------

    C_s_so_no = Trim(Request("txtConSoNo"))    
    iStrPrevKey = Trim(Request("lgStrPrevKey"))    
    If iStrPrevKey <> "" then
		C_s_so_seq = iStrPrevKey
    Else
		C_s_so_seq = 0
    End If    
	
    Set iPS3G152 = Server.CreateObject("PS3G152.cSListSoDtl")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If

    Call iPS3G152.S_LIST_SO_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, C_s_so_no, C_s_so_seq, _
                                    EG1_exp_grp, EG2_exp_grp)
    
    '-------------------------
    ' ������������� ������.
    '-------------------------
    If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "203100" Then     
        If CheckSYSTEMError(Err,True) = True Then        
%>     
        <Script Language=vbscript>
 			parent.SetDefaultVal
 			Call parent.SetToolbar("11000000000011")
 			parent.frm1.txtConSoNo.focus
        </Script>
<%
        End If
        Response.End		                                                      
    End If    

    '----------------------------
	' ������������� ǥ���Ѵ�.
	'----------------------------
%>
<Script Language=vbscript>
	With parent		
<%
		Dim lgCurrency																			'�׻� �ŷ�ȭ�� �켱 
		lgCurrency = ConvSPChars(EG2_exp_grp(EG2_cur))
%>
		.frm1.txtCurrency.value = "<%=lgCurrency%>"
		parent.CurFormatNumericOCX	
		.frm1.txtSoldToParty.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_cd3))%>"           '�ֹ�ó 
		.frm1.txtSoldToPartyNm.value	= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_nm3))%>"           '�ֹ�ó�� 
		.frm1.txtCustPoNo.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_cust_po_no))%>"	   '���ֹ���ȣ 
		.frm1.HReqDlvyDt.value			=          "<%=UNIDateClientFormat(EG2_exp_grp(EG2_req_dlvy_dt))%>"	 '---������	
		.frm1.txtNetAmt.Text			= "<%=UNINumClientFormatByCurrency(EG2_exp_grp(EG2_net_amt),lgCurrency,ggAmtOfMoneyNo)%>" '���ּ��ݾ� 
		.frm1.txtHNetAmt.value			= "<%=UNINumClientFormatByCurrency(EG2_exp_grp(EG2_net_amt),lgCurrency,ggAmtOfMoneyNo)%>" '---���ּ��ݾ� 
		.frm1.txtHVATAmt.value			=      "<%=UNINumClientFormatByTax(EG2_exp_grp(EG2_vat_amt),lgCurrency,ggAmtOfMoneyNo)%>" '---�ΰ����� 
		.frm1.txtCurrency.value			= "<%=lgCurrency%>"
		.frm1.txtShipToParty.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_bp_cd2))%>"           '---��ǰó 
		.frm1.txtVatIncFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag))%>"     '�ΰ������� 
		.frm1.txtVatIncFlagNm.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag_nm))%>"  '�ΰ������и� 
		.frm1.txtHVatRate.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_rate))%>"         '---
		.frm1.txtHMaintNo.value 		= "<%=ConvSPChars(EG2_exp_grp(EG2_maint_no))%>"         '---
		.frm1.HPriceFlag.value 			= "<%=ConvSPChars(EG2_exp_grp(EG2_price_flag))%>"       '---
		.frm1.txtHConfirmFlg.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_cfm_flag))%>"         '---
		.frm1.txtSoType.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_so_type))%>"		    '---		
		.frm1.txtHPreSONo.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_pre_doc_no))%>"       '---�������ֹ�ȣ 
		.frm1.HRetItemFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_ret_item_flag))%>"    '---
		.frm1.txtHVATType.value			= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_type))%>"		    '---
		.frm1.txtHVATIncFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag))%>"	    '---	
		.frm1.txtHVATIncFlagNm.value	= "<%=ConvSPChars(EG2_exp_grp(EG2_vat_inc_flag_nm))%>"  '---
		.frm1.txtHPayTermsCd.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_pay_meth))%>"		    '---
		.frm1.txtHDealType.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_deal_type))%>"	    '---
		.frm1.txtHDnReqFlag.value		= "<%=ConvSPChars(EG2_exp_grp(EG2_dn_req_flag))%>"
		
		If "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "Y" Then
			.frm1.RdoConfirm.value = "N"
			.frm1.btnConfirm.value = "Ȯ�����"
		ElseIf "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "N" Then
			.frm1.RdoConfirm.value = "Y"
			.frm1.btnConfirm.value = "Ȯ��ó��"
		Else
			.frm1.RdoConfirm.value = "Y"
			.frm1.btnConfirm.value = "Ȯ��ó��"
		End IF
		
		.frm1.txtHSoNo.value    = "<%=ConvSPChars(Request("txtConSoNo"))%>"		
		.frm1.HExportFlag.value = "<%=ConvSPChars(EG2_exp_grp(EG2_export_flag))%>"			    '���⿩��		
		.frm1.HCiFlag.value     = "<%=ConvSPChars(EG2_exp_grp(EG2_ci_flag))%>"					'�������	
		.frm1.txtSoDt.value     = "<%=UNIDateClientFormat(EG2_exp_grp(EG2_so_dt))%>"			'������ 
 		
		If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet		
		.lgIntFlgMode = OPMD_UMODE				
		parent.HideLotRetField
		
	End With
</Script>		
<%    
  
    '-------------------------
    ' ���ֳ��������� ������.
    '-------------------------
    If CheckSYSTEMError(Err,True) = True Then
        Set iPS3G152 = Nothing		       
%>
		<Script Language=vbscript>
		With parent.frm1
			If .vspdData.MaxRows = 0 Then
				.btnConfirm.disabled   = True
				.btnConfirm.value      = "Ȯ��ó��"
				.btnDNCheck.disabled   = True
				.btnATPCheck.disabled  = True
				.btnCTPCheck.disabled  = True
				.btnAvlStkRef.disabled = True
			End If

			If Trim(.txtHPreSONo.value) <> "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11101011000111"
			ElseIf Trim(.txtHPreSONo.value) = "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11101111001111"
			ElseIf UCase(Trim(.HRetItemFlag.value)) <> "Y" Then
				parent.SetToolbar "11101111001111"
			Else
				parent.SetToolbar "11101111001111"
			End If
			
			parent.SetDefaultPlant
			parent.ChangePlantColor
            .txtConSoNo.focus
		End With
		</Script>
<%

        Response.End                                              
        'Exit Sub
    End If   

    Set iPS3G152 = Nothing
		
	'----------------------------
	' ���ֳ��������� ǥ���Ѵ�.
	'----------------------------
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim iLngMaxRow       
    Dim iLngRow          
    Dim strTemp
    Dim istrData
	    
	With parent
		iLngMaxRow = .frm1.vspdData.MaxRows
<%        
		For iLngRow = 0 To UBound(EG1_exp_grp,1)
		    If iLngRow < C_SHEETMAXROWS_D  Then
		    Else
		       iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq)) 
               Exit For
            End If 	           
%>           
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_cd))%>"     'ǰ���ڵ� 
			istrData = istrData & Chr(11)                                                             'ǰ���ڵ��˾� 
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_nm))%>"     'ǰ��� 
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_spec))%>"        '�԰�	
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_unit))%>"     '���� 
			istrData = istrData & Chr(11)                                                             '�����˾� 
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_tracking_no))%>" 'Tracking No.
			istrData = istrData & Chr(11)                                                             'Tracking No.�˾� 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(cdbl(EG1_exp_grp(iLngRow, EG1_so_qty)) - cdbl(EG1_exp_grp(iLngRow, EG1_req_qty)), ggQty.DecPoint, 0)%>" '�����ܷ� 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_qty), ggQty.DecPoint, 0)%>"          '���� 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_so_price), lgCurrency, ggUnitCostNo)%>" '�ܰ� 
		    istrData = istrData & Chr(11) & "0"                                                       '�ܰ�üũ��ư			
		    
			Select Case "<%=EG1_exp_grp(iLngRow, EG1_price_flag)%>"                                   '�ܰ����� (���ܰ�/���ܰ�)
			Case "Y"
				istrData = istrData & Chr(11) & "���ܰ�"
			Case "N"
				istrData = istrData & Chr(11) & "���ܰ�"
			Case Else
				istrData = istrData & Chr(11)
			End Select

        	If "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>" =  "2" Then   	    	  
		     	istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(cdbl(EG1_exp_grp(iLngRow, EG1_net_amt)) + cdbl(EG1_exp_grp(iLngRow, EG1_vat_amt)), lgCurrency, ggAmtOfMoneyNo)%>"
			Else
				istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>"							
			End If		

			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>" '���ּ��ݾ�(�ŷ�ȭ��)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByTax(EG1_exp_grp(iLngRow, EG1_vat_amt),lgCurrency,ggAmtOfMoneyNo)%>"   'VAT �ݾ� 
    		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_cd))%>"     '�����ڵ� 
			istrData = istrData & Chr(11)                                                                               '�����˾�			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_nm))%>"     '�����									
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_dlvy_dt))%>"      '������			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_cd))%>"        '��ǰó			
			istrData = istrData & Chr(11)                                                                               '��ǰó�˾�			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_nm))%>"        '��ǰó��			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_hs_no))%>"        'HS Code			
			istrData = istrData & Chr(11)                                                                               'Hs Code Popup
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_more_rate), ggExchRate.DecPoint, 0)%>" '�����������(+)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_less_rate), ggExchRate.DecPoint, 0)%>" '�����������(-)
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type))%>"        'VAT Type
			istrData = istrData & Chr(11)                                                                                  'VAT Type Popup
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type_nm))%>"     'VAT Name
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_vat_rate), ggExchRate.DecPoint, 0)%>" 'VAT Rate
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"    'VAT���Ա��� 			
			'istrData = istrData & Chr(11) &                  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag_nm))%>" 'VAT���Ա��и� 
			Select Case "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"
			Case "1"
				istrData = istrData & Chr(11) & "����"
			Case "2"
				istrData = istrData & Chr(11) & "����"
			Case Else
				istrData = istrData & Chr(11)
			End Select
			
     		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type))%>"        '��ǰ���� 
			istrData = istrData & Chr(11)                                                                                  '��ǰ�����˾� 
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type_nm))%>"     '��ǰ������			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_no))%>"          'Lot No			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_seq))%>"         'Lot Seq			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_no))%>"           '�������Ϲ�ȣ			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_seq))%>"          '�������ϼ���			
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_promise_dt))%>"      '���Ͽ�û��											
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_qty), ggQty.DecPoint, 0)%>" '��������(��)			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_cd))%>"           'â���ڵ�			
			istrData = istrData & Chr(11)                                                                                  'â���˾�			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_nm))%>"           'â���			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_remark))%>"          '���			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_status))%>"       '�����������			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bill_qty), ggQty.DecPoint, 0)%>"      '�������			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_base_qty), ggQty.DecPoint, 0)%>"   '������			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_base_qty), ggQty.DecPoint, 0)%>"'��������			
			istrData = istrData & Chr(11) & ""                                                                             '��������			
			istrData = istrData & Chr(11) & ""                                                                             '�ֹ�������			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_host))%>"        'APSHost			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_port))%>"        'APSPort			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ctp_times))%>"       'CTPTimes			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_flag))%>"            'CTPCheckFlag			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq))%>"          '���ּ���			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_no))%>"      '�������ֹ�ȣ			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_seq))%>"     '�������ּ��� 
			'--- v3.0 Tracking No ����ä�� 
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_tracking_flag))%>"
			
			If "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_oldnet_amt), lgCurrency, ggAmtOfMoneyNo)%>" <> "" Then
			istrData = istrData & Chr(11) &					 "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_oldnet_amt), lgCurrency, ggAmtOfMoneyNo)%>" '���ּ��ݾ�(�ŷ�ȭ��)
			Else
			istrData = istrData & Chr(11) &					 "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>" '���ּ��ݾ�(�ŷ�ȭ��)
			End If
			
			If "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_originalnet_amt), lgCurrency, ggAmtOfMoneyNo)%>" <> "" Then
			istrData = istrData & Chr(11) &					 "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_originalnet_amt), lgCurrency, ggAmtOfMoneyNo)%>" '���ּ��ݾ�(�ŷ�ȭ��)
			Else
			istrData = istrData & Chr(11) &					 "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>" '���ּ��ݾ�(�ŷ�ȭ��)
			End If
			
			istrData = istrData & Chr(11) & iLngMaxRow + <%=iLngRow%>			
			istrData = istrData & Chr(11) & Chr(12)			
<%      
		Next
		

%>    
		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip istrData
	
		.lgStrPrevKey = "<%=iStrNextKey%>"
		
    	.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtConSoNo"))%>"   ' Request���� hidden input���� �Ѱ��� 
					
		If "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "Y" Then         'Ȯ����ư ó�� 
			
			If "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "N" Then  'Or "<%=EG2_exp_grp(EG2_so_sts)%>" = 1
				.frm1.btnDNCheck.value = "���Ͽ�ûó��"
				.frm1.btnDNCheck.disabled = True              '���Ϲ�ư ó�� 
			
			ElseIf "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "Y" Then
				If "<%=EG2_exp_grp(EG2_dn_req_flag)%>" = "N" And "<%=EG2_exp_grp(EG2_so_sts)%>" = 2 Then			
					.frm1.RdoDnReq.value = "N"
					.frm1.btnDNCheck.value = "���Ͽ�ûó��"
					.frm1.btnDNCheck.disabled = False
				ElseIf "<%=EG2_exp_grp(EG2_dn_req_flag)%>" = "N" And "<%=EG2_exp_grp(EG2_so_sts)%>" = 1 Then
					.frm1.btnDNCheck.value = "���Ͽ�û���"
					.frm1.btnDNCheck.disabled = True
				ElseIf "<%=EG2_exp_grp(EG2_dn_req_flag)%>" = "Y" And "<%=EG2_exp_grp(EG2_so_sts)%>" = 1 Then
					.frm1.RdoDnReq.value = "Y"
					.frm1.btnDNCheck.value = "���Ͽ�û���"
					.frm1.btnDNCheck.disabled = False
				End If
			End IF								
				.frm1.btnATPCheck.disabled = True                 'ATP CHECK��ư ó�� 

		Else				
			.frm1.btnDNCheck.disabled = True                  '���Ϲ�ư ó�� 							
			.frm1.btnATPCheck.disabled = False                'ATP CHECK��ư ó��	
		End IF
		
		.frm1.txtPlant.value	= "<%=ConvSPChars(EG1_exp_grp(0, EG1_plant_cd))%>"  '��ó�� ǰ���� ������ �Է��Ѵ�.
	    .frm1.txtPlantNm.value	= "<%=ConvSPChars(EG1_exp_grp(0, EG1_plant_nm))%>"
		.DbQueryOk
	
	End With
<%		response.write  istrData & "::"%>

</Script>
<%		
															

Case CStr(UID_M0002)																'��: ���� ��û�� ���� 

	Dim iErrorPosition
	Dim pvCB
    Dim pvICustomXML
    Dim prOCustomXML
    Dim itxtSpread

    iErrorPosition  = ""                                                           '��: Set to space

    itxtSpread		= ""

    For ii = 1 To Request.Form("txtCUSpread").Count
        itxtSpread = itxtSpread & Request.Form("txtCUSpread")(ii)
    Next
    
    If itxtSpread = "" Then
       Response.End
    End If   

    I1_command  = "SAVE"
    I2_s_so_hdr(l2_s_so_no) = Trim(Request("txtHSoNo"))
	    	
    Set iPS3G121 = Server.CreateObject("PS3G121.cSSoDtlSvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If    

	pvCB = "F"
	
    Call iPS3G121.S_SO_DTL_SVR(pvCB, gStrGlobalCollection, I1_command, I2_s_so_hdr, _
                              itxtSpread, iErrorPosition, , prOCustomXML)   
		
	If Trim(iErrorPosition) <> "" Then
			
		If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
			Set iPS3G121 = Nothing	
			%>
			<Script Language=vbscript>
				Dim ii
			    For ii = 1 To parent.divTextArea.children.length
			        parent.divTextArea.removeChild(parent.divTextArea.children(0))
			    Next    
				Call Parent.SubSetErrPos("<%=iErrorPosition%>")
			</Script>
			<%																
			Response.End
		End If
	
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Response.End 
		End If
	End If    

    Set iPS3G121 = Nothing 			
%>
<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
</Script>
<%																

Case "DNCheck"																'��: ���� ���Ͽ�ûó�� ���� 

    Err.Clear                                                               '��: Protect system from crashing
    Dim iPS3G117
    Dim flag
   
    I1_s_so_hdr = Trim(Request("txtSoNo"))
    flag = Trim(Request("RdoDnReq"))
       
    Set iPS3G117 = Server.CreateObject("PS3G117.cSCreateDnBySoSvr")          
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If    

    Call iPS3G117.S_CREATE_DN_BY_SO_SVR(gStrGlobalCollection, I1_s_so_hdr, Flag)    
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G117 = Nothing	
       Response.End		                                               
    End If
	
	Set iPS3G117 = Nothing	

%>
<Script Language=vbscript>
	parent.DbSaveOk()
</Script>		
<%


Case "PRICE"																'��: ���� ���Ͽ�ûó�� ���� 

    Err.Clear                                                               '��: Protect system from crashing
    
    Dim I4_s_so_dtl
    Const S321_I4_so_unit = 0
    Const S321_I4_so_qty = 1
    ReDim I4_s_so_dtl(1)
    
    DIm E1_s_so_dt
    Public Const S321_E1_so_price = 0
    Public Const S321_E1_bonus_qty = 1
    
    Dim pS31121PR
    
    Dim I1_ief_supplied_select_char    
    Dim I2_b_item_item_cd    
    Dim I3_s_so_hdr_so_no 
    
    I1_ief_supplied_select_char = Trim(Request("lsPriceQty"))   
    
    I2_b_item_item_cd = Trim(Request("lsItemCode"))
    
    I3_s_so_hdr_so_no = Trim(Request("txtHSoNo"))
    
    I4_s_so_dtl(S321_I4_so_unit) = Trim(Request("lsSoUnit"))
    I4_s_so_dtl(S321_I4_so_qty) = UNIConvNum(Request("lsSoQty"),0)	

    Set pS31121PR = Server.CreateObject("PS3G112.cSGetSoPriceSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
    
    E1_s_so_dtl = pS31121PR.S_GET_SO_PRICE_SVR (gStrGlobalCollection,I1_ief_supplied_select_char, I2_b_item_item_cd, _
												I3_s_so_hdr_so_no,I4_s_so_dtl)
       
	If CheckSYSTEMError(Err,True) = True Then
       Set pS31121PR = Nothing		                                                 '��: Unload Comproxy DLL
       Response.End 
    End If   

    Set pS31121PR = Nothing	  
%>
<Script Language=vbscript>

	With parent																'��: ȭ�� ó�� ASP �� ��Ī�� 

		.frm1.vspdData.Row = <%=Request("PRow")%>

		Select Case "<%=Trim(Request("lsPriceQty"))%>"
		Case "A"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		Case "P"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
		Case "Q"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		End Select

	End With

</Script>		
<%
'========================================================================================================
' Name : btnCONFIRM
' Desc : Ȯ��ó��/��� ���� 
'========================================================================================================
Case "btnCONFIRM"																	'��: Ȯ��ó�� ��û�� ���� 
									
    Err.Clear																		'��: Protect system from crashing
    Dim iPS3G150
    Redim I1_s_so_hdr(1)
	
	Dim iErrPosition
	
	I1_s_so_hdr(0) = Trim(Request("txtHSoNo"))
	I1_s_so_hdr(1) = Trim(Request("RdoConfirm"))
   
    Set iPS3G150 = Server.CreateObject("PS3G150.cSConfirmSalesOrderSvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G150.S_CONFIRM_SALES_ORDER_SVR(gStrGlobalCollection, I1_s_so_hdr, iErrPosition)   
    
    If Trim(iErrPosition) = "" Then
		If CheckSYSTEMError(Err,True) = True Then
			Set iPS3G150 = Nothing 
		    Response.End		  
		End If 
	Else
		If CheckSYSTEMError2(Err, True, "����" & iErrPosition ,"","","","") = True Then  	
			 Set iPS3G150 = Nothing
		     Response.End
		End If
	End If

    Set iPS3G150 = Nothing 		  
%>
<Script Language=vbscript>
	parent.DbSaveOk()	
</Script>
<%					


Case "ItemByHsCode"															'��: ǰ�񺰿� ���� HS CODE Change

	Dim iPB3C104
	
    Dim I1_b_item    
    Dim prE1_b_item
    Const prE1_item_cd = 0
    Const prE1_item_nm = 1
    Const prE1_formal_nm = 2
    Const prE1_spec = 3
    Const prE1_basic_unit = 4
    Const prE1_item_acct = 5
    Const prE1_item_class = 6
    Const prE1_phantom_flg = 7
    Const prE1_hs_cd = 8
    Const prE1_hs_unit = 9
    Const prE1_unit_weight = 10
    Const prE1_unit_of_weight = 11
    Const prE1_draw_no = 12
    Const prE1_item_image_flg = 13
    Const prE1_blanket_pur_flg = 14
    Const prE1_base_item_cd = 15
    Const prE1_proportion_rate = 16
    Const prE1_valid_flg = 17
    Const prE1_valid_from_dt = 18
    Const prE1_valid_to_dt = 19
    Const prE1_vat_type = 20
    Const prE1_vat_rate = 21
    
    
	Err.Clear
	
	I1_b_item = Trim(Request("ItemCd"))
	
    Set iPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")     
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If
    
    Call iPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, I1_b_item, , , , , prE1_b_item)	
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPB3C104 = Nothing	
       Response.End		                                               
       'Exit Sub
    End If	

%>

<Script Language="vbscript">
		With parent.frm1.vspdData
			.Row 	= "<%=Request("CRow")%>"
			.Col 	= parent.C_ItemName
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_item_nm))%>"
			.Col 	= parent.C_ItemSpec
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_spec))%>"
			
			If Trim(parent.frm1.HExportFlag.value) = "Y" Then 
				.Col 	= parent.C_HsNo
				.text	= "<%=ConvSPChars(prE1_b_item(prE1_hs_cd))%>"
			End If
			
			.Col 	= parent.C_SoUnit
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_basic_unit))%>"

			.Col	= parent.C_VatType

			If .text = "" Then
				If Len(parent.frm1.txtHVATType.value) Then
					.text	= parent.frm1.txtHVATType.value
				Else
					.text	= "<%=ConvSPChars(prE1_b_item(prE1_vat_type))%>"
				End If
			End If
			
			.Col	= parent.C_VatIncFlag

			If .text = "" Then 
				If Len(parent.frm1.txtHVATIncFlag.value) Then
					.Col	= parent.C_VatIncFlag
					.text	= parent.frm1.txtHVATIncFlag.value

					.Col	= parent.C_VatIncFlagNm
					Select Case parent.frm1.txtHVATIncFlag.value
					Case "1"
						parent.frm1.vspdData.Text = "����"
					Case "2"
						parent.frm1.vspdData.Text = "����"
					End Select
				End If			
			End If
			Call parent.SetVatType(<%=Request("CRow")%>)
			
			parent.lsPriceQty = "Q"
			Call parent.GetItemPrice(<%=Request("CRow")%>)
			Call parent.PricePadChange(<%=Request("CRow")%>)
		End With	
</Script>
<%
    Set iPB3C104 = Nothing

'========================================================================================================
' Name : CheckCreditlimit
' Desc : �����ѵ� �ʰ� ���� üũ 
'========================================================================================================
Case "CheckCreditlimit"														
    Err.Clear														'��: Protect system from crashing

	Dim objPS3G113	
	Dim iArrData
	Dim iDblOverLimitAmt
	
	Redim iArrData(2)
    
    iArrData(0) = Trim(Request("txtCaller"))
    iArrData(1) = Trim(Request("txtHSONo"))
    iArrData(2) = Request("txtTotalAmt")
	
	Set objPS3G113 = Server.CreateObject("PS3G113.cChkCreditLimit")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call objPS3G113.ChkCreditLimitSvr(gStrGlobalCollection, iArrData, iDblOverLimitAmt)
    
	Set objPS3G113 = Nothing	
			
	If Err.number = 0 Then
		Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
		Response.Write("Call parent.ConfirmSO()" & vbCr)
		Response.Write("</SCRIPT>" & vbCr)

    Else
   
		' �����ѵ��� �ʰ��� ���(���ó��)
		If InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201929") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			Response.Write("Dim iReturnVal" & vbCr)
			' �����ѵ��� %1 %2 ��ŭ �ʰ��Ͽ����ϴ�. �����Ͻðڽ��ϱ�?
			Response.Write("iReturnVal = parent.DisplayMsgBox(""201929"", parent.parent.VB_YES_NO, parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr )
			Response.Write("If iReturnVal = vbYes Then Call parent.ConfirmSO()" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
			
		ElseIf InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201722") > 0 Then

			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			'�����ѵ��� %1 %2 ��ŭ �ʰ��Ͽ����ϴ�.			
			Response.Write("Call parent.DisplayMsgBox(""201722"", ""X"", parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		Else
			Call CheckSYSTEMError(Err,True)
		End If
	End If
	
	Response.End
End Select


'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>
