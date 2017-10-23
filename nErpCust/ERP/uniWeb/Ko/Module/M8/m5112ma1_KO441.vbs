
Option Explicit																	'��: indicates that All variables must be declared in advance




Const ivType = "ST"

Dim C_PlantCd						  '���� 
Dim c_PlantPopUP                      '���� �˾� 
Dim C_PlantNm						  '����� 
Dim C_ItemCd						  'ǰ�� 
Dim C_ItemPopup                       'ǰ�� �˾� 
Dim C_ItemNm                          'ǰ��� 

Dim C_SpplSpec                        'ǰ�� �԰� �߰� 

Dim C_IvQty1                          '���Լ��� 
Dim C_Unit                            '���� 
Dim C_UnitPopup                       '�����˾� 
Dim C_Cost                            '���Դܰ� 
Dim C_IvAmt		                      '���رݾ� 
Dim C_NetAmt                          '���Աݾ� => "���Լ��ݾ�"���� ����(2002-06-19)
Dim C_OrgNetAmt                       '���Աݾ� 

'�߰� 
Dim C_IOFlg		                       '���Ա��и� 
Dim C_IOFlgCd	                       '���Ա����ڵ� 

Dim C_VatType                          'vat
Dim C_VatPopup                         'vat�˾� 
Dim C_VatNm                            'vat�� 
Dim C_VatRate                          'vat�� 


Dim C_VatDocAmt                       'VAT�ݾ� 
Dim C_OrgVatDocAmt                    'VAT�ݾ� 
Dim C_IvLocAmt						  '�����ڱ��ݾ� 
Dim C_NetLocAmt                       '�����ڱ��ݾ� => "�����ڱ����ݾ�"���� ����(2002-06-19) 
Dim C_VatLocAmt                       'VAT�ڱ��ݾ� 
Dim C_Remark						  '����߰� -> 2005.12.19
Dim C_OrderQty                        '���ּ��� 
Dim C_OrderCost                       '���ִܰ� 

Dim C_GmQty                           '�԰���� 
Dim C_IvQty2                          '���ԿϷ���� 
Dim C_PoNo                            '���ֹ�ȣ 
Dim C_PoSeq                           '���ּ��� 
Dim C_MvmtRcptNo                      '�԰��ȣ 
Dim C_GmNo                            '���ó����ȣ 
Dim C_GmSeq                           '���ó������ 
Dim C_IvSeq                           '���Լ��� 
Dim C_OldQty                          'hidden

Dim C_MvmtNo                          'hidden
Dim C_MvmtIvQty                       '����Լ��� 
Dim C_oldIvQty1   
Dim C_vat_rvs_flg                     '���� ���� �񱳰� 
Dim C_chkVatDocAmt                    'vat ���� ���� 
Dim C_ref_vatrate_flg 
Dim C_TrackingNo  
'2007.04.16 �߰� 
Dim C_TrackingPopup
Dim C_ChgNetAmt                       '���Աݾ�(HIDDEN ����Ǵ� ��)
Dim C_ChgVatDocAmt                    'VAT�ݾ�(HIDDEN ����Ǵ� ��)
'2002.09.10 �߰� 
Dim C_PoAmt                           '���ֱݾ� 
Dim C_MvmtAmt                         '�԰�ݾ� 
Dim C_TotIvDocAmt                     '����Աݾ� 
Dim C_upt_amt_flg                     '���ֱݾ� ���ſ��� 
Dim C_prcflg                          '�ܰ�(ǥ�شܰ�(S),�̵���մܰ�(M))
Dim C_PoVatAmt                        '���� vat�ݾ� 
Dim C_TotIvVatAmt                     '����� vat�ݾ� 
Dim C_PoIvQty     
Dim C_retflg      
Dim C_ref_flg      
Dim C_Stateflg	
Dim C_PoVatIncFlg
'2003-02-20�߰� : �Ѽ����ڿ䱸����.
Dim C_LCNo
'####��LC�߰�#####
Dim C_LCSeqNo
'����(2003.03.24)***
Dim C_LcFlg								'LC���� 
Dim C_XchRt								'���Գ���ȯ��(L/Cȯ�� �����ϵ��� ��) - 2003.09.19

'//�����ΰ�� ����ϴ� ��� 
'Dim C_NOT_USED1
Dim C_Cost_Ref    
Dim C_IvAmt_Ref 
                
'==========================================  1.2.2 Global ���� ����  =====================================

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey

'ȭ�鿡 ������ �ʴ� �ǿ� ���� �ݾ� 
Dim lgTotalIvAmt		
Dim lgTotalNetAmt		
Dim lgTotalDocAmt

Dim IsOpenPop          
Dim lblnWinEvent
Dim interface_Account
Dim arrCollectVatType

'============================  initSpreadPosVariables() =================================================
Sub initSpreadPosVariables()  
	C_PlantCd     		= 1                 '���� 
	c_PlantPopUP  		= 2                 '���� �˾� 
	C_PlantNm     		= 3                 '����� 
	C_ItemCd      		= 4                 'ǰ�� 
	C_ItemPopup   		= 5                 'ǰ�� �˾� 
	C_ItemNm      		= 6                 'ǰ��� 
	C_SpplSpec    		= 7                 'ǰ�� �԰� �߰� 
	C_IvQty1      		= 8                 '���Լ��� 
	C_Unit        		= 9                 '���� 
	C_UnitPopup   		= 10                '�����˾� 
	C_Cost        		= 11                '���Դܰ� 
	C_IvAmt		  		= 12                '���رݾ� 
	C_NetAmt      		= 13                '���Աݾ� => "���Լ��ݾ�"���� ����(2002-06-19)
	C_OrgNetAmt   		= 14                '���Աݾ� 
	'�߰�         		                
	C_IOFlg		  		= 15                '���Ա��и� 
	C_IOFlgCd	  		= 16                '���Ա����ڵ� 
	C_VatType     		= 17                'vat
	C_VatPopup    		= 18                'vat�˾� 
	C_VatNm       		= 19                'vat�� 
	C_VatRate     		= 20                'vat�� 
	C_VatDocAmt   		= 21                'VAT�ݾ� 
	C_OrgVatDocAmt		= 22                'VAT�ݾ� 
	C_IvLocAmt	  		= 23                '�����ڱ��ݾ� 
	C_NetLocAmt   		= 24                '�����ڱ��ݾ� => "�����ڱ����ݾ�"���� ����(2002-06-19) 
	C_VatLocAmt   		= 25                'VAT�ڱ��ݾ� 
	C_Remark			= 26            
	C_OrderQty    		= 27                '���ּ��� 
	C_OrderCost   		= 28                '���ִܰ� 
	C_GmQty       		= 29                '�԰���� 
	C_IvQty2      		= 30                '���ԿϷ���� 
	C_PoNo        		= 31                '���ֹ�ȣ 
	C_PoSeq       		= 32                '���ּ��� 
	C_MvmtRcptNo  		= 33                '�԰��ȣ 
	C_GmNo        		= 34                '���ó����ȣ 
	C_GmSeq       		= 35                '���ó������ 
	C_IvSeq       		= 36                '���Լ��� 
	C_OldQty      		= 37                'hidden
	C_MvmtNo      		= 38                'hidden
	C_MvmtIvQty   		= 39                '����Լ��� 
	C_oldIvQty1   		= 40            
	C_vat_rvs_flg 		= 41                '���� ���� �񱳰� 
	C_chkVatDocAmt		= 42                'vat ���� ���� 
	C_ref_vatrate_flg 	= 43            
	C_TrackingNo  		= 44
	'2007-04-16 added            
	C_TrackingPopup		= 45
	C_ChgNetAmt   		= 46                '���Աݾ�(HIDDEN ����Ǵ� ��)
	C_ChgVatDocAmt		= 47                'VAT�ݾ�(HIDDEN ����Ǵ� ��)
	'2002.09.10 �߰�                    
	C_PoAmt       		= 48                '���ֱݾ� 
	C_MvmtAmt     		= 49                '�԰�ݾ� 
	C_TotIvDocAmt 		= 50                '����Աݾ� 
	C_upt_amt_flg 		= 51                '���ֱݾ� ���ſ��� 
	C_prcflg      		= 52                '�ܰ�(ǥ�شܰ�(S),�̵���մܰ�(M))
	C_PoVatAmt    		= 53                '���� vat�ݾ� 
	C_TotIvVatAmt 		= 54                '����� vat�ݾ� 
	C_PoIvQty     		= 55
	C_retflg      		= 56
	C_ref_flg     		= 57
	C_Stateflg	  		= 58
	C_PoVatIncFlg 		= 59
	C_LCNo		  		= 60          
	'####��LC�߰�(2003.03.14)#####
	C_LCSeqNo	  		= 61
	C_LcFlg		  		= 62
	C_XchRt		  		= 63				'���Գ���ȯ�� - 2003.09.19

End Sub

'--------------------------------------------------------------------
'		Cookie ����Լ� 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
Dim strTemp
		Dim IntRetCD

	If Kubun = 1 Then

	    If lgIntFlgMode <> parent.OPMD_UMODE Then            'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie "txtIvNo" , FilterVar(UCase(Trim(frm1.txtIvNo.value)), "", "SNM")
		
		Call PgmJump(BIZ_PGM_JUMP_ID)
		 
	ElseIf Kubun = 0 Then
		
		strTemp = ReadCookie("txtIvNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtIvNo.Value = strTemp
		
		WriteCookie "txtIvNo" , ""

		if Trim(strTemp) <> "" then
			
			frm1.txtQuerytype.value = "Auto"
			frm1.txthdnIvNo.value = strTemp

			frm1.hdnPoNo.Value = ReadCookie("txtPoNo")
			Call dbquery()
		end if
		WriteCookie "txtPoNo" , ""
	ElseIf Kubun = 2 Then

	    If lgIntFlgMode <> parent.OPMD_UMODE Then                   'Check if there is retrived data
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie "txtIvNo" , FilterVar(UCase(Trim(frm1.txtIvNo.value)), "", "SNM")
		
		Call PgmJump(BIZ_PGM_JUMP_ID2)		
	End IF
	
End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	frm1.vspdData.MaxRows = 0
	lgSortKey         = 1                                       '��: initializes sort direction
    
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtSoNo1.className = "TD6" 
	Call SetToolBar("1110000000001111")
    frm1.btnPosting.value = "Ȯ��"
    frm1.btnPosting.disabled = true
    frm1.btnGlSel.disabled = true
    frm1.ChkPrepay.Checked =   false                 '���ޱݿ��� ���� check box
    frm1.txtIvNo.focus 
    Set gActiveElement = document.activeElement
    interface_Account = GetSetupMod(parent.gSetupMod, "a")
	if interface_Account = "N" then
		'btnintAcc.style.display = "none"
		frm1.btnPosting.disabled = true
	End if
End Sub


'================================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20070420",,parent.gAllowDragDropSpread 
	
	With frm1.vspdData
		
		.ReDraw = false
		'���Գ���ȯ�� �߰� - 2003.09.19
		.MaxCols = C_XchRt + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit 		C_PlantCd, "����", 7,,,4,2
		ggoSpread.SSSetButton 		C_PlantPopup
		ggoSpread.SSSetEdit 		C_PlantNm, "�����", 0
		ggoSpread.SSSetEdit 		C_ItemCd, "ǰ��", 18,,,18,2
		ggoSpread.SSSetButton 		C_ItemPopup
		ggoSpread.SSSetEdit 		C_ItemNm, "ǰ���", 20 
    
		ggoSpread.SSSetEdit			C_SpplSpec, "ǰ��԰�", 20        'ǰ��԰� �߰� 
    
		SetSpreadFloatLocal 		C_IvQty1, "���Լ���",15,1,3
		ggoSpread.SSSetEdit 		C_Unit, "����",7,,,3,2
		ggoSpread.SSSetButton 		C_UnitPopup
		SetSpreadFloatLocal			C_Cost, "���Դܰ�", 15, 1, 4
    
		SetSpreadFloatLocal			C_IvAmt, "�ݾ�",15,1,2        '�߰� 
    
		SetSpreadFloatLocal			C_NetAmt, "���Լ��ݾ�",15,1,2
		SetSpreadFloatLocal			C_OrgNetAmt, "���Լ��ݾ�",15,1,2    

		'�߰� 
		ggoSpread.SSSetCombo		C_IOFlg,"VAT���Կ���", 15,2
		ggoSpread.SetCombo      "����" & vbTab & "����",C_IOFlg
		ggoSpread.SSSetEdit 		C_IOFlgCd, "VAT���Կ����ڵ�", 15,2
		ggoSpread.SetCombo      "1" & vbTab & "2",C_IOFlgCd
		ggoSpread.SSSetEdit 		C_VatType, "VAT", 7,,,4,2
		ggoSpread.SSSetButton 		C_VatPopup
		ggoSpread.SSSetEdit 		C_VatNm, "VAT��", 20 
		SetSpreadFloatLocal			C_VatRate, "VAT��",15,1,5

		SetSpreadFloatLocal			C_VatDocAmt, "VAT�ݾ�",15,1,2
		SetSpreadFloatLocal			C_OrgVatDocAmt, "OrgVatDocAmt",15,1,2          
    
		SetSpreadFloatLocal			C_IvLocAmt, "�ڱ��ݾ�",15,1,2   '�߰� 
    
		SetSpreadFloatLocal			C_NetLocAmt, "�����ڱ����ݾ�",15,1,2 
		SetSpreadFloatLocal			C_VatLocAmt, "VAT�ڱ��ݾ�",15,1,2
		ggoSpread.SSSetEdit 		C_Remark, "���", 20
		SetSpreadFloatLocal			C_OrderQty, "���ּ���",15,1,3
		SetSpreadFloatLocal			C_OrderCost,"���ִܰ�",15,1,4

		SetSpreadFloatLocal			C_GmQty, "�԰����",15,1,3
		SetSpreadFloatLocal			C_IvQty2, "���ԿϷ����",15,1,3
		ggoSpread.SSSetEdit 		C_PoNo, "���ֹ�ȣ",15
		ggoSpread.SSSetEdit 		C_PoSeq, "���ּ���",10
		ggoSpread.SSSetEdit 		C_MvmtRcptNo, "�԰��ȣ",15
		ggoSpread.SSSetEdit 		C_GmNo, "���ó����ȣ",15   
		ggoSpread.SSSetEdit 		C_GmSeq, "���ó������",15
    
		ggoSpread.SSSetEdit 		C_IvSeq, "���Լ���", 10    
		SetSpreadFloatLocal 		C_OldQty, "OldQty",15,1,3				'hidden
		ggoSpread.SSSetEdit 		C_MvmtNo, "MvmtNo",10					'hidden
		SetSpreadFloatLocal 		C_MvmtIvQty, "MvmtIvQty",15,1,3			'hidden
		SetSpreadFloatLocal			C_oldIvQty1,  "oldIvQty1",15,1,3 
		ggoSpread.SSSetEdit			C_vat_rvs_flg,"vat_rvs_flg",5			'vat ���� ���� 
		SetSpreadFloatLocal			C_chkVatDocAmt,"chkVatDocAmt",15,1,2
		ggoSpread.SSSetEdit			C_ref_vatrate_flg,"ref_vatrate_flg",5   'vat ���� ���� 
		
		ggoSpread.SSSetEdit			C_TrackingNo, "Tracking No",20
		'2007-04-16 added
		ggoSpread.SSSetButton 		C_TrackingPopup
		SetSpreadFloatLocal			C_ChgNetAmt, "C_ChgNetAmt",15,1,2  
		SetSpreadFloatLocal			C_ChgVatDocAmt, "C_ChgVatDocAmt",15,1,2   	
	
		'2002.09.10 �߰� 
		SetSpreadFloatLocal			C_PoAmt, "PoAmt",15,1,2
		SetSpreadFloatLocal 		C_MvmtAmt, "MvmtAmt",15,1,2
		SetSpreadFloatLocal 		C_TotIvDocAmt, "IvDocAmt",15,1,2
		ggoSpread.SSSetEdit			C_upt_amt_flg, "upt_amt_flg",5
		ggoSpread.SSSetEdit			C_prcflg, "prcflg",5
		SetSpreadFloatLocal			C_PoVatAmt, "PoVatAmt",15,1,2
		SetSpreadFloatLocal			C_TotIvVatAmt, "IvVatAmt",15,1,2
				
		SetSpreadFloatLocal			C_PoIvQty, "PoIvQty",15,1,3
	  
		ggoSpread.SSSetEdit			C_retflg, "retflg",5    
		ggoSpread.SSSetEdit			C_ref_flg, "ref_flg",5    
		ggoSpread.SSSetEdit			C_Stateflg, "Stateflg",20 
		ggoSpread.SSSetEdit			C_PoVatIncFlg, "PoVatIncFlg",5   
		'####��LC�߰�(2003.03.14)#####
		ggoSpread.SSSetEdit 		C_LCNo, "LOCAL L/C��ȣ",15
		ggoSpread.SSSetEdit 		C_LCSeqNo, "LOCAL L/C����",15
		ggoSpread.SSSetEdit 		C_LcFlg, "LcFlg",5
		'���Գ���ȯ�� �߰� - 2003.09.19
		SetSpreadFloatLocal			C_XchRt, "ȯ��",10,1,5

		Call ggoSpread.MakePairsColumn(C_PlantCd,c_PlantPopUP)
		Call ggoSpread.MakePairsColumn(C_ItemCd, C_ItemPopup)
		Call ggoSpread.MakePairsColumn(C_Unit, C_UnitPopup)
		Call ggoSpread.MakePairsColumn(C_IOFlg, C_IOFlgCd, "1")
		Call ggoSpread.MakePairsColumn(C_VatType, C_VatPopup, "1")
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True) 
		Call ggoSpread.SSSetColHidden(C_OldQty, C_ref_vatrate_flg, True)
		Call ggoSpread.SSSetColHidden(C_IOFlgCd, C_VatRate, True)
		Call ggoSpread.SSSetColHidden(C_OrgNetAmt, C_OrgNetAmt, True)
		Call ggoSpread.SSSetColHidden(C_OrgVatDocAmt, C_OrgVatDocAmt, True)
		Call ggoSpread.SSSetColHidden(C_ChgNetAmt, C_ChgVatDocAmt, True)
		Call ggoSpread.SSSetColHidden(C_PoAmt, C_PoVatIncFlg, True)
		Call ggoSpread.SSSetColHidden(C_LcFlg, C_LcFlg, True)
						
		Call SetSpreadLock
		
		.ReDraw = true

    End With
    
End Sub

'=============================  GetSpreadColumnPos()  ==============================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCd     = iCurColumnPos(1)                     '���� 
			c_PlantPopUP  = iCurColumnPos(2)                     '���� �˾� 
			C_PlantNm     = iCurColumnPos(3)                     '����� 
			C_ItemCd      = iCurColumnPos(4)                     'ǰ�� 
			C_ItemPopup   = iCurColumnPos(5)                     'ǰ�� �˾� 
			C_ItemNm      = iCurColumnPos(6)                     'ǰ��� 
			C_SpplSpec    = iCurColumnPos(7)                     'ǰ�� �԰� �߰� 
			C_IvQty1      = iCurColumnPos(8)                     '���Լ��� 
			C_Unit        = iCurColumnPos(9)                     '���� 
			C_UnitPopup   = iCurColumnPos(10)                    '�����˾� 
			C_Cost        = iCurColumnPos(11)                    '���Դܰ� 
			C_IvAmt		  = iCurColumnPos(12)                    '���رݾ� 
			C_NetAmt      = iCurColumnPos(13)                    '���Աݾ� => "���Լ��ݾ�"���� ����(2002-06-19)
			C_OrgNetAmt   = iCurColumnPos(14)                    '���Աݾ� 
			'�߰� 
			C_IOFlg		  = iCurColumnPos(15)                     '���Ա��и� 
			C_IOFlgCd	  = iCurColumnPos(16)                     '���Ա����ڵ� 
			C_VatType     = iCurColumnPos(17)                     'vat
			C_VatPopup    = iCurColumnPos(18)                     'vat�˾� 
			C_VatNm       = iCurColumnPos(19)                     'vat�� 
			C_VatRate     = iCurColumnPos(20)                     'vat�� 
			C_VatDocAmt   = iCurColumnPos(21)                    'VAT�ݾ� 
			C_OrgVatDocAmt= iCurColumnPos(22)                    'VAT�ݾ� 
			C_IvLocAmt	  = iCurColumnPos(23)                    '�����ڱ��ݾ� 
			C_NetLocAmt   = iCurColumnPos(24)                    '�����ڱ��ݾ� => "�����ڱ����ݾ�"���� ����(2002-06-19) 
			C_VatLocAmt   = iCurColumnPos(25)                    'VAT�ڱ��ݾ� 
			C_Remark	  = iCurColumnPos(26)
			C_OrderQty    = iCurColumnPos(27)                    '���ּ��� 
			C_OrderCost   = iCurColumnPos(28)                    '���ִܰ� 
			C_GmQty       = iCurColumnPos(29)                    '�԰���� 
			C_IvQty2      = iCurColumnPos(30)                    '���ԿϷ���� 
			C_PoNo        = iCurColumnPos(31)                    '���ֹ�ȣ 
			C_PoSeq       = iCurColumnPos(32)                    '���ּ��� 
			C_MvmtRcptNo  = iCurColumnPos(33)                    '�԰��ȣ 
			C_GmNo        = iCurColumnPos(34)                    '���ó����ȣ 
			C_GmSeq       = iCurColumnPos(35)                    '���ó������ 
			C_IvSeq       = iCurColumnPos(36)                    '���Լ��� 
			C_OldQty      = iCurColumnPos(37)                    'hidden
			C_MvmtNo      = iCurColumnPos(38)                    'hidden
			C_MvmtIvQty   = iCurColumnPos(39)                    '����Լ��� 
			C_oldIvQty1   = iCurColumnPos(40)
			C_vat_rvs_flg = iCurColumnPos(41)                    '���� ���� �񱳰� 
			C_chkVatDocAmt= iCurColumnPos(42)                    'vat ���� ���� 
			C_ref_vatrate_flg = iCurColumnPos(43)
			C_TrackingNo  = iCurColumnPos(44)
			'2007-04-16 added
			C_TrackingPopup = iCurColumnPos(45)
			C_ChgNetAmt   = iCurColumnPos(46)                    '���Աݾ�(HIDDEN ����Ǵ� ��)
			C_ChgVatDocAmt= iCurColumnPos(47)                    'VAT�ݾ�(HIDDEN ����Ǵ� ��)
			'2002.09.10 �߰� 
			C_PoAmt       = iCurColumnPos(48)                    '���ֱݾ� 
			C_MvmtAmt     = iCurColumnPos(49)                    '�԰�ݾ� 
			C_TotIvDocAmt = iCurColumnPos(50)                    '����Աݾ� 
			C_upt_amt_flg = iCurColumnPos(51)                    '���ֱݾ� ���ſ��� 
			C_prcflg      = iCurColumnPos(52)                    '�ܰ�(ǥ�شܰ�(S),�̵���մܰ�(M))
			C_PoVatAmt    = iCurColumnPos(53)                    '���� vat�ݾ� 
			C_TotIvVatAmt = iCurColumnPos(54)                    '����� vat�ݾ� 
			C_PoIvQty     = iCurColumnPos(55)
			C_retflg      = iCurColumnPos(56)
			C_ref_flg     = iCurColumnPos(57)
			C_Stateflg	  = iCurColumnPos(58)
			C_PoVatIncFlg = iCurColumnPos(59)
			'2003-02-20 �Ѽ����ڿ䱸���� ksh
			C_LCNo		  = iCurColumnPos(60)
			'####��LC�߰�(2003.03.14)#####	
			C_LCSeqNo	  = iCurColumnPos(61)	
			C_LcFlg		  = iCurColumnPos(62)
			C_XchRt		  = iCurColumnPos(63)		'���Գ���ȯ�� - 2003.09.19
				
    End Select    
End Sub
'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
    With frm1
    
    ggoSpread.SpreadLock 		C_IvSeq , -1
    ggoSpread.SpreadUnLock 		C_PlantCd , -1
    ggoSpread.SSSetRequired		C_PlantCd , -1
    ggoSpread.SpreadLock		C_PlantNm , -1
    ggoSpread.SpreadUnLock 		C_ItemCd , -1
    ggoSpread.SSSetRequired		C_ItemCd , -1
    ggoSpread.spreadlock 		C_ItemNm , -1
	ggoSpread.spreadlock		C_SpplSpec,-1         'ǰ��԰� �߰� 
	ggoSpread.spreadUnlock 		C_IvQty1, -1
	ggoSpread.SSSetRequired		C_IvQty1,-1
    ggoSpread.spreadlock 		C_Unit, -1
    ggoSpread.spreadlock 		C_Cost, -1
    
    ggoSpread.spreadUnlock 		C_IvAmt, -1
	ggoSpread.SSSetRequired		C_IvAmt,-1
    
    'ggoSpread.spreadUnlock 	C_NetAmt, -1
    ggoSpread.SSSetProtected	C_NetAmt,-1
    ggoSpread.SSSetProtected	C_OrgNetAmt,-1    
    'ggoSpread.spreadUnlock 	C_NetAmt, -1
	'ggoSpread.SSSetRequired	C_NetAmt,-1
	
	'�߰�(12��)
    ggoSpread.SpreadUnLock		C_VatType, -1
	ggoSpread.SSSetRequired		C_VatType, -1
	ggoSpread.spreadlock 	    C_VatNm, -1
	ggoSpread.spreadlock		C_VatRate, -1


	If Trim(frm1.hdnVatType.Value) = "" Then
 	    ggoSpread.SSSetProtected	C_VatDocAmt,  -1
	Else			    
		ggoSpread.SpreadUnLock		C_VatDocAmt, -1
		ggoSpread.SSSetRequired		C_VatDocAmt, -1
	End If	
	
 	ggoSpread.SSSetProtected	C_OrgVatDocAmt,  -1	
    
    ggoSpread.spreadUnlock 	    C_IvLocAmt, -1  '�����ڱ��ݾ� 
	ggoSpread.SSSetProtected	C_IvLocAmt,-1
    
    ggoSpread.spreadUnlock 		C_NetLocAmt, -1
	ggoSpread.SSSetRequired		C_NetLocAmt,-1
	
	If Trim(frm1.hdnVatType.Value) = "" Then
	    ggoSpread.SSSetProtected	C_VatLocAmt,  -1
	Else			    
		ggoSpread.SpreadUnLock		C_VatLocAmt, -1
		ggoSpread.SSSetProtected	C_VatLocAmt, -1
	End If	
    ggoSpread.spreadlock 		C_OrderQty , -1
    ggoSpread.spreadlock 		C_GmQty , -1
    ggoSpread.spreadlock 		C_IvQty2, -1
    ggoSpread.spreadlock		C_MvmtRcptNo, -1
    ggoSpread.spreadlock		C_GmNo, -1
    ggoSpread.spreadlock 		C_GmSeq , -1
    ggoSpread.spreadlock 		C_PoSeq , -1
    
    ggoSpread.spreadlock 		C_OldQty , -1
    ggoSpread.spreadlock 		C_MvmtNo , -1
    ggoSpread.spreadlock 		C_MvmtIvQty , -1
    ggoSpread.spreadlock		C_TrackingNo , -1
    
    ggoSpread.spreadlock		C_ChgNetAmt , -1
    ggoSpread.spreadlock		C_ChgVatDocAmt , -1    
    ggoSpread.spreadlock		C_Stateflg , -1
    ggoSpread.spreadlock		C_PoVatIncFlg , -1
    ggoSpread.SSSetProtected	C_PoVatIncFlg + 1,  -1	

    ggoSpread.spreadlock		C_LCNo , -1
	'####��LC�߰�(2003.03.14)#####	
	ggoSpread.spreadlock		C_LCSeqNo , -1
	ggoSpread.spreadlock 		C_XchRt, -1

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStarRow, Byval pvEndRow)
    Dim index
    With ggoSpread
		
	    .SSSetRequired		C_PlantCd, pvStarRow, pvEndRow
	    .SSSetProtected		C_PlantNm, pvStarRow, pvEndRow
	    .SSSetRequired		C_ItemCd, pvStarRow, pvEndRow
	    .SSSetProtected		C_ItemNm, pvStarRow, pvEndRow
	    .SSSetProtected		C_SpplSpec, pvStarRow, pvEndRow
	    .SSSetRequired		C_IvQty1, pvStarRow, pvEndRow
	    .SSSetRequired		C_Unit, pvStarRow, pvEndRow

	    .SSSetRequired		C_Cost, pvStarRow, pvEndRow
	    .SSSetRequired		C_IvAmt, pvStarRow, pvEndRow
	    .SSSetProtected		C_NetAmt, pvStarRow, pvEndRow
	    .SSSetProtected		C_OrgNetAmt, pvStarRow, pvEndRow	    

		'�߰� 
	    .SSSetRequired		C_IOFlg, pvStarRow, pvEndRow
		.SSSetRequired		C_IOFlgCd, pvStarRow, pvEndRow
		
		.SSSetRequired		C_VatType, pvStarRow, pvEndRow
		.SSSetProtected		C_VatNm, pvStarRow, pvEndRow
		.SSSetProtected	    C_VatRate, pvStarRow, pvEndRow
		
		For index = pvStarRow to pvEndRow
			frm1.vspdData.Row=index
			frm1.vspdData.Col=C_LCFlg
			if Trim(frm1.hdnVatType.Value) = "" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
 			    .SSSetProtected	C_VatDocAmt, index, index		'VAT�ݾ� 
 			    .SSSetProtected	C_VatLocAmt, index, index
 			    .SSSetProtected	C_IOFlg, index, index
 			    .SSSetProtected	C_IOFlgCd, index, index
			else
				.SSSetRequired	C_VatDocAmt, index, index	
				.SSSetProtected	C_VatLocAmt, index, index				
			end if
		Next
		
 	    .SSSetProtected		C_OrgVatDocAmt, pvStarRow, pvEndRow     'VAT�ݾ�		
 	    
		.SSSetProtected		C_IvLocAmt, pvStarRow, pvEndRow	    
	    .SSSetProtected		C_NetLocAmt, pvStarRow, pvEndRow
			    
	    .SSSetProtected		C_OrderQty, pvStarRow, pvEndRow         '���ּ��� 
	    .SSSetProtected		C_OrderCost, pvStarRow, pvEndRow


	    .SSSetProtected		C_GmQty, pvStarRow, pvEndRow
	    .SSSetProtected		C_IvQty2, pvStarRow, pvEndRow
	    .SSSetProtected		C_PoNo, pvStarRow, pvEndRow
	    .SSSetProtected		C_PoSeq, pvStarRow, pvEndRow
	    .SSSetProtected		C_MvmtRcptNo, pvStarRow, pvEndRow
	    .SSSetProtected		C_GmNo, pvStarRow, pvEndRow   
	    .SSSetProtected		C_GmSeq, pvStarRow, pvEndRow
	    .SSSetProtected		C_IvSeq, pvStarRow, pvEndRow
        .SSSetProtected     C_TrackingNo, pvStarRow, pvEndRow
        .SSSetProtected     C_PoVatIncFlg + 1, pvStarRow, pvEndRow

        .SSSetProtected     C_LCNo, pvStarRow, pvEndRow
        '####��LC�߰�(2003.03.14)#####	
        .SSSetProtected     C_LCSeqNo, pvStarRow, pvEndRow
        .SSSetProtected     C_XchRt, pvStarRow, pvEndRow	'���Գ���ȯ�� - 2003.09.19
				
    End With
    
End Sub
'==========================================  2.2.6 SetRdSpreadColor()  ==================================  
'������ ��� 
Sub SetSpreadColorRef(ByVal lRow)
    
    frm1.vspdData.Row=lRow
    ggoSpread.SSSetProtected        C_PlantCd, lRow, lRow            '���� 
    ggoSpread.SSSetProtected        c_PlantPopUP, lRow, lRow
    ggoSpread.SSSetProtected        C_PlantNm, lRow, lRow
    ggoSpread.SSSetProtected        C_ItemCd, lRow, lRow             'ǰ�� 
    ggoSpread.SSSetProtected        C_ItemPopUP, lRow, lRow
    ggoSpread.SSSetProtected        C_ItemNm, lRow, lRow
    ggoSpread.SSSetProtected		C_SpplSpec, lRow, lRow
    ggoSpread.SSSetRequired         C_IvQty1, lRow, lRow             '���Լ��� 
    ggoSpread.SSSetProtected        C_Unit, lRow, lRow               '���� 
    ggoSpread.SSSetProtected        C_UnitPopup, lRow, lRow
    '**��LC ���� ���� (2003.03.19-Lee,Eun Hee)
    frm1.vspdData.Col=C_LCFlg
	If Trim(frm1.hdnRetflg.value) = "Y" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
		ggoSpread.SSSetProtected        C_Cost, lRow, lRow
	Else
		ggoSpread.SpreadUnLock          C_Cost , lRow ,C_Cost , lRow     '�ܰ� 
		ggoSpread.SSSetRequired         C_Cost, lRow, lRow
	End If      
    '**��ǰ�� ��� �ݾ׼����Ұ�(2003.08.01)
	If Trim(frm1.hdnRetflg.value) = "Y" Then
		ggoSpread.SSSetProtected		C_IvAmt, lRow,lRow
	Else
		ggoSpread.SpreadUnLock			C_IvAmt, lRow,C_IvAmt,lRow
		ggoSpread.SSSetRequired			C_IvAmt, lRow,lRow
	End If    
        
    ggoSpread.SSSetProtected		C_NetAmt, lRow, lRow
    ggoSpread.SSSetProtected		C_OrgNetAmt, lRow, lRow        
        
    ggoSpread.SSSetProtected		C_IvLocAmt, lRow, lRow
    ggoSpread.SSSetRequired        C_NetLocAmt, lRow, lRow            '�����ڱ��ݾ� 
    
	'�߰� 
	ggoSpread.SSSetRequired		    C_IOFlg, lRow, lRow
	ggoSpread.SSSetRequired		    C_IOFlgCd, lRow, lRow
		
	ggoSpread.SSSetRequired			C_VatType, lRow, lRow
	ggoSpread.SpreadUnLock          C_VatPopup, lRow, C_VatPopup, lRow
	ggoSpread.SSSetProtected		C_VatNm, lRow, lRow
	ggoSpread.SSSetProtected		C_VatRate, lRow, lRow
	'LC���� ��� VAT�ݾ� ���� �Ұ�(2003.09.26)
	frm1.vspdData.Col=C_LCFlg	
	' Issue for 8547 by Byun Jee Hyun 2004-08-09
	'if Trim(frm1.hdnVatType.Value) = "" Or Trim(frm1.hdnRetflg.value) = "Y" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
	if Trim(frm1.hdnVatType.Value) = "" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
	' End of Issue for 8547
        ggoSpread.SSSetProtected        C_VatDocAmt, lRow, lRow       'VAT�ݾ� 
        ggoSpread.SSSetProtected        C_VatLocAmt, lRow, lRow       'VAT�ڱ��ݾ� 
        ggoSpread.SSSetProtected		C_IOFlg, lRow, lRow
		ggoSpread.SSSetProtected		C_IOFlgCd, lRow, lRow   
    else
        ggoSpread.SSSetRequired			C_VatDocAmt, lRow, lRow 
        ggoSpread.SSSetRequired		    C_IOFlg, lRow, lRow
		ggoSpread.SSSetRequired		    C_IOFlgCd, lRow, lRow    
    end if      
    ggoSpread.SSSetProtected        C_OrgVatDocAmt, lRow, lRow       'VAT�ݾ�            
                    
    ggoSpread.SSSetProtected        C_OrderQty, lRow, lRow            '���ּ��� 
    ggoSpread.SSSetProtected        C_OrderCost, lRow, lRow
        
    ggoSpread.SSSetProtected        C_GmQty, lRow, lRow
    ggoSpread.SSSetProtected        C_IvQty2, lRow, lRow
    ggoSpread.SSSetProtected        C_PoNo, lRow, lRow
    ggoSpread.SSSetProtected        C_PoSeq, lRow, lRow
    ggoSpread.SSSetProtected        C_MvmtRcptNo, lRow, lRow 
    ggoSpread.SSSetProtected        C_GmNo, lRow, lRow  
    ggoSpread.SSSetProtected        C_GmSeq, lRow, lRow
    ggoSpread.SSSetProtected        C_IvSeq, lRow, lRow  
    ggoSpread.SSSetProtected        C_TrackingNo, lRow, lRow    
    ggoSpread.SSSetProtected        C_PoVatIncFlg + 1, lRow, lRow  
        
    ggoSpread.SSSetProtected        C_LCNo, lRow, lRow  
    '####��LC�߰�(2003.03.14)#####	
    ggoSpread.SSSetProtected        C_LCSeqNo, lRow, lRow
    ggoSpread.SSSetProtected        C_XchRt, lRow, lRow    '���Գ���ȯ�� - 2003.09.19
    
    ggoSpread.SpreadUnLock          C_Remark, lRow, C_Remark, lRow
            
    Call SetSpreadHiddenAttrByCurrency(lRow)
    
End Sub
'==========================================  2.2.6 SetRdSpreadColor()  ==================================  
'Ȯ���� �Ȱ�� ȣ�� 
Sub SetRdSpreadColor(ByVal lRow)
	Dim index	
		for index = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = index
			If UCase(parent.gCurrency) = UCase(Trim(frm1.txtCur.value)) Then 
				frm1.vspdData.Col = C_IvLocAmt	: frm1.vspdData.ColHidden = True 
				frm1.vspdData.Col = C_NetLocAmt	: frm1.vspdData.ColHidden = True 
				frm1.vspdData.Col = C_VatLocAmt	: frm1.vspdData.ColHidden = True 
			Else 
				frm1.vspdData.Col = C_IvLocAmt	: frm1.vspdData.ColHidden = False 
				frm1.vspdData.Col = C_NetLocAmt	: frm1.vspdData.ColHidden = False
				frm1.vspdData.Col = C_VatLocAmt	: frm1.vspdData.ColHidden = False
			End If
		
		Next
		
		ggoSpread.SpreadLock        C_IvSeq, -1
		ggoSpread.SpreadLock        C_PlantCd,  -1 
		ggoSpread.SpreadLock        c_PlantPopUP,  -1 
		ggoSpread.SpreadLock        C_PlantNm,  -1
		ggoSpread.SpreadLock        C_ItemCd,  -1 
		ggoSpread.SpreadLock        C_ItemPopup,  -1
		ggoSpread.SpreadLock        C_ItemNm,  -1
		ggoSpread.SpreadLock		C_SpplSpec,  -1 
		ggoSpread.SpreadLock        C_IvQty1,  -1
		ggoSpread.SpreadLock        C_Unit,  -1 
		ggoSpread.SpreadLock        C_UnitPopup,  -1 
		ggoSpread.SpreadLock        C_Cost,  -1
		ggoSpread.SpreadLock		C_IvAmt,  -1 
		ggoSpread.SpreadLock        C_NetAmt,  -1
		ggoSpread.SpreadLock        C_OrgNetAmt,  -1      
		ggoSpread.SpreadLock		C_IOFlg,  -1
		ggoSpread.SpreadLock		C_IOFlgCd,  -1
		ggoSpread.SpreadLock        C_VatType,  -1
		ggoSpread.SpreadLock        C_VatPopup, -1 
		ggoSpread.SpreadLock        C_VatNm,  -1 
		ggoSpread.SpreadLock        C_VatRate,  -1 
		ggoSpread.SpreadLock        C_VatDocAmt, -1
		ggoSpread.SpreadLock        C_OrgVatDocAmt,  -1 	 
		ggoSpread.SpreadLock		C_IvLocAmt,  -1 
		ggoSpread.SpreadLock        C_NetLocAmt,  -1 
		ggoSpread.SpreadLock        C_VatLocAmt,  -1
		ggoSpread.SpreadLock        C_OrderQty,  -1
		ggoSpread.SpreadLock        C_GmQty,  -1 
		ggoSpread.SpreadLock        C_IvQty2,  -1
		ggoSpread.SpreadLock        C_MvmtRcptNo,  -1 
		ggoSpread.SpreadLock        C_GmNo,  -1  
		ggoSpread.SpreadLock        C_GmSeq,  -1
		ggoSpread.SpreadLock        C_PoSeq,  -1 
		ggoSpread.SpreadLock        C_TrackingNo,  -1
		ggoSpread.SpreadLock        C_PoVatIncFlg + 1,  -1      
		ggoSpread.SpreadLock	    C_LCNo, -1 
		ggoSpread.SpreadLock	    C_LCSeqNo, -1
		ggoSpread.SpreadLock	    C_XchRt, -1			'���Գ���ȯ�� - 2003.09.19
		ggoSpread.SpreadLock		C_Remark,  -1	'��� - 2005.12.19
	    
End Sub
'==========================================  2.2.6 QueryAtSetSpreadColor()  ============================== 
Sub QueryAtSetSpreadColor(ByVal lRow)
    Dim index
    With frm1
    	    	
		For index = Cint(.hdnmaxrow.value) + 1 to .vspdData.MaxRows
			.vspdData.Row=index
			If UCase(parent.gCurrency) = UCase(Trim(frm1.txtCur.value)) Then 
				.vspdData.Col = C_IvLocAmt	: .vspdData.ColHidden = True 
				.vspdData.Col = C_NetLocAmt	: .vspdData.ColHidden = True 
				.vspdData.Col = C_VatLocAmt	: .vspdData.ColHidden = True 
			Else 
				.vspdData.Col = C_IvLocAmt	: .vspdData.ColHidden = False 
				.vspdData.Col = C_NetLocAmt	: .vspdData.ColHidden = False
				.vspdData.Col = C_VatLocAmt	: .vspdData.ColHidden = False
			End If
		
		    ggoSpread.SSSetProtected		C_PlantCd, index,index
			ggoSpread.SSSetProtected		c_PlantPopUP, index,index
		    ggoSpread.SSSetProtected		C_PlantNm, index,index
		    ggoSpread.SSSetProtected		C_ItemCd, index,index
			ggoSpread.SSSetProtected		C_ItemPopup, index,index
		    ggoSpread.SSSetProtected		C_ItemNm, index,index
		    ggoSpread.SSSetProtected		C_SpplSpec, index,index
		    ggoSpread.SpreadUnLock			C_IvQty1, index,C_IvQty1, index
		    ggoSpread.SSSetRequired			C_IvQty1, index,index
		    ggoSpread.SSSetProtected		C_Unit, index,index
		    ggoSpread.SSSetProtected		C_UnitPopup,  index,index
		   
			'**��LC ���� ���� (2003.03.19-Lee,Eun Hee)-��Local LC���� �����Ͽ��� ��� �ܰ� ���� �Ұ� 
			'**��LC ���� ���� (2003.03.19-Lee,Eun Hee)-Local LC�� �� �԰��� ���� �����Ͽ��� ��� �ܰ� ���� �Ұ� 
			
			.vspdData.Col=C_LCFlg
		    if Trim(.hdnRetflg.value) = "Y" Or Trim(.vspdData.Text) = "A" Or Trim(.vspdData.Text) = "B" then            '��ǰ�ΰ�� �ܰ� �����Ұ� 
		        ggoSpread.SSSetProtected	C_Cost, index,index
		    else
		        ggoSpread.SpreadUnLock 		C_Cost , index, C_Cost, index
		        ggoSpread.SSSetRequired		C_Cost, index,index
		    end if
		    '**��ǰ�� ��� �ݾ׼����Ұ�(2003.08.01)
		    if Trim(.hdnRetflg.value) = "Y" Then
				ggoSpread.SSSetProtected		C_IvAmt, index,index
		    Else
				ggoSpread.SpreadUnLock			C_IvAmt, index,C_IvAmt,index
				ggoSpread.SSSetRequired			C_IvAmt, index,index
		    End If
		    
			ggoSpread.SSSetProtected		C_NetAmt, index,index
			ggoSpread.SSSetProtected		C_OrgNetAmt, index,index						
			
		    ggoSpread.SSSetRequired			C_VatType, index,index
		    ggoSpread.SpreadUnLock			C_VatPopup, index,C_VatPopup,index
		    ggoSpread.SSSetProtected 	    C_VatNm, index,index
		    ggoSpread.SSSetProtected		C_VatRate, index,index

			'LC���� ��� VAT�ݾ� ���� �Ұ� 
			.vspdData.Col=C_LCFlg
			' Issue for 8547 by Byun Jee Hyun 2004-08-09
		    'if Trim(.hdnRetflg.value) = "Y" Or Trim(frm1.hdnVatType.Value) = "" Or Trim(.vspdData.Text) = "A" Or Trim(.vspdData.Text) = "B" then
		    if Trim(frm1.hdnVatType.Value) = "" Or Trim(.vspdData.Text) = "A" Or Trim(.vspdData.Text) = "B" then
		    ' End of Issue for 8547
 		        ggoSpread.SSSetProtected		C_VatDocAmt, index,index
 		        ggoSpread.SSSetProtected        C_VatLocAmt, index,index       'VAT�ڱ��ݾ� 
 		        ggoSpread.SSSetProtected		C_IOFlg, index,index
				ggoSpread.SSSetProtected		C_IOFlgCd,  index,index
			else	
				ggoSpread.SpreadUnLock			C_IOFlg, index,C_IOFlg, index
				ggoSpread.SSSetRequired			C_IOFlg, index,index
				ggoSpread.SpreadUnLock			C_IOFlgCd, index,C_IOFlgCd,index
				ggoSpread.SSSetRequired			C_IOFlgCd, index,index		    
				ggoSpread.SpreadUnLock			C_VatDocAmt, index,C_VatDocAmt, index
				ggoSpread.SSSetRequired			C_VatDocAmt, index,index
		    end if		
 		    ggoSpread.SSSetProtected		C_OrgVatDocAmt, index,index
		    
				    
		    ggoSpread.SSSetProtected		C_OrderQty, index,index
		    ggoSpread.SSSetProtected		C_OrderCost, index,index
		    
		    ggoSpread.SSSetProtected		C_GmQty, index,index
		    ggoSpread.SSSetProtected		C_IvQty2, index,index
		    ggoSpread.SSSetProtected		C_PoNo,  index,index
		    ggoSpread.SSSetProtected		C_PoSeq,  index,index
		    ggoSpread.SSSetProtected		C_MvmtRcptNo, index,index 
		    ggoSpread.SSSetProtected		C_GmNo, index,index   
		    ggoSpread.SSSetProtected		C_GmSeq, index,index
		    ggoSpread.SSSetProtected		C_IvSeq, index,index
		    ggoSpread.SSSetProtected        C_TrackingNo, index,index
		    ggoSpread.SSSetProtected        C_PoVatIncFlg + 1, index,index

		    ggoSpread.SSSetProtected        C_LCNo, index,index
		    '####��LC�߰�(2003.03.14)#####	
		    ggoSpread.SSSetProtected        C_LCSeqNo, index,index
		    ggoSpread.SSSetProtected        C_XchRt, index,index		'���Գ���ȯ�� - 2003.09.19
		    
			if Trim(UCase(frm1.hdnImportflg.Value)) <> "Y" And Trim(UCase(frm1.hdnPostingflg.Value)) <> "Y" And Trim(frm1.hdnRetflg.value) <> "Y" Then
				ggoSpread.SpreadUnLock          C_IvLocAmt , index ,C_IvLocAmt , index    
				ggoSpread.SSSetRequired         C_IvLocAmt, index, index
				ggoSpread.SpreadUnLock          C_VatLocAmt , index ,C_VatLocAmt , index    
				ggoSpread.SSSetRequired         C_VatLocAmt, index, index
			Else
				ggoSpread.SSSetProtected		C_IvLocAmt, index,index
 				ggoSpread.SSSetProtected		C_VatLocAmt, index,index
			End if
			
			'VAT������ ���ų� L/C���� ���� VAT�ڱ��ݾ� ����Ұ�(2003.09.26)
			frm1.vspdData.Col=C_LCFlg
			If Trim(frm1.hdnVatType.Value) = "" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
			    ggoSpread.SSSetProtected        C_VatLocAmt, index, index       'VAT�ڱ��ݾ� 
			End If
				
		next
    	
    End With
End Sub
'==========================================  2.2.6 SetSpreadHiddenAttrByCurrency()  ======================
Sub SetSpreadHiddenAttrByCurrency(ByVal lRow)
	With frm1.vspdData
		
		If UCase(parent.gCurrency) = UCase(Trim(frm1.txtCur.value)) Then 
			.Col = C_IvLocAmt	: .ColHidden = True 
			.Col = C_NetLocAmt	: .ColHidden = True 
			.Col = C_VatLocAmt	: .ColHidden = True 
		Else 
			.Col = C_IvLocAmt	: .ColHidden = False 
			.Col = C_NetLocAmt	: .ColHidden = False
			.Col = C_VatLocAmt	: .ColHidden = False

			if Trim(UCase(frm1.hdnImportflg.Value)) <> "Y" And Trim(UCase(frm1.hdnPostingflg.Value)) <> "Y" And Trim(frm1.hdnRetflg.value) <> "Y" Then
				ggoSpread.SpreadUnLock          C_IvLocAmt , lRow ,C_IvLocAmt , lRow    
				ggoSpread.SSSetRequired         C_IvLocAmt, lRow, lRow
				ggoSpread.SpreadUnLock          C_VatLocAmt , lRow ,C_VatLocAmt , lRow    
				ggoSpread.SSSetRequired         C_VatLocAmt, lRow, lRow
			Else
				ggoSpread.SSSetProtected		C_IvLocAmt, lRow,lRow
				ggoSpread.SSSetProtected		C_NetLocAmt, lRow,lRow
 				ggoSpread.SSSetProtected		C_VatLocAmt, lRow,lRow
			End if
			
			'VAT������ ���ų� L/C���� ���� VAT�ڱ��ݾ� ����Ұ�(2003.09.26)
			frm1.vspdData.Col=C_LCFlg
			If Trim(frm1.hdnVatType.Value) = "" Or Trim(frm1.vspdData.Text) = "A" Or Trim(frm1.vspdData.Text) = "B" Then
			    ggoSpread.SSSetProtected        C_VatLocAmt, lRow, lRow       'VAT�ڱ��ݾ� 
			End If
			
		End If
	End with	
End Sub	
'==============================  OpenIvNo()  =======================================================
Function OpenIvNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		
		If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
		lblnWinEvent = True
		
		arrParam(0) = ivType
		
		
		iCalledAspName = AskPRAspName("m5111pa1_KO441")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m5111pa1_KO441", "X")
			lgIsOpenPop = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False
	
		If strRet(0) = "" Then
			frm1.txtIvNo.focus	
			Set gActiveElement = document.activeElement
			Exit Function
		Else
			frm1.txtIvNo.value = strRet(0)
			frm1.txtIvNo.focus	
			Set gActiveElement = document.activeElement
		End If	
		
End Function



'------------------------------------------  OpenPoRef()  ----------------------------------------------
Function OpenPoRef()

	Dim strRet
	Dim arrParam(14)
	
	if lgIntFlgMode = parent.OPMD_CMODE then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if 

	if UCase(Trim(frm1.hdnPostingFlg.Value)) = "Y" then            'Ȯ���̸� ���� �Ұ� 
		Call DisplayMsgBox("17a009","X","X","X")                   'ȸ��ó�������̹Ƿ� ���� �Ҽ� �����ϴ� 
		Exit Function
	End if

	if UCase(Trim(frm1.hdnImportFlg.Value)) = "Y" then             '�������°� ���԰�� Y
		Call DisplayMsgBox("17A012", "X","���԰�","���ֳ�������")
		Exit Function
	End if
	'hdnImportFlg(�����ΰ�� Y),  hdnExceptflg(�����ΰ�� Y), hdnRetFlg(��ǰ�ΰ�� Y)
'	if UCase(Trim(frm1.hdnRetFlg.Value)) <> "Y" or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
	'if Trim(UCase(frm1.hdnRcptFlg.Value)) = "Y" or  Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
     if Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then	
		Call DisplayMsgBox("17a015", "X","��������" & ":" & frm1.txtIvTypeCd.value & "(" & frm1.txtIvTypeNm.value & ")","���ֳ�������" )
		Exit Function
	End if
			
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnSppl.value)
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.hdnGrp.Value)
	arrParam(3) = ""
	arrParam(4) = ""		'Clsflg
	arrParam(5) = ""		'Releaseflg
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = "IV"		'Rcptflg
	arrParam(9) = ""
	arrParam(10) = "Y"		'IVflg
	arrParam(11) = Trim(frm1.hdnIvType.Value)	'IVType
	arrParam(12) = ""	'PoType
	arrParam(13) = UCase(Trim(frm1.hdnPoNo.value))  'pono
	arrParam(14) = UCase(Trim(frm1.txtCur.value))  'pocur

	strRet = window.showModalDialog("../m8/m3112ra6_KO441.asp?txtCurrency=" + frm1.txtCur.value, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetPoRef(strRet)
	End If	
		
End Function

'��ǥ��ȸ Ŭ���� ȣ�� 
'------------------------------------------  OpenGLRef()  ----------------------------------------------
Function OpenGLRef()
	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)

   If frm1.hdnGlType.Value = "A" Then               'ȸ����ǥ�˾� 
   		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "T" Then          '������ǥ�˾� 
		iCalledAspName = AskPRAspName("a5130ra1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	lblnWinEvent = False
	
End Function

'���ֳ������� ������ ���� ����ش� 
'------------------------------------------  SetPoRef()  ----------------------------------------------
Function SetPoRef(strRet)

Dim Index1,index2,Index3,Count1,Count2
Dim IntIflg
Dim strMessage
Dim temp1,temp2
Dim VatRt 
Dim VatType
Dim temp,changeVatflg
Dim TempRow

Const C_PoNo_Ref		= 0
Const C_PoSeq_Ref		= 1
Const C_PlantCd_Ref		= 2
Const C_PlantNm_Ref		= 3
Const C_ItemCd_Ref		= 4
Const C_ItemNm_Ref		= 5
Const C_SpplSpec_Ref    = 6                         'ǰ�� �԰� �߰� 
Const C_Qty_Ref			= 7
Const C_NotQty_Ref	    = 8
Const C_Unit_Ref		= 9
Const C_Prc_Ref			= 10
Const C_Amt_Ref			= 11
Const C_Cur_Ref			= 12
Const C_VatType_Ref		= 13
Const C_VatNm_Ref		= 14 
Const C_VatRate_Ref		= 15
Const C_PoDt_Ref        = 16
Const C_DlvyDt_Ref		= 17
Const C_SLCd_Ref		= 18
Const C_SLNm_Ref		= 19
Const C_TrackingNo_Ref  = 22
Const C_IoFg_Ref		= 23  
Const C_TotIvDocAmt_Ref = 24                        '�� ���Աݾ� 
Const C_upt_amt_flg_Ref = 25                        '���ֱݾ� ���ſ���	
Const C_PoVatAmt_Ref    = 26                        '���� vat�ݾ� 
Const C_TotIvVatAmt_Ref = 27    	
Const C_vat_rvs_flg_Ref = 28                        'vat ���� ���� 
Const C_PoIvQty_Ref     = 29
Const C_Retflg_Ref      = 30

    VatType = frm1.hdnVatType.value
    VatRt  = UNICDbl(frm1.hdnVatRt.value)
	
	Count1 = Ubound(strRet,1)  'row ���� 
	Count2 = UBound(strRet,2)  'col ���� 
	
	strMessage = ""
	IntIflg=true
	
	with frm1
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
		TempRow = .vspdData.MaxRows					'����Ʈ max��				
	
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
		
		If TempRow <> 0 Then
			for Index3=1 to TempRow			     'count1		'���� ReqNo�� ������ Row�� �߰����� �ʴ´�.
				.vspdData.Row = index3
				.vspdData.Col=C_PoNo
				temp1 = .vspdData.Text
				.vspdData.Col=C_PoSeq
				temp2 = .vspdData.Text
				if temp1 = strRet(index1,C_PoNo_Ref) And temp2 = strRet(index1,C_PoSeq_Ref) then
					strMessage = strMessage & strRet(index1,C_PoNo_Ref) & "-" & strRet(index1,C_PoSeq_Ref) & ";"
					intIflg=False
					Exit for
				End if 
			Next
		End If
		
		if IntIflg <> False then
	         
	        ggoSpread.InsertRow
			.vspdData.Row=.vspdData.ActiveRow 
				
			'Call SetSpreadColorRef(.vspdData.Row)
			
			'**���Գ������� ȯ�� ������(2003.09.19)** --> ���� �������� 
			.vspdData.Col=C_XchRt
			.vspdData.Text= frm1.hdnXch.value
			
			.vspdData.Col=C_PlantCd
			.vspdData.Text=strRet(index1,C_PlantCd_Ref)
			ggoSpread.spreadlock C_PlantCd,.vspdData.Row,C_PlantCd,.vspdData.Row
		
			.vspdData.Col=C_PlantNm
			.vspdData.Text=strRet(index1,C_PlantNm_Ref)
		
			.vspdData.Col=C_itemCd
			.vspdData.Text=strRet(index1,C_ItemCd_Ref)
			ggoSpread.spreadlock C_ItemCd,.vspdData.Row,C_ItemCd,.vspdData.Row
		
			.vspdData.Col=C_itemNm
			.vspdData.Text=strRet(index1,C_ItemNm_Ref)
		
			.vspdData.Col=C_SpplSpec
			.vspdData.Text=strRet(index1,C_SpplSpec_Ref)
		
			.vspdData.Col=C_OrderQty                  '���ּ��� 
			.vspdData.Text=strRet(index1,C_Qty_Ref)
			.vspdData.Col=C_IvQty2                    '���ԿϷ���� 
			.vspdData.Text=strRet(index1,C_Qty_Ref)	
		
			.vspdData.Col=C_IvQty1                    '���Լ��� 
			.vspdData.Text=strRet(index1,C_NotQty_Ref)
			
			.vspdData.Col=C_oldIvQty1                 '���Լ��� 
			.vspdData.Text= 0

			.vspdData.Col=C_IvQty2                    '����Լ��� = ���ּ��� - �̸��Լ��� 
			temp = UNICDbl(.vspdData.Text) - UNICDbl(strRet(index1,C_NotQty_Ref))
			.vspdData.Text= UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
			
			.vspdData.Col=C_Unit
			.vspdData.Text=strRet(index1,C_Unit_Ref)
			ggoSpread.spreadlock C_Unit,.vspdData.Row,C_ItemCd,.vspdData.Row
		
			.vspdData.Col=C_Cost                      '���Դܰ� 
			.vspdData.Text=strRet(index1,C_Prc_Ref)
           
			If UCase(Trim(strRet(index1,C_Retflg_Ref))) = "Y" then  '��ǰ���� 
				.vspdData.Text =  UNIFormatNumber(UNICDbl(.vspdData.Text) * -1,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
				ggoSpread.SSSetProtected		C_Cost, .vspdData.Row
			End if
			
			.vspdData.Col=C_retflg                    '��ǰ���� 
			.vspdData.Text=strRet(index1,C_Retflg_Ref)
			
			.vspdData.Col=C_OrderCost                 '���ִܰ� 
			.vspdData.Text=strRet(index1,C_Prc_Ref)
		
					
			'.vspdData.Col=C_OrgNetAmt                     '���Աݾ�(HIDDEN)
			'.vspdData.Text=strRet(index1,C_Amt_Ref)					
			'@����@(2003.02.17)
			.vspdData.Col = C_PoVatIncFlg
			.vspdData.Text = strRet(index1,C_IoFg_Ref)	
			'�߰�     
            if strRet(index1,C_IoFg_Ref) = "" then
                strRet(index1,C_IoFg_Ref) = .hdvatFlg.value  'header�� �ִ� ���� �־��ش� 
            end if
                     
            .vspdData.Col = C_IOFlg
			if strRet(index1,C_IoFg_Ref) = "2" Then
			    .vspdData.value = 1		'vat���� 
			ElseIf strRet(index1,C_IoFg_Ref) = "1" Then
				.vspdData.value = 0		'vat������ 
			End If
			'���ܼ������ڵ� 
			.vspdData.Col = C_IOFlgCd
			.vspdData.Text = strRet(index1,C_IoFg_Ref)
    
			.vspdData.Col=C_VatType                   'vat type
			.vspdData.Text=VatType

			.vspdData.Col=C_VatRate                  'vat�� 
			.vspdData.Text=UNIFormatNumber(VatRt, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
						
			.vspdData.Col=C_PoNo
			.vspdData.Text=strRet(index1,C_PoNo_Ref)
	
		    .vspdData.Col=C_PoSeq
			.vspdData.Text=strRet(index1,C_PoSeq_Ref)
		
            .vspdData.Col=C_TrackingNo
            .vspdData.Text=strRet(index1,C_TrackingNo_Ref)

            '2002.09.10 �߰� 
            .vspdData.Col=C_PoAmt
            .vspdData.Text=strRet(index1,C_Amt_Ref)              '�߰� 
            
            .vspdData.Col=C_TotIvDocAmt
            .vspdData.Text=strRet(index1,C_TotIvDocAmt_Ref)       '�߰� 
            
            .vspdData.Col=C_upt_amt_flg
            .vspdData.Text=strRet(index1,C_upt_amt_flg_Ref)    '�߰� 
 
            .vspdData.Col=C_PoVatAmt
            .vspdData.Text=strRet(index1,C_PoVatAmt_Ref)       '�߰� 
            
            .vspdData.Col=C_TotIvVatAmt
            .vspdData.Text=strRet(index1,C_TotIvVatAmt_Ref)       '�߰� 
				
			.vspdData.Col=C_vat_rvs_flg
            .vspdData.Text=strRet(index1,C_vat_rvs_flg_Ref)       '�߰� 
            
            .vspdData.Col=C_PoIvQty           
            .vspdData.Text=strRet(index1,C_PoIvQty_Ref)       '�߰� 
            
            if VatRt <> UNICDbl(strRet(index1,C_VatRate_Ref)) then 
                .vspdData.Col=C_ref_vatrate_flg
                .vspdData.Text = "N"    
            else
                .vspdData.Col=C_ref_vatrate_flg
                .vspdData.Text = "Y"    
            end if
	    
			.vspdData.Col = 1
			.vspdData.Action = 0
				

            .vspdData.Col=C_upt_amt_flg
            .vspdData.Text="N"       '�߰�	
	
			'---- Hidden �ʵ��� ������ ����Ͽ� �����ϱ� ���� �Լ� --------------------------------------------------------------------
			'---- ���� HIDDEN�ʵ��� ���� �ٷ� �Ʒ� vspddata_Change���� ���ǹǷ�, 
			'---- vspddata_change�� ȣ���ϱ� ����  ChangeAmtOrg�� ȣ��Ǿ�� �Ѵ� 
			
            .vspdData.Col = C_prcflg
            .vspdData.text = "S"
            If UCase(Trim(strRet(index1,C_Retflg_Ref))) = "Y" then  '��ǰ���� 
                .vspdData.Col = C_ref_flg
                .vspdData.text = "Y"
                Call vspdData_Change(C_Cost, .vspdData.ActiveRow)
            else
                Call changeNetAmt(C_Cost_Ref,.vspdData.ActiveRow)
            end if
            changeVatflg = "Y"
		
		    .vspdData.Col = C_ref_flg
            .vspdData.text = ""
			'---- C_OrgNetAmt(HIDDEN)�� C_NetAmt�� ���� �Է�--------------------------------------------------------------------
			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)			
			temp = UNICDbl(Trim(.vspdData.Text))
			
			.vspdData.Col = C_OrgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text =UNIConvNumPCToCompanyByCurrency(temp,parent.gCurrency,parent.ggAmtOfMoneyNo,"X","X") 

			.vspdData.Col = C_ChgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(temp,parent.gCurrency,parent.ggAmtOfMoneyNo,"X","X") 			

			'---- ���� VAT�ݾ��� HIDDEN �ʵ忡 ����--------------------------------------------------------------------
			dim tmpVatDocAmt
			.vspdData.Col=C_VatDocAmt
			tmpVatDocAmt = UNICDbl(.vspdData.Text)

			.vspdData.Col=C_OrgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
			
			.vspdData.Col=C_ChgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")												
			'------------------------------------------------------------------------
	        
	        .vspdData.Col=C_chkVatDocAmt
            .vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")												
             
			call ChangeAmtOrg(C_IvAmt,.vspdData.ActiveRow,changeVatflg)	
			'---------------------------------------------------------------------------------------------------------------------------
            frm1.vspdData.Col = C_NetAmt
			NetAmt = UNICDbl(frm1.vspdData.Text)

			Call vspdData_Change(C_Cost_Ref, .vspdData.ActiveRow)
			'next

			'----������ ����� ���ڸ� ����ϴ� �κ�--------------------------------------------------------------------------------------
			'----������ NetAMT�� VatDocAmt ��ŭ ����� ������----------------------------------------------------------------------------
			dim VatDocAmt
			dim NetAmt,VatIncFlag
			frm1.vspdData.Col = C_VatDocAmt
			VatDocAmt = UNICDbl(frm1.vspdData.Text)	 				
	
			frm1.vspdData.Col = C_OrgVatDocAmt
			tmpVatDocAmt = UNICDbl(.vspdData.Text)	 
	
			frm1.vspdData.Col = C_NetAmt
			NetAmt = UNICDbl(frm1.vspdData.Text)	 							
            
 			frm1.vspdData.Col = C_IOFlgCd
			VatIncFlag = frm1.vspdData.text	
			
			'#�� LC�߰�(2003.03.17)#######
			.vspdData.Col=C_LCNo
            .vspdData.Text= ""						       '�߰� 
			
			.vspdData.Col=C_LCSeqNo
            .vspdData.Text= ""								'�߰�  
            
            .vspdData.Col=C_LcFlg
			.vspdData.Text= ""								'�߰� 
			
			Call SetSpreadColorRef(.vspdData.Row)
			
		Else
			IntIFlg=True
		End if 
	next
	'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
	Call HSumAmtNewCalc()
	'Call SetSpreadColorRef(-1)
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005","X",strmessage,"���ֹ�ȣ" & "," & "���ּ���")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	.vspdData.ReDraw = True
	
	End with
	
End Function

'------------------------------------------  OpenGrRef()  --------------------------------------------
Function OpenGrRef()

	Dim strRet
	Dim arrParam(9)
	Dim iCalledAspName
	
	if lgIntFlgMode = parent.OPMD_CMODE then				'��: Indicates that current mode is Update mode
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	
	if UCase(Trim(frm1.hdnPostingFlg.Value)) = "Y" then
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
	
	if UCase(Trim(frm1.hdnImportFlg.Value)) = "Y" then
		Call DisplayMsgBox("17A012", "X","���԰�","���������")
		Exit Function
	End if
	'hdnExceptflg(�����ΰ�� Y)
	'if Trim(UCase(frm1.hdnRcptFlg.Value)) = "N" or UCase(Trim(frm1.hdnRetFlg.Value)) = "Y" or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
	if (Trim(UCase(frm1.hdnPoNo.Value)) <> ""  and Trim(UCase(frm1.hdnRcptType.Value)) = "" and Trim(UCase(frm1.hdnIssueType.Value)) = "") or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
		Call DisplayMsgBox("17a015", "X","��������" & ":" & frm1.txtIvTypeCd.value & "(" & frm1.txtIvTypeNm.value & ")" ,"�԰�������" )
		Exit Function
	End if
	
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnMvmtType.Value)
	arrParam(1) = Trim(frm1.hdnSppl.Value)
	arrParam(2) = Trim(frm1.hdnGrp.Value)
	arrParam(3) = "IV"
	arrParam(4) = Trim(frm1.hdnIvType.Value)
	arrParam(5) = UCase(Trim(frm1.hdnPoNo.value))
    arrParam(6) = UCase(Trim(frm1.txtCur.value))
    '����(2003.03.24)
    arrParam(7) = UCase(Trim(frm1.hdnLcKind.value))
    arrParam(8) = UCase(Trim(frm1.hdnPayMeth.value))
    '�߰�(2005.10.28)
    arrParam(9) = UCase(Trim(frm1.txtIvDt.text))
	
	iCalledAspName = AskPRAspName("m4111ra1_KO441")
MSGBOX 	iCalledAspName 
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m4111ra1_KO441", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetGrRef(strRet)
	End If	
		
End Function


Function OpenExceptGrRef()

	Dim strRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	if lgIntFlgMode = parent.OPMD_CMODE then				'��: Indicates that current mode is Update mode
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	
	if UCase(Trim(frm1.hdnPostingFlg.Value)) = "Y" then
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
	
	if UCase(Trim(frm1.hdnImportFlg.Value)) = "Y" then
		Call DisplayMsgBox("17A012", "X","���԰�","�������������")
		Exit Function
	End if
	'hdnExceptflg(�����ΰ�� Y)
	'if Trim(UCase(frm1.hdnRcptFlg.Value)) = "N" or UCase(Trim(frm1.hdnRetFlg.Value)) = "Y" or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
	if (Trim(UCase(frm1.hdnPoNo.Value)) <> ""  and Trim(UCase(frm1.hdnRcptType.Value)) = "" and Trim(UCase(frm1.hdnIssueType.Value)) = "") or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
		Call DisplayMsgBox("17a015", "X","��������" & ":" & frm1.txtIvTypeCd.value & "(" & frm1.txtIvTypeNm.value & ")" ,"�������������" )
		Exit Function
	End if
	
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnMvmtType.Value)
	arrParam(1) = Trim(frm1.hdnSppl.Value)
	arrParam(2) = Trim(frm1.hdnGrp.Value)
	arrParam(3) = "IV"
	arrParam(4) = Trim(frm1.hdnIvType.Value)
	arrParam(5) = UCase(Trim(frm1.hdnPoNo.value))
'	arrParam(5) = ""
    arrParam(6) = UCase(Trim(frm1.txtCur.value))
    '����(2003.03.24)
    arrParam(7) = UCase(Trim(frm1.hdnLcKind.value))
    arrParam(8) = UCase(Trim(frm1.hdnPayMeth.value))
	
	iCalledAspName = AskPRAspName("m4112ra1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m4112ra1_KO441", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetGrRef(strRet)
	End If	
		
End Function

'------------------------------------------  SetGrRef()  --------------------------------------------
Function SetGrRef(strRet)

Dim Index1,index2,Index3,Count1 ',Count2
Dim IntIflg
Dim strMessage
Dim VatRt
Dim VatType,vatflg
Dim Row	
Dim temp, TempRow

Const C_PlantCd_Ref		= 0
Const C_PlantNm_Ref		= 1
Const C_ItemCd_Ref		= 2
Const C_ItemNm_Ref		= 3
Const C_SpplSpec_Ref    = 4  
Const C_PoQty_Ref		= 5
Const C_MvmtQty_Ref	    = 6
Const C_NoIvQty_Ref		= 7
Const C_Unit_Ref		= 8
Const C_VatType_Ref		= 9
Const C_VatNm_Ref		= 10
Const C_VatRate_Ref		= 11
Const C_RcptDt_Ref      = 12
Const C_RcptNo_Ref		= 13
Const C_GmNo_Ref		= 14
Const C_GmSeqNo_Ref		= 15
Const C_PoNo_Ref  		= 16
Const C_PoSeqNo_Ref		= 17  
Const C_PoPrc_Ref		= 18  
Const C_PoDocAmt_Ref	= 19  
Const C_IvQty_Ref		= 20  
Const C_MvmtNo_Ref		= 21
Const C_TrackingNo_Ref	= 22
Const C_VatIncFlag_Ref	= 23
Const C_MvmtDocAmt_Ref 	= 24       
Const C_TotIvMNum_Ref 	= 25 
Const C_AmtUptFlg_Ref   = 26
Const C_PrcCtrlFlg_Ref 	= 27
Const C_VatDocAmt_Ref 	= 28
Const C_SumVatDoc_Ref   = 29
Const C_VatAmtRvsFlg_Ref= 30
Const C_IvQty1_Ref   	= 31
Const C_RetFlg_Ref   	= 32
Const C_Lc_No_Ref		= 33
Const C_Lc_Seq_Ref		= 34
Const C_LcPrice_Ref		= 35
Const C_XchRt_Ref		= 36		'#���Գ������� ȯ�� ���� (2003.09.21)

	VatType = frm1.hdnVatType.value
    VatRt  = UNICDbl(frm1.hdnVatRt.value)
	
	Count1 = Ubound(strRet,1)

	strMessage = ""
	IntIflg=true
	
	with frm1
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
		TempRow = .vspdData.MaxRows					'����Ʈ max��				

	for index1 = 0 to Count1 '.vspdData.MaxRows

		If TempRow <> 0 Then

			For Index3 = 1 To TempRow				'���� No�� ������ Row�� �߰����� �ʴ´�.
				.vspdData.Row=Index3
				.vspdData.Col=C_MvmtNo
				if .vspdData.Text = strRet(index1,21) then
					strMessage = strMessage & strRet(Index1,21) & ";"
					intIflg=False
					Exit for
				End if 
			Next
		
		End If
	    '0����,1�����,2ǰ��,3ǰ���, 4���ּ���,5�԰����,6�̸��Լ���,7����,8�԰���,9�԰��ȣ(rcptmvmtno),10���ó����ȣ(gmno),
	    '11���ó������,12���ֹ�ȣ,13���ּ���,14���ִܰ�,15�԰�ܰ�,16����Լ���,17�԰��ȣ(mvmtno)
		
		if IntIflg <> False then

	         ggoSpread.InsertRow
			.vspdData.Row=.vspdData.ActiveRow 
			'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
			'Call SetSpreadColorRef(.vspdData.Row)
			
			'***����(2003.03.24)-Lee Eun Hee
			If Trim(strRet(index1,C_Lc_No_Ref)) <> "" Then
				.vspdData.Col=C_LcFlg
				.vspdData.Text="B"       '�߰� (�� Local LC�� ���)
			End If
			
			'**���Գ������� ȯ�� ������(2003.09.19)
			.vspdData.Col=C_LcFlg
			If Trim(.vspdData.Text) = "B" Then
			.vspdData.Col=C_XchRt
			.vspdData.Text= strRet(index1,C_XchRt_Ref)
			Else
			.vspdData.Col=C_XchRt
			.vspdData.Text= frm1.hdnXch.value
			End If
									
			.vspdData.Col=C_PlantCd
			.vspdData.Text=strRet(index1,C_PlantCd_Ref)
	
			.vspdData.Col=C_PlantNm
			.vspdData.Text=strRet(index1,C_PlantNm_Ref)
			
			.vspdData.Col=C_ItemCd
			.vspdData.Text=strRet(index1,C_ItemCd_Ref)
			
			.vspdData.Col=C_ItemNm
			.vspdData.Text=strRet(index1,C_ItemNm_Ref)
			
			.vspdData.Col=C_SpplSpec
		    .vspdData.Text=strRet(index1,C_SpplSpec_Ref)
			
			if Trim(strRet(index1,C_MvmtQty_Ref)) = "" then
				strRet(index1,C_MvmtQty_Ref) = "0"
			End if			
			if Trim(strRet(index1,C_PoPrc_Ref)) = "" then  '���Դܰ� 
				strRet(index1,C_PoPrc_Ref) = "0"
			End if
			
			if Trim(strRet(index1,C_LcPrice_Ref)) = "" then  '���Դܰ� 
				strRet(index1,C_LcPrice_Ref) = "0"
			End if
					
			if Trim(strRet(index1,C_IvQty_Ref)) = "" then  '����Լ��� 
				strRet(index1,C_IvQty_Ref) = "0"
			End if
			
			'**����(�̸��Լ���=�԰����-����Լ���-(after_LC����-after_LC����Լ���))-2003.03.18-Lee,Eun Hee
			.vspdData.Col=C_IvQty1
			if Trim(strRet(index1,C_NoIvQty_Ref)) = "" then
				strRet(index1,C_NoIvQty_Ref) = "0"
			End if
			.vspdData.Text=strRet(index1,C_NoIvQty_Ref)
			
			.vspdData.Col=C_Unit
			.vspdData.Text=strRet(index1,C_Unit_Ref)
			
			.vspdData.Col=C_LcFlg
			If UCase(Trim(.vspdData.Text)) = "B" or UCase(Trim(.vspdData.Text)) = "A" Then
				'Local LC�� �԰������ϴ� ��� 
				.vspdData.Col=C_Cost                  '���Դܰ� 
				.vspdData.Text=strRet(index1,C_LcPrice_Ref)
			Else
				.vspdData.Col=C_Cost                  '���Դܰ� 
				.vspdData.Text=strRet(index1,C_PoPrc_Ref)
			End If

			If UCase(Trim(strRet(index1,C_RetFlg_Ref))) = "Y" then
				.vspdData.Text =  UNICDbl(.vspdData.Text) * -1
			     ggoSpread.SSSetProtected		C_Cost, .vspdData.Row
			End if
			
			.vspdData.Col=C_retflg                    '��ǰ���� 
			.vspdData.Text=strRet(index1,C_RetFlg_Ref)

			'If UCase(Trim(.vspdData.Text)) = "Y" then
			'     .vspdData.Text =  UNIFormatNumber(UNICDbl(.vspdData.Text) * -1,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
			'End if
			
			.vspdData.Col=C_OrderQty
			if Trim(strRet(index1,C_PoQty_Ref)) = "" then
				strRet(index1,C_PoQty_Ref) = "0"
			End if
			.vspdData.Text=strRet(index1,C_PoQty_Ref)
			
			.vspdData.Col=C_GmQty                   '�԰���� 
			.vspdData.Text=strRet(index1,C_MvmtQty_Ref)
			
			'@����@(2003.02.17)
			.vspdData.Col = C_PoVatIncFlg
			.vspdData.Text = strRet(index1,C_VatIncFlag_Ref)	
			'�߰�(12��)

            'VAT���Ա��и� 
			
			If strRet(index1,C_VatIncFlag_Ref) = "" Then
			   strRet(index1,C_VatIncFlag_Ref) = .hdvatFlg.value  'header�� �ִ� ���� �־��ش� 
			End If
			
			.vspdData.Col = C_IOFlg
			vatflg = strRet(index1,C_VatIncFlag_Ref)
			If strRet(index1,C_VatIncFlag_Ref) = "2" Then
				.vspdData.value = 1
			ElseIf strRet(index1,C_VatIncFlag_Ref) = "1" Then
				.vspdData.value = 0
			End If

			'VAT���Ա����ڵ� 
			.vspdData.Col = C_IOFlgCd
			.vspdData.Text = strRet(index1,C_VatIncFlag_Ref)
            
            .vspdData.Col=C_VatType                 'vat 
            .vspdData.Text=VatType
            .vspdData.Col=C_VatRate                 'vat �� 

			.vspdData.Text=UNIFormatNumber(VatRt, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

            
			.vspdData.Col=C_IvQty2					'���ԿϷ����=����Լ���		
			.vspdData.Text=strRet(index1,C_IvQty_Ref)
			
			.vspdData.Col=C_OrderCost
			.vspdData.Text=strRet(index1,C_PoPrc_Ref)        '���ִܰ� 
			
			.vspdData.Col=C_PoNo
			.vspdData.Text=strRet(index1,C_PoNo_Ref)        '���ֹ�ȣ 
		
			.vspdData.Col=C_PoSeq
			.vspdData.Text=strRet(index1,C_PoSeqNo_Ref)        '���ּ��� 
			
			.vspdData.Col=C_MvmtRcptNo              '�԰��ȣ 
			.vspdData.Text=strRet(index1,C_RcptNo_Ref)
			
			.vspdData.Col=C_GmNo
			.vspdData.Text=strRet(index1,C_GmNo_Ref)
			
			.vspdData.Col=C_GmSeq                   '���ó������ 
			.vspdData.Text=strRet(index1,C_GmSeqNo_Ref)
			 
			.vspdData.Col=C_MvmtNo                 'MVMTNO(hidden)
			.vspdData.Text=strRet(index1,C_MvmtNo_Ref)
			
	        .vspdData.Col=C_MvmtIvQty
            .vspdData.Text=strRet(index1,C_IvQty_Ref)       '�߰� 
             
            .vspdData.Col=C_oldIvQty1
             temp = 0
            .vspdData.Text= UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)

	        .vspdData.Col=C_TrackingNo
            .vspdData.Text=strRet(index1,C_TrackingNo_Ref)       '�߰� 
            
			.vspdData.Col = 1
			.vspdData.Action = 0
	
            '2002.09.10 �߰� 
            .vspdData.Col=C_PoAmt
            .vspdData.Text=strRet(index1,C_PoDocAmt_Ref)       '�߰� 
            
            .vspdData.Col=C_MvmtAmt
            .vspdData.Text=strRet(index1,C_MvmtDocAmt_Ref)       '�߰� 
            '.vspdData.Text=UNIFormatNumber(uniCdbl(strRet(index1,C_MvmtDocAmt_Ref)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
            '.vspdData.Text=UNIConvNumPCToCompanyByCurrency(UniCdbl(strRet(index1,C_MvmtDocAmt_Ref)) ,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")      '�߰� 
            
            'SumNetAmt = UNIConvNumPCToCompanyByCurrency(SumNetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")

            .vspdData.Col=C_TotIvDocAmt
            .vspdData.Text=strRet(index1,C_TotIvMNum_Ref)       '�߰� 
            
            
            .vspdData.Col=C_upt_amt_flg
            .vspdData.Text=strRet(index1,C_AmtUptFlg_Ref)       '�߰� 
            
            .vspdData.Col=C_prcflg
            .vspdData.Text=strRet(index1,C_PrcCtrlFlg_Ref)       '�߰� 

            .vspdData.Col=C_PoVatAmt
            .vspdData.Text=strRet(index1,C_VatDocAmt_Ref)       '�߰� 
            
            .vspdData.Col=C_TotIvVatAmt
            .vspdData.Text=strRet(index1,C_SumVatDoc_Ref)       '�߰� 
  
			.vspdData.Col=C_vat_rvs_flg
            .vspdData.Text=strRet(index1,C_VatAmtRvsFlg_Ref)       '�߰� 
            
            .vspdData.Col=C_PoIvQty           
            .vspdData.Text=strRet(index1,C_IvQty1_Ref)       '�߰� 
            
             If VatRt <> UNICDbl(strRet(index1,C_VatRate_Ref)) then
                 .vspdData.Text="N"
                 .vspdData.Col=C_ref_vatrate_flg
                 .vspdData.Text = "N"    
             Else
                .vspdData.Col=C_ref_vatrate_flg
                .vspdData.Text = "Y"    
             End If
			'**����(2003.03.25)
			'.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			'NetAmt =  UNICDbl(strRet(index1,C_PoPrc_Ref)) * UNICDbl(strRet(index1,C_NoIvQty_Ref))
			'.vspdData.Text= UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			
			.vspdData.Col = C_LcFlg
			If UCase(Trim(.vspdData.Text)) = "B" or UCase(Trim(.vspdData.Text)) = "A" Then
			'Local LC�� �԰������ϴ� ��� 
			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			NetAmt =  UNICDbl(strRet(index1,C_LcPrice_Ref)) * UNICDbl(strRet(index1,C_NoIvQty_Ref))
			.vspdData.Text= UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			Else
			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			NetAmt =  UNICDbl(strRet(index1,C_PoPrc_Ref)) * UNICDbl(strRet(index1,C_NoIvQty_Ref))
			.vspdData.Text= UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			
			End If
		
			If UCase(Trim(strRet(index1,C_RetFlg_Ref))) = "Y" then  '��ǰ���� 
                .vspdData.Col = C_ref_flg
                .vspdData.text = "Y"
                Call vspdData_Change(C_Cost, .vspdData.ActiveRow)
			else
	            call changeNetAmt(C_Cost_Ref,.vspdData.ActiveRow)
			end if
			
		    .vspdData.Col = C_ref_flg
            .vspdData.text = ""
			
			'---- Hidden �ʵ��� ������ ����Ͽ� �����ϱ� ���� �Լ� --------------------------------------------------------------------
			'---- ���� HIDDEN�ʵ��� ���� �ٷ� �Ʒ� vspddata_Change���� ���ǹǷ�, 
			'---- vspddata_change�� ȣ���ϱ� ����  ChangeAmtOrg�� ȣ��Ǿ�� �Ѵ� 
			
						'---- C_OrgNetAmt(HIDDEN)�� C_NetAmt�� ���� �Է�--------------------------------------------------------------------
			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)			
			temp = UNICDbl(Trim(.vspdData.Text))
			
			.vspdData.Col = C_OrgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(temp,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")

			.vspdData.Col = C_ChgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(temp,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")			

			'---- ���� VAT�ݾ��� HIDDEN �ʵ忡 ����--------------------------------------------------------------------
			Dim tmpVatDocAmt
			.vspdData.Col=C_VatDocAmt
			tmpVatDocAmt = UNICDbl(.vspdData.Text)

			.vspdData.Col=C_OrgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")			
			 
			.vspdData.Col=C_ChgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")

			.vspdData.Col=C_chkVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")	

			Call ChangeAmtOrg(C_IvAmt,.vspdData.ActiveRow,"Y")	

			Call vspdData_Change(C_Cost_Ref,.vspdData.ActiveRow)    '�����ڱ��ݾ�,VAT�ڱ��ݾ� ó�� 
		

			'----������ ����� ���ڸ� ����ϴ� �κ�--------------------------------------------------------------------------------------
			'----������ NetAMT�� VatDocAmt ��ŭ ����� ������----------------------------------------------------------------------------
			Dim VatDocAmt
			Dim NetAmt
			frm1.vspdData.Col = C_VatDocAmt
			VatDocAmt = UNICDbl(frm1.vspdData.Text)	 				

			frm1.vspdData.Col = C_OrgVatDocAmt
			tmpVatDocAmt = UNICDbl(.vspdData.Text)	 							
			
			frm1.vspdData.Col = C_NetAmt
			NetAmt = UNICDbl(frm1.vspdData.Text)	 								 							
			
			'#�� LC�߰�#######
			.vspdData.Col = C_LCNo		'�߰� 
            .vspdData.Text = strRet(index1,C_Lc_No_Ref)		
			
			.vspdData.Col = C_LCSeqNo		'�߰� 
            .vspdData.Text = strRet(index1,C_Lc_Seq_Ref)										

			Call SetSpreadColorRef(.vspdData.Row)
			
		Else
			IntIFlg=True
		End If 
	Next
	
	'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
	Call HSumAmtNewCalc()
	'Call SetSpreadColorRef(-1)
	
	If strMessage<>"" then
		Call DisplayMsgBox("17a005","X",strmessage,"�԰��ȣ")
	End if
	 
	.vspdData.ReDraw = True
	
	End With
	
End Function
'If ��LC�߰�(2003.03.14)##################################################################################
'------------------------------------------  OpenLLCRef()  -------------------------------------------------
Function OpenLLCRef()

	Dim strRet
	Dim arrParam(7)
	Dim iCalledAspName
	Dim IntRetCD
	if lgIntFlgMode = parent.OPMD_CMODE then				'��: Indicates that current mode is Update mode
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if

	if UCase(Trim(frm1.hdnPostingFlg.Value)) = "Y" then
		Call DisplayMsgBox("17a009","X","X","X")
		Exit Function
	End if
	
	if UCase(Trim(frm1.hdnImportFlg.Value)) = "Y" then
		Call DisplayMsgBox("17A012", "X","���԰�","LOCAL L/C��������")
		Exit Function
	End if
	
	'hdnExceptflg(�����ΰ�� Y)
	'if Trim(UCase(frm1.hdnRcptFlg.Value)) = "N" or UCase(Trim(frm1.hdnRetFlg.Value)) = "Y" or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
	if (Trim(UCase(frm1.hdnPoNo.Value)) <> ""  and Trim(UCase(frm1.hdnRcptType.Value)) = "" and Trim(UCase(frm1.hdnIssueType.Value)) = "") or Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" then
		Call DisplayMsgBox("17a015", "X","��������" & ":" & frm1.txtIvTypeCd.value & "(" & frm1.txtIvTypeNm.value & ")" ,"LOCAL L/C��������" )
		Exit Function
	End if
	
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnSppl.Value)
	arrParam(1) = Trim(frm1.hdnGrp.Value)
	arrParam(2) = "IV"
	arrParam(3) = Trim(frm1.hdnIvType.Value)
	arrParam(4) = UCase(Trim(frm1.hdnPoNo.value))
    arrParam(5) = UCase(Trim(frm1.txtCur.value))
    '����(2003.03.26)
    arrParam(6) = UCase(Trim(frm1.hdnLcKind.value))
    arrParam(7) = UCase(Trim(frm1.hdnPayMeth.value))
	
	iCalledAspName = AskPRAspName("M3212RA5_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212RA5_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam,document), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lblnWinEvent = False
	
	'lgOpenFlag = False
	If isEmpty(strRet) Then Exit Function				'�������� ã�� �� ���� �����߻���.
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetLLCRef(strRet)
	End If	
		
End Function

'------------------------------------------  SetLLCRef()  ----------------------------------------------
Function SetLLCRef(strRet)

    Dim Index1,index2,Index3,Count1,Count2
    Dim IntIflg
    Dim strMessage
    Dim strMessage2
    Dim strtemp1,strtemp2
    Dim temp , TempRow
    Dim VatRt
	Dim VatType,vatflg
	Dim DocAmt, NetAmt
	Dim strLcNo, strLcSeqNo
	
	Const C_LcNo_Ref		= 0
	Const C_LCSeq_Ref		= 1
	Const C_OpenDt_Ref		= 2
	Const C_PlantCd_Ref		= 3
	Const C_PlantNm_Ref		= 4
	Const C_ItemCd_Ref		= 5
	Const C_ItemNm_Ref		= 6
	Const C_Spec_Ref		= 7	
	Const C_PoQty_Ref		= 8
	Const C_Qty_Ref			= 9
	Const C_MvmtQty_Ref		= 10
	Const C_NotQty_Ref	    = 11
	Const C_Unit_Ref		= 12
	Const C_Price_Ref		= 13
	Const C_DocAmt_Ref		= 14	
	Const C_PoNo_Ref		= 15
	Const C_PoSeq_Ref		= 16
	Const C_MvmtRcptNo_Ref	= 17
	Const C_GmNo_Ref		= 18
	Const C_GmSeqNo_Ref		= 19
	Const C_TrackingNo_Ref	= 20
	Const C_PoPrc_Ref		= 21
	Const C_PoDocAmt_Ref	= 22
	Const C_IvQty_Ref		= 23
	Const C_MvmtNo_Ref		= 24
	Const C_AmtUptFlg_Ref	= 25
	Const C_XchRate_Ref		= 26	'L/Cȯ�� �߰� 

	VatType = frm1.hdnVatType.value
    VatRt  = UNICDbl(frm1.hdnVatRt.value)
    
	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2) 
	strMessage = ""
	strMessage2 = ""
	IntIflg=true
	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
			
		.vspdData.ReDraw = False
		TempRow = .vspdData.MaxRows					'����Ʈ max��		
		
	For index1 = 0 To Count1
		
		If TempRow <> 0 Then

			For Index3=1 To TempRow		'���� No�� ������ Row�� �߰����� �ʴ´�.
				.vspdData.Row=Index3
				.vspdData.Col=C_LCNo
				strLcNo = .vspdData.Text
				.vspdData.Col=C_LCSeqNo
				strLcSeqNo = .vspdData.Text
				if strLcNo = strRet(index1,C_LcNo_Ref) And strLcSeqNo = strRet(index1,C_LCSeq_Ref) then
					strMessage = strMessage & strRet(Index1,C_LcNo_Ref) & ";"
					intIflg=False
					Exit for
				End if 
			Next
		End If
		
		if IntIflg <> False then
			'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee
	         ggoSpread.InsertRow
			
			.vspdData.Row=.vspdData.ActiveRow 
			'Call SetSpreadColorRef(.vspdData.ActiveRow)
			
			'***����(2003.10.07)***
			'Local LC�������� �� LLC�� �� LLC�� ��� �����Ѵ�. LC_FLG�� ���� 
			If Trim(strRet(index1,C_MvmtNo_Ref)) <> "" Then
			.vspdData.Col = C_LcFlg
			.vspdData.Text = "A"       '(�� Local LC�� ���)
			Else
			.vspdData.Col = C_LcFlg
			.vspdData.Text = "B"       '(�� Local LC�� ���)
			End If
				
			.vspdData.Col=C_LCNo
			.vspdData.Text=strRet(index1,C_LcNo_Ref)
			.vspdData.Col=C_LCSeqNo
			.vspdData.Text=strRet(index1,C_LCSeq_Ref)
			.vspdData.Col=C_PlantCd
			.vspdData.Text=strRet(index1,C_PlantCd_Ref)
			.vspdData.Col=C_PlantNm
			.vspdData.Text=strRet(index1,C_PlantNm_Ref)
			.vspdData.Col=C_ItemCd
			.vspdData.Text=strRet(index1,C_ItemCd_Ref)
			.vspdData.Col=C_ItemNm
			.vspdData.Text=strRet(index1,C_ItemNm_Ref)
				
			.vspdData.Col=C_SpplSpec
			.vspdData.Text=strRet(index1,C_Spec_Ref)
				
			
			if Trim(strRet(index1,C_MvmtQty_Ref)) = "" then
				strRet(index1,C_MvmtQty_Ref) = "0"
			End if			
			if Trim(strRet(index1,C_PoPrc_Ref)) = "" then  '���Դܰ� 
				strRet(index1,C_PoPrc_Ref) = "0"
			End if
					
			if Trim(strRet(index1,C_IvQty_Ref)) = "" then  '����Լ��� 
				strRet(index1,C_IvQty_Ref) = "0"
			End if
			
			if Trim(strRet(index1,C_NotQty_Ref)) = "" then  '�̸��Լ��� 
				strRet(index1,C_IvQty_Ref) = "0"
			End if
			 
			.vspdData.Col=C_IvQty1                '���Լ��� = �԰����-����Լ��� 
			.vspdData.Text=strRet(index1,C_NotQty_Ref)
			.vspdData.Col=C_Unit
			.vspdData.Text=strRet(index1,C_Unit_Ref)
			
			.vspdData.Col=C_Cost                  '���Դܰ�=Local LC�ܰ� 
			.vspdData.Text=strRet(index1,C_Price_Ref)
			
			.vspdData.Col=C_retflg                    '��ǰ���� 
			.vspdData.Text="N"
			
			.vspdData.Col=C_OrderQty				'���ּ��� 
			if Trim(strRet(index1,C_PoQty_Ref)) = "" then
				strRet(index1,C_PoQty_Ref) = "0"
			End if
			.vspdData.Text=strRet(index1,C_PoQty_Ref)
			
			.vspdData.Col=C_GmQty                   '�԰���� 
			.vspdData.Text=strRet(index1,C_MvmtQty_Ref)
			
			.vspdData.Col = C_PoVatIncFlg
			.vspdData.Text = .hdvatFlg.value

            'VAT���Ա��и� 
			
			.vspdData.Col = C_IOFlg
			vatflg = .hdvatFlg.value
			If vatflg = "2" Then
				.vspdData.value = 1
			ElseIf vatflg = "1" Then
				.vspdData.value = 0
			End If

			'VAT���Ա����ڵ� 
			.vspdData.Col = C_IOFlgCd
			.vspdData.Text = vatflg
            
            .vspdData.Col=C_VatType                 'vat 
            .vspdData.Text=VatType
            .vspdData.Col=C_VatRate                 'vat �� 
			.vspdData.Text= "0" '������~~
			
			.vspdData.Col=C_ref_vatrate_flg
            .vspdData.Text = "L"
                
			.vspdData.Col=C_IvQty2					'���ԿϷ����=����Լ���		
			.vspdData.Text=strRet(index1,C_IvQty_Ref)
			
			.vspdData.Col=C_OrderCost
			.vspdData.Text=strRet(index1,C_PoPrc_Ref)        '���ִܰ� 
			
			.vspdData.Col=C_PoNo
			.vspdData.Text=strRet(index1,C_PoNo_Ref)        '���ֹ�ȣ 
		
			.vspdData.Col=C_PoSeq
			.vspdData.Text=strRet(index1,C_PoSeq_Ref)        '���ּ��� 
			
			.vspdData.Col=C_MvmtRcptNo              '�԰��ȣ 
			.vspdData.Text=strRet(index1,C_MvmtRcptNo_Ref)
			
			.vspdData.Col=C_GmNo
			.vspdData.Text=strRet(index1,C_GmNo_Ref)
			
			.vspdData.Col=C_GmSeq                   '���ó������ 
			.vspdData.Text=strRet(index1,C_GmSeqNo_Ref)
			 
			.vspdData.Col=C_MvmtNo                 'MVMTNO(hidden)
			.vspdData.Text=strRet(index1,C_MvmtNo_Ref)
			
	        .vspdData.Col=C_MvmtIvQty
            .vspdData.Text=strRet(index1,C_IvQty_Ref)       '�߰� 
             
            .vspdData.Col=C_oldIvQty1
             temp = 0
            .vspdData.Text= UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)

	        .vspdData.Col=C_TrackingNo
            .vspdData.Text=strRet(index1,C_TrackingNo_Ref)       '�߰� 
            
			.vspdData.Col = 1
			.vspdData.Action = 0
	
            '2002.09.10 �߰� 
            .vspdData.Col=C_PoAmt
            .vspdData.Text=strRet(index1,C_PoDocAmt_Ref)       '�߰� 

            .vspdData.Col=C_upt_amt_flg
            .vspdData.Text=strRet(index1,C_AmtUptFlg_Ref)       '�߰� 

			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = LC�ܰ� * (�԰����-����Լ���)
			NetAmt =  UNICDbl(strRet(index1,C_Price_Ref)) * UNICDbl(strRet(index1,C_NotQty_Ref))
			.vspdData.Text= UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			
			.vspdData.Col=C_VatDocAmt
			tmpVatDocAmt = 0
			.vspdData.Text= UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
		    
		    .vspdData.Col=C_VatLocAmt
			.vspdData.Text= UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")'vatloc ���� 
		    
		    .vspdData.Col = C_ref_flg
            .vspdData.text = "Y"
						
			'**���Գ������� ȯ�� ������(2003.09.19)**
			'Local L/C ���μ����� ��ģ ��� Local L/Cȯ���� �ڱ��ݾ��� ����ؾ� �Ѵ�.
			.vspdData.Col = C_XchRt
			.vspdData.Text = strRet(index1,C_XchRate_Ref) '200308 ����������ġ 
			
			'+++++++++++++++++++++++++++++++++++++
			'DocAmt ��� 
			If vatflg = "1" Then
			    DocAmt = NetAmt
			Else
			    DocAmt = NetAmt + UNICDbl(UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X"))
			End If    
                 
			.vspdData.Col = C_IvAmt
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			
			'+++++++++++++++++++++++++++++++++++++
			'---- Hidden �ʵ��� ������ ����Ͽ� �����ϱ� ���� �Լ� --------------------------------------------------------------------
			'---- ���� HIDDEN�ʵ��� ���� �ٷ� �Ʒ� vspddata_Change���� ���ǹǷ�, 
			'---- vspddata_change�� ȣ���ϱ� ����  ChangeAmtOrg�� ȣ��Ǿ�� �Ѵ� 
			
			'---- C_OrgNetAmt(HIDDEN)�� C_NetAmt�� ���� �Է�--------------------------------------------------------------------
			.vspdData.Col = C_NetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)			
			temp = UNICDbl(Trim(.vspdData.Text))
			
			.vspdData.Col = C_OrgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(temp,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")

			.vspdData.Col = C_ChgNetAmt					'Invoice�ݾ� = ���ִܰ� * (�԰����-����Լ���)
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(temp,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")			

			'---- ���� VAT�ݾ��� HIDDEN �ʵ忡 ����--------------------------------------------------------------------
			Dim tmpVatDocAmt
			.vspdData.Col=C_VatDocAmt
			tmpVatDocAmt = 0

			.vspdData.Col=C_OrgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")			
			 
			.vspdData.Col=C_ChgVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")
			'------------------------------------------------------------------------
			.vspdData.Col=C_chkVatDocAmt
			.vspdData.Text=UNIConvNumPCToCompanyByCurrency(tmpVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")	

			Call ChangeAmtOrg(C_IvAmt,.vspdData.ActiveRow,"Y")	
			'---------------------------------------------------------------------------------------------------------------------------

			Call vspdData_Change(C_Cost_Ref,.vspdData.ActiveRow)    '�����ڱ��ݾ�,VAT�ڱ��ݾ� ó�� 
			
			Call SetSpreadColorRef(.vspdData.Row)
			'ggoSpread.SSSetProtected        C_Cost, .vspdData.ActiveRow, .vspdData.ActiveRow

		Else
			IntIFlg=True
		End if 
	Next
	
	'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
	Call HSumAmtNewCalc()
	'Call SetSpreadColorRef(-1)
	'ggoSpread.SSSetProtected        C_Cost, -1, -1
	
	If strMessage<>"" then
		Call DisplayMsgBox("17a005","X",strmessage,"LOCAL L/C��ȣ")
	End if
		
	.vspdData.ReDraw = True
		
	End with
	
End Function
'End If##########################################################################################
'------------------------------------------  OpenRetRef()  -------------------------------------------------
'	Name : OpenRetRef()
'	Description : ���ܹ�ǰ������� 
'---------------------------------------------------------------------------------------------------------
Function OpenRetRef()
	Dim strRet
	Dim arrParam(15)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = parent.OPMD_CMODE then				'��: Indicates that current mode is Update mode
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if


'	if Not(UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "Y") then
'		Call DisplayMsgBox("17A012", "X","���������" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","���ܹ�ǰ�������" )
'		frm1.txtGrNo.focus	
'		Exit Function
'	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	'===============����====================
	arrParam(0) = Trim(frm1.txtSpplCd.value)
	'===============����====================

	iCalledAspName = AskPRAspName("M4132RA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4132RA1_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		frm1.txtIvNo.focus	
		Exit Function
	Else
		Call SetRetRef(strRet)
	End If	
End Function
''==============================================================================================================================
Function SetRetRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim comtemp1,comtemp2,temp
	Dim iInsRow
	
	Const C_PlantCd_Ref		= 0		' ���� 
	Const C_PlantNm_Ref		= 1
	Const C_ItemCd_Ref		= 2		' ǰ�� 
	Const C_ItemNm_Ref		= 3
	Const C_MvmtQty_Ref		= 4		' ��ǰ������ 
	Const C_TotRetQty_ref	= 5		' ���԰���� 
	Const C_IvQty_Ref		= 6 	' ���Լ��� 
	Const C_Unit_Ref		= 7 	' ���� 
	Const C_MvmtDt_Ref		= 8 	' ��ǰ����� 
	Const C_MvmtRcptNo_Ref	= 9 	' ��ǰ����ȣ 
	Const C_GmNo_Ref		= 10	' ���ó����ȣ 
	Const C_GmSeqNo_Ref		= 11	' ���ó������ 
	Const C_TrackingNo_Ref  = 12	' Tracking No
	Const C_Lot_No_Ref		= 13	' Lot No.
	Const C_Lot_Seq_Ref		= 14	' Lot No. ���� 
	Const C_SpplSpec_Ref    = 15	' ǰ��԰� 
	Const C_Prc_Ref 		= 16	' ���ִܰ� 
	Const C_Amt_Ref 		= 17
	Const C_SlCd_Ref		= 18	' â�� 
	Const C_SlNm_Ref		= 19
	Const C_Trackingflg_Ref = 20	' TRACKINGFLG
	Const C_MvmtNo_Ref		= 21	' ����ȣ 
	Const C_ItemPrc_Ref		= 22	' ǰ��ܰ� 

	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		intStartRow = .vspdData.MaxRows + 1
		
		.vspdData.Redraw = False
		
		TempRow = .vspdData.MaxRows					'����Ʈ max�� 
		
		For index1 = 0 to Count1
		
			If TempRow <> 0 Then
	
				For Index3 = 1 To TempRow				'���� No�� ������ Row�� �߰����� �ʴ´�.
					.vspdData.Row=Index3
					.vspdData.Col=C_MvmtNo
					if .vspdData.Text = strRet(index1,C_MvmtNo_Ref) then
						strMessage = strMessage & strRet(Index1,C_MvmtNo_Ref) & ";"
						intIflg=False
						Exit for
					End if 
				Next
			
			End If

			.vspdData.Row = .vspdData.ActiveRow

			If IntIflg <> False then
				.vspdData.MaxRows = CLng(TempRow) + CLng(index1) + 1
				iInsRow = CLng(TempRow) + CLng(index1) + 1
	
				Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_TrackingNo,	iInsRow, "*")

				Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Ref))
				Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Ref))
				Call .vspdData.SetText(C_itemCd		,	iInsRow, strRet(index1,C_ItemCd_Ref))
				Call .vspdData.SetText(C_itemNm		,	iInsRow, strRet(index1,C_ItemNm_Ref))
				Call .vspdData.SetText(C_Cost		,	iInsRow, strRet(index1,C_Prc_Ref) * (-1))	' ��ǰ.
				Call .vspdData.SetText(C_retflg		,	iInsRow, "Y")

				Call .vspdData.SetText(C_SpplSpec	,	iInsRow, strRet(index1,C_SpplSpec_Ref))
				Call .vspdData.SetText(C_Unit		,	iInsRow, strRet(index1,C_Unit_Ref))
				Call .vspdData.SetText(C_MvmtRcptNo	,	iInsRow, strRet(index1,C_MvmtRcptNo_Ref))
				Call .vspdData.SetText(C_GmNo		,	iInsRow, strRet(index1,C_GmNo_Ref))
				Call .vspdData.SetText(C_GmSeq		,	iInsRow, strRet(index1,C_GmSeqNo_Ref))
				Call .vspdData.SetText(C_TrackingNo	,	iInsRow, strRet(index1,C_TrackingNo_Ref))
				Call .vspdData.SetText(C_MvmtNo		,	iInsRow, strRet(index1,C_MvmtNo_Ref))
				Call .vspdData.SetText(C_XchRt		,	iInsRow, .hdnXch.value)
				Call .vspdData.SetText(C_IOFlgCd	,	iInsRow, .hdvatFlg.value)
				Call .vspdData.SetText(C_VatType	,	iInsRow, .hdnVatType.value)
	            .vspdData.Row = iInsRow
	            .vspdData.Col = C_IOFlg
				if .hdvatFlg.value = "2" Then
				    .vspdData.value = 1		'vat���� 
				ElseIf .hdvatFlg.value = "1" Then
					.vspdData.value = 0		'vat������ 
				End If
					
				' ��ǰ������ - ���԰���� - ���Լ��� 
				IF strRet(index1,C_TotRetQty_ref) <> "" Then
					temp= UNICDbl(strRet(index1,C_MvmtQty_Ref)) - UNICDbl(strRet(index1,C_TotRetQty_ref))
					Call .vspdData.SetText(C_IvQty1,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				End If
				IF strRet(index1,C_IvQty_Ref) <> "" Then
					temp = UNICDbl(GetSpreadText(.vspdData,C_IvQty1,iInsRow,"X","X")) - UNICDbl(strRet(index1,C_IvQty_Ref))
					Call .vspdData.SetText(C_IvQty1,	iInsRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				End If
				
				Call vspddata_change(C_IvQty1, iInsRow)

				Call SetSpreadColorRef(iInsRow)
				
				'ggoSpread.SSSetProtected		C_Cost, iInsRow
			Else
				IntIFlg=True
			End if 
		    
		Next
	
		intEndRow = .vspdData.MaxRows
		
		Call HSumAmtNewCalc()

		if strMessage<>"" then
			Call DisplayMsgBox("17a005","X",strmessage,"������ȣ")
			.vspdData.ReDraw = True
			Exit Function
		End if
		
		.vspdData.ReDraw = True
	
	End with

End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	frm1.vspdData.Col = C_PlantCd	
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 	 
	
	if  Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		Exit Function
	End if

	IsOpenPop = True
	
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(0) = Trim(frm1.vspdData.Text)
	
	frm1.vspdData.Col=C_ItemCd
	arrParam(1) = Trim(frm1.vspdData.Text)
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"	
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���	
	arrField(2) = 3 ' -- Spec
    
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_ItemCd:		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ItemNm:		frm1.vspdData.Text = arrRet(1)
		frm1.vspdData.Col = C_SpplSpec:		frm1.vspdData.Text = arrRet(2)
		Call SetActiveCell(frm1.vspdData,C_IvQty1,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function
'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_PlantCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_PlantCd:		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_PlantNm:		frm1.vspdData.Text = arrRet(1)
		Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
	
End Function
 '------------------------------------------  OpenUnit()  -------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ִ���"						' �˾� ��Ī 
	arrParam(1) = "B_Unit_OF_MEASURE"					' TABLE ��Ī 
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 
	arrParam(2) = Trim(frm1.vspdData.text)				' Code Condition%>
	arrParam(4) = ""									' Where Condition%>
	arrParam(5) = "���ִ���"						' TextBox ��Ī 
	
    arrField(0) = "Unit"								' Field��(0)
    arrField(1) = "Unit_Nm"								' Field��(1)
    
    arrHeader(0) = "���ִ���"						' Header��(0)
    arrHeader(1) = "���ִ�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_Unit,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col = C_Unit:		frm1.vspdData.Text = arrRet(0)
		Call SetActiveCell(frm1.vspdData,C_Cost,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function

'------------------------------------------  OpenVat()  -------------------------------------------------
Function OpenVat()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function 

	IsOpenPop = True
 
    frm1.vspdData.Col=C_VatType
	frm1.vspdData.Row=frm1.vspdData.ActiveRow 

	arrParam(0) = "VAT����"				
	arrParam(1) = "B_MINOR,b_configuration"	
	arrParam(2) = Trim(frm1.vspdData.Text)		
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "	
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT����"					
	
    arrField(0) = "b_minor.MINOR_CD"			
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"	
    
    arrHeader(0) = "VAT����"					
    arrHeader(1) = "VAT���¸�"				
    arrHeader(2) = "VAT��"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_VatType,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetVat(arrRet)
		Call SetActiveCell(frm1.vspdData,C_VatDocAmt,frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function
'------------------------------------------  SetVat()  -------------------------------------------------
Function SetVat(byval arrRet)
    Dim price,changeVatflg
    changeVatflg = "N"
	With frm1
		.vspdData.Col = C_VatType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_VatNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_VatRate
		.vspdData.Text = arrRet(2)
		
	End With
    Call ChangeAmt(C_NetAmt, frm1.vspdData.ActiveRow,changeVatflg)
	Call vspdData_Change(C_VatType , frm1.vspdData.ActiveRow )
	lgBlnFlgChgValue = True
End Function

'2007-04-16 added
Function OpenTrackingNo()

	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	If Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		IsOpenPop = False
		Exit Function
	End if
    
    arrParam(0) = ""
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""
	
	arrParam(4) = ""
	arrParam(5) = " and A.tracking_no not in (" & FilterVar("*", "''", "S") & " ) " 
	arrParam(6) = "M" 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_TrackingNo
		frm1.vspdData.Text = arrRet
	End If	

End Function

'2007-04-16 Modified
'=========================  SetUnitCost() ======================================================
Sub SetUnitCost( Row )
	Dim strssText1, strssText2, strssText3, strVal
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iCodeArr
    
	frm1.vspdData.Row = Row
    frm1.vspdData.Col = C_PlantCd
    strssText1 = Trim(frm1.vspdData.text)
    frm1.vspdData.Col = C_ItemCd
    strssText2 = Trim(frm1.vspdData.text)
    frm1.vspdData.Col = C_Unit
    strssText3 = Trim(frm1.vspdData.text)
    
    If strssText1 = "" Or strssText2 = "" Then
		Exit Sub
	End If
	
	Call CommonQueryRs(" TRACKING_FLG "," B_ITEM_BY_PLANT (NOLOCK) "," PLANT_CD = " & FilterVar(strssText1, "''", "S")  _
				& " AND ITEM_CD = " & FilterVar(strssText2, "''", "S") _
				,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				
	If Err.number <> 0 Or Trim(lgF0) = "" Then
		Call DisplayMsgBox("122600","X","X","X")
		Err.Clear 
		Call frm1.vspdData.SetText(C_ItemCd, Row, "")
		Call frm1.vspdData.SetText(C_ItemNm, Row, "")
		Call frm1.vspdData.SetText(C_Unit, Row, "")
		Call frm1.vspdData.SetText(C_Unit, Row, "")
		Call SetActiveCell(frm1.vspdData,C_ItemCd,Row,"M","X","X")
		Exit Sub
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	ggoSpread.Source = frm1.vspdData
	
	If UCase(Trim(lgF0(0))) = "Y" And UCase(Trim(frm1.hdnExceptflg.value)) = "Y" Then
		ggoSpread.SpreadUnLock          C_TrackingNo , Row ,C_TrackingPopup , Row     '�ܰ� 
		ggoSpread.SSSetRequired			C_TrackingNo, Row, Row
	Else
		ggoSpread.SpreadLock            C_TrackingNo , Row ,C_TrackingPopup , Row     '�ܰ� 
		ggoSpread.SSSetProtected		C_TrackingNo, Row, Row
		ggoSpread.SSSetProtected		C_TrackingPopup, Row, Row
	End If
	
	
	strVal = Row 
	strVal = strVal & parent.gColSep & strssText1
	strVal = strVal & parent.gColSep & strssText2
	strVal = strVal & parent.gColSep & strssText3 
	strVal = strVal & parent.gColSep & frm1.txtSpplCd.value 
	strVal = strVal & parent.gColSep & frm1.txtCur.value 
	strVal = strVal & parent.gColSep & frm1.txtIvDt.text 
	
	frm1.txtMode.value = "LookupUnitCost"
	frm1.txtSpread.value = strVal
	
	If Trim(frm1.txtSpread.value) <> "" Then
		If LayerShowHide(1) = False Then Exit Sub
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	End if

End Sub
'=========================  SetSpreadFloatLocal() ==================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	     
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '����������� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub

'================================== =====================================================
' Function Name : InitCollectType  �߰� 
' Function Desc : �Һ������ڵ�/��/�� �����ϱ� 
' ������� Ű���忡�� �Һ������ڵ带 ����� �Һ�������,�Һ���,���Աݾ�,NetAmount�� �����Ű�� �Լ� 
'========================================================================================

Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub
'=========================  GetCollectTypeRef() ==================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub
'=========================  SetVatType() ==================================================
Sub SetVatType()
	Dim VatType, VatTypeNm, VatRate
    
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_VatType

	VatType = Trim(frm1.vspdData.text)
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

    frm1.vspdData.Col = C_VatNm              'vat�� 
	frm1.vspdData.text = VatTypeNm
    
	frm1.vspdData.Col = C_VatRate            'vat�� 
	frm1.vspdData.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	
	lgBlnFlgChgValue = True
End Sub


'=========================  vspdData_MouseDown() ==================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'=========================  FncSplitColumn() ==================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function
'=========================  vspdData_Click() ==================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	 
	If interface_Account = "Y" Then
		If Trim(UCase(frm1.hdnImportflg.Value)) = "Y" Or lgIntFlgMode <> Parent.OPMD_UMODE Then
			Call SetPopupMenuItemInf("0000111111")
		ElseIf Trim(UCase(frm1.hdnImportflg.Value)) = "N" And frm1.vspdData.MaxRows < 1 Then
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
				Call SetPopupMenuItemInf("1001111111")
			Else
				Call SetPopupMenuItemInf("0001111111")
			End If
		Else
			If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
				Call SetPopupMenuItemInf("0000111111")
			Else
				If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
					Call SetPopupMenuItemInf("1101111111")
				Else
					Call SetPopupMenuItemInf("0101111111")
				End If
			End If
		End If
	Else
		If Trim(UCase(frm1.hdnImportflg.Value)) = "Y" Or lgIntFlgMode <> Parent.OPMD_UMODE Then
			Call SetPopupMenuItemInf("0000111111")
		ElseIf Trim(UCase(frm1.hdnImportflg.Value)) = "N" And frm1.vspdData.MaxRows < 1 Then
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
				Call SetPopupMenuItemInf("1001111111")
			Else
				Call SetPopupMenuItemInf("0001111111")
			End If
		Else
			If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
				Call SetPopupMenuItemInf("0000111111")
			Else
				If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
					Call SetPopupMenuItemInf("1101111111")
				Else
					Call SetPopupMenuItemInf("0101111111")
				End If
			End If
		End If
	End If
		
   
   gMouseClickStatus = "SPC"  
   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
		
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   
    frm1.vspdData.Row = Row
	  
End Sub
'=========================  vspdData_ColWidthChange() ==================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'=========================  vspdData_ScriptDragDropBlock() =============================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'=========================  PopSaveSpreadColumnInf() ==================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=========================  PopRestoreSpreadColumnInf() ==================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
	'Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	

	If Trim(UCase(frm1.hdnImportflg.Value)) = "Y" or Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
		Call SetRdSpreadColor(1)

	Elseif Trim(UCase(frm1.hdnPostingflg.Value)) = "N" Then
		Call QueryAtSetSpreadColor(1)

	End if
			
	
End Sub
'==========================================  ChangeAmtOrg()  ========================================
'	Name : ChangeAmtOrg()
'	Description : 
'	�԰�����, ���� ������ C_NetAmt , C_VatDocAmt�� ���ʰ��� ������ �ִ� 
'	C_OrgNetAmt , C_OrgVatDocAmt�� Hidden �ʵ忡 ���� �־� �ֱ� ���� ȣ���ϴ� �Լ� 
'========================================================================================================= 
Function ChangeAmtOrg(Col,Row,changeVatflg)

'---- vspddata_change �� C_COST�� �ش��ϴ� �κа� ������--------------------------------------------------------------------

		Dim Qty, Price
		Dim IvAmt, VatDocAmt, VatLocAmt, XchRt, VatRt, IvLocAmt,vat1,vat2 ,NetLocAmt,VatIncFlag,DocAmt		
		Dim tmpVatRate,tmpIvAmt,tmpChgNetAmt,tmpNetAmt,tmpChgIvAmt,tmpVatDocAmt,tmpChgVatDocAmt,tmpSumIvAmt
		Dim IvQty2,OrderQty1,MvmtIvQty1,TotalIvQty1,PoVatAmt1,IvVatAmt1
		Dim vat_rvs_flg
		ggoSpread.Source = frm1.vspdData

	    VatRt = UNICDbl(frm1.hdnVatRt.value)

		Frm1.vspdData.Row = Row
		Frm1.vspdData.Col = 0
					
		If Frm1.vspdData.text = ggoSpread.DeleteFlag Then Exit Function

		ggoSpread.UpdateRow Row

		Frm1.vspdData.Row = Row
		Frm1.vspdData.Col = Col
				  
		Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

		frm1.vspdData.Col = C_ivQty1
		If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
			Qty = 0
		Else
			Qty = UNICDbl(frm1.vspdData.Text)
		End If
		
		frm1.vspdData.Col = C_Cost
		If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
			Price = 0
		Else
			Price = UNICDbl(frm1.vspdData.Text)
		End If
		
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_vat_rvs_flg
	    vat_rvs_flg = Trim(frm1.vspdData.Text)
		
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_MvmtNo
		

'---- changeAmt �� C_IvAmt�� �ش��ϴ� �κа� ������--------------------------------------------------------------------
	'Dim IvAmt, VatDocAmt, VatLocAmt, VatRt, IvLocAmt,vat1,vat2 ,NetLocAmt,VatIncFlag,DocAmt

	'frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_IvAmt
	If frm1.vspdData.text <> "" Then
	    DocAmt = UNICDbl(frm1.vspdData.text)
	Else
	    DocAmt = 0
	End If

 	frm1.vspdData.Col = C_IOFlgCd
	VatIncFlag = frm1.vspdData.text
     	
     	Select Case Col 
		Case C_IvAmt
            
            If changeVatflg = "N" Then
			    vat1 = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / (VatRt + 100),frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'CInt(DocAmt * VatRt / (VatRt + 100))  'vat ���� vat �ݾ� 
			    vat2 = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / 100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat ���� vat �ݾ� 
            Else
                frm1.vspdData.Row = Row
                frm1.vspdData.Col = C_VatDocAmt
                vat1 = frm1.vspdData.text
                vat2 = frm1.vspdData.text
            End If
            
            
			If VatIncFlag = "2" Then          'vat �����ΰ�� 
				
				frm1.vspdData.Col = C_OrgNetAmt
				frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(DocAmt - UNICDbl(vat1),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")

				frm1.vspdData.Col = C_OrgVatDocAmt 'vat �ݾ� 
				frm1.vspdData.text = vat1
			
			    frm1.vspdData.Col = C_chkVatDocAmt '�������� flg seting �񱳰� 
				frm1.vspdData.text = vat1
				
				frm1.vspdData.Col = C_NetAmt
				
			Else                              'vat �����ΰ��			       
				frm1.vspdData.Col = C_OrgNetAmt   '���Աݾ� 
				frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")

				frm1.vspdData.Col = C_OrgVatDocAmt 'vat �ݾ� 
				frm1.vspdData.text = vat2
				
			    frm1.vspdData.Col = C_chkVatDocAmt '�������� flg seting �񱳰� 
				frm1.vspdData.text = vat2
			End If
		End Select 
		

End Function
'=============================================================================================
'  Function Name  : ChangeAmt()
'  Function Desc. : �ݾ׺���� ���Աݾ�, ���ݾ�, Vat�ڱ��ݾװ� �����ڱ����ݾ��� �ٽ� �����.
'  History        : 2003.03.26-Lee Eun Hee
'============================================================================================
Function ChangeAmt(Col,Row,chvatflg)
	Dim IvAmt, VatDocAmt, VatLocAmt, XchRt, VatRt, IvLocAmt,vat1,vat2 ,NetLocAmt,VatIncFlag,DocAmt
	Dim tmpVatRate,tmpIvAmt,tmpChgNetAmt,tmpNetAmt,tmpChgIvAmt,tmpVatDocAmt,tmpChgVatDocAmt,tmpSumIvAmt
	
	frm1.vspdData.Row = Row
	
	'Local L/C ���μ����� ��ģ ��� L/Cȯ���� �ڱ��ݾ��� ����ؾ� ��.(2003.09.19) - LEH
	frm1.vspdData.Col = C_XchRt
	If Trim(frm1.vspdData.text) = "" Then
		XchRt = UNICDbl(frm1.hdnXch.value)
	Else
		XchRt = UNICDbl(frm1.vspdData.text)
	End If
	VatRt = UNICDbl(frm1.hdnVatRt.value)
	
	frm1.vspdData.Col = C_IvAmt
	If frm1.vspdData.text <> "" Then
	    DocAmt = UNICDbl(frm1.vspdData.text)
	Else
	    DocAmt = 0
	End If

 	frm1.vspdData.Col = C_IOFlgCd
	VatIncFlag = frm1.vspdData.text
	
	'frm1.vspdData.Col = C_VatRate
	'VatRt = UNICDbl(frm1.vspdData.value)
     	
     	Select Case Col 
			Case C_IvAmt,C_IvAmt_Ref
      
				If chvatflg = "N" Then 
				    vat1 = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / (VatRt + 100),frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'CInt(DocAmt * VatRt / (VatRt + 100))  'vat ���� vat �ݾ� 
				    vat2 = UNIConvNumPCToCompanyByCurrency((DocAmt * VatRt) / 100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'DocAmt * VatRt / 100                  'vat ���� vat �ݾ� 
				Else
				    frm1.vspdData.Row = Row
				    frm1.vspdData.Col = C_VatDocAmt
				    vat1 = frm1.vspdData.text
				    vat2 = frm1.vspdData.text
				End If
				'***����(2003.03.21)-Lee Eun Hee***
				frm1.vspdData.Col = C_LCFlg
				If Trim(frm1.vspdData.text) = "A" Or Trim(frm1.vspdData.text) = "B" Then
					vat1 = "0"
					vat2 = "0"
				End If
				'**********************************
				
		        If VatIncFlag = "2" Then          'vat �����ΰ�� 

                    frm1.vspdData.Col = C_NetAmt
                    frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(DocAmt - UNICDbl(vat1),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X", "X")

			        frm1.vspdData.Col = C_VatDocAmt 'vat �ݾ� 
			        frm1.vspdData.text = vat1
			        
			        frm1.vspdData.Col = C_chkVatDocAmt '�������� flg seting �񱳰� 
				    frm1.vspdData.text = vat1
			        
			        If Trim(frm1.hdnDiv.value) = "*" Then
			            IvLocAmt = DocAmt * XchRt               '�����ڱ��ݾ� 
			           
			            NetLocAmt = (DocAmt- UNICDbl(vat1)) * XchRt      '�����ڱ� �ݾ� 
			           
			            frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
		                frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(vat1) * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'vatloc ���� 
	    	                     
			        ElseIf Trim(frm1.hdnDiv.value) = "/" Then
			            IvLocAmt = DocAmt / XchRt       '�����ڱ��ݾ� 
			           
			            NetLocAmt = (DocAmt - UNICDbl(vat1)) / XchRt      '�����ڱ� �ݾ� 
			          
			            frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
	     	            frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(vat1 / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")'vatloc ���� 
			        Else 
			            IvLocAmt = DocAmt
			           			       
			            NetLocAmt = DocAmt - UNICDbl(vat1)
			           
			            frm1.vspdData.Col = C_VatLocAmt
			            frm1.vspdData.text = vat1
			        End If
			       
			           frm1.vspdData.Col = C_IvLocAmt
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(IvLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
			    
			           frm1.vspdData.Col = C_NetLocAmt
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(NetLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
			           
 						'------------------------------------------------------------------------------------------------------
 						frm1.vspdData.Col = C_IOFlgCd
						VatIncFlag = frm1.vspdData.text	              
	 			           
			   Else                              'vat �����ΰ��			       
			       frm1.vspdData.Col = C_NetAmt   '���Աݾ� 
			       frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			       frm1.vspdData.Col = C_VatDocAmt 'vat �ݾ� 
			       frm1.vspdData.text = vat2
			       
			       frm1.vspdData.Col = C_chkVatDocAmt '�������� flg seting �񱳰� 
				   frm1.vspdData.text = vat2
			       
			       If Trim(frm1.hdnDiv.value) = "*" Then
			           IvLocAmt = DocAmt * XchRt       '�����ڱ��ݾ� 
			          
			           NetLocAmt = DocAmt * XchRt      '�����ڱ� �ݾ� 
			           frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(vat2) * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo , "X")'vatloc ���� 
			       
			       Elseif Trim(frm1.hdnDiv.value) = "/" Then
			           IvLocAmt = DocAmt / XchRt       '�����ڱ��ݾ� 
			           
			           NetLocAmt = DocAmt / XchRt      '�����ڱ� �ݾ� 
			           
			           frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(vat2) / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo , "X")'vatloc ���� 
			       Else 
			           IvLocAmt = DocAmt
			           
			           NetLocAmt = DocAmt
			           
			           frm1.vspdData.Col = C_VatLocAmt
			           frm1.vspdData.text = vat2
			       End If
			       
			           frm1.vspdData.Col = C_IvLocAmt
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(IvLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
			    
			           frm1.vspdData.Col = C_NetLocAmt
			           frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(NetLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
			   End If

			Case C_NetAmt                                      '���Աݾ� 
				frm1.vspdData.Col = C_NetAmt
				If frm1.vspdData.Text <> "" Then
					DocAmt = UNICDbl(frm1.vspdData.Text)
				Else				
					DocAmt = 0
				End IF
                
                If chvatflg = "N" Then
                    vat1 = UNIConvNumPCToCompanyByCurrency((DocAmt * UNICDbl(VatRt)) / (VatRt + 100),frm1.txtCur.value,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo , "X")'CInt(DocAmt * VatRt / (VatRt + 100))  'vat ���� vat �ݾ� 
	                vat2 = UNIConvNumPCToCompanyByCurrency((DocAmt * UNICDbl(VatRt)) / 100,frm1.txtCur.value,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo , "X")'(DocAmt * VatRt) / 100                  'vat ���� vat �ݾ� 
                Else
		            frm1.vspdData.Row = Row
		            frm1.vspdData.Col = C_VatDocAmt
		            vat1 = frm1.vspdData.text
		            vat2 = frm1.vspdData.text
		        End If                      
                
                If VatIncFlag = "2" Then          'vat �����ΰ�� 
                    VatDocAmt = UNICDbl(vat1)             
                    IvAmt = DocAmt + VatDocAmt    'vat�����ΰ�� ���رݾ� 
                Else
                    VatDocAmt = UNICDbl(vat2)              
                    IvAmt = DocAmt                'vat�����ΰ�� ���رݾ� 
                End If
				                			    
				If Trim(frm1.hdnDiv.value) = "*" Then
					IvLocAmt  = UNIConvNumPCToCompanyByCurrency(IvAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
					NetLocAmt = UNIConvNumPCToCompanyByCurrency(DocAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
                    VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo ,"X")'vatloc ���� 
				ElseIf Trim(frm1.hdnDiv.value) = "/" Then
					IvLocAmt  = UNIConvNumPCToCompanyByCurrency(IvAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
					NetLocAmt = UNIConvNumPCToCompanyByCurrency(DocAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
					VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo ,"X")'vatloc ���� 
				Else
					IvLocAmt  = UNIConvNumPCToCompanyByCurrency(IvAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo ,"X")
					NetLocAmt = UNIConvNumPCToCompanyByCurrency(DocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo ,"X")
					VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")
				End If
				
				frm1.vspdData.Col = C_IvAmt                    '���رݾ� 
		        frm1.vspdData.Text = IvAmt	

                frm1.vspdData.Col = C_VatDocAmt                'vat �ݾ� 
		        frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(VatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")	
	            
	            frm1.vspdData.Col = C_chkVatDocAmt '�������� flg seting �񱳰� 
				frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(VatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")	
	                
                frm1.vspdData.Col = C_IvLocAmt                 '�����ڱ��ݾ� 
		        frm1.vspdData.Text = IvLocAmt	
				
				frm1.vspdData.Col = C_NetLocAmt                '�����ڱ��ݾ� 
		        frm1.vspdData.Text = NetLocAmt	

				frm1.vspdData.Col = C_VatLocAmt                'VAT �ڱ��ݾ� 
				frm1.vspdData.Text = VatLocAmt

				     			
			Case C_NetLocAmt
				
				frm1.vspdData.Col = C_NetLocAmt
				If frm1.vspdData.Text <> "" Then
					IvLocAmt = UNICDbl(frm1.vspdData.Text)
				Else				
					IvLocAmt = 0
				End IF
  
				frm1.vspdData.Col = C_NetLocAmt
				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(IvLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo ,"X")
	
				'frm1.vspdData.Col = C_VatLocAmt
				'frm1.vspdData.Text = VatLocAmt
		     Case C_VatDocAmt

		          frm1.vspdData.Col = C_VatDocAmt
		          VatDocAmt = UNICDbl(frm1.vspdData.Text)  
	         	 
	         	  If Trim(frm1.hdnDiv.value) = "*" Then
	                  VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")'vatloc ���� 
				  ElseIf Trim(frm1.hdnDiv.value) = "/" Then
	                  VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")'vatloc ���� 
				  Else
	                  VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
				  End If
	              
	              frm1.vspdData.Col = C_VatLocAmt
	              frm1.vspdData.Text = VatLocAmt	    
	              
 				  frm1.vspdData.Col = C_IOFlgCd
				  VatIncFlag = frm1.vspdData.text	              

				  If VatIncFlag = "2" Then          'vat �����ΰ�� 

					  frm1.vspdData.Col = C_NetAmt
					  frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency((DocAmt-VatDocAmt),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X","X")
				  	 
						If Trim(frm1.hdnDiv.value) = "*" Then
							NetLocAmt = UNIConvNumPCToCompanyByCurrency((DocAmt- VatDocAmt) * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
						ElseIf Trim(frm1.hdnDiv.value) = "/" Then
							NetLocAmt = UNIConvNumPCToCompanyByCurrency((DocAmt- VatDocAmt) / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
						Else
							NetLocAmt = UNIConvNumPCToCompanyByCurrency((DocAmt- VatDocAmt),parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo ,"X")
						End If
						
						frm1.vspdData.Col = C_NetLocAmt                '�����ڱ��ݾ� 
						frm1.vspdData.Text = NetLocAmt	

			     End If
			                     
		End Select 

End Function	
'=============================================================================================
'  Function Name  : ChangeLocAmt()
'  Function Desc. : Iv Local Amt.�� ����� Vat�ڱ��ݾװ� �����ڱ����ݾ��� �ٽ� �����.
'  History        : 2003.03.26-Lee Eun Hee
'============================================================================================
Function ChangeLocAmt(Col, Row)

	Dim IvLocAmt, VatLocAmt, VatDocAmt, tmpNetLocAmt, vat1,vat2, VatIncFlag, VatRt, XchRt

	frm1.vspdData.Row = Row
	'Local L/C ���μ����� ��ģ ��� L/Cȯ���� �ڱ��ݾ��� ����ؾ� ��.(2003.09.19) - LEH
	frm1.vspdData.Col = C_XchRt
	If Trim(frm1.vspdData.text) = "" Then
		XchRt = UNICDbl(frm1.hdnXch.value)
	Else
		XchRt = UNICDbl(frm1.vspdData.text)
	End If
	
	VatRt = UNICDbl(frm1.hdnVatRt.value)
	
	frm1.vspdData.Col = C_IvLocAmt
	If frm1.vspdData.text <> "" Then
	    IvLocAmt = UNICDbl(frm1.vspdData.text)
	Else
	    IvLocAmt = 0
	End If
	
	frm1.vspdData.Col = C_VatLocAmt
	If frm1.vspdData.text <> "" Then
	    VatLocAmt = UNICDbl(frm1.vspdData.text)
	Else
	    VatLocAmt = 0
	End If
	
	frm1.vspdData.Col = C_VatDocAmt
	If frm1.vspdData.text <> "" Then
	    VatDocAmt = UNICDbl(frm1.vspdData.text)
	Else
	    VatDocAmt = 0
	End If
		
 	frm1.vspdData.Col = C_IOFlgCd
	VatIncFlag = frm1.vspdData.text
	
	Select Case Col 
		Case C_IvLocAmt
			

			If Trim(frm1.hdnDiv.value) = "*" Then
	            VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")'vatloc ���� 
			ElseIf Trim(frm1.hdnDiv.value) = "/" Then
	            VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")'vatloc ���� 
			Else
	            VatLocAmt = UNIConvNumPCToCompanyByCurrency(VatDocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
			End If
						
			If VatIncFlag = "2" Then          'vat �����ΰ�� 

			    frm1.vspdData.Col = C_NetLocAmt
			    frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(IvLocAmt - CDbl(VatLocAmt),parent.gCurrency,parent.ggAmtOfMoneyNo,"X", "X")

			    frm1.vspdData.Col = C_VatLocAmt 'vat �ݾ� 
			    frm1.vspdData.text = VatLocAmt

			Else
				frm1.vspdData.Col = C_NetLocAmt
			    frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(IvLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,"X", "X")

			    frm1.vspdData.Col = C_VatLocAmt 'vat �ݾ� 
			    frm1.vspdData.text = VatLocAmt
			End If
		
		Case C_VatLocAmt
		
			If VatIncFlag = "2" Then          'vat �����ΰ�� 
  	
			    frm1.vspdData.Col = C_NetLocAmt
			    frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(IvLocAmt - VatLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo,"X","X")
			End If	
		
	End Select

End Function
'����(2003-01-15)-������	
'==========================================   HSumAmtNewCalc()  ===============================
'	Name : HSumAmtNewCalc()
'	Description : detail �ݾ��� ���Ҷ� ��ȸ�� �Ѿ׺��� Event �ռ� 
'==============================================================================================
Function HSumAmtNewCalc()

	Dim iIndex
	Dim SumIvAmt, SumNetAmt,SumVatDocAmt
	Dim IvAmt, NetAmt, VatDocAmt
	
	SumIvAmt = lgTotalIvAmt
	SumNetAmt = lgTotalNetAmt
	SumVatDocAmt = lgTotalDocAmt	
			
	With frm1.vspdData
	
		If .Maxrows >= 1 then 
			For iIndex = 1 to .Maxrows
				.Row = iIndex
				.Col = 0
				If Trim(.text) <> ggoSpread.DeleteFlag then 			
				
					'VAT�ݾ� 
					.Col = C_VatDocAmt
					VatDocAmt	=	 unicdbl(.text)						
					SumVatDocAmt = SumVatDocAmt + VatDocAmt
					
					'�Ѹ��Աݾ� 
					.Col = C_IvAmt
					IvAmt	=	 unicdbl(.text)
					'�ΰ��������� ���Աݾ�+VAT�ݾ��� �Ѹ��Աݾ�					
					.Col = C_IOFlgCd					
					If Trim(.text) = "1" then 
						IvAmt = IvAmt + VatDocAmt
					End if
											
					SumIvAmt = SumIvAmt + IvAmt
					
					'���Աݾ� 
					.Col = C_NetAmt
					NetAmt	=	 unicdbl(.text)						
					SumNetAmt = SumNetAmt + NetAmt
									
				End if
			Next
		Else
			SumIvAmt = 0
			SumNetAmt = 0
			SumVatDocAmt = 0
		End if
			
	End with				
			
	frm1.txtivAmt.Text = UNIConvNumPCToCompanyByCurrency(SumIvAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
	frm1.txtnetAmt.Text = UNIConvNumPCToCompanyByCurrency(SumNetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
	frm1.txtvatAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatDocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")

End Function

'==========================================   changeNetAmt()  ======================================
'	Name : changeNetAmt()
'	Description : ���� ����� net �ݾ� ���� 
'========================================================================================================= 
Sub changeNetAmt(ByVal Col,ByVal Row)
    Dim GmQty,OldIvQty1,IvQty1,MvmtIvQty
    Dim OrderQty,IvQty,TotalIvQty,PoIvQty
    Dim NetAmt,PoAmt,MvmtAmt,IvDocAmt,DocAmt,Cost
    Dim reference,vat_flg,prcflg,vat_rvs_flg,PoVatIncFlg,ref_Vatrate_flg
    Dim vatAmt,vatrate,PoVatAmt,IvVatAmt,XchRt,MvmtVatAmt
    
    
    With frm1
    .vspdData.Row = Row
    .vspdData.Col = C_VatRate
    vatrate = UNICDbl(.vspdData.Text)   
	
	'����vat���� ����vat�� ��(2003.09.09)
	.vspdData.Col = C_ref_vatrate_flg
	ref_Vatrate_flg = Trim(.vspdData.text)
	
    .vspdData.Col = C_GmQty
    GmQty = UNICDbl(.vspdData.Text)     '�԰���� 
				 
    .vspdData.Col = C_oldIvQty1         '��ȸ�� ���Լ��� 
	OldIvQty1 = UNICDbl(.vspdData.Text)
				 
	.vspdData.Col = C_IvQty1            '���Լ��� 
	IvQty = UNICDbl(.vspdData.Text)
	IvQty1 = IvQty - OldIvQty1          ' new - old
	
				 
	.vspdData.Col = C_OrderQty          '���ּ��� 
    OrderQty = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = C_MvmtIvQty         '����Լ���	
	MvmtIvQty = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = C_PoIvQty
	PoIvQty = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = 0					'�Է�/���� �÷��� 
	
	If col = C_Cost_Ref And Trim(.vspdData.text) = ggoSpread.InsertFlag Then
	    TotalIvQty = MvmtIvQty + IvQty      '�Ѹ��Լ��� = ����Լ��� + ���Լ��� 
    Else	'��ȸ�� ������.	
        TotalIvQty = MvmtIvQty + IvQty1
    End If
    
    .vspdData.Col = C_PoAmt             '���ֱݾ� 
    PoAmt = UNICDbl(.vspdData.Text)
    
    .vspdData.Col = C_PoVatAmt          '���� vat�ݾ� 
    PoVatAmt = UNICDbl(.vspdData.Text)

    .vspdData.Col = C_MvmtAmt           '�԰�ݾ� 
	MvmtAmt = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = C_TotIvDocAmt          '������ѱݾ� 
	IvDocAmt = UNICDbl(.vspdData.Text)
	
    .vspdData.Col = C_TotIvVatAmt          '����� vat�ݾ� 
	IvVatAmt = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = C_IOFlgCd           'vat ���Կ��� 
	vat_flg = Trim(.vspdData.Text)
	
	.vspdData.Col = C_upt_amt_flg       '���ֱݾ� ���ſ��� 
	reference = Trim(.vspdData.Text)
	
    .vspdData.Col = C_prcflg            '�ܰ�flg
	prcflg = Trim(.vspdData.Text)
	'����(2003.05.12)
	.vspdData.Col = C_Cost
	Cost = UNICDbl(.vspdData.Text)
	
	.vspdData.Col = C_vat_rvs_flg
	vat_rvs_flg = Trim(.vspdData.Text)
	'@����@(2003.02.17)
	.vspdData.Col = C_PoVatIncFlg
	PoVatIncFlg = Trim(.vspdData.Text)

	'**����(2003.03.24)+++++++++++++++++++++++++++++++++++++
	.vspdData.Col = C_LcFlg
    If UCase(Trim(.vspdData.Text)) = "B" or UCase(Trim(.vspdData.Text)) = "A" Then
		'������!!!
		NetAmt = Cost * IvQty	'LC�ܰ� * ���Լ��� 
		.vspdData.Col = C_NetAmt
        .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
		NetAmt = UNICDbl(.vspdData.Text)
		vatAmt = 0
		'DocAmt ��� 
		If vat_flg = "1" Then
		    DocAmt = NetAmt '-- byun jee hyun 
		Else
		    DocAmt = NetAmt + UNICDbl(UNIConvNumPCToCompanyByCurrency(vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X"))
		End If    
                 
		.vspdData.Col = C_IvAmt
		.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
		.vspdData.Col = C_VatDocAmt
	    .vspdData.Text = UNIConvNumPCToCompanyByCurrency(vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")

		Exit Sub
    End If 

	'###Po_Amt���#####
	If PoVatIncFlg <> vat_flg Or (PoVatAmt = 0 And vatrate <> 0 ) Or ref_Vatrate_flg = "N" Then
		If (PoVatIncFlg = "1" or PoVatIncFlg = "") and vat_flg = "2" Then	'������ ���Ա����� "����"�� ��� 
			PoVatAmt = (PoAmt * vatrate) / (vatrate + 100)  'vat ���� vat �ݾ� 
			PoAmt = PoAmt - PoVatAmt
			MvmtVatAmt = (MvmtAmt * vatrate) / (vatrate + 100)  'vat ���� vat �ݾ� 
			MvmtAmt = MvmtAmt - MvmtVatAmt
		ElseIf PoVatIncFlg = "2" And vat_flg = "1" Then
			PoAmt = PoAmt + PoVatAmt
			MvmtAmt = PoAmt * (GmQty/OrderQty)
			PoVatAmt = (PoAmt * vatrate) / 100					'vat ���� vat �ݾ� 
		
		'���Ժΰ������� ������ �ΰ������� �ٸ���� ���Ժΰ������� ������.(2003.09.09)
		ElseIf PoVatIncFlg = "2" And vat_flg = "2" Then	
			PoAmt = PoAmt + PoVatAmt
			PoVatAmt = (PoAmt * vatrate) / (vatrate + 100)  'vat ���� vat �ݾ� 
			PoAmt = PoAmt - PoVatAmt
			MvmtAmt = PoAmt * (GmQty/OrderQty)
		Else	'(PoVatIncFlg = "" and vat_flg = "1") or (PoVatIncFlg = "1" And vat_flg = "1")
			PoVatAmt = (PoAmt * vatrate) / 100					'vat ���� vat �ݾ� 
		End If
	End If
	'#####################
	.vspdData.Col = C_MvmtRcptNo        '�԰��ȣ 
	
	If (MvmtIvQty + IvQty1) > GmQty And Trim(.vspdData.Text) <> "" Then	'������ 
        If OrderQty = TotalIvQty Then        ' �Ѹ��Լ��� = ���ּ��� 
            NetAmt = PoAmt - IvDocAmt        ' ���ֱݾ� - ������ѱݾ� 
            
            '**2003.03�� ��ġ(KJH) ************
            .vspdData.Col = C_NetAmt
            .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
            
            NetAmt = UNICDbl(.vspdData.Text)
'-------------------------------
'  update by JT.Kim  20040601 
			If vat_rvs_flg <> "N" And PoVatAmt <> 0 Then
				vatAmt = PoVatAmt - IvVatAmt     '����vat - �����vat
			ElseIf vat_flg = "1" Then
			    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			ElseIf vat_flg = "2" Then
			    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			End If
'------------------------------            
        Else
			'**2003.1�� ��ġ************
			if prcflg = "S" then	'ǥ�شܰ� 
				NetAmt = PoAmt * (IvQty/OrderQty)	'���ֱݾ� * (���Լ���/���ּ���)
			Else
				NetAmt = MvmtAmt * (IvQty/GmQty)	'�԰�ݾ� * (���Լ���/�԰����)
			End if

            '**2003.03�� ��ġ(KJH) ************
            .vspdData.Col = C_NetAmt
            .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
            NetAmt = UNICDbl(.vspdData.Text)
            
'-------------------------------
'  update by JT.Kim  20040601 
			If vat_flg = "1" Then
			    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			ElseIf vat_flg = "2" Then
			    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			End If
'-------------------------------
            
        End If

        'DocAmt ��� 
        If vat_flg = "1" Then
            DocAmt = NetAmt
        Else
            DocAmt = NetAmt + UNICDbl(UNIConvNumPCToCompanyByCurrency(vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X"))
        End If        
    
    Else
       If prcflg = "S" Then       
			.vspdData.Col = 0					'�Է�/���� �÷��� 
   			If col = C_Cost_Ref And Trim(.vspdData.text) = ggoSpread.InsertFlag Then
			    TotalIvQty = PoIvQty + IvQty      '�Ѹ��Լ��� = ����Լ��� + ���Լ��� 
			Else	'��ȸ�� ������.	
			    TotalIvQty = PoIvQty + IvQty1
			End If
			
           If TotalIvQty = OrderQty Then          '�Ѹ��Լ��� = ���ּ��� 
'-------------------------------
'  update by JT.Kim  20040708 : ������ 20040601 �������� ��������/ �������ϰ��� ���ֿ� ���Լ����� �ٸ���츸 ����ǵ��� �ϴ� ����� 
'								������ ���� ��쿡�� ����ǵ��� ����(���ܰ� ���� �� �԰� �� ���ܰ� Ȯ���Ͽ� ���Խ� �̵���մܰ����� �԰�ݾ��� �������� ������ 
'								���ִܰ� * ���� <> �԰�ݾ� ������ ���� ������ 
'                               �� case�� ǥ�شܰ����� �Ȱ��� �ݾ�= ���� * �ܰ� �� �������� ���� 
'               NetAmt = PoAmt - IvDocAmt        '���ֱݾ� - �� �����ѱݾ� 
               NetAmt = Cost * IvQty
'-------------------------------
				'**2003.03�� ��ġ(KJH) ************
				.vspdData.Col = C_NetAmt
				.vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				NetAmt = UNICDbl(.vspdData.Text)

'-------------------------------
'  update by JT.Kim  20040601 
				If vat_rvs_flg <> "N" And PoVatAmt <> 0 Then
				    vatAmt = PoVatAmt - IvVatAmt     '����vat - �����vat	            
				ElseIf vat_flg = "1" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				ElseIf vat_flg = "2" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				End If

'--------------------------------
	       Else
               NetAmt = PoAmt * (IvQty/OrderQty) '���ֱݾ� * (���Լ���/���ּ���)
	           .vspdData.Col = C_NetAmt
               .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
               NetAmt = UNICDbl(.vspdData.Text)
'-------------------------------
'  update by JT.Kim  20040601 
				If vat_flg = "1" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				ElseIf vat_flg = "2" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				End If
'--------------------------------
	       End If
       ElseIf prcflg = "M" Then
           
           If TotalIvQty = GmQty Then            '�Ѹ��Լ��� = �԰���� 
				If OldIvQty1 > GmQty Then 
					NetAmt = MvmtAmt - UNICDbl(UNIConvNumPCToCompanyByCurrency(IvDocAmt * GmQty / OldIvQty1,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X"))
					'�԰�ݾ� - �԰�п� ���� ���Աݾ� 
				Else                
'-------------------------------
'  update by JT.Kim  20040708 : ������ 20040601 �������� ��������/ �������ϰ��� ���ֿ� ���Լ����� �ٸ���츸 ����ǵ��� �ϴ� ����� 
'								������ ���� ��쿡�� ����ǵ��� ����(���ܰ� ���� �� �԰� �� ���ܰ� Ȯ���Ͽ� ���Խ� �̵���մܰ����� �԰�ݾ��� �������� ������ 
'								���ִܰ� * ���� <> �԰�ݾ� ������ ���� ������ 
'					NetAmt = MvmtAmt - IvDocAmt       '�԰�ݾ� - �� �����ѱݾ� 
					NetAmt = Cost * IvQty
'-------------------------------
				End If 		
				'**2003.03�� ��ġ(KJH) ************	
	           .vspdData.Col = C_NetAmt
               .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
               NetAmt = UNICDbl(.vspdData.Text)
           
'-------------------------------
'  update by JT.Kim  20040601 - ���ݰ� ���ݾ��� �ܰ� * ���� ���� ���� ����� 
				If (PoIvQty + IvQty) = OrderQty And vat_rvs_flg <> "N" And PoVatAmt <> 0 Then
				    vatAmt = PoVatAmt - IvVatAmt  '����vat - �����vat
					If vat_flg = "1" Then
					    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
					ElseIf vat_flg = "2" Then
					    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
					End If
				ElseIf vat_flg = "1" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				ElseIf vat_flg = "2" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				End If
 
'				If vat_flg = "1" Then
'				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
'				    vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
'				ElseIf vat_flg = "2" Then
'				    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
'				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
'				End If


'-------------------------------
	       Else
               NetAmt = MvmtAmt * (IvQty/GmQty)  '�԰�ݾ� * (���Լ���/�԰����)
	           '**2003.03�� ��ġ(KJH) ************
	           .vspdData.Col = C_NetAmt
               .vspdData.Text = UNIConvNumPCToCompanyByCurrency(NetAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
     
               NetAmt = UNICDbl(.vspdData.Text)
' update by JT.Kim 20040615 
					If vat_flg = "1" Then
					    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
					ElseIf vat_flg = "2" Then
					    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
					End If

'  update by JT.Kim  20040601 
				If vat_flg = "1" Then
					vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				ElseIf vat_flg = "2" Then
				    vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				    NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
				End If

'-------------------------------
	       End If
       '**����(2003.03.20)
       Else
		.vspdData.Col = C_NetAmt
		NetAmt = UNICDbl(.vspdData.Text)
'-------------------------------
'  update by JT.Kim  20040601 
		If vat_flg = "1" Then
			vatAmt = UNIConvNumPCToCompanyByCurrency((NetAmt * vatrate)/100,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
		ElseIf vat_flg = "2" Then
			vatAmt = UNIConvNumPCToCompanyByCurrency((Cost * IvQty * vatrate)/(100 + vatrate),frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
			NetAmt = UNIConvNumPCToCompanyByCurrency(Cost * IvQty - vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
		End If
'-------------------------------
       
       End If
      'DocAmt ��� 
       If vat_flg = "1" Then
           DocAmt = NetAmt
       Else
           DocAmt = NetAmt + UNICDbl(UNIConvNumPCToCompanyByCurrency(vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X"))
       End If    
                 
    End If
    
    .vspdData.Col = C_IvAmt
    .vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X" , "X")
    .vspdData.Col = C_VatDocAmt
    .vspdData.Text = UNIConvNumPCToCompanyByCurrency(vatAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")

    'XchRt = UNICDbl(frm1.hdnXch.value)
    'Local L/C ���μ����� ��ģ ��� L/Cȯ���� �ڱ��ݾ��� ����ؾ� ��.(2003.09.19) - LEH
	frm1.vspdData.Col = C_XchRt
	If Trim(frm1.vspdData.text) = "" Then
		XchRt = UNICDbl(frm1.hdnXch.value)
	Else
		XchRt = UNICDbl(frm1.vspdData.text)
	End If
	
    If Trim(frm1.hdnDiv.value) = "*" Then
	    frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
        frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(vatAmt * XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")'vatloc ���� 
	ElseIf Trim(frm1.hdnDiv.value) = "/" Then
        frm1.vspdData.Col = C_VatLocAmt 'vat �ڱ��ݾ� 
	    frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(vatAmt / XchRt,parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo ,"X")'vatloc ���� 
	Else 
	    frm1.vspdData.Col = C_VatLocAmt
	    frm1.vspdData.text = vatAmt
	End If

    End With
End Sub
'==========================================   Posting()  ======================================
'	Name : Posting()
'	Description : Ȯ����ư,Ȯ����ҹ�ư�� Event �ռ� 
'========================================================================================================= 
Sub Posting()
    Dim IntRetCD 
    
    Err.Clear                                                         '��: Protect system from crashing
    
    If ggoSpread.SSCheckChange = True Then
		Call DisplayMsgBox("189217","X","X","X")
		Exit sub
	End if
	
	If Trim(frm1.hdnPostDt.value) = "" Then
		Call DisplayMsgBox("17A002","X" , "������","X")
		Exit Sub
	End If
	
    If frm1.hdnPostingFlg.Value = "Y" Then
    
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		Call DbSave("Posting")

	ElseIf frm1.hdnPostingFlg.Value = "N" then
		
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			frm1.btnPosting.disabled = False	'20040315  
			Exit Sub
		Else 
				frm1.btnPosting.disabled = True		'20040315 
		End If
		
		Call DbSave("Posting")
		
	End If
	
End Sub
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1   
		'�ѱݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtIvAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
        '�Ѹ��Աݾ� 
        ggoOper.FormatFieldByObjectOfCur .txtnetAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
        '��VAT�ݾ� 
        ggoOper.FormatFieldByObjectOfCur .txtvatAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
        '���Աݾ� +vat �ݾ� 
        'ggoOper.FormatFieldByObjectOfCur .txtsumnetAmt,	.txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000,parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'���Դܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_Cost,-1, .txtCur.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		'�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_IvAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		'���Աݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_ChgNetAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		'VAT�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_VatDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_OrgVatDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_ChgVatDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_chkVatDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		'po�ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_OrderCost,-1, .txtCur.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		'�����߰� 
		ggoSpread.SSSetFloatByCellOfCur C_PoAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_MvmtAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_TotIvDocAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur C_PoVatAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
		ggoSpread.SSSetFloatByCellOfCur C_TotIvVatAmt,-1, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		

	End With

End Sub

'==========================================  3.1.1 Form_Load()  ======================================


'========================================  Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel , UnloadMode )
  
End Sub

'======================================  vspdData_Change()  ===================================
Sub vspdData_Change(ByVal Col , ByVal Row )

Dim Qty, Price, DocAmt,VatIncFlag,VatDocAmt,changeVatflg,vat_rvs_flg,retflg,ref_flg
Dim ref_Vatrate_flg '(����vat���� ����vat�� ��)(2003.09.09)
Dim OrderQty1,TotalIvQty1,PoVatAmt1,IvVatAmt1,IvQty2,MvmtIvQty1
Dim MvmtRcptNo, LcNo
    changeVatflg = "N"
    ggoSpread.Source = frm1.vspdData
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = 0
	
	If Frm1.vspdData.text = ggoSpread.DeleteFlag Then Exit Sub

    ggoSpread.UpdateRow Row

    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = C_vat_rvs_flg
	vat_rvs_flg = Trim(Frm1.vspdData.Text)
	
	Frm1.vspdData.Col = C_retflg
	retflg = Trim(Frm1.vspdData.Text)
	
    Frm1.vspdData.Col = C_ref_flg    '�������� ���������� Y �Ѱ��Լ��� �ѹ��� ȣ����ϱ����� 
    ref_flg = Trim(Frm1.vspdData.text) 
	
	'����vat���� ����vat�� ��(2003.09.09)
	Frm1.vspdData.Col = C_ref_vatrate_flg
	ref_Vatrate_flg = Trim(frm1.vspdData.text)
	
	Frm1.vspdData.Col = Col
  
	Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

	Select Case col
	
	Case C_ivQty1,C_Cost,C_Cost_Ref       '���Լ���,���Դܰ�,��������,�԰������ΰ��(C_Cost)= ���Աݾ� 

		frm1.vspdData.Col = C_ivQty1
		If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
			Qty = 0
		Else
			Qty = UNICDbl(frm1.vspdData.Text)
		End If
		
		
		frm1.vspdData.Col = C_Cost
		If Trim(frm1.vspdData.Text) = "" Or IsNull(frm1.vspdData.Text) Then
			Price = 0
		Else
			Price = UNICDbl(frm1.vspdData.Text)
		End If
		'***��LC���� ����(2003.03.19)-Lee, Eun Hee
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_MvmtRcptNo
		MvmtRcptNo = frm1.vspdData.Text
		frm1.vspdData.Col = C_LCNo
		LcNo = frm1.vspdData.Text
		
		If Trim(MvmtRcptNo) <> "" And Trim(LcNo) = "" And col = C_ivQty1 And retflg <> "Y"  Then  '�԰������̰� ���������ΰ�� 
			Call changeNetAmt(Col,Row)
		    changeVatflg = "Y"
		ElseIf Trim(LcNo) <> ""  And col = C_ivQty1 And retflg <> "Y"  Then	'��LC�������̰� ���� ������ ��� 
			DocAmt = Qty * Price          '(���Լ���) * (�ܰ�)
		    frm1.vspdData.Col = C_IvAmt   '���Աݾ� 
		    frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X","X") 
			changeVatflg = "Y"
		ElseIf Trim(MvmtRcptNo) <> "" And col = C_Cost_Ref Then
		    changeVatflg = "Y"
		Else
			'����(12/18)
		    If Col = C_ivQty1 And retflg <> "Y" And frm1.hdnExceptflg.value <> "Y" Then '���ܸ����� ��� 0���� ������ ���� ���� ���� 
				frm1.vspdData.Row = Row
                frm1.vspdData.Col = C_prcflg
                frm1.vspdData.text = "S"
                Call changeNetAmt(Col,Row)
                changeVatflg = "N"			'@@����@@
            ElseIf col = C_Cost_Ref Then
				changeVatflg = "Y"
		    Else
		        DocAmt = Qty * Price          '(���Լ���) * (�ܰ�)
		        frm1.vspdData.Col = C_IvAmt   '���Աݾ� 
		        frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo,"X","X") 
                'vat ���� 
                If Col <> C_Cost And retflg <> "Y" And frm1.hdnExceptflg.value <> "Y" Then '����(12/18)
                    Frm1.vspdData.Row = Row
                    
                    Frm1.vspdData.Col = C_IvQty1            '���Լ��� 
	                IvQty2 = UNICDbl(frm1.vspdData.Text)
				 
	                Frm1.vspdData.Col = C_OrderQty          '���ּ��� 
                    OrderQty1 = UNICDbl(frm1.vspdData.Text)
	
	                Frm1.vspdData.Col = C_MvmtIvQty         '����Լ��� 
	                MvmtIvQty1 = UNICDbl(frm1.vspdData.Text)
	
	                TotalIvQty1 = MvmtIvQty1 + IvQty2         '�Ѹ��Լ��� = ����Լ��� + ���Լ��� 
   
                    frm1.vspdData.Col = C_prcflg
                    
                    
                    If OrderQty1 = TotalIvQty1 And vat_rvs_flg <> "N" And ref_Vatrate_flg = "Y" Then        ' �Ѹ��Լ��� = ���ּ��� 
                        
                        frm1.vspdData.Col = C_PoVatAmt
                        PoVatAmt1 = UNICDbl(frm1.vspdData.Text)
                        
                        frm1.vspdData.Col = C_TotIvVatAmt
                        IvVatAmt1 = UNICDbl(frm1.vspdData.Text)
                        
                        frm1.vspdData.Col = C_VatDocAmt
                        frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(PoVatAmt1 - IvVatAmt1,frm1.txtCur.value,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo , "X")
                        changeVatflg = "Y"
                    Else	'vat�ݾ� ���(2003.09.09)
                    
                    End If
                End If
            End If
        End If
        'frm1.vspdData.Text = UNIFormatNumberByCurrecny(DocAmt,frm1.txtCur.value,parent.ggAmtOfMoneyNo)
		
		
		Dim tmpVatIncFlag
 		frm1.vspdData.Col = C_IOFlgCd
		tmpVatIncFlag = frm1.vspdData.text	              
        
        'If ref_flg = "Y" And retflg = "Y" Then Exit Sub '����(12/18)
        
		If col = C_Cost_Ref And tmpVatIncFlag = "2" Then
			Call ChangeAmt(C_IvAmt_Ref,Row,changeVatflg)		
		Else
			Call ChangeAmt(C_IvAmt,Row,changeVatflg)
		End If
		'����(2003-01-15)-������ 
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
		If col <> C_Cost_Ref Then
           Call HSumAmtNewCalc()
        End If

	    
	Case C_IvAmt                     
		Call ChangeAmt(C_IvAmt,Row,changeVatflg)
	    '����(2003-01-15)-������ 
        Call HSumAmtNewCalc()
	Case C_NetAmt                      'VAT�ݾ�,�ڱ��ݾ�,VAT�ڱ��ݾ׺��� 
		Call ChangeAmt(C_NetAmt,Row,changeVatflg)
	    '����(2003-01-15)-������ 
        Call HSumAmtNewCalc()
	Case C_NetLocAmt                   '�ڱ��ݾ�,VAT�ڱ��ݾ׺��� 
		Call ChangeAmt(C_NetLocAmt,Row,changeVatflg)
        '����(2003-01-15)-������ 
        Call HSumAmtNewCalc()
    '**2003.3�� ��ġ****
    Case C_IvLocAmt                   '�ڱ��ݾ׺��� 
		Call ChangeLocAmt(C_IvLocAmt,Row)
        Call HSumAmtNewCalc()    
    Case C_VatLocAmt                   '�ڱ��ݾ׺��� 
		Call ChangeLocAmt(C_VatLocAmt,Row)
        Call HSumAmtNewCalc()  
	'********************
	Case C_VatType                    'vat Ÿ���� �ٲ�� vat�� ��ȯ 
	    Call SetVatType()
        Call ChangeAmt(C_NetAmt,Row,changeVatflg)
	    '����(2003-01-15)-������ 
        Call HSumAmtNewCalc()
	Case C_VatDocAmt
	    Call ChangeAmt(C_VatDocAmt,Row,changeVatflg)         	
	    '����(2003-01-15)-������ 
        Call HSumAmtNewCalc() 
	'***2003.1�� ��ġ*********
	Case C_PlantCd
		Call SetUnitCost( Row )
	Case C_ItemCd
		if Trim(frm1.vspdData.text) = "" then
			frm1.vspdData.Col = C_ItemNm
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_SpplSpec
			frm1.vspdData.text = ""
		end if
		Call SetUnitCost( Row )    
	'**************************
	End Select
		
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex,changeVatflg
    changeVatflg = "N"
	
	If Frm1.vspdData.text = ggoSpread.DeleteFlag Then Exit Sub		
	
	With frm1.vspdData
		.Row = Row
    	.Col = Col
		intIndex = .Value
		.Col = C_IOFlg+1
		.Value = intIndex+1
    End With
    Call ChangeAmt(C_IvAmt,Row,changeVatflg)
    '����(2003-01-15)-������ 
    Call HSumAmtNewCalc()
End Sub


Sub vspdData_DblClick(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	Call JumpPgm()

End Sub

Function JumpPgm()
	
	Dim pvSelmvid, pvFB_fg,pvKeyVal,StrNVar,StrNPgm,pvSingle
	
	if frm1.vspddata.Maxrows  < 1 then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	ggoSpread.Source = frm1.vspdData
	
	frm1.vspddata.row = 0
    frm1.vspddata.col = frm1.vspddata.Activecol

    Select case frm1.vspddata.value
    
   									
	Case "[�԰��ȣ]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow
		frm1.vspddata.COL =C_MvmtNo
		if 	TRIM(frm1.vspddata.value) <> "" then
		
				pvKeyVal =   frm1.vspddata.value
				pvSingle  =	""
				pvFB_fg = "B"
				pvSelmvid = "RCPT_NO"
	
					Call Jump_Pgm (	pvSelmvid, _
									pvFB_fg, _
									pvSingle,  _
									pvKeyVal)										
	
			
		End if
	
	
	End select
End Function


'================================  vspdData_ButtonClicked()  ============================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp
Dim intPos1
   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 Then
        
        .Col = Col - 1
        .Row = Row
        
        Select Case Col
			Case C_PlantPopup
        		Call OpenPlant()
        		'***2003.1�� ��ġ***
        		Call SetUnitCost( Row )
        	Case C_ItemPopup
        		Call OpenItem()
        		Call SetUnitCost( Row )
        	Case C_UnitPopup
        		Call OpenUnit()
        	Case C_VatPopup             '�߰� 
			    Call OpenVat()
			'2007-04-16�߰�    
			Case C_TrackingPopup
				Call OpenTrackingNo()    			    
       End Select
       
    Else
    	Exit sub
    End If
    
    End With
End Sub

'================================  vspdData_TopLeftChange()  ============================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgStrPrevKey <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================  FncQuery()  ============================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
	
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables
    															'��: Initializes local global variables
    If Not chkFieldByCell(frm1.txtIvNo,"A",1) Then									'��: This function check indispensable field
       Exit Function
    End If
    
    
    frm1.txtQuerytype.value = "Query"
    
    If DbQuery = False Then Exit Function
       
    FncQuery = True																'��: Processing is OK
    
End Function

'================================  FncNew()  ============================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
        
    Err.Clear                                                               '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")
    'Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call LockHTMLField(frm1.txtIvNo, "R")
    Call LockHTMLField(frm1.txtIvTypeCd, "P")
    Call LockHTMLField(frm1.txtIvTypeNm, "P")
    Call LockObjectField(frm1.txtIvDt, "P")
    Call LockHTMLField(frm1.ChkPrepay, "P")
    Call LockHTMLField(frm1.txtSpplCd, "P")
    Call LockHTMLField(frm1.txtSpplNm, "P")
    Call LockHTMLField(frm1.txtGrpCd, "P")
    Call LockHTMLField(frm1.txtGrpNm, "P")
    Call LockObjectField(frm1.txtivAmt, "P")
    Call LockObjectField(frm1.txtXchRt, "P")
    Call LockObjectField(frm1.txtnetAmt, "P")
    Call LockObjectField(frm1.txtvatAmt, "P")
    
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    
    FncNew = True                                                           '��: Processing is OK

End Function

'================================  FncDelete()  ============================================
Function FncDelete() 
   
Dim IntRetCD

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function

    
    FncDelete = False                                                       '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If
    
    Call ggoOper.ClearField(Document, "A")
    
    FncDelete = True                                                        '��: Processing is OK
    '����(2003-01-15)-������ 
    Call HSumAmtNewCalc()

End Function

'================================  FncSave()  ============================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear   

	if CheckRunningBizProcess = true then
		exit function
	end if                                                            '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If
    '**����(2003.10.16)**
    'If Not chkField(Document, "2") Then                                  '��: Check contents area
    '   Exit Function
    'End If

	If Not ggoSpread.SSDefaultCheck         Then            
	   Exit Function
	End If

    If DbSave("toolbar") = False Then Exit Function                         '��: Save db data
    
    If frm1.txthdnIvNo.value <> frm1.txtIvNo.value then
		frm1.txtIvNo.value =	frm1.txthdnIvNo.value		
	End If
    
    FncSave = True                                                          '��: Processing is OK
    
End Function

'================================  FncCopy()  ============================================
Function FncCopy() 

	On Error Resume Next 

	if frm1.vspdData.Maxrows < 1	then exit function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
       
    frm1.vspdData.ReDraw = False    
    'SetSpreadColor frm1.vspdData.ActiveRow
    
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    '����(2003-01-15)-������ 
    frm1.vspdData.ReDraw = True
    
    Call HSumAmtNewCalc()
    
    
End Function

'================================  FncCancel()  ============================================
Function FncCancel() 
	dim startindex
	dim endindex

	startindex = frm1.vspdData.SelBlockRow
	endindex = frm1.vspdData.SelBlockRow2
	
	ReDim NetAmt(endindex - startindex)
	ReDim VatDocAmt(endindex - startindex)
    ReDim OrgNetAmt(endindex - startindex)
    Redim OrgVatDocAmt(endindex - startindex)
    ReDim ChgNetAmt(endindex - startindex)
    Redim ChgVatDocAmt(endindex - startindex)    
    Redim delFlag(endindex - startindex)    
	Dim current_index
    Dim maxRow,maxRow1
    Dim i
    
	If frm1.vspdData.Maxrows < 1	Then Exit Function

	ggoSpread.Source = frm1.vspdData

	For i = startindex To endindex
		frm1.vspdData.Row = i '//frm1.vspdData.ActiveRow

		current_index = i - startindex

		'���Լ��ݾ� 
		frm1.vspdData.Col = C_NetAmt
		NetAmt(current_index) = UNICDbl(frm1.vspdData.Text)	
	
		'���Լ��ݾ�(HIDDEN)		
		frm1.vspdData.Col = C_OrgNetAmt	
		OrgNetAmt(current_index) = UNICDbl(frm1.vspdData.Text)
	
		'VAT �ݾ�	
		frm1.vspdData.Col = C_VatDocAmt
		VatDocAmt(current_index) = UNICDbl(frm1.vspdData.Text)	 	
	
		'VAT �ݾ�(HIDDEN)		
		frm1.vspdData.Col = C_OrgVatDocAmt	
		OrgVatDocAmt(current_index) = UNICDbl(frm1.vspdData.Text)	 	
		

		'VAT �ݾ�	
		frm1.vspdData.Col = C_ChgNetAmt
		ChgNetAmt(current_index) = UNICDbl(frm1.vspdData.Text)	 	
	
		'VAT �ݾ�(HIDDEN)		
		frm1.vspdData.Col = C_ChgVatDocAmt	
		ChgVatDocAmt(current_index) = UNICDbl(frm1.vspdData.Text)	 	
		
		frm1.vspdData.Col = 0
		delFlag(current_index) = frm1.vspdData.Text

	Next
	

			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = 0
			current_index = frm1.vspdData.ActiveRow - startindex
			
			ggoSpread.Source = frm1.vspdData
		'����(2003-01-15)-������ 

		 frm1.vspdData.Col = C_Stateflg
		 frm1.vspdData.text = ""				
		 ggoSpread.EditUndo 
	
	'����(2003-01-15)-������	 
	Call HSumAmtNewCalc()		          			 

End Function
'================================  FncInsertRow()  ============================================
Function FncInsertRow(ByVal pvRowCnt) 
 	Dim IntRetCD
    Dim imRow, index
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
		
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
	    ggoSpread.Source = .vspdData
	    
	    ggoSpread.InsertRow .vspdData.ActiveRow, imRow

	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    .vspdData.ReDraw = True
        '(2003.09.19)
        For index = .vspdData.ActiveRow To (.vspdData.ActiveRow + imRow - 1)
        .vspdData.Row = index
        .vspdData.Col = C_IOFlg          'header�� �ִ� ���� �־��ش� 
        
        If Trim(.hdvatFlg.value) = "2" Then
		    .vspdData.value = 1
		    .vspdData.Col = C_IOFlgCd
		    .vspdData.value = 2
		ElseIf Trim(.hdvatFlg.value) = "1" Then	'���� 
			.vspdData.value = 0
			.vspdData.Col = C_IOFlgCd
		    .vspdData.value = 1
		End If
		'���Գ���ȯ�� �߰� - 2003.09.19
		.vspdData.Col = C_XchRt
		.vspdData.Text= frm1.hdnXch.value
		
		'-- Issue for 8548 by ByunJeeHyun 2004-08-10
		'Call ggoSpread.SSSetColHidden(C_VatType, C_VatRate, False)
		
		.vspdData.Row = index
        .vspdData.Col = C_VatType          'header�� �ִ� ���� �־��ش� 
        .vspdData.value = .hdnVatType.value
        
        .vspdData.Row = index
        .vspdData.Col = C_VatRate          'header�� �ִ� ���� �־��ش� 
        .vspdData.value = .hdnVatRt.value
		Next
		
    End With
	
	Set gActiveElement = document.ActiveElement
	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
       
    
End Function


'================================  FncDeleteRow()  ============================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim NetAmt,VatDocAmt
    Dim index
    Dim idel
    ggoSpread.Source = frm1.vspdData
    
    If frm1.vspdData.Maxrows < 1	Then Exit Function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
	
	For index = .SelBlockRow To .SelBlockRow2
		.Row = index
		.Col = C_Stateflg
		idel = .text
		.Col = 0

		If Trim(.text) <> ggoSpread.InsertFlag And Trim(idel) <> "D" Then

			'���Լ��ݾ� 
			frm1.vspdData.Col = C_NetAmt
			NetAmt = UNICDbl(frm1.vspdData.Text)	
	
			'VAT �ݾ�	
			frm1.vspdData.Col = C_VatDocAmt
			VatDocAmt = UNICDbl(frm1.vspdData.Text)	
			
			.Col = C_Stateflg	
			frm1.vspdData.text = "D"		   
	   End If
	Next
	
	lDelRows = ggoSpread.DeleteRow	
	
    lgBlnFlgChgValue = True
    
    End With
    
    '����(2003-01-15)-������ 
    Call HSumAmtNewCalc()

End Function

'================================  FncPrint()  ============================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
End Function

'================================  FncPrev()  ============================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'================================  FncNext()  ============================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'================================  FncExcel()  ============================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(parent.C_MULTI)										'��: ȭ�� ���� 
End Function

'================================  FncFind()  ============================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncFind(parent.C_MULTI , False)                                 '��:ȭ�� ����, Tab ���� 
End Function
'================================  FncExit()  ============================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")    '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'================================  DbQuery()  =============================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	
	Dim strVal
    
    With frm1
    
	If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtIvNo=" & .txthdnIvNo.value
	Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtIvNo=" & Trim(.txtIvNo.value)
	End If
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtQuerytype=" & .txtQuerytype.value
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		'����(2003.06.10)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCur.value)
    .hdnmaxrow.value = .vspdData.MaxRows
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'============================  DbQueryOk()  ================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    'Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    Call LockHTMLField(frm1.txtIvNo, "R")
    Call LockHTMLField(frm1.txtIvTypeCd, "P")
    Call LockHTMLField(frm1.txtIvTypeNm, "P")
    Call LockObjectField(frm1.txtIvDt, "P")
    Call LockHTMLField(frm1.ChkPrepay, "P")
    Call LockHTMLField(frm1.txtSpplCd, "P")
    Call LockHTMLField(frm1.txtSpplNm, "P")
    Call LockHTMLField(frm1.txtGrpCd, "P")
    Call LockHTMLField(frm1.txtGrpNm, "P")
    Call LockObjectField(frm1.txtivAmt, "P")
    Call LockObjectField(frm1.txtXchRt, "P")
    Call LockObjectField(frm1.txtnetAmt, "P")
    Call LockObjectField(frm1.txtvatAmt, "P")
    
	If interface_Account = "Y" Then	
		if Trim(UCase(frm1.hdnImportflg.Value)) = "Y" Then
			Call SetToolBar("11100000000111")
	        Call SetRdSpreadColor(1)
			frm1.btnPosting.disabled = True
		ElseIf Trim(UCase(frm1.hdnImportflg.Value)) = "N" And frm1.vspdData.MaxRows < 1 Then
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
				Call SetToolBar("11101101001111")
			Else
				Call SetToolBar("11101001000111")
			End If
			frm1.btnPosting.disabled = True
		Else
			If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
				Call SetToolBar("11100000000111")
				call SetRdSpreadColor(1)
			    
			Else
				If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
					Call SetToolBar("11101111001111")
				Else
					Call SetToolBar("11101011000111")
				End If
				
				Call QueryAtSetSpreadColor(1)
			End If
			frm1.btnPosting.disabled = False
		End If
		
		If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
			frm1.btnPosting.value = "Ȯ�����"
			frm1.btnGlSel.disabled = False
		Else
			frm1.btnPosting.value = "Ȯ��"
		    frm1.btnGlSel.disabled = True
		End If
	    If frm1.hdnGlType.Value = "A" Then
	       frm1.btnGlSel.value = "ȸ����ǥ��ȸ"
	    ElseIf frm1.hdnGlType.Value = "T" Then
	       frm1.btnGlSel.value = "������ǥ��ȸ"
	    ElseIf frm1.hdnGlType.Value = "B" Then
	       frm1.btnGlSel.value = "��ǥ��ȸ"
	    End If		
	Else
		If Trim(UCase(frm1.hdnImportflg.Value)) = "Y" Then
			Call SetToolBar("11100000000111")
			call SetRdSpreadColor(1)
			frm1.btnPosting.disabled = True
		ElseIf Trim(UCase(frm1.hdnImportflg.Value)) = "N" And frm1.vspdData.MaxRows < 1 Then
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
				Call SetToolBar("11101101001111")
			Else
				Call SetToolBar("11101001000111")
			End If
			frm1.btnPosting.disabled = True
		Else
			If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
				Call SetToolBar("11100000000111")
	
				call SetRdSpreadColor(1)
			Else
				If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
					Call SetToolBar("11101111001111")
				Else
					Call SetToolBar("11101011000111")
				End If
			
				Call QueryAtSetSpreadColor(1)
			End If
			frm1.btnPosting.disabled = False
		End If
		'***���� 2003.03.26 (ȸ������ ��� Ȯ��,��� �����ϵ��� ����)-Lee Eun Hee
		If Trim(UCase(frm1.hdnPostingflg.Value)) = "Y" Then
			frm1.btnPosting.value = "Ȯ�����"
			frm1.btnGlSel.disabled = True
		Else
			frm1.btnPosting.value = "Ȯ��"
		    frm1.btnGlSel.disabled = True
		End If	
		
	End If
	
	'*********************************************************************
'����(2003-01-15)-������ 
	Dim iIndex,VatDocAmt,SumVatDocAmt,IvAmt,SumIvAmt,NetAmt,SumNetAmt
	
	With frm1.vspdData
	
		If .Maxrows >= 1 then 
			For iIndex = 1 to .Maxrows
				.Row = iIndex
				.Col = 0						
				
				If Trim(.text) <> ggoSpread.DeleteFlag then 						
					'VAT�ݾ� 
					.Col = C_VatDocAmt
					VatDocAmt	=	 unicdbl(.text)						
					SumVatDocAmt = SumVatDocAmt + VatDocAmt
								
					'�Ѹ��Աݾ� 
					.Col = C_IvAmt
					IvAmt	=	 unicdbl(.text)
					'�ΰ��������� ���Աݾ�+VAT�ݾ��� �Ѹ��Աݾ�					
					.Col = C_IOFlgCd					
					If Trim(.text) = "1" then 
						IvAmt = IvAmt + VatDocAmt
					End if
														
					SumIvAmt = SumIvAmt + IvAmt
								
					'���Աݾ� 
					.Col = C_NetAmt
					NetAmt	=	 unicdbl(.text)						
					SumNetAmt = SumNetAmt + NetAmt
				End If					
			Next
			frm1.vspdData.focus
		Else
			SumIvAmt = 0
			SumNetAmt = 0
			SumVatDocAmt = 0
			frm1.txtIvNo.focus
		End if
	End with
	'ȭ�鿡 �Ⱥ��̴� �ݾ��� ����� ���رݾ��� ��.
	lgTotalIvAmt	= unicdbl(frm1.txtIvAmt.Text) - SumIvAmt
	lgTotalNetAmt	= unicdbl(frm1.txtNetAmt.Text) - SumNetAmt
	lgTotalDocAmt	= unicdbl(frm1.txtVatAmt.Text) - SumVatDocAmt
'*******************************************************************
	Call RemovedivTextArea
	
End Function

'===========================  DbSave()  ===================================================
Function DbSave(byval btnflg)
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim ColSep, RowSep
	Dim msgCreditlimit
	Dim GmQty
	Dim MvmtIvQty
	Dim IvQty1,OldIvQty1
	Dim chkVatAmt

	Dim iVatDocAmt
	Dim iChkVatDocAmt
	Dim iRefVatRateFlg
	
	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]

	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size

    DbSave = False                                                          '��: Processing is NG
    
    'On Error Resume Next                                                   '��: Protect system from crashing

	ColSep = parent.gColSep														'��: Column �и� �Ķ��Ÿ 
	RowSep = parent.gRowSep														'��: Row �и� �Ķ��Ÿ 
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ʱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '�ʱ� ������ ����[����]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	With frm1
	.txtMode.value = parent.UID_M0002
		
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
	If btnflg = "Posting" Then
		.txtMode.value = "Release" 				'��: Ȯ�� ��ư 
	End If

    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        Select Case .vspdData.Text
        
        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
        
			If .vspdData.Text = ggoSpread.InsertFlag Then
				strVal = strVal & "C" & ColSep				'��: C=Create
			ElseIf .vspdData.Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & ColSep				'��: U=Update
			End If
			
			.vspdData.Col = C_IvQty1
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","���Լ���","X")
				Call LayerShowHide(0)
				Exit Function
			End If
				
        	.vspdData.Col = C_Cost
			If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
				Call DisplayMsgBox("970021","X","���Դܰ�","X")
				Call LayerShowHide(0)
				Exit Function
			End If
			
			'2003.03 KJH  �ڱ��ݾ� üũ 
			If UCase(parent.gCurrency) <> UCase(Trim(frm1.txtCur.value)) Then
				.vspdData.Col = C_IvLocAmt
				If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
					Call DisplayMsgBox("970021","X","�ڱ��ݾ�","X")
					Call LayerShowHide(0)
					Exit Function
				End If
			End if			
			
			.vspdData.Col = C_GmQty
			If Trim(UNICDbl(.vspdData.Text)) <> "0" Then
			    GmQty = UNICDbl(.vspdData.Text)     '�԰���� 
				 
			    .vspdData.Col = C_oldIvQty1         '��ȸ�� ���Լ��� 
			    OldIvQty1 = UNICDbl(.vspdData.Text)
				 
			    .vspdData.Col = C_IvQty1            '���Լ��� 
			    IvQty1 = UNICDbl(.vspdData.Text) - OldIvQty1 ' new - old
				 
			    .vspdData.Col = C_MvmtIvQty         '�԰� 
			    MvmtIvQty = UNICDbl(.vspdData.Text)
                   
	            If (MvmtIvQty + IvQty1) > GmQty Then
				    
			        msgCreditlimit = DisplayMsgBox("175222", parent.VB_YES_NO, lRow & "��:", "X")
	                If msgCreditlimit = vbYes Then 
                    Else
                        Exit Function
                    End If
			    
			    End If
			
			End If
			
			.vspdData.Col = C_PlantCd:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			'##Interface����##-���Գ������� ȯ�� ���� (2003.09.21)
			.vspdData.Col = C_XchRt:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_PlantNm:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_ItemCd:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_ItemPopup:	strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_ItemNm:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_SpplSpec:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvQty1:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_Unit:			strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_UnitPopup:	strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_Cost:			strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_IvAmt:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_NetAmt:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_IOFlg:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IOFlgCd:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_VatType:		
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
				strVal = strVal & Trim(frm1.hdnVatType.value) & ColSep 'hdr�� vat type���� 
			Else
				strVal = strVal & Trim(.vspdData.Text) & ColSep
			End If
			.vspdData.Col = C_VatPopup:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_VatNm:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_VatRate '19
			If Trim(UCase(frm1.hdnExceptflg.Value)) = "Y" Then
			strVal = strVal & UNICDbl(frm1.hdnVatRt.value) & ColSep 'hdr�� vat rate���� 
			Else
			strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			End If
			.vspdData.Col = C_VatDocAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_IvLocAmt:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_NetLocAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_VatLocAmt:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_OrderQty:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_OrderCost:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_GmQty:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_IvQty2:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_PoNo:			strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_PoSeq:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_MvmtRcptNo:	strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_GmNo:			strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_GmSeq:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_IvSeq:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_OldQty:		strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_MvmtNo:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_MvmtIvQty:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_oldIvQty1:	strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
			.vspdData.Col = C_TrackingNo:	strVal = strVal & Trim(.vspdData.Text) & ColSep

			.vspdData.Col=C_VatDocAmt:		iVatDocAmt=UNICDbl(Trim(.vspdData.Text))
			.vspdData.Col=C_chkVatDocAmt:	iChkVatDocAmt=UNICDbl(Trim(.vspdData.Text))
			.vspdData.Col=C_ref_vatrate_flg: iRefVatRateFlg=Trim(.vspdData.Text)
			
			If iVatDocAmt = iChkVatDocAmt And iRefVatRateFlg = "Y" Then
				strVal = strVal & "Y" & ColSep
			Else 
				strVal = strVal & "N" & ColSep
			End If
			.vspdData.Col = C_LCNo:			strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_LCSeqNo:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_LcFlg:		strVal = strVal & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_Remark:		strVal = strVal & Trim(.vspdData.Text) & ColSep	'����߰� 
			strVal = strVal & lRow & RowSep
		
		Case ggoSpread.DeleteFlag
			
			strDel = strDel & "D" & ColSep				'��: D=Delete
			.vspdData.Col = C_IvSeq:		strDel = strDel & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_PoNo:			strDel = strDel & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_PoSeq:		strDel = strDel & Trim(.vspdData.Text) & ColSep
			.vspdData.Col = C_MvmtNo:		strDel = strDel & Trim(.vspdData.Text) & ColSep
			strDel = strDel & lRow & RowSep
			
		End Select  
		lGrpCnt = lGrpCnt + 1		         
        '=====================
        .vspdData.Col = 0
		Select Case .vspdData.Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
		                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
		       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
		      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
		       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
		         
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select  
        strVal = ""
        strDel = ""
        '=====================
       
    Next
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If 

	If lGrpCnt > 1 Or btnflg = "Posting" Then
		If LayerShowHide(1) = False Then
			Exit function
		End If
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	End If
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'======================================  RemovedivTextArea()  =================================
Function RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'================================  DbSaveOk()  ============================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	Call InitVariables
	Call MainQuery()

End Function

'================================  DbDelete()  ============================================
Function DbDelete() 
End Function

'============================================================================================================
' Name : SubGetGlNo
' Desc : Get Gl_no : 2003.03 KJH ��ǥ��ȣ �������� ���� ���� 
'============================================================================================================
Sub SubGetGlNo()
	Dim lgStrFrom
	Dim strTempGlNo, strGlNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	On Error Resume Next
	Err.Clear 
	
	lgStrFrom =  " ufn_a_GetGlNo( " & FilterVar(frm1.txthdnIvNo.Value, "''", "S") & " )"
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", lgStrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If lgF0 <> "" then 
		strTempGlNo = Split(lgF0, Chr(11))
		strGlNo		= Split(lgF1, Chr(11))
					
		If strGlNo(0) = "" and strTempGlNo(0) = "" then 
			frm1.hdnGlNo.Value		=	""
			frm1.hdnGlType.value	=	"B"
		Elseif strGlNo(0) = "" and strTempGlNo(0) <> "" then
			frm1.hdnGlNo.Value		=	strTempGlNo(0) 
			frm1.hdnGlType.value	=	"T"
		Elseif strGlNo(0) <> "" then 
			frm1.hdnGlNo.Value		=	strGlNo(0) 
			frm1.hdnGlType.value	=	"A"
		End If
	Else
		frm1.hdnGlNo.Value		=	""
		frm1.hdnGlType.value	=	"B"
	End if
	
End Sub

