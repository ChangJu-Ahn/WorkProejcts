<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M5212ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ����B/L ������� ASP														*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : ȭ�� design												*
'*							  2. 2000/04/11 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc ����  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 ���� Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
 Option Explicit
	
	Dim interface_Account

	Const BIZ_PGM_QRY_ID = "m5212mb1.asp"		   '��ȸ 
	Const BIZ_PGM_SAVE_ID = "m5212mb2.asp"		   '����,����,���� 
	Const BIZ_PGM_POSTQRY_ID = "m5211mb5.asp"	   'Ȯ����ư Ŭ���� 
	Const BL_HEADER_ENTRY_ID = "m5211ma1"		   'B/L��� ���� 
	Const CC_LAN_ENTRY_ID = "m5213ma1"			
	Const BIZ_PGM_JUMP_ID_PUR_CHARGE = "M6111MA2"  '����� ���� 
	Const BIZ_PGM_JUMP_ID_IV_Paymen = "M5113MA1"  '���޳������ ���� 

	Const CID_POST  = 5211                           'Ȯ�� 

	Dim lgBlnFlgChgValue		
	Dim lgIntGrpCount			
	Dim lgIntFlgMode			

	Dim lgStrPrevKey
	Dim lgSortKey
	Dim lgLngCurRows
	Dim gblnWinEvent
	
	Dim C_ItemCd								'ǰ���ڵ�			
	Dim C_ItemNm								'ǰ��� 
	Dim C_SPEC									'�԰� 
	Dim C_Unit 									'���� 
	Dim C_Qty 									'B/L���� 
	Dim C_Price 								'�ܰ� 
	Dim C_DocAmt 								'�ݾ׾� 
	Dim C_LocAmt                                '�ڱ��ݾ� 
	Dim C_GrossWeight 							'���߷� 
	Dim C_Volume 								'���� 
	Dim C_HsCd 									'HS��ȣ 
	Dim C_HsNm 									'HS��					
	Dim C_BlSeq									'B/L���� 
	Dim C_PoNo 									'P/O��ȣ 
	Dim C_PoSeq 								'P/O���� 
	Dim C_LcDocNo								'L/C��ȣ 
	Dim C_LcSeq 								'L/C���� 
	Dim C_OverTolerance							'OverTolerance
    Dim C_underTolerance 						'underTolerance
    Dim C_LcNo 									'L/C������ȣ 
	Dim C_TrackingNo	                        'TrackingNo(�߰�)
	Dim C_Remark								'���(�߰�)
	'��ǰ��ݾװ���� ���� �߰�(2003.05)
	Dim C_OrgDocAmt		'��ȭ�� ���� 
	Dim C_OrgDocAmt1	'��ȸ�� �ʱⰪ ���� 
	
	
	Dim C_Qty_Ref

'########################################################################################################
'���޳������ ������ ȣ�� 
Function LoadIvPayment()
	Dim strHdrOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then          
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    WriteCookie "txtIvNo" , Trim(frm1.hdnIvNo.value)
		
	PgmJump(BIZ_PGM_JUMP_ID_IV_Paymen)

End Function

'����� ������ ȣ�� 
Function LoadChargeHdr()

	Dim IntRetCD

    If lgIntFlgMode <> Parent.OPMD_UMODE Then   
        Call displaymsgbox("900002","X","X","X")
        Exit Function
    End if
    '***����(2003.02.25)_Lee,Eun Hee***
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    	
    WriteCookie "Process_Step" , "VB"
	WriteCookie "Po_No" , Trim(frm1.txtBLNo.value)        'B/L������ȣ 
	WriteCookie "Pur_Grp", Trim(frm1.hdnGrpCd.Value)
	WriteCookie "Po_Cur", Trim(frm1.txtCurrency.Value)    'B/L�ݾ� 
	WriteCookie "Po_Xch", Trim(frm1.txtXchRate.Value)
	
	Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
				
End Function

<!--
'==========================================  2.1.1 InitVariables()  =====================================
-->
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE			
	lgBlnFlgChgValue = False			
	lgIntGrpCount = 0					
	lgStrPrevKey = ""					
	lgLngCurRows = 0 					
	gblnWinEvent = False
End Function

<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
Sub SetDefaultVal()
	Call SetToolbar("1110000000001111")	
	frm1.txtDocAmt.text = UNIFormatNumber(UNICDbl(0), ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.btnPosting.Disabled = True
	frm1.btnPosting.value = "Ȯ��"
	frm1.btnGlSel.disabled = true
	frm1.txtBLNo.Focus
	frm1.ChkPrepay.Checked =   false                 '���ޱݿ��� ���� check box
	Set gActiveElement = document.activeElement
	interface_Account = GetSetupMod(Parent.gSetupMod, "a")
End Sub

<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

<!--
'=============================================  2.5.1 LoadBlHdr()  ======================================
-->
Function LoadBlHdr()
	Dim strHdrOpenParam
	Dim IntRetCD

    If lgIntFlgMode <> Parent.OPMD_UMODE Then  
        Call displaymsgbox("900002","X","X","X")
        Exit Function
    End if
    '***����(2003.02.25)_Lee,Eun Hee***
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	WriteCookie "BlNo", UCase(Trim(frm1.txtBLNo.value))			

	PgmJump(BL_HEADER_ENTRY_ID)

End Function
<!--
'==========================================  2.2.3 InitSpreadPosVariables()  ===================================
-->
Sub InitSpreadPosVariables()
	 C_ItemCd			= 1							'ǰ���ڵ�			
	 C_ItemNm			= 2							'ǰ��� 
	 C_SPEC				= 3
	 C_Unit 			= 4							'���� 
	 C_Qty 				= 5							'B/L���� 
	 C_Price 			= 6							'�ܰ� 
	 C_DocAmt 			= 7							'�ݾ� 
	 C_LocAmt			= 8                         '�ڱ��ݾ� 
	 C_GrossWeight		= 9 						'���߷� 
	 C_Volume			= 10						'���� 
	 C_HsCd				= 11 						'HS��ȣ 
	 C_HsNm				= 12 						'HS��					
	 C_BlSeq			= 13	 					'B/L���� 
	 C_PoNo				= 14 						'P/O��ȣ 
	 C_PoSeq			= 15						'P/O���� 
	 C_LcDocNo			= 16						'L/C��ȣ 
	 C_LcSeq 			= 17						'L/C���� 
	 C_OverTolerance	= 18						'OverTolerance
     C_underTolerance	= 19						'underTolerance
     C_LcNo 			= 20						'L/C������ȣ 
	 C_TrackingNo		= 21						'Tracking_No
	 C_Remark			= 22						'��� 
	 C_OrgDocAmt		= 23						'��ȭ�ϴ� ���ݾ�(2003.05)
	 C_OrgDocAmt1		= 24						'��ȸ�� ���ݾ�(2003.05)
End Sub

<!--
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
Sub InitSpreadSheet()
		
	Call InitSpreadPosVariables()
		
    With frm1
    
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20051222",,Parent.gAllowDragDropSpread  
		.vspdData.MaxCols = C_OrgDocAmt1 + 1
		.vspdData.Col = .vspdData.MaxCols
			
		.vspdData.MaxRows = 0
		.vspdData.ReDraw = False
			
		Call GetSpreadColumnPos("A")			
			
		ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 20, 0
		ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 20, 0
		ggoSpread.SSSetEdit		C_SPEC, "ǰ��԰�", 20, 0
		ggoSpread.SSSetEdit		C_Unit, "����", 10, 2
		SetSpreadFloatLocal		C_Qty, "B/L����", 15, 1, 3
		SetSpreadFloatLocal		C_Price, "�ܰ�", 15, 1, 4
		SetSpreadFloatLocal		C_DocAmt, "�ݾ�", 15, 1, 2
		SetSpreadFloatLocal		C_LocAmt, "�ڱ��ݾ�",15,1,2         '13�� �߰� 
		SetSpreadFloatLocal		C_GrossWeight, "���߷�", 15, 1, 3
		SetSpreadFloatLocal		C_Volume, "������", 15, 1, 3
		ggoSpread.SSSetEdit		C_HsCd, "HS��ȣ", 20, 0
		ggoSpread.SSSetEdit		C_HsNm, "HS��", 20, 0
		ggoSpread.SSSetEdit		C_BlSeq, "B/L����",  10, 2
		ggoSpread.SSSetEdit		C_PoNo, "���ֹ�ȣ", 20, 0
		ggoSpread.SSSetEdit		C_PoSeq, "���ּ���", 10, 2
		ggoSpread.SSSetEdit		C_LcDocNo, "L/C��ȣ", 20, 0
		ggoSpread.SSSetEdit		C_LcSeq, "L/C����", 10, 2
		SetSpreadFloatLocal		C_OverTolerance, "�����������(+)", 15, 1, 5
		SetSpreadFloatLocal		C_UnderTolerance, "�����������(-)", 15, 1, 5
		ggoSpread.SSSetEdit		C_LcNo, "L/C��ȣ",5,0
		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No",20
		ggoSpread.SSSetEdit		C_Remark, "���",20
		SetSpreadFloatLocal		C_OrgDocAmt, "C_OrgDocAmt", 15, 1, 2
		SetSpreadFloatLocal		C_OrgDocAmt1, "C_OrgDocAmt1", 15, 1, 2
			
		Call ggoSpread.SSSetColHidden(C_LcNo, C_LcNo, true)
		Call ggoSpread.SSSetColHidden(C_OrgDocAmt, C_OrgDocAmt1, true)
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols, true)
			
		Call SetSpreadLock()
			
		.vspdData.ReDraw = True
			
	End With
End Sub

<!--
'==========================================  2.2.4 GetSpreadColumnPos()  =====================================
-->
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
		
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				
			C_ItemCd		= iCurColumnPos(1)							'ǰ���ڵ�			
			C_ItemNm		= iCurColumnPos(2)							'ǰ��� 
			C_SPEC			= iCurColumnPos(3)							'�԰� 
			C_Unit 			= iCurColumnPos(4)							'���� 
			C_Qty 			= iCurColumnPos(5)							'B/L���� 
			C_Price 		= iCurColumnPos(6)							'�ܰ� 
			C_DocAmt 		= iCurColumnPos(7)							'�ݾ� 
			C_LocAmt		= iCurColumnPos(8)							'�ڱ��ݾ� 
			C_GrossWeight	= iCurColumnPos(9) 							'���߷� 
			C_Volume		= iCurColumnPos(10)							'���� 
			C_HsCd			= iCurColumnPos(11) 						'HS��ȣ 
			C_HsNm			= iCurColumnPos(12) 						'HS��					
			C_BlSeq			= iCurColumnPos(13)	 						'B/L���� 
			C_PoNo			= iCurColumnPos(14) 						'P/O��ȣ 
			C_PoSeq			= iCurColumnPos(15)							'P/O���� 
			C_LcDocNo		= iCurColumnPos(16)							'L/C��ȣ 
			C_LcSeq 		= iCurColumnPos(17)							'L/C���� 
			C_OverTolerance = iCurColumnPos(18)							'OverTolerance
			C_underTolerance= iCurColumnPos(19)							'underTolerance
			C_LcNo 			= iCurColumnPos(20)							'L/C������ȣ 
			C_TrackingNo	= iCurColumnPos(21)							'Tracking_No
			C_Remark		= iCurColumnPos(22)							'��� 
			C_OrgDocAmt		= iCurColumnPos(23)							'�����ݾ�(2003.05)
			C_OrgDocAmt1	= iCurColumnPos(24)							'�����ݾ�(2003.05)
	End Select
End Sub
	

<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
Sub SetSpreadLock()
	    
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
	
	    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
		ggoSpread.SpreadLock C_ItemCd,			-1,		C_ItemCd,			-1
		ggoSpread.SpreadLock C_ItemNm,			-1,		C_ItemNm,			-1
		ggoSpread.SpreadLock C_SPEC,			-1,		C_SPEC,				-1
		ggoSpread.SpreadLock C_Unit,			-1,		C_Unit,				-1
		ggoSpread.SpreadUnLock C_Qty,			-1,		C_Qty,				-1
		ggoSpread.SpreadUnLock C_Price,			-1,		C_Price,			-1
		'ggoSpread.SpreadLock C_DocAmt,			-1,		C_DocAmt,			-1
		
		'2003.3 ��ġ �Է��ʼ� -KJH
		'ggoSpread.SpreadLock C_LocAmt,			-1,		C_LocAmt,			-1
		ggoSpread.SpreadUnLock C_LocAmt,		-1,		C_LocAmt,			-1
		
		ggoSpread.SpreadUnLock C_GrossWeight,	-1,		C_GrossWeight,		-1
		ggoSpread.SpreadUnLock C_Volume,		-1,		C_Volume,			-1
		ggoSpread.SpreadLock C_HsCd,			-1,		C_HsCd,				-1
		ggoSpread.SpreadLock C_HsNm,			-1,		C_HsNm,				-1
		ggoSpread.SpreadLock C_BlSeq,		    -1,		C_BlSeq,			-1
		ggoSpread.SpreadLock C_PoNo,			-1,		C_PoNo,				-1
		ggoSpread.SpreadLock C_PoSeq,			-1,		C_PoSeq,			-1
		ggoSpread.SpreadLock C_LcDocNo,			-1,		C_LcDocNo,			-1
		ggoSpread.SpreadLock C_LcSeq,			-1,		C_LcSeq,			-1
		ggoSpread.SpreadLock C_OverTolerance,	-1,		C_OverTolerance,	-1
		ggoSpread.SpreadLock C_UnderTolerance,	-1,		C_UnderTolerance,	-1
		ggoSpread.SpreadLock C_TrackingNo,		-1,		C_TrackingNo,		-1
		ggoSpread.SpreadLock C_Remark,		 	-1,		C_Remark,		 	-1
		
		.vspdData.ReDraw = True
	End With
	  
End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
		'.Redraw = False

	    ggoSpread.SSSetProtected frm1.vspddata.maxcols, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCd, pvStarRow, pvEndRow              'ǰ�� 
		ggoSpread.SSSetProtected C_ItemNm, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_SPEC, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_Unit, pvStarRow, pvEndRow                '���� 
		ggoSpread.SSSetRequired  C_Qty, pvStarRow, pvEndRow                 'B/L���� 
		ggoSpread.SSSetRequired  C_Price, pvStarRow, pvEndRow               '�ܰ� 
		ggoSpread.SSSetRequired  C_DocAmt, pvStarRow, pvEndRow              '�ݾ� 
		'2003.3 ��ġ �Է��ʼ�-KJH
		ggoSpread.SSSetRequired  C_LocAmt, pvStarRow, pvEndRow              '�ݾ�        12�߰� 
		'ggoSpread.SSSetProtected C_LocAmt, pvStarRow, pvEndRow              '�ݾ�        12�߰� 
			
		ggoSpread.SSSetRequired  C_GrossWeight, pvStarRow, pvEndRow         '���߷� 
		ggoSpread.SSSetRequired  C_Volume, pvStarRow, pvEndRow              '������ 
		ggoSpread.SSSetProtected C_HsCd, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_HsNm, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_BlSeq, pvStarRow, pvEndRow               'B/L���� 
		ggoSpread.SSSetProtected C_PoNo, pvStarRow, pvEndRow                '���ֹ�ȣ 
		ggoSpread.SSSetProtected C_PoSeq, pvStarRow, pvEndRow               '���ּ��� 
		ggoSpread.SSSetProtected C_LcDocNo, pvStarRow, pvEndRow             'L/C��ȣ 
		ggoSpread.SSSetProtected C_LcSeq, pvStarRow, pvEndRow               'L/C���� 
		ggoSpread.SSSetProtected C_OverTolerance, pvStarRow, pvEndRow       '�����������(+)
		ggoSpread.SSSetProtected C_UnderTolerance, pvStarRow, pvEndRow      '�����������(-)
		ggoSpread.SSSetProtected C_TrackingNo, pvStarRow, pvEndRow          'TrackingNo
		ggoSpread.SpreadUnlock 	C_Remark,	-1,	C_Remark,	-1
			
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
		'.Col = 1
		'.Row = .ActiveRow
		'.Action = 0
		'.EditMode = True
			
		'.ReDraw = True
	End With
End Sub
<!--
'==========================================  2.2.5 SetSpreadColor2()  ====================================
-->
Sub SetSpreadColor2()
	Dim Col
	ggoSpread.Source = frm1.vspdData
		
    With frm1.vspdData
	    
		.Redraw = False
	
		For Col = C_ItemCd To C_TrackingNo
			ggoSpread.SSSetProtected Col, -1, -1
		Next
				
		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True

		.ReDraw = True
	End With
End Sub	

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBlNoPop()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenBlNoPop()
	Dim strRet,IntRetCD
	Dim iCalledAspName 
		
		
	If gblnWinEvent = True Or UCase(frm1.txtBLNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M5211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		
	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtBLNo.focus 
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBLNo.value = strRet
		frm1.txtBLNo.focus 
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPODtlRef()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenPODtlRef()
	Dim arrRet,IntRetCD
	Dim arrParam(9)
	Dim iCalledAspName 
	Dim IsOpenPop
		
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call displaymsgbox("900002","X","X","X")
		Exit Function
	End if 

	if Trim(frm1.txtPost.Value) = "Y" then
		Call displaymsgbox("17a009","X","X","X")     'ȸ��ó�������̹Ƿ� ���� �Ҽ� �����ϴ� 
		Exit Function
	End if
		
	arrParam(0) = UCase(Trim(frm1.hdnPoNo.value))           'B/L������ȣ 
	arrParam(1) = UCase(Trim(frm1.hdnPayMethCd.Value))
	arrParam(2) = Trim(frm1.hdnPayMethNm.Value)
	arrParam(3) = UCase(Trim(frm1.hdnIncotermsCd.Value))
	arrParam(4) = Trim(frm1.hdnIncotermsNm.Value)
	arrParam(5) = UCase(Trim(frm1.txtCurrency.Value))
	arrParam(6) = UCase(Trim(frm1.txtBeneficiary.Value))    'B/L�ݾ� 
	arrParam(7) = Trim(frm1.txtBeneficiaryNm.Value)         '������ 
	arrParam(8) = UCase(Trim(frm1.hdnGrpCd.Value))          '�����ڸ� 
	arrParam(9) = Trim(frm1.hdnGrpNm.Value)
		
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M3112RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA2", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If arrRet(0, 0) = "" Then
		frm1.txtBLNo.focus 
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPODtlRef(arrRet)
	End If	
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetPODtlRef()  +++++++++++++++++++++++++++++++++++++++++
-->
Function SetPODtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j, intEndRow
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt, Row1, intCnt2
	Dim temp_Qty,temp_DocAmt, temp1,temp2
	Dim strMessage

	Const C_Ref_ItemCd			= 0
	Const C_Ref_ItemN			= 1
	Const C_Ref_PoQty			= 2
	Const C_Ref_LcQty			= 3
	Const C_Ref_BlQty			= 4
	Const C_Ref_Spec			= 5
	Const C_Ref_Unit			= 6
	Const C_Ref_Price			= 7
	Const C_Ref_PoNo			= 8
	Const C_Ref_PoSeq			= 9
	Const C_Ref_LcNo			= 10
	Const C_Ref_LcSeq			= 11
	Const C_Ref_HsCd			= 12
	Const C_Ref_DocAmt			= 13
	Const C_Ref_LocAmt			= 14
	Const C_Ref_OverTolerance	= 15
	Const C_Ref_UnderTolerance	= 16
	Const C_Ref_TrackingNo		= 17
	Const C_Ref_LcBlQty 		= 18
				
	With frm1 
		    			
	    .vspdData.focus
		ggoSpread.Source = .vspdData
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
		.vspdData.ReDraw = False	

		TempRow = .vspdData.MaxRows					'����Ʈ max��				
		intLoopCnt = Ubound(arrRet, 1)				'�˾����� ������ row count			
			
		blnEqualFlg = False
		intCnt2 = 0
			
		For intCnt = 1 to intLoopCnt            'row count ��ŭ ���鼭(�˾����� ���� Ű������)

			If TempRow <> 0 Then
				'���� ����Ÿ�� �ִ��� ���ϴ� ���� ���� 
				For j = 1 To TempRow                '���� data ��ŭ ���鼭 �� 
					.vspdData.Row = j
					.vspdData.Col = C_PoNo
					temp1 = .vspdData.Text
					.vspdData.Col = C_PoSeq
					temp2 = .vspdData.Text
					If temp1 = arrRet(intCnt - 1, C_Ref_PoNo) and temp2 = arrRet(intCnt - 1, C_Ref_PoSeq) Then
						strMessage = strMessage & arrRet(intCnt - 1, C_Ref_PoNo) & "-" & arrRet(intCnt - 1, C_Ref_PoSeq) & ";"
						blnEqualFlg = True
						Exit For
					Else
						blnEqualFlg = False
					End If
				Next
				'�� ���� �� 
			End If

			If blnEqualFlg = False Then        '������ ������ grid�� �Կ��ش� 
			
				'������ ���� ��ȣ�� �ִ� ���� ���� �Ǿ����� ���� ���� ���� �߰��Ǿ�� �Ѵ�.
				intCnt2 = intCnt2 + 1	
				.vspdData.MaxRows = CLng(TempRow) + CLng(intCnt2)
				.vspdData.Row = CLng(TempRow) + CLng(intCnt2)
				
				Row1 = .vspdData.Row
				'lc����� �ǵ� �������� ���ܵǵ��� ����(2003.07.25)
				temp_Qty = UNIFormatNumber((UNICDbl(arrRet(intCnt - 1, C_Ref_PoQty)) - UNICDbl(arrRet(intCnt - 1, C_Ref_BlQty))-UNICDbl(arrRet(intCnt - 1, C_Ref_LcQty))+UNICDbl(arrRet(intCnt - 1, C_Ref_LcBlQty))),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				temp_DocAmt = UNIFormatNumber((UNICDbl(arrRet(intCnt - 1, C_Ref_PoQty)) - UNICDbl(arrRet(intCnt - 1, C_Ref_BlQty))) * UNICDbl(arrRet(intCnt - 1, C_Ref_Price)),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
				
				Call .vspdData.SetText(0		,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_ItemCd	,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
				Call .vspdData.SetText(C_ItemNm	,	Row1, arrRet(intCnt - 1, C_Ref_ItemN))
				Call .vspdData.SetText(C_Spec	,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
				Call .vspdData.SetText(C_Unit	,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
				Call .vspdData.SetText(C_Qty	,	Row1, temp_Qty)
				Call .vspdData.SetText(C_Price	,	Row1, arrRet(intCnt - 1, C_Ref_Price))
				Call .vspdData.SetText(C_DocAmt	,	Row1, temp_DocAmt)
				Call .vspdData.SetText(C_HsCd	,	Row1, arrRet(intCnt - 1, C_Ref_HsCd))
				Call .vspdData.SetText(C_HsNm	,	Row1, "")
				Call .vspdData.SetText(C_LcSeq	,	Row1, "")
				Call .vspdData.SetText(C_PoNo	,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
				Call .vspdData.SetText(C_PoSeq	,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
				Call .vspdData.SetText(C_OverTolerance,	Row1, arrRet(intCnt - 1, C_Ref_OverTolerance))
				Call .vspdData.SetText(C_UnderTolerance,Row1, arrRet(intCnt - 1, C_Ref_UnderTolerance))
				Call .vspdData.SetText(C_TrackingNo,	Row1, arrRet(intCnt - 1, C_Ref_TrackingNo))
									
				Call vspdData_Change(C_Qty_Ref, .vspdData.Row)		'���� �ٲ���Ƿ� changeȣ�� 

				'SetSpreadColor CLng(TempRow)+CLng(intCnt), CLng(TempRow)+CLng(intCnt)
			End If
		Next
			
		intEndRow = .vspdData.MaxRows
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee
		Call TotalSum
		Call SetSpreadColor(TempRow+1,intEndRow)	
					
		if strMessage<>"" then
			Call displaymsgbox("17a005","X",strmessage,"���ֹ�ȣ" & "," & "���ּ���")
			.vspdData.ReDraw = True
			Exit Function
		End if
			
		.vspdData.ReDraw = True

	End With
		
End Function
	
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCDtlRef()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCDtlRef()
	Dim arrRet
	Dim arrParam(13)
	Dim iCalledAspName
	Dim IntRetCD
		
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call displaymsgbox("900002","X","X","X")
		Exit Function
	End if 

	if Trim(frm1.txtPost.Value) = "Y" then
		Call displaymsgbox("17a009","X","X","X")
		Exit Function
	End if
		
	arrParam(0) = UCase(Trim(frm1.hdnLcNo.value))
	arrParam(1) = UCase(Trim(frm1.hdnPayMethCd.Value))
	arrParam(2) = Trim(frm1.hdnPayMethNm.Value)
	arrParam(3) = UCase(Trim(frm1.hdnIncotermsCd.Value))
	arrParam(4) = Trim(frm1.hdnIncotermsNm.Value)
	arrParam(5) = UCase(Trim(frm1.txtCurrency.Value))
	arrParam(6) = UCase(Trim(frm1.txtBeneficiary.Value))
	arrParam(7) = Trim(frm1.txtBeneficiaryNm.Value)
	arrParam(8) = UCase(Trim(frm1.hdnGrpCd.Value))
	arrParam(9) = Trim(frm1.hdnGrpNm.Value)
		
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M3212RA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212RA3", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	If arrRet(0, 0) = "" Then
		Exit Function
	Else
		Call SetLCDtlRef(arrRet)
	End If	
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetLCDtlRef()  +++++++++++++++++++++++++++++++++++++++++
-->
Function SetLCDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j, Row1, intEndRow
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt, intCnt2
	Dim temp_Qty,temp_DocAmt, temp1,temp2
	Dim strMessage
    Dim XchRt
        
	Const C_Ref_ItemCd			= 0
	Const C_Ref_ItemNm			= 1
	Const C_Ref_LcQty			= 2
	Const C_Ref_LlQty			= 3
	Const C_Ref_Spec			= 4 
	Const C_Ref_Unit			= 5
	Const C_Ref_Price			= 6
	Const C_Ref_LcNo			= 7
	Const C_Ref_LcSeq			= 8
	Const C_Ref_PoNo			= 9
	Const C_Ref_PoSeq			= 10
	Const C_Ref_TrackingNo		= 11
	Const C_Ref_HsCd			= 12
	Const C_Ref_HsNm			= 13
	Const C_Ref_OverTolerance	= 14
	Const C_Ref_UnderTolerance	= 15
	Const C_Ref_LcNo2			= 16

	With frm1 
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	

		TempRow = .vspdData.MaxRows								
		intLoopCnt = Ubound(arrRet, 1)							
			
		blnEqualFlg = False
		intCnt2 = 0
		
		For intCnt = 1 to intLoopCnt + 1

			If TempRow <> 0 Then
				For j = 1 To TempRow
					.vspdData.Row = j
					.vspdData.Col = C_LcDocNo
					temp1 = .vspdData.Text
					.vspdData.Col = C_LcSeq
					temp2 = .vspdData.Text
						
					If temp1 = arrRet(intCnt - 1, C_Ref_LcNo) and temp2 = arrRet(intCnt - 1, C_Ref_LcSeq) Then
						strMessage = strMessage & arrRet(intCnt - 1, C_Ref_LcNo) & "," & arrRet(intCnt - 1, C_Ref_LcSeq) & ";"
						blnEqualFlg = True
						Exit For
					Else
						blnEqualFlg = False
					End If
				Next
			End If


			If blnEqualFlg = False Then
				'������ ���� ��ȣ�� �ִ� ���� ���� �Ǿ����� ���� ���� ���� �߰��Ǿ�� �Ѵ�.
				intCnt2 = intCnt2 + 1	
				.vspdData.MaxRows = CLng(TempRow) + CLng(intCnt2)
				.vspdData.Row = CLng(TempRow) + CLng(intCnt2)
				
				Row1 = .vspdData.Row
				
				temp_Qty = UNIFormatNumber(UNICDbl(arrRet(intCnt - 1, C_Ref_LcQty)) - UNICDbl(arrRet(intCnt - 1, C_Ref_LlQty)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				temp_DocAmt = UNIFormatNumber((UNICDbl(arrRet(intCnt - 1, C_Ref_LcQty)) - UNICDbl(arrRet(intCnt - 1, C_Ref_LlQty))) * UNICDbl(arrRet(intCnt - 1, C_Ref_Price)),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit) 
				
				Call .vspdData.SetText(0		,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_ItemCd	,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
				Call .vspdData.SetText(C_ItemNm	,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
				Call .vspdData.SetText(C_Spec	,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
				Call .vspdData.SetText(C_Unit	,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
				Call .vspdData.SetText(C_Qty	,	Row1, temp_Qty)
				Call .vspdData.SetText(C_Price	,	Row1, arrRet(intCnt - 1, C_Ref_Price))
				Call .vspdData.SetText(C_DocAmt	,	Row1, temp_DocAmt)
				Call .vspdData.SetText(C_HsCd	,	Row1, arrRet(intCnt - 1, C_Ref_HsCd))
				Call .vspdData.SetText(C_HsNm	,	Row1, arrRet(intCnt - 1, C_Ref_HsNm))
				Call .vspdData.SetText(C_LcDocNo,	Row1, arrRet(intCnt - 1, C_Ref_LcNo))
				Call .vspdData.SetText(C_LcSeq	,	Row1, arrRet(intCnt - 1, C_Ref_LcSeq))
				Call .vspdData.SetText(C_PoNo	,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
				Call .vspdData.SetText(C_PoSeq	,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
				'Tolerance(�����ʵ��̹Ƿ�) Format ���� ����(2003.06.13)
				Call .vspdData.SetText(C_OverTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_OverTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_UnderTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_UnderTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_LcNo	,	Row1, arrRet(intCnt - 1, C_Ref_LcNo))
				Call .vspdData.SetText(C_TrackingNo,	Row1, arrRet(intCnt - 1, C_Ref_TrackingNo))
									
				Call vspdData_Change(C_Qty_Ref, Row1)
										
			End If
		Next
		
		intEndRow = .vspdData.MaxRows	
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee
		Call TotalSum
		Call SetSpreadColor(TempRow+1,intEndRow)
		ggoSpread.spreadlock C_Price,  TempRow+1,C_Price,	intEndRow
			
		if strMessage<>"" then
			Call displaymsgbox("17a005","X",strmessage,"L/C��ȣ" & "," & "L/C����")
			.vspdData.ReDraw = True
			Exit Function
		End if
			
		.vspdData.ReDraw = True

	End With
End Function

'��ǥ��ȸ Ŭ���� ȣ�� 
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True
	
	arrParam(0) = FilterVar(Trim(frm1.hdnGlNo.value),"","SNM")
	arrParam(1) = FilterVar(Trim(frm1.hdnIvNo.value),"","SNM")          '���Թ�ȣ 
	'arrParam(2) = Trim(frm1.txtGrpCd.value)
	'arrParam(3) = Trim(frm1.txtGrpNm.value)
	
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
     	Call displaymsgbox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	gblnWinEvent = False
	
End Function

<!--
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
-->
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
    End Select
         
End Sub
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'�ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt1,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

	End With

End Sub
<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
		
	frm1.txtBLNo.Value = ReadCookie("BlNo")
	frm1.hdnPoNo.Value = ReadCookie("PoNo")
		
	frm1.hdnQueryType.Value = "autoQuery"
		
	WriteCookie "BlNo",""
	WriteCookie "PoNo",""
		
	if Trim(frm1.txtBLNo.Value) <> "" then Call dbQuery
			
End Function

<!--
'============================================  2.5.1 SetSpreadLockAfterQuery()  ======================================
-->
Function SetSpreadLockAfterQuery()
	ggoSpread.source = frm1.vspdData
	ggoSpread.SpreadLock 1,1,frm1.vspdData.MaxCols,frm1.vspdData.MaxRows
End Function
<!--
'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header ȭ�����κ��� �Ѱܹ��� parameter setting(Cookie ���)				=
'========================================================================================================
-->
Sub TotalSum()
	Dim SumTotal, lRow
		
	SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	ggoSpread.source = frm1.vspdData
	For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text = ggoSpread.InsertFlag then
			frm1.vspdData.Col = C_DocAmt
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		end if
	Next
	frm1.txtDocAmt.Text = UNIConvNumPCToCompanyByCurrency(Cstr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

End Sub
'########################################################################################
'============================================  2.5.1 TotalSumNew()  ======================================
'=	Name : TotalSumNew()																					=
'=	Description : Master L/C Header ȭ�����κ��� �Ѱܹ��� parameter setting(Cookie ���)				=
'========================================================================================================
Sub TotalSumNew(ByVal row)
		
    Dim SumTotal, lRow, tmpGrossAmt

	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_DocAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = C_OrgDocAmt							
	SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))

        
    frm1.txtDocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub
'######################################################################################


<!--
'=============================================  2.5.1 LoadCCLan()  ======================================
-->
Function LoadCCLan()
	Dim strCCLanOpenParam

	WriteCookie "CCNo", UCase(Trim(frm1.txtBLNo.value))	

	PgmJump(CC_LAN_ENTRY_ID)

End Function
	
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	Call LoadInfTB19029	
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")	
	Call InitSpreadSheet
	Call InitVariables
	Call SetDefaultVal
	Call OpenCookie()
	
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	   
End Sub

<!--
'=======================================  3.2.1 btnBLNoOnClick()  ======================================
-->
Sub btnBLNoOnClick()
	Call OpenBlNoPop()
End Sub

<!--
'======================================  btnPosting_OnClick()   ==================================
-->
Sub btnPosting_OnClick()
	Dim Answer
	Dim strVal
	Dim strBLNo
	
    Err.Clear                                       

	strBLNo = frm1.txtBLNo.value 
		
	If strBLNo = "" Then							
		Call displaymsgbox("900002","X","X","X")
		
	Else
		
		if Trim(frm1.txtPost.Value) = "N" then
			Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")   '�۾��� ���� �Ͻðڽ��ϱ�?
			If Answer = vbNo Then
				frm1.btnPosting.disabled = False	'20040315          
				Exit Sub
			Else 
				frm1.btnPosting.disabled = True		'20040315   
			End If	
		else
			Answer = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
			If Answer = vbNo Then
				frm1.btnPosting.disabled = False	'200308    
				Exit Sub
			Else 		
				frm1.btnPosting.disabled = True		'200308 	
			End If
		End if
		
		If Answer = VBYES Then
			If  LayerShowHide(1) = False Then
		   	Exit Sub
		End If
			strVal = BIZ_PGM_POSTQRY_ID & "?txtMode=" & CID_POST			
			strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)		
				if Trim(frm1.txtPost.Value) = "N" then
				strVal = strVal & "&txtPost=" & "C"
			elseif Trim(frm1.txtPost.Value) = "Y" then
				strVal = strVal & "&txtPost=" & "D"
			End if
			strVal = strVal & "&txtBlIssueDt=" & Trim(frm1.txtIssueDt.Text)
			strVal = strVal & "&txtIvNo=" & Trim(frm1.hdnIvNo.Value)
				
				
			Call RunMyBizASP(MyBizASP, strVal)						

		End IF
	End IF
End Sub
'======================================  vspdData_Click()   ==================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
     
	gMouseClickStatus = "SPC"  
    
	Set gActiveSpdSheet = frm1.vspdData
	
	'Call SetPopupMenuItemInf("0101111111") 
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		if UCase(Trim(frm1.txtPost.Value)) = "Y" then
			Call SetPopupMenuItemInf("0000111111")
		Else
			If frm1.vspddata.maxRows > 0 Then
				Call SetPopupMenuItemInf("0101111111")
			Else
				Call SetPopupMenuItemInf("0001111111")
			End If
		End if
	Else	
			Call SetPopupMenuItemInf("0000111111")
	End If

		   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If   
	
End Sub
'======================================  vspdData_MouseDown()   ==================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
 
End Sub    

'======================================  vspdData_ScriptDragDropBlock()   ==================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'======================================  FncSplitColumn()   ==================================
Function FncSplitColumn()
     If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function


<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
Sub vspdData_Change(ByVal Col, ByVal Row )
	
	Dim Qty, Price, DocAmt, LocAmt, XchRate
	    
    XchRate = UNICDbl(frm1.txtXchRate.value) 
		
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
		
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
	   If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
	      Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
	   End If
	End If
		
	' Call CheckMinNumSpread(frm1.vspdData, Col, Row)        

	lgBlnFlgChgValue = True

	Select Case col
	Case C_Qty, C_Price, C_Qty_Ref                '�����ΰ�� ���� ���� ���� 0���� setting
			
		frm1.vspdData.Col = C_Qty
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			Qty = 0
		Else
			Qty = UNICDbl(frm1.vspdData.Text)
		End If
			
		frm1.vspdData.Col = C_Price
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			Price = 0
		Else
			Price = UNICDbl(frm1.vspdData.Text)
		End If
			
		DocAmt = Qty * Price            '�ݾ��� �ܰ� * ���� 
			
		frm1.vspdData.Col = C_DocAmt
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(CDBL(DocAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
		    
	    '13�� �߰� 
	    frm1.vspdData.Col = C_DocAmt
	    If frm1.vspdData.Text <> "" Then
		    DocAmt = UNICDbl(frm1.vspdData.Text)
		Else				
			DocAmt = 0
		End IF
		    
	     If Trim(frm1.hdnDiv.value) = "*" then
			LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt) * UNICDbl(Trim(frm1.txtXchRate.Value)),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    elseif Trim(frm1.hdnDiv.value) = "/" then
	        LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt) / UNICDbl(Trim(frm1.txtXchRate.Value)),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    else
	        LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    end if
		    
		    
	    frm1.vspdData.Col = C_LocAmt
	    frm1.vspdData.Text = LocAmt
		'13�� �߰� �� 
		'����(ȭ�鼺�ɰ�������)-2003.04.03-Lee Eun Hee	
		If col <> C_Qty_Ref Then
			Call TotalSumNew(Row)					'��ǰ��ݾ��հ� 
		End If
		'�ѱݾװ���� ���� �ʿ�(2003.05)
		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt
		
	'13�� �߰� 
	Case C_DocAmt
		    
	    frm1.vspdData.Col = C_DocAmt
	    If frm1.vspdData.Text <> "" Then
		    DocAmt = UNICDbl(frm1.vspdData.Text)
		Else				
			DocAmt = 0
		End IF
		    
	    If Trim(frm1.hdnDiv.value) = "*" then
	        LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt) * UNICDbl(Trim(frm1.txtXchRate.Value)),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    elseif Trim(frm1.hdnDiv.value) = "/" then
	        LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt) / UNICDbl(Trim(frm1.txtXchRate.Value)),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    else
	        LocAmt  = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
	    end if
		    
	    frm1.vspdData.Col = C_LocAmt
	    frm1.vspdData.Text = LocAmt
		    
	    Call TotalSumNew(Row)					'�հ�ȣ�� 
	    '�ѱݾװ���� ���� �ʿ�(2003.05)
		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt
	'13�� �߰� �� 
	end select
End Sub

<!--
'========================================  3.3.2 vspdData_ColWidthChange()  ==================================
-->
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
		
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
	
<!--
'========================================  3.3.2 vspdData_DblClick()  ==================================
-->
Sub vspdData_DblClick(ByVal Col, ByVal Row)
'	 Dim iColumnName
    
'	 If Row <= 0 Then
'		Exit Sub  
'	 End if
		 
'	 If frm1.vspdData.MaxRow = 0 Then
'		Exit Sub  
'	 End if
End Sub
<!--
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
-->

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then	
				DbQuery
			End If
		End If
	End With
End Sub
	
<!--
'========================================  3.3.3 vspdData_TopLeftChange()  ==================================
-->
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

<!--
'=========================================  5.1.1 FncQuery()  ===========================================
-->
Function FncQuery()
	Dim IntRetCD

	FncQuery = False					

	Err.Clear							

	ggoSpread.Source = frm1.vspdData
		
	If lgBlnFlgChgValue = True Then
		IntRetCD = displaymsgbox("900013", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData					
	Call InitVariables											

	frm1.hdnQueryType.Value = "Query"

	If DbQuery = False Then Exit Function

	FncQuery = True
	Set gActiveElement = document.activeElement		
End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
Function FncNew()
	Dim IntRetCD 

	FncNew = False          
	ggoSpread.Source = frm1.vspdData
		
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = displaymsgbox("900015", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")							
	Call ggoOper.LockField(Document, "N")							
	Call InitVariables				
		
	Call SetDefaultVal

	FncNew = True													
	Set gActiveElement = document.activeElement
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
		
	ggoSpread.Source = frm1.vspdData
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then							
		Call displaymsgbox("900002","X","X","X")
		Exit Function
	End If

	If DbDelete = False Then Exit Function

	FncDelete = True										
	Set gActiveElement = document.activeElement	
End Function
	
<!--
'===========================================  5.1.4 FncSave()  ==========================================
-->
Function FncSave()
	Dim IntRetCD
		
	FncSave = False											
		
	Err.Clear
						
	If CheckRunningBizProcess = True Then
    	Exit Function
	End If											
		
	ggoSpread.Source = frm1.vspdData                        
	If ggoSpread.SSCheckChange = False Then                 
	    IntRetCD = displaymsgbox("900001","X","X","X")      
	    Exit Function
	End If
    
	ggoSpread.Source = frm1.vspdData                        
	If Not ggoSpread.SSDefaultCheck         Then            
	   Exit Function
	End If
		
	If DbSave = False Then Exit Function
		
	If frm1.txtHBLNo.value <> frm1.txtBLNo.value then
		frm1.txtBLNo.value =	frm1.txtHBLNo.value		
	End If												
		
	FncSave = True	
	Set gActiveElement = document.activeElement									
End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
Function FncCopy()
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData
		
	If lgBlnFlgChgValue = True Then
		IntRetCD = displaymsgbox("900017", Parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE										
	frm1.vspdData.ReDraw = False
	if frm1.vspdData.Maxrows < 1	then exit function

	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow 

	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
Function FncCancel() 
	Dim SumTotal,tmpGrossAmt,orgtmpGrossAmt, Row, CUDflag
	if frm1.vspdData.Maxrows < 1	then exit function
	'�ѱݾװ�����(2003.05.28)
	'---------------------------------------------
    SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	Row = frm1.vspdData.SelBlockRow
		
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_DocAmt
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = C_OrgDocAmt1
	orgtmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = 0
	CUDflag = frm1.vspdData.Text
				
    If CUDflag = ggoSpread.UpdateFlag Then
        SumTotal = SumTotal + (orgtmpGrossAmt - tmpGrossAmt )
    ElseIf CUDflag = ggoSpread.InsertFlag  Then
        SumTotal = SumTotal - tmpGrossAmt
    End If

	frm1.txtDocAmt.Text = SumTotal
	'--------------------------------------------
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo	
	Set gActiveElement = document.activeElement			
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
		
	On Error Resume NExt
		
	FncInsertRow = False
		
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = AskSpdSheetAddRowCoun() 
	Else
		imRow = AskSpdSheetAddRowCoun() 
			
		If imRow = "" Then
			Exit Function
		End If
	End If
		 
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData

		'.vspdData.EditMode = True

		.vspdData.ReDraw = False
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True

		Set gActiveElement = document.ActiveElement 
			
		If Err.Number  = 0 Then
			FncInsertRow = True
		End If
    End With
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	
	if frm1.vspdData.Maxrows < 1	then exit function
		
	With frm1.vspdData 
	
		.focus
		ggoSpread.Source = frm1.vspdData

		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
	End With
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
Function FncPrint()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
Function FncPrev() 
	ggoSpread.Source = frm1.vspdData
	If lgIntFlgMode <> Parent.OPMD_UMODE Then	
		Call displaymsgbox("900002","X","X","X")
		Exit Function
	ElseIf lgPrevNo = "" Then			
		Call displaymsgbox("900011","X","X","X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
	ggoSpread.Source = frm1.vspdData
	If lgIntFlgMode <> Parent.OPMD_UMODE Then		
		Call displaymsgbox("900002","X","X","X")
		Exit Function
	ElseIf lgNextNo = "" Then				
		Call displaymsgbox("900012","X","X","X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncExport(Parent.C_SINGLEMULTI)
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
Function FncFind()
	ggoSpread.Source = frm1.vspdData
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
	Set gActiveElement = document.activeElement
End Function

'=========================================  PopSaveSpreadColumnInf()  ============================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=========================================  PopRestoreSpreadColumnInf()  ============================
Sub PopRestoreSpreadColumnInf()
	Dim index
    ggoSpread.Source = gActiveSpdSheet
       
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	
    If UCase(Trim(frm1.txtPost.Value)) = "Y" Then 
		Call SetSpreadColor2()
	Else
		Call SetSpreadColor(-1, -1)
		Call CurFormatNumSprSheet()
		
		For index = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = index
			frm1.vspdData.Col = C_LcDocNo
		
			If Trim(frm1.vspdData.Text) <> "" then                'LC��ȣ ���� ������ �ܰ��� lock ó�� 
				ggoSpread.spreadlock		C_Price,	 index,		C_Price,	 index
				ggoSpread.SSSetProtected C_Price,	    -1,		C_Price,		-1
			End If
		Next
	End If
	
End Sub
<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = displaymsgbox("900016", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
	Set gActiveElement = document.activeElement
End Function
<!--
'=============================================  5.2.1 DbQuery()  ========================================
-->
Function DbQuery()
	Dim strVal

	Err.Clear													

	DbQuery = False												

    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
		
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001				
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtHBLNo.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001				
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey		
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
	strVal = strVal & "&txtQueryType=" & Trim(frm1.hdnQueryType.value)
		
	Call RunMyBizASP(MyBizASP, strVal)									
	
	DbQuery = True														
End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave() 
	Dim lRow, ColSep
	Dim lGrpCnt
	Dim strVal, strDel
	Dim intIndex
	Dim strItemCd, strUnit, strQty, strPrice, strDocAmt, strLocAmt, strGrossWeight, strVolume
	Dim strHsCd, strBlSeq, strPoNo, strPoSeq, strLcDocNo, strLcSeq, strOverTol, strUnderTol, strLcNo, strTrackingNo
	Dim strRemark
	
	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]

	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size	
			
	DbSave = False				
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ʱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '�ʱ� ������ ����[����]
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
		   
    If  LayerShowHide(1) = False Then
       	Exit Function
    End If
	
	ColSep = Parent.gColSep	
	
	With frm1
		
		.txtMode.value = Parent.UID_M0002

		lGrpCnt = 0

		strVal = ""
		strDel = ""

		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
			Case ggoSpread.DeleteFlag
				strDel = "D" & ColSep	'0

		        .vspdData.Col = C_BlSeq			'1
		        strDel = strDel & Trim(.vspdData.Text) & ColSep & lRow & Parent.gRowSep

		        lGrpCnt = lGrpCnt + 1 
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				
				If .vspdData.Text=ggoSpread.InsertFlag Then
					strVal = "C" & ColSep	'0
				Else
					strVal = "U" & ColSep
				End If          
		
			.vspdData.Col  = C_Qty	
			IF UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
				Call displaymsgbox("970021","X","B/L����","X")
				Call SetActiveCell(frm1.vspdData,C_Qty,lRow,"M","X","X")
				Call LayerShowHide(0)
				Exit Function
			End IF
					
			'2007.2 ��ġ �ڱ��ݾ� �Է��ʼ� ����- KSJ
			'.vspdData.Col  = C_Price
			'If UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
			'	Call displaymsgbox("970021","X","�ܰ�","X")
			'	Call SetActiveCell(frm1.vspdData,C_Price,lRow,"M","X","X")
			'	Call LayerShowHide(0)
			'	Exit Function
			'End If
				
			'2003.3 ��ġ �ڱ��ݾ� �Է��ʼ�- KJH
			'.vspdData.Col  = C_DocAmt	
			'IF UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
			'	Call displaymsgbox("970021","X","�ݾ�","X")
			'	Call SetActiveCell(frm1.vspdData,C_DocAmt,lRow,"M","X","X")
			'	Call LayerShowHide(0)
			'	Exit Function
			'End If
					
			'.vspdData.Col  = C_LocAmt
			'IF UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
			'	Call displaymsgbox("970021","X","�ڱ��ݾ�","X")
			'	Call SetActiveCell(frm1.vspdData,C_LocAmt,lRow,"M","X","X")
			'	Call LayerShowHide(0)
			'	Exit Function
			'End IF	
			'2003.3 ��ġEnd - KJH
			'2007.2 ��ġEnd �ڱ��ݾ� �Է��ʼ� ����- KSJ
							
			.vspdData.Col  = C_GrossWeight
			IF UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
				Call displaymsgbox("970021","X","���߷�","X")
				Call SetActiveCell(frm1.vspdData,C_GrossWeight,lRow,"M","X","X")
				Call LayerShowHide(0)
				Exit Function
			End IF
							
			.vspdData.Col  = C_Volume
			IF UNIConvNum(Trim(.vspdData.Text),0) <= 0 then
				Call displaymsgbox("970021","X","������","X")
				Call SetActiveCell(frm1.vspdData,C_Volume,lRow,"M","X","X")
				Call LayerShowHide(0)
				Exit Function
			End IF	
							
			.vspdData.Col = C_ItemCd	'1
			strItemCd = Trim(.vspdData.Text)
			.vspdData.Col = C_Unit 	'2
			strUnit = Trim(.vspdData.Text)
			.vspdData.Col = C_Qty  	'3
			strQty = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col = C_Price   	'4
			strPrice = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col =  C_DocAmt   	'5
			strDocAmt = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col = C_LocAmt  	'6
			strLocAmt = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col = C_GrossWeight  	'7
			strGrossWeight = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col = C_Volume  	'8
			strVolume = UNIConvNum((.vspdData.Text), 0)
			.vspdData.Col = C_HsCd   	'9
			strHsCd = Trim(.vspdData.Text)
			.vspdData.Col = C_BlSeq  	'10
			strBlSeq = Trim(.vspdData.Text)
			.vspdData.Col = C_PoNo   	'11
			strPoNo = Trim(.vspdData.Text)
			.vspdData.Col = C_PoSeq   	'12
			strPoSeq = Trim(.vspdData.Text)
			.vspdData.Col = C_LcDocNo  	'13
			strLcDocNo = Trim(.vspdData.Text)
			.vspdData.Col = C_LcSeq   	'14
			strLcSeq = Trim(.vspdData.Text)
			.vspdData.Col = C_OverTolerance	  	'15
			strOverTol = Trim(.vspdData.Text)
			.vspdData.Col = C_underTolerance   	'16
			strUnderTol = Trim(.vspdData.Text)
			.vspdData.Col = C_LcNo   	'17
			strLcNo = Trim(.vspdData.Text)
			.vspdData.Col = C_TrackingNo   	'18
			strTrackingNo = Trim(.vspdData.Text)
			.vspdData.Col = C_Remark   	
			strRemark = Trim(.vspdData.Text)
	       
			strVal = strVal & strItemCd & ColSep &strUnit & ColSep &strQty & ColSep &strPrice & ColSep &strDocAmt & ColSep &strLocAmt & ColSep & strGrossWeight & ColSep & strVolume & ColSep & _   
							strHsCd & ColSep & strBlSeq & ColSep &strPoNo & ColSep &strPoSeq & ColSep &strLcDocNo & ColSep &strLcSeq & ColSep & strOverTol & ColSep & _
							strUnderTol & ColSep & strLcNo & ColSep & strTrackingNo & ColSep  
			strVal = strVal & ColSep & "0" & ColSep & ColSep & ColSep & ColSep & ColSep & "0" & ColSep & "0" & ColSep & ColSep & ColSep & ColSep
			strVal = strVal & "0" & ColSep & "0" & ColSep & "0" & ColSep & "0" & ColSep & ColSep & ColSep & "0" & ColSep & "0" & ColSep & "0" & ColSep & ColSep & ColSep & ColSep
			strVal = strVal & "0" & ColSep & "0" & ColSep & strRemark & ColSep
			strVal = strVal & lRow & Parent.gRowSep
			
			lGrpCnt = lGrpCnt + 1
  
		End Select 
		
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
	
		'.txtMaxRows.value	= lGrpCnt
		'.txtSpread.value	= strDel
		

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)	

	End With

	DbSave = True												
End Function

'======================================  RemovedivTextArea()  =================================
Function RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function	
<!--
'=============================================  5.2.3 DbDelete()  =======================================
-->
Function DbDelete()
End Function
	
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()											
	Dim index
		
	lgIntFlgMode = Parent.OPMD_UMODE

	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")	
	Call RemovedivTextArea
		
	'**����(2003.03.26)-ȸ������ ��� Ȯ��,Ȯ����� �����ϵ��� ������.
	if UCase(Trim(frm1.txtPost.Value)) = "Y" then
		Call SetSpreadLockAfterQuery()                            'Ȯ���̸� ���� �Ұ��� ��ü�� lock
		Call SetToolbar("11100000000111")
		frm1.btnPosting.value = "Ȯ�����"
		If interface_Account <> "N" Then
			frm1.btnGlSel.disabled = False
		Else
			frm1.btnGlSel.disabled = True
		End If
	Else
		frm1.vspdData.Redraw = false
		SetSpreadColor -1	, -1
		For index = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = index
			frm1.vspdData.Col = C_LcDocNo
			if Trim(frm1.vspdData.Text) <> "" then                'LC��ȣ ���� ������ �ܰ��� lock ó�� 
				ggoSpread.spreadlock C_Price,index,C_Price,index
			end if
		Next
			
		frm1.vspdData.Redraw = true	
		Call SetToolbar("11101011000111")
		frm1.btnPosting.value = "Ȯ��"
		frm1.btnGlSel.disabled = true
	End if
		
    if frm1.hdnGlType.Value = "A" Then
       frm1.btnGlSel.value = "ȸ����ǥ��ȸ"
    elseif frm1.hdnGlType.Value = "T" Then
       frm1.btnGlSel.value = "������ǥ��ȸ"
    elseif frm1.hdnGlType.Value = "B" Then
       frm1.btnGlSel.value = "��ǥ��ȸ"
    end if	
		
		
	if frm1.vspdData.MaxRows > 0 then	
		frm1.btnPosting.Disabled = False
		frm1.vspdData.focus
	else
		Call SetToolbar("11101001000111")
		frm1.btnPosting.Disabled = True
		frm1.txtBLNo.Focus
	End if


End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()							
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()						
End Function
	
	
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_00%>>
		<TR>
			<TD <%=HEIGHT_TYPE_00%>></TD>
		</TR>
		<TR HEIGHT=23>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_10%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L����</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenPODtlRef">���ֳ�������</A> | <A href="vbscript:OpenLCDtlRef">L/C��������</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR HEIGHT=*>
			<TD WIDTH=100% CLASS="Tab11">
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=29 MAXLENGTH=18 TAG="12XXXU" ALT="B/L ������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnBLNoOnClick()"></TD>
										<TD CLASS=TD6>&nbsp;</TD>
										<TD CLASS=TD6>&nbsp;</TD>
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU">
														 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>B/L������</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=B/L������ NAME="txtIssueDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="24" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="��ȭ"></TD>
												<TD>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" CLASS=FPDS140 tag="24X2" ALT="B/L�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>	
											</TR>
										</TABLE>
									</TD>														 
									<TD CLASS=TD5 NOWRAP>���ޱݿ���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"></TD>
								
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
	    <TR>
	      <TD <%=HEIGHT_TYPE_01%>></TD>
	    </TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
					    <BUTTON NAME="btnPosting" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
					    <BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">��ǥ��ȸ</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadBlHdr()">B/L���</a> | <A href="vbscript:LoadChargeHdr()">�����</A> | <A href="vbscript:LoadIvPayment()">���޳������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
			</TD>
		</TR>
	</TABLE>
<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHBLNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnLcNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPost" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
