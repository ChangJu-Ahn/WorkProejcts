
Option Explicit                               

Const BIZ_PGM_ID = "s4112mb1_temp.asp"            '��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID = "s4111ma1_KO441"            '��: JUMP�� �����Ͻ� ���� ASP�� 

Const btnClick = "btnClick"              '��:��ưŬ���� ���ڰ� 

'��: Spread Sheet�� Column�� ��� 
Dim C_ItemCd		'ǰ�� 
Dim C_ItemNm		'ǰ��� 
Dim C_Spec			'ǰ��԰� 
Dim C_TrackingNo    'Tracking No
Dim C_DnUnit		'���� 
Dim C_DnQty			'����û���� 
Dim C_DnBonusQty    '����û������ 
Dim C_PickQty       'Picking���� 
Dim C_PickBonusQty  'Picking������ 
Dim C_LotNo			'LOT No
Dim C_LotNoPopup    'LOT NoPopup
Dim C_LotSeq		'LOT No ���� 
Dim C_OnStkQty		'��� 
Dim C_BasicUnit		'������ 
Dim C_CartonNo		'Carton No

Dim C_GiAmt			'���ݾ� 
Dim C_GiAmtLoc      '���(�ڱ�)�ݾ� 
Dim C_DepositAmt    '�����ݾ� 
Dim C_VatAmt		'�ΰ����ݾ� 
Dim C_VatAmtLoc     '�ΰ���(�ڱ�)�ݾ� 

Dim C_QMItemFlag  
Dim C_QmFlag		'�˻籸�� 
Dim C_QmNoPopup  

Dim C_PlantCd       '���� 
Dim C_PlantPopup    '����Popup
Dim C_SlCd			'â�� 
Dim C_SlCdPopup     'â��Popup
Dim C_TolMoreQty    '��������뷮(+)
Dim C_TolLessQty    '��������뷮(-)
Dim C_CIQty			'������� 
Dim C_SoNo			'���ֹ�ȣ 
Dim C_SoSeq			'���ּ��� 
Dim C_SoSchdNo		'��ǰ���� 
Dim C_LcNo			'L/C��ȣ 
Dim C_LcSeq			'L/C���� 
Dim C_RetType		'��ǰ���� 
Dim C_RetTypeNm     '��ǰ������ 
Dim C_Remark		'��� 
Dim C_LotReqmtFlag  'Lot��ǰ���� 
Dim C_LotFlag		'Lot�������� 
Dim C_DnSeq			'���ϼ��� 
Dim C_RelBillNo
Dim C_RelBillCnt
Dim C_DnReqNo       '���Ͽ�û��ȣ 
Dim C_DnReqSeq      '���Ͽ�û���� 
Dim C_RCPT_LOT_NO		'�԰� Lot No.	
Dim C_CUST_LOT_NO		'�� Lot No.	
Dim C_OUT_NO
Dim C_TRANS_TIME
Dim C_OUT_TYPE_SUB
Dim C_CREATE_TYPE
Dim C_REF_GUBUN
'2008-06-16 6:13���� :: hanc
Dim C_pgm_name      
Dim C_pgm_price     

Const	C_REF2_ITEM_CD	=	0
Const	C_REF2_ITEM_NM	=	1
Const	C_REF2_ITEM_ACCT	=	2
Const	C_REF2_SPEC	=	3
Const	C_REF2_PLANT_CD	=	4
Const	C_REF2_PLANT_NM	=	5
Const	C_REF2_SL_CD	=	6
Const	C_REF2_SL_NM	=	7
Const	C_REF2_DN_REQ_SEQ	=	8
Const	C_REF2_REQ_QTY	=	9
Const	C_REF2_REQ_BONUS_QTY	=	10
Const	C_REF2_UNISSUED_QTY	=	11
Const	C_REF2_GI_QTY	=	12
Const	C_REF2_GI_BONUS_QTY	=	13
Const	C_REF2_GI_UNIT	=	14
Const	C_REF2_POST_GI_FLAG	=	15
Const	C_REF2_TOL_MORE_QTY	=	16
Const	C_REF2_TOL_LESS_QTY	=	17
Const	C_REF2_LOT_NO	=	18
Const	C_REF2_LOT_SEQ	=	19
Const	C_REF2_CC_QTY	=	20
Const	C_REF2_REMARK	=	21
Const	C_REF2_TRACKING_NO	=	22
Const	C_REF2_GI_AMT_LOC	=	23
Const	C_REF2_QM_FLAG	=	24
Const	C_REF2_VAT_AMT_LOC	=	25
Const	C_REF2_VAT_AMT	=	26
Const	C_REF2_GI_AMT	=	27
Const	C_REF2_EXT1_QTY	=	28
Const	C_REF2_EXT2_QTY	=	29
Const	C_REF2_EXT1_AMT	=	30
Const	C_REF2_EXT2_AMT	=	31
Const	C_REF2_EXT1_CD	=	32
Const	C_REF2_EXT2_CD	=	33
Const	C_REF2_EXT3_QTY	=	34
Const	C_REF2_EXT3_AMT	=	35
Const	C_REF2_EXT3_CD	=	36
Const	C_REF2_DEPOSIT_AMT	=	37
Const	C_REF2_PRICE	=	38
Const	C_REF2_VAT_RATE	=	39
Const	C_REF2_VAT_INC_FLAG	=	40
Const	C_REF2_VAT_TYPE	=	41
Const	C_REF2_DN_REQ_NO	=	42
Const	C_REF2_LC_NO	=	43
Const	C_REF2_LC_SEQ	=	44
Const	C_REF2_SO_NO	=	45
Const	C_REF2_SO_SEQ	=	46
Const	C_REF2_LC_FLAG	=	47
Const	C_REF2_RET_ITEM_FLAG	=	48
Const	C_REF2_SO_SCHD_NO	=	49
Const	C_REF2_LOT_FLG	=	50
Const	C_REF2_SHIP_INSPEC_FLG	=	51
Const	C_REF2_GOOD_ON_HAND_QTY	=	52
Const	C_REF2_RET_TYPE	=	53
Const	C_REF2_RET_TYPE_NM	=	54
Const	C_REF2_CARTON_NO	=	55
Const	C_REF2_REL_BILL_NO	=	56
Const	C_REF2_REL_BILL_CNT	=	57
Const	C_REF2_BASIC_UNIT	=	58



'=========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey
Dim lgLngStartRow

Dim IsOpenPop      ' Popup

'=========================================
Sub FormatField()
    ' ��¥ OCX Foramt ���� 
    Call FormatDATEField(frm1.txtActualGIDt)
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' ��¥ OCX
        Call LockObjectField(.txtActualGIDt, "P")

        If pvFlag = "N" Then
			Call LockHTMLField(.txtInvMgr, "P")	
			Call LockHTMLField(.chkArFlag, "P")	
			Call LockHTMLField(.chkVatFlag, "P")	
        End If
    End With

End Sub
'=========================================
Sub initSpreadPosVariables()
	C_ItemCd	    = 1    'ǰ�� 
	C_ItemNm		= 2    'ǰ��� 
	C_Spec			= 3    'ǰ��԰� 
	C_TrackingNo	= 4    'Tracking No
	C_DnUnit		= 5    '���� 
	C_DnQty			= 6    '����û���� 
	C_DnBonusQty	= 7    '����û������ 
	C_PickQty		= 8    'Picking���� 
	C_PickBonusQty  = 9    'Picking������ 
	C_LotNo			= 10    'LOT No
	C_LotNoPopup	= 11   'LOT NoPopup
	C_LotSeq		= 12   'LOT No ���� 
	C_OnStkQty		= 13   '��� 
	C_BasicUnit		= 14	' ������ 
	C_CartonNo		= 15
	
	C_GiAmt			= 16   '���ݾ� 
	C_GiAmtLoc		= 17   '���(�ڱ�)�ݾ� 
	C_DepositAmt	= 18   '�����ݾ� 
	C_VatAmt		= 19   '�ΰ����ݾ� 
	C_VatAmtLoc		= 20   '�ΰ���(�ڱ�)�ݾ� 

	C_QMItemFlag	= 21 
	C_QmFlag		= 22   '�˻籸�� 
	C_QmNoPopup		= 23

	C_PlantCd		= 24   '���� 
	C_PlantPopup	= 25   '����Popup
	C_SlCd			= 26   'â�� 
	C_SlCdPopup		= 27   'â��Popup
	C_TolMoreQty	= 28   '��������뷮(+)
	C_TolLessQty	= 29   '��������뷮(-)
	C_CIQty			= 30   '������� 
	C_SoNo			= 31   '���ֹ�ȣ 
	C_SoSeq			= 32   '���ּ��� 
	C_SoSchdNo		= 33   '��ǰ���� 
	C_LcNo			= 34   'L/C��ȣ 
	C_LcSeq			= 35   'L/C���� 
	C_RetType		= 36   '��ǰ���� 
	C_RetTypeNm		= 37   '��ǰ������ 
	C_Remark		= 38   '��� 
	C_LotReqmtFlag  = 39   'Lot��ǰ���� 
	C_LotFlag		= 40   'Lot�������� 
	C_DnSeq			= 41   '���ϼ��� 
	C_RelBillNo     = 42
	C_RelBillCnt    = 43
	C_DnReqNo       = 44   '���Ͽ�û��ȣ 
	C_DnReqSeq      = 45   '���Ͽ�û���� 
	C_RCPT_LOT_NO		= 46
	C_CUST_LOT_NO		= 47
	C_OUT_NO				= 48
	C_TRANS_TIME		= 49
	C_OUT_TYPE_SUB	= 50
	C_CREATE_TYPE		= 51
	C_REF_GUBUN 		= 52
	'2008-06-16 7:52���� :: hanc
    C_pgm_name          = 53
    C_pgm_price         = 54

End Sub

'=========================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'=========================================
Sub SetDefaultVal()

	frm1.txtConDnNo.focus
	frm1.btnPosting.disabled = True
	frm1.btnPostCancel.disabled = True
	frm1.btnPosting.value = "���ó��"
	frm1.btnPostCancel.value = "���ó�����"  
	 
	lgBlnFlgChgValue = False
	frm1.chkARflag.checked = False
	frm1.chkVatFlag.checked = False
	Call chkVatFlag_OnClick()

    frm1.txtSumPicking.value    =   CStr(0)
    
End Sub

'=========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	With ggoSpread

		.Source = frm1.vspdData
		.Spreadinit "V20030902",,parent.gAllowDragDropSpread    
		frm1.vspdData.ReDraw = false

		frm1.vspdData.MaxCols = C_pgm_price + 1                  '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		frm1.vspdData.Col = frm1.vspdData.MaxCols               '��: ������Ʈ�� ��� Hidden Column
		frm1.vspdData.ColHidden = True

		frm1.vspdData.MaxRows = 0

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("7","5","0")

		.SSSetFloat C_DnSeq,"���ϼ���" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetEdit C_ItemCd, "ǰ��", 18,,,18,2
		.SSSetEdit C_ItemNm, "ǰ���", 25
		.SSSetEdit C_Spec, "�԰�", 30
		.SSSetEdit C_TrackingNo, "Tracking No", 18,,,25,2
		.SSSetFloat C_DnQty,"���Ͽ�û����" ,15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetEdit C_DnUnit, "����", 8,2,,5,2
		.SSSetFloat C_DnBonusQty,"������" ,15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetFloat C_PickQty,"PICKING����" ,15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetFloat C_PickBonusQty,"��PICKING����" ,15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"    

		'���ݾ� 
		.SSSetFloat C_GiAmt,"���ݾ�",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		'���(�ڱ�)�ݾ� 
		.SSSetFloat C_GiAmtLoc,"����ڱ��ݾ�",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		'�����ݾ� 
		.SSSetFloat C_DepositAmt,"�����ݾ�",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		'�ΰ����ݾ� 
		.SSSetFloat C_VatAmt,"VAT�ݾ�",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		'�ΰ���(�ڱ�)�ݾ� 
		.SSSetFloat C_VatAmtLoc,"VAT�ڱ��ݾ�",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		  
		'�˻籸�� 
		.SSSetEdit C_QMItemFlag, "�˻�ǰ����", 10
		.SSSetEdit C_QmFlag, "�˻籸��", 15
		.SSSetButton C_QmNoPopup

		.SSSetEdit C_PlantCd, "����", 8,,,4,2     
		.SSSetButton C_PlantPopup
		.SSSetEdit C_SlCd, "â��", 8,,,7,2     
		.SSSetButton C_SlCdPopup
		
		.SSSetFloat C_TolMoreQty,"��������뷮(+)" ,15,parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetFloat C_TolLessQty,"��������뷮(-)" ,15,parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetEdit C_LotNo, "LOT NO", 12,,,25,2
		.SSSetButton C_LotNoPopup

		.SSSetFloat C_LotSeq,"LOT NO ����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetFloat C_OnStkQty,"���" ,15,parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetEdit C_BasicUnit, "������", 10,2,,5,2
		.SSSetEdit C_CartonNo, "Carton No", 15,,,10,2
		.SSSetFloat C_CIQty,"�������" ,15,parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
		.SSSetEdit C_SoNo, "[���ֹ�ȣ]", 18,,,18,2
		.SSSetFloat C_SoSeq,"���ּ���" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetFloat C_SoSchdNo, "��ǰ����", 15,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetEdit C_LcNo, "L/C��ȣ", 18
		.SSSetFloat C_LcSeq,"L/C����" ,12,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		.SSSetEdit C_Remark, "���", 60,,,120
		.SSSetEdit C_LotReqmtFlag, "LOT��ǰ����", 1
		.SSSetEdit C_LotFlag, "LOT��������", 1
		.SSSetEdit C_RetType, "��ǰ����", 10
		.SSSetEdit C_RetTypeNm, "��ǰ������", 20
		.SSSetEdit C_RelBillNo, "RelBillNo", 20
		.SSSetEdit C_RelBillCnt, "RelBillCnt", 20

		.SSSetEdit C_DnReqNo, "���Ͽ�û��ȣ", 18
		.SSSetFloat C_DnReqSeq,"���Ͽ�û����" ,12,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
		
		.SSSetEdit C_RCPT_LOT_NO, "�԰� Lot No.", 18
		.SSSetEdit C_CUST_LOT_NO, "�� Lot No.", 18
		
		.SSSetEdit C_OUT_NO, "C_OUT_NO", 18
		.SSSetEdit C_TRANS_TIME, "C_TRANS_TIME", 18
		.SSSetEdit C_OUT_TYPE_SUB, "C_OUT_TYPE_SUB", 18
		.SSSetEdit C_CREATE_TYPE, "C_CREATE_TYPE", 18
		.SSSetEdit C_REF_GUBUN, "C_REF_GUBUN", 18

		.SSSetEdit C_pgm_name, "PGM NAME", 50
		.SSSetFloat C_pgm_price,"PGM ����ܰ�" ,15,parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"


		call .MakePairsColumn(C_LotNo,C_LotNoPopup)
		call .MakePairsColumn(C_QmFlag,C_QmNoPopup)
		call .MakePairsColumn(C_SlCd,C_SlCdPopup)

		Call ggoSpread.SSSetColHidden(C_DnSeq,C_DnSeq,True)
		Call .SSSetColHidden(C_PlantCd,C_PlantPopup,True)
		Call .SSSetColHidden(C_LotReqmtFlag,C_LotReqmtFlag,True)
		Call .SSSetColHidden(C_LotFlag,C_LotFlag,True)
		Call .SSSetColHidden(C_GiAmt,C_GiAmt,True)
		Call .SSSetColHidden(C_VatAmt,C_VatAmt,True)
		Call .SSSetColHidden(C_VatAmtLoc,C_VatAmtLoc,True)
		Call .SSSetColHidden(C_RelBillNo,C_RelBillNo,True)
		Call .SSSetColHidden(C_RelBillCnt,C_RelBillCnt,True)
		Call .SSSetColHidden(C_Remark,C_Remark,True)

		Call .SSSetColHidden(C_OUT_NO,C_OUT_NO,True)
		Call .SSSetColHidden(C_TRANS_TIME,C_TRANS_TIME,True)
		Call .SSSetColHidden(C_OUT_TYPE_SUB,C_OUT_TYPE_SUB,True)
		Call .SSSetColHidden(C_CREATE_TYPE,C_CREATE_TYPE,True)
		Call .SSSetColHidden(C_REF_GUBUN,C_REF_GUBUN,True)
		
		frm1.vspdData.ReDraw = true
  
    End With
    
End Sub

'=========================================
Sub SetSpreadLock()
End Sub

'=========================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	Dim iRow

    With ggoSpread
		.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
		.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
		.SSSetProtected C_Spec, pvStartRow, pvEndRow
		.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow        
		.SSSetRequired C_DnQty, pvStartRow, pvEndRow
		.SSSetProtected C_DnUnit, pvStartRow, pvEndRow
'		.SSSetRequired C_DnBonusQty, pvStartRow, pvEndRow
		.SSSetProtected C_OnStkQty, pvStartRow, pvEndRow
		.SSSetProtected C_BasicUnit, pvStartRow, pvEndRow

		.SSSetRequired C_PickQty, pvStartRow, pvEndRow

		.SSSetProtected C_GiAmt, pvStartRow, pvEndRow
		.SSSetProtected C_GiAmtLoc, pvStartRow, pvEndRow
		.SSSetProtected C_VatAmt, pvStartRow, pvEndRow
		.SSSetProtected C_VatAmtLoc, pvStartRow, pvEndRow
		.SSSetProtected C_DepositAmt, pvStartRow, pvEndRow
		.SSSetProtected C_QMItemFlag, pvStartRow, pvEndRow
		.SSSetProtected C_QmFlag, pvStartRow, pvEndRow

		.SSSetProtected C_PlantCd, pvStartRow, pvEndRow
		.SSSetRequired  C_SlCd, pvStartRow, pvEndRow
		.SSSetProtected C_CIQty, pvStartRow, pvEndRow
		.SSSetProtected C_SoNo, pvStartRow, pvEndRow
		.SSSetProtected C_SoSeq, pvStartRow, pvEndRow
		.SSSetProtected C_SoSchdNo, pvStartRow, pvEndRow
		.SSSetProtected C_LcNo, pvStartRow, pvEndRow
		.SSSetProtected C_LcSeq, pvStartRow, pvEndRow

		.SSSetProtected C_TolMoreQty, pvStartRow, pvEndRow
		.SSSetProtected C_TolLessQty, pvStartRow, pvEndRow
			  
		.SSSetProtected C_RetType, pvStartRow, pvEndRow
		.SSSetProtected C_RetTypeNm, pvStartRow, pvEndRow

		.SSSetProtected C_DnReqNo, pvStartRow, pvEndRow
		.SSSetProtected C_DnReqSeq, pvStartRow, pvEndRow
		.SSSetProtected C_RCPT_LOT_NO, pvStartRow, pvEndRow
		.SSSetProtected C_CUST_LOT_NO, pvStartRow, pvEndRow
		
		' ��ǰ�� ���� Lot ��ȣ�� ������ �� ���� 
'		If Trim(frm1.txtHRetFlag.value) = "Y" Then   
'			frm1.vspdData.Col = C_RetType	: frm1.vspdData.ColHidden = False
'			frm1.vspdData.Col = C_RetTypeNm	: frm1.vspdData.ColHidden = False
'			.SSSetProtected C_LotNo, pvStartRow, pvEndRow
'			.SSSetProtected C_LotSeq, pvStartRow, pvEndRow
'			.SpreadLock C_LotNoPopup, pvStartRow, C_LotNoPopup, pvEndRow
'		Else
			frm1.vspdData.Col = C_RetType	: frm1.vspdData.ColHidden = True
			frm1.vspdData.Col = C_RetTypeNm	: frm1.vspdData.ColHidden = True

			For iRow = pvStartRow To pvEndRow
				frm1.vspdData.Row = iRow	:	frm1.vspdData.Col = C_LotFlag
				' Lot ���� ǰ�� ��� Lot ���� ���� ���� 
				If frm1.vspdData.Text = "Y" Then
					.SpreadUnLock C_LotNo, iRow, C_LotNo, iRow
					.SpreadUnLock C_LotSeq, iRow, C_LotSeq, iRow
					.SSSetRequired C_LotNo, iRow, iRow
					.SSSetRequired C_LotSeq, iRow, iRow
					.SpreadUnLock C_LotNoPopup, iRow, C_LotNoPopup, iRow
				Else
					.SpreadLock C_LotNo, iRow, C_LotNo, iRow
					.SpreadLock C_LotSeq, iRow, C_LotSeq, iRow
					.SSSetProtected C_LotNo, iRow, iRow
					.SSSetProtected C_LotSeq, iRow, iRow
					.SpreadLock C_LotNoPopup, iRow, C_LotNoPopup, iRow
				End If
			Next
'		End If
    End With
End Sub

'========================================
Function OpenConDnDtl()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
	   
	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4111PA1")
			
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111PA1", "x")
		IsOpenPop = False
		exit Function
	end if
	  
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtConDnNo.focus
	  
	If strRet <> "" Then
		frm1.txtConDnNo.value = strRet 
	End If 
 
End Function

'========================================
' ����� 
'=========================================
Sub OpenInvMgrPopUp()

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop Then Exit Sub

	With frm1
		If .txtInvMgr.readOnly Then	Exit Sub

		IsOpenPop = True

		iArrParam(1) = "dbo.B_MINOR"
		iArrParam(2) = Trim(.txtInvMgr.value)
		iArrParam(3) = ""											
		iArrParam(4) = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & ""
				
		iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
		iArrHeader(0) = .txtInvMgr.alt						
		iArrHeader(1) = .txtInvMgrNm.alt						

		.txtInvMgr.focus
	End With
	
	iArrParam(0) = iArrHeader(0)							' �˾� Title
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		frm1.txtInvMgr.value = iArrRet(0)
		frm1.txtInvMgrNm.value = iArrRet(1)
	End If	
End Sub

'========================================
Function OpenLotNoPopup(Byval iWhere)
 Dim iCalledAspName
 Dim arrRet
 Dim Param1
 Dim Param2
 Dim Param3
 Dim Param4
 Dim Param5
 Dim Param6, Param7, Param8, Param9

 Dim lgLcNo, lgLcSeq, lgItemCd

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True
 
 With frm1

  .vspdData.Row = iWhere

  .vspdData.Col = C_LotNo : lgLcNo = Trim(.vspdData.text)
  .vspdData.Col = C_LcSeq : lgLcSeq = Trim(.vspdData.text)
  .vspdData.Col = C_ItemCd : lgItemCd = Trim(.vspdData.text)

  .vspdData.Col = C_LotReqmtFlag
  If Trim(.vspdData.text) = "Y" Then        '���� config���� ret_item_falg�� Y(��ǰ)�̸� 

   Dim arrParam(5), arrField(6), arrHeader(6)

   arrParam(0) = "��ǰ LOT NO"       
   arrParam(1) = "S_DN_HDR DNHDR, S_DN_DTL DNDTL, " _
       & "S_SO_TYPE_CONFIG SOTYPE"    
   arrParam(2) = lgLcNo         
   arrParam(3) = lgLcSeq         ' Name Condition
   arrParam(4) = "DNHDR.DN_NO = DNDTL.DN_NO " _
       & "AND DNHDR.SO_TYPE = SOTYPE.SO_TYPE " _
       & "AND SOTYPE.RET_ITEM_FLAG = " & FilterVar("N", "''", "S") & "  " _
       & "AND DNHDR.POST_FLAG = " & FilterVar("Y", "''", "S") & "  " _
       & "AND DNHDR.SHIP_TO_PARTY =  " & FilterVar(.txtShipToParty.value, "''", "S") & " " _
       & "AND DNDTL.ITEM_CD =  " & FilterVar(lgItemCd , "''", "S") & "" 
   arrParam(5) = "��ǰ LOT NO"       

   arrField(0) = "DNDTL.LOT_NO"       
   arrField(1) = "ED" & parent.gColSep & "DNDTL.LOT_SEQ"
   arrField(2) = "DD" & parent.gColSep & "DNHDR.ACTUAL_GI_DT"
   arrField(3) = "DNHDR.DN_NO"        
   arrField(4) = "ED" & parent.gColSep & "DNDTL.DN_SEQ"
    
   arrHeader(0) = "LOT NO"        
   arrHeader(1) = "LOT SEQ"       
   arrHeader(2) = "��������"       ' Header��(2)
   arrHeader(3) = "���Ϲ�ȣ"       ' Header��(3)
   arrHeader(4) = "���ϼ���"       ' Header��(4)

   arrRet = window.showModalDialog("../../comasp/commonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False

   If Trim(arrRet(0)) <> "" Then
    .vspdData.Col = C_LotNo : .vspdData.Text = arrRet(0)
    .vspdData.Col = C_LotSeq : .vspdData.Text = arrRet(1)
    Call vspdData_Change(.vspdData.Col, .vspdData.Row)   ' ������ �о�ٰ� �˷��� 
    lgBlnFlgChgValue = True
   End If

  Else

   .vspdData.Col = C_SlCd
   Param1 = .vspdData.text
   .vspdData.Col = C_ItemCd
   Param2 = .vspdData.text
   .vspdData.Col = C_TrackingNo
   Param3 = .vspdData.text
   .vspdData.Col = C_PlantCd
   Param4 = .vspdData.text

   Param5 = "J"

   .vspdData.Col = C_LotNo
   Param6 = .vspdData.text

   Param7 = ""

   .vspdData.Col = C_ItemNm
   Param8 = .vspdData.text
   
   .vspdData.Col = C_DnUnit
   Param9 = .vspdData.text

	iCalledAspName = AskPRAspName("I2212RA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "I2212RA1", "x")
		gblnWinEvent = False
		exit Function
	end if

   arrRet = window.showModalDialog(iCalledAspName, Array(window.parent , Param1, Param2,Param3,Param4,Param5,Param6,Param7,Param8, Param9), _
    "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
   IsOpenPop = False

   If Trim(arrRet(0)) <> "" Then
    .vspdData.Col = C_LotNo : .vspdData.Text = arrRet(3)
    .vspdData.Col = C_LotSeq : .vspdData.Text = arrRet(4)

    Dim lsDnQty,lsDnBonusQty, lsPickQty,lsPickBonusQty, lsTotDnQty, lsTotPickQty, lsAvlQty

    .vspdData.Col = C_DnQty : lsDnQty = UNICDbl(Trim(.vspdData.text))
    .vspdData.Col = C_DnBonusQty : lsDnBonusQty = UNICDbl(Trim(.vspdData.text))
    .vspdData.Col = C_PickQty : lsPickQty = UNICDbl(Trim(.vspdData.text))
    .vspdData.Col = C_PickBonusQty : lsPickBonusQty = UNICDbl(Trim(.vspdData.text))

'    lsTotDnQty = @@@UNICDbl(lsDnQty) + @@@UNICDbl(lsDnBonusQty)
    lsTotPickQty = UNIFormatNumber(UNICDbl(lsPickQty) + UNICDbl(lsPickBonusQty), ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)

    lsAvlQty = UNICDbl(arrRet(5))
    
    If lsAvlQty >= uniCDbl(lsTotPickQty) Then
     '.vspdData.Col = C_PickQty : .vspdData.Text = lsPickQty
     '.vspdData.Col = C_PickBonusQty : .vspdData.Text = lsPickBonusQty
    ElseIf lsAvlQty < uniCDbl(lsTotPickQty) Then
     If lsAvlQty <= lsPickQty Then
      .vspdData.Col = C_PickQty : .vspdData.Text = lsAvlQty
      .vspdData.Col = C_PickBonusQty : .vspdData.Text = 0
     ElseIf lsAvlQty > lsPickQty Then
      .vspdData.Col = C_PickQty : .vspdData.Text = lsPickQty
      .vspdData.Col = C_PickBonusQty : .vspdData.Text = UNIFormatNumber(UNICDbl(lsAvlQty) - UNICDbl(lsPickQty),  ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
     End If
    End If

    Call vspdData_Change(.vspdData.Col, .vspdData.Row)
    lgBlnFlgChgValue = True
   End If

  End If

 End With
 
End Function

'========================================
Function OpenDnDtl(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrTemp(2)

	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere  
		Case 1 '���� 
			arrParam(1) = "b_plant plant, b_item_by_plant item_plant" 
			arrParam(2) = strCode        
			arrParam(3) = ""         
			arrParam(4) = "plant.plant_cd=item_plant.plant_cd" 
			arrParam(5) = "����"       
			 
			arrField(0) = "plant.plant_cd"      
			arrField(1) = "plant.plant_nm"      
				   
			arrHeader(0) = "����"       
			arrHeader(1) = "�����"       

		Case 2 'â�� 
			Dim strValue
				 
			strValue = Split(strCode,gColSep)

			arrParam(1) = "B_STORAGE_LOCATION"     
			arrParam(2) = strValue(0)       
			arrParam(3) = ""         

			If strValue(1) <> "" Then
				arrParam(4) = "PLANT_CD =" + FilterVar(strValue(1), " ", "S")  
			End If

			arrParam(5) = "â��"       
			 
			arrField(0) = "SL_CD"        
			arrField(1) = "SL_NM"        
				   
			arrHeader(0) = "â��"       
			arrHeader(1) = "â���"        
	End Select

	arrParam(0) = arrParam(5)        

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If Err.number <> 0 Then Err.Clear 
		Exit Function
	Else
		Call SetDnDtl(arrRet, iWhere)
	End If 
End Function

'========================================
Function OpenSODtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim strParam

	On Error Resume Next

	If Trim(frm1.txtPlannedGIDt.value) = "" Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	If Trim(frm1.txtHRefRoot.value) = "DR" Then
		Msgbox "���� ��û ������ ���� �����Դϴ�.",vbInformation, parent.gLogoName
		Exit Function
	End If
	
	If Len(Trim(frm1.txtGINo.value)) Then
		Msgbox "���ó���� ǰ���� ���ֳ����� ���� �� �� �����ϴ�",vbInformation, parent.gLogoName
		Exit Function
	End If

'	iCalledAspName = AskPRAspName("S3112AA1_ko441")     '20080219::hanc
	iCalledAspName = AskPRAspName("S3112AA1_temp")     '20080617::hanc

	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S3112AA1", "x")
		gblnWinEvent = False
		exit Function
	end if

	strParam =	Trim(frm1.txtSoNo.value) & parent.gColSep & _
				Trim(frm1.txtPlannedGIDt.value) & parent.gColSep & _
				Trim(frm1.txtDnType.Value) & parent.gColSep & _
				Trim(frm1.txtShipToParty.Value) & parent.gColSep & _
				Trim(frm1.txtShipToPartyNm.Value) & parent.gColSep & _
				Trim(frm1.txtSoType.Value) & parent.gColSep & _
				Trim(frm1.txtHRetFlag.Value) & parent.gColSep & _
				Trim(frm1.txtPlantCd.Value) & parent.gColSep & _  
				Trim(frm1.txtDnType.value)                '20080225::HANC::�������� �߰�

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strParam), _
	"dialogWidth=850px; dialogHeight=620px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0, 0) = "" Then
		If Err.number <> 0 Then Err.Clear 
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If 
End Function
 
Function OpenSODtlRef1()
	Dim iCalledAspName
	Dim arrRet
	Dim strParam

	On Error Resume Next

	If Trim(frm1.txtPlannedGIDt.value) = "" Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	If Trim(frm1.txtHRefRoot.value) = "DR" Then
		Msgbox "���� ��û ������ ���� �����Դϴ�.",vbInformation, parent.gLogoName
		Exit Function
	End If
	
	If Len(Trim(frm1.txtGINo.value)) Then
		Msgbox "���ó���� ǰ���� ���ֳ����� ���� �� �� �����ϴ�",vbInformation, parent.gLogoName
		Exit Function
	End If

	iCalledAspName = AskPRAspName("S3112AA1")     '20080219::hanc

	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S3112AA1", "x")
		gblnWinEvent = False
		exit Function
	end if

	strParam =	Trim(frm1.txtSoNo.value) & parent.gColSep & _
				Trim(frm1.txtPlannedGIDt.value) & parent.gColSep & _
				Trim(frm1.txtDnType.Value) & parent.gColSep & _
				Trim(frm1.txtShipToParty.Value) & parent.gColSep & _
				Trim(frm1.txtShipToPartyNm.Value) & parent.gColSep & _
				Trim(frm1.txtSoType.Value) & parent.gColSep & _
				Trim(frm1.txtHRetFlag.Value) & parent.gColSep & _
				Trim(frm1.txtPlantCd.Value) & parent.gColSep & _  
				Trim(frm1.txtDnType.value)                '20080225::HANC::�������� �߰�

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strParam), _
	"dialogWidth=850px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0, 0) = "" Then
		If Err.number <> 0 Then Err.Clear 
		Exit Function
	Else
		Call SetSODtlRef1(arrRet)
	End If 
End Function

'========================================
Function OpenDnReqDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim strParam

'	On Error Resume Next

	If Trim(frm1.txtPlannedGIDt.value) = "" Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	If Trim(frm1.txtHRefRoot.value) = "SO" Then
		Msgbox "���� ������ ���� �����Դϴ�.",vbInformation, parent.gLogoName
		Exit Function
	End If

	If Len(Trim(frm1.txtGINo.value)) Then
		Msgbox "���ó���� ǰ���� ���Ͽ�û�� ���� �� �� �����ϴ�",vbInformation, parent.gLogoName
		Exit Function
	End If


	iCalledAspName = AskPRAspName("S4512RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4512RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If

	strParam =	Trim(frm1.txtSoNo.value) & parent.gColSep & _
				Trim(frm1.txtShipToParty.Value) & parent.gColSep & _
				Trim(frm1.txtShipToPartyNm.Value) & parent.gColSep & _
				Trim(frm1.txtPlantCd.Value)
				
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,strParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	' Popup���� Cancel�� ��� 
	If UBOUND(arrRet, 1) = 0 Then	
		If Err.Number <> 0 Then	Err.Clear 
	Else
		Call SetDNReqDtlRef(arrRet)
	End if
End Function


'========================================
Function OpenQMDtlRef(Row)
	Dim iCalledAspName
	Dim strRet
	Dim arrValue(2)
	Dim ItemCode
	Dim DnSeq

	On Error Resume Next

	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Function

	If Len(Trim(frm1.txtGINo.value)) Then
		Exit Function
	End If

	frm1.vspdData.Row = Row

	frm1.vspdData.Col = C_QMItemFlag
	  
	If frm1.vspdData.text = "N" Then 
		Call DisplayMsgBox("220731", "X", "X", "X")
		Exit Function
	End If
	   
	arrValue(0) = Trim(frm1.txtConDnNo.value)

	frm1.vspdData.Col = C_DnSeq
	arrValue(1) = frm1.vspdData.text
	  
	frm1.vspdData.Col = C_ItemCd
	arrValue(2) = frm1.vspdData.text

	iCalledAspName = AskPRAspName("S4112RA9")
			
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4112RA9", "x")
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrValue), _
	"dialogWidth=780px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet = "" Then
		If Err.number <> 0 Then Err.Clear 
	End If 
End Function
	 
'========================================
Function SetDnDtl(Byval arrRet,ByVal iWhere)

	With frm1

	Select Case iWhere
		Case 1 '���� 
			.vspdData.Col = C_PlantCd
			.vspdData.Text = arrRet(0)

		Case 2 'â�� 
			.vspdData.Col = C_SlCd
			.vspdData.Text = arrRet(0)
	   
	End Select
	  
	Call vspdData_Change(.vspdData.Col, .vspdData.Row)   ' ������ �Ͼ�ٰ� �˷��� 

	End With

	lgBlnFlgChgValue = True
 
End Function

'========================================
Function SetSODtlRef(pvArrRet)
On Error Resume Next

	Dim iStrVal, txtSpread
	Dim iLngStartRow, iLngLoopCnt, iLngCnt
	Dim iPickingQty
		    
Const C_ACTUAL_GI_DT_REF = 0
Const C_ITEM_CD_REF = 1
Const C_ITEM_NM_REF = 2
Const C_GI_QTY_REF = 3
Const C_BonusQty_REF = 4
Const C_GI_UNIT_REF = 5
Const C_OnStkQty_REF = 6
Const C_BasicUnit_REF = 7
Const C_SoNo_REF = 8
Const C_SoSeq_REF = 9
Const C_SoSchdNo_REF = 10
Const C_TrackingNo_REF = 11
Const C_SHIP_TO_PARTY_REF = 12
Const C_SHIP_TO_PARTY_NM_REF = 13
Const C_PLANT_CD_REF = 14
Const C_PlantNm_REF = 15
Const C_SlCd_REF = 16
Const C_SlNm_REF = 17
Const C_TolMoreQty_REF = 18
Const C_TolLessQty_REF = 19
Const C_LcNo_REF = 20
Const C_LcSeq_REF = 21
Const C_LotFlag_REF = 22
Const C_LOT_NO_REF = 23
Const C_LOT_SEQ_REF = 24
Const C_RetItemFlag_REF = 25
Const C_RetType_REF = 26
Const C_RetTypeNm_REF = 27
Const C_SPEC_REF = 28
Const C_Remark_REF = 29
Const C_OUT_NO_REF = 30
Const C_TRANS_TIME_REF = 31
Const C_CREATE_TYPE_REF = 32
Const C_CUSTOM_LOT_NO_REF = 33
Const C_RCPT_LOT_NO_REF = 34
'2008-06-16 7:30���� :: hanc
Const C_pgm_name_REF = 35
Const C_pgm_price_REF = 36

    iPickingQty =   0

With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData   
		.ReDraw = False 

		iLngStartRow = .MaxRows + 1            '��: ��������� MaxRows 
		iLngLoopCnt = Ubound(pvArrRet, 1)           '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� 

		For iLngCnt = 0 to iLngLoopCnt - 1
			'If UCase(Trim(pvArrRet(iLngCnt, 22))) = "Y" Then
			'	txtSpread = txtSpread & pvArrRet(iLngCnt, 8) & Chr(11)
			'	txtSpread = txtSpread & pvArrRet(iLngCnt, 9) & Chr(12)
			'Else
				.MaxRows = .MaxRows + 1
				.Row = .MaxRows
	

				.Col = 0							:	.Text = ggoSpread.InsertFlag
				.Col = C_SoNo					:	.text = pvArrRet(iLngCnt, C_SoNo_REF)			'���ֹ�ȣ 
				.Col = C_SoSeq      	:	.text = pvArrRet(iLngCnt, C_SoSeq_REF)			'���ּ��� 
				.Col = C_SoSchdNo   	:	.text = pvArrRet(iLngCnt, C_SoSchdNo_REF)			'������������ 
				.Col = C_ItemCd				:	.text = pvArrRet(iLngCnt, C_ITEM_CD_REF)			'ǰ�� 
				.Col = C_ItemNm				:	.text = pvArrRet(iLngCnt, C_ITEM_NM_REF)			'ǰ��� 
				.Col = C_Spec					:	.text = pvArrRet(iLngCnt, C_SPEC_REF)			'�԰� 
				.Col = C_TrackingNo 	:	.text = pvArrRet(iLngCnt, C_TrackingNo_REF)			'Tracking No
				.Col = C_DnUnit				:	.text = pvArrRet(iLngCnt, C_GI_UNIT_REF)			'���� 
				.Col = C_DnQty				:	.text = pvArrRet(iLngCnt, C_GI_QTY_REF)			'�������� 
				.Col = C_DnBonusQty 	:	.text = pvArrRet(iLngCnt, C_BonusQty_REF)			'���������� 
				.Col = C_OnStkQty			:	.text = pvArrRet(iLngCnt, C_OnStkQty_REF)			'��� 
				.Col = C_BasicUnit		:	.text = pvArrRet(iLngCnt, C_BasicUnit_REF)			'������ 
				.Col = C_PickQty			:	.text = pvArrRet(iLngCnt, C_GI_QTY_REF)			'Picking���� 
				.Col = C_PickBonusQty	:	.text = pvArrRet(iLngCnt, C_BonusQty_REF)			'Picking������ 
				.Col = C_PlantCd			:	.text = pvArrRet(iLngCnt, C_PLANT_CD_REF)			'�����ڵ� 
				.Col = C_SlCd					:	.text = pvArrRet(iLngCnt, C_SlCd_REF)			'â���ڵ� 
				.Col = C_TolMoreQty		:	.text = pvArrRet(iLngCnt, C_TolMoreQty_REF)			'��������뷮(+)
				.Col = C_TolLessQty		:	.text = pvArrRet(iLngCnt, C_TolLessQty_REF)			'��������뷮(-)
				.Col = C_LcNo					:	.text = pvArrRet(iLngCnt, C_LcNo_REF)			'L/C��ȣ 
				.Col = C_LcSeq				:	.text = pvArrRet(iLngCnt, C_LcSeq_REF)			'L/C���� 
				.Col = C_Remark				:	.text = pvArrRet(iLngCnt, C_Remark_REF)			'��� 
				.Col = C_LotReqmtFlag	:	.text = pvArrRet(iLngCnt, C_RetItemFlag_REF)			'��ǰ���� 
				.Col = C_RetType			:	.text = pvArrRet(iLngCnt, C_RetType_REF)     '��ǰ���� 
				.Col = C_RetTypeNm		:	.text = pvArrRet(iLngCnt, C_RetTypeNm_REF)     '��ǰ������ 
				.Col = C_LotFlag			:	.text = pvArrRet(iLngCnt, C_LotFlag_REF)			'Lot �������� 
				.Col = C_DnSeq				:	.Text = 0
				.Col = C_CIQty				:	.Text = 0
				.Col = C_OUT_NO				:	.text = pvArrRet(iLngCnt, C_OUT_NO_REF)			'
				.Col = C_TRANS_TIME		:	.text = pvArrRet(iLngCnt, C_TRANS_TIME_REF)			'
				.Col = C_CREATE_TYPE	:	.text = pvArrRet(iLngCnt, C_CREATE_TYPE_REF)			'
				.Col = C_RCPT_LOT_NO	:	.text = pvArrRet(iLngCnt, C_RCPT_LOT_NO_REF)			'
				.Col = C_CUST_LOT_NO	:	.text = pvArrRet(iLngCnt, C_CUSTOM_LOT_NO_REF)			'

				.Col = C_pgm_name	:	.text = pvArrRet(iLngCnt, C_pgm_name_REF)			'2008-06-16 7:32���� :: hanc
				.Col = C_pgm_price	:	.text = pvArrRet(iLngCnt, C_pgm_price_REF)			'2008-06-16 7:32���� :: hanc

				.Col = C_REF_GUBUN		:	.Text = "2"
				
				iPickingQty =   CDbl(iPickingQty) + CDbl(pvArrRet(iLngCnt, C_GI_QTY_REF))

				'====================================================================  
				' 02.23 SMJ
				' -- ��ǰ�� ��� ���������� Lot no, lot seq�� �����´�.   
				'====================================================================  
				If UCase(Trim(frm1.txtHRetFlag.value)) = "Y" Then
					.Col = C_LotNo		:		.Text = pvArrRet(iLngCnt, C_LOT_NO_REF)
					.Col = C_LotSeq		:		.Text = pvArrRet(iLngCnt, C_LOT_SEQ_REF)
				Else
					' Lot ���� ǰ�� �ƴ� ��� Lot��ȣ�� '*'�� ó���Ѵ�.
					' 20040813 SMJ lot_flag ��ġ�� �߸��� 23->22�� ���� 
					
					If UCase(Trim(pvArrRet(iLngCnt, C_LotFlag_REF))) = "Y" Then
						.Col = C_LotNo		:		.Text = pvArrRet(iLngCnt, C_LOT_NO_REF)
						.Col = C_LotSeq		:		.Text = pvArrRet(iLngCnt, C_LOT_SEQ_REF)
					Else
						.Col = C_LotNo		:		.Text = "*"
					End If
					.Col = C_LotSeq			:	.Text = 0
				End If
			'End If
		Next

		Call SetSpreadColor(iLngStartRow, .MaxRows)

		' Focus ó�� 
		Call SubSetErrPos(iLngStartRow)

		.ReDraw = True    

	End With

    'frm1.txtSumPicking.value    =   CStr(iPickingQty)
    
    Call SumPicking()

	'If Trim(txtSpread) <> "" Then
	'	iStrVal = BIZ_PGM_ID & "?txtMode=GetIssueFromMES"    
	'	iStrVal = iStrVal & "&txtSpread=" & txtSpread
	'	
	'	Call RunMyBizASP(MyBizASP, iStrVal)            
	'End If

	lgBlnFlgChgValue = True
End Function
 
Function SetSODtlRef1(pvArrRet)
On Error Resume Next

	Dim iLngStartRow, iLngLoopCnt, iLngCnt
		    
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData   
		.ReDraw = False 

		iLngStartRow = .MaxRows + 1            '��: ��������� MaxRows 
		iLngLoopCnt = Ubound(pvArrRet, 1)           '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� 

		For iLngCnt = 0 to iLngLoopCnt - 1
			.MaxRows = .MaxRows + 1
			.Row = .MaxRows

			.Col = 0			:		.Text = ggoSpread.InsertFlag
			.Col = C_SoNo		:		.text = pvArrRet(iLngCnt, 8)			'���ֹ�ȣ 
			.Col = C_SoSeq      :		.text = pvArrRet(iLngCnt, 9)			'���ּ��� 
			.Col = C_SoSchdNo   :		.text = pvArrRet(iLngCnt, 10)			'������������ 
			.Col = C_ItemCd		:		.text = pvArrRet(iLngCnt, 1)			'ǰ�� 
			.Col = C_ItemNm		:		.text = pvArrRet(iLngCnt, 2)			'ǰ��� 
			.Col = C_Spec		:		.text = pvArrRet(iLngCnt, 28)			'�԰� 
			.Col = C_TrackingNo :		.text = pvArrRet(iLngCnt, 11)			'Tracking No
			.Col = C_DnUnit		:		.text = pvArrRet(iLngCnt, 5)			'���� 
			.Col = C_DnQty		:		.text = pvArrRet(iLngCnt, 3)			'�������� 
			.Col = C_DnBonusQty :		.text = pvArrRet(iLngCnt, 4)			'���������� 
			.Col = C_OnStkQty	:		.text = pvArrRet(iLngCnt, 6)			'��� 
			.Col = C_BasicUnit	:		.text = pvArrRet(iLngCnt, 7)			'������ 
			.Col = C_PickQty	:		.text = pvArrRet(iLngCnt, 3)			'Picking���� 
			.Col = C_PickBonusQty	:	.text = pvArrRet(iLngCnt, 4)			'Picking������ 
			.Col = C_PlantCd		:	.text = pvArrRet(iLngCnt, 14)			'�����ڵ� 
			.Col = C_SlCd			:	.text = pvArrRet(iLngCnt, 16)			'â���ڵ� 
			.Col = C_TolMoreQty		:	.text = pvArrRet(iLngCnt, 18)			'��������뷮(+)
			.Col = C_TolLessQty		:	.text = pvArrRet(iLngCnt, 19)			'��������뷮(-)
			.Col = C_LcNo			:	.text = pvArrRet(iLngCnt, 20)			'L/C��ȣ 
			.Col = C_LcSeq			:	.text = pvArrRet(iLngCnt, 21)			'L/C���� 
			.Col = C_Remark			:	.text = pvArrRet(iLngCnt, 29)			'��� 
			.Col = C_LotReqmtFlag	:	.text = pvArrRet(iLngCnt, 25)			' ��ǰ���� 
			.Col = C_RetType		:	.text = pvArrRet(iLngCnt, 26)     		'��ǰ���� 
			.Col = C_RetTypeNm		:	.text = pvArrRet(iLngCnt, 27)     		'��ǰ������ 
			.Col = C_LotFlag		:	.text = pvArrRet(iLngCnt, 22)			'Lot �������� 
			.Col = C_DnSeq			:	.Text = 0
			.Col = C_CIQty			:	.Text = 0

			'====================================================================  
			' 02.23 SMJ
			' -- ��ǰ�� ��� ���������� Lot no, lot seq�� �����´�.   
			'====================================================================  
			If UCase(Trim(frm1.txtHRetFlag.value)) = "Y" Then
				.Col = C_LotNo		:		.Text = pvArrRet(iLngCnt, 23)
				.Col = C_LotSeq		:		.Text = pvArrRet(iLngCnt, 24)
			Else
				' Lot ���� ǰ�� �ƴ� ��� Lot��ȣ�� '*'�� ó���Ѵ�.
				' 20040813 SMJ lot_flag ��ġ�� �߸��� 23->22�� ���� 
				
				If UCase(Trim(pvArrRet(iLngCnt, 22))) = "Y" Then
					.Col = C_LotNo		:		.Text = ""
				Else
					.Col = C_LotNo		:		.Text = "*"
				End If
				.Col = C_LotSeq			:	.Text = 0
			End If
		Next

		Call SetSpreadColor(iLngStartRow, .MaxRows)

		' Focus ó�� 
		Call SubSetErrPos(iLngStartRow)

		.ReDraw = True    

	End With

	lgBlnFlgChgValue = True
End Function
 
'========================================
Function SetDNReqDtlRef(pvArrRet)
On Error Resume Next

	Dim iLngStartRow, iLngLoopCnt, iLngCnt
		    
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData   
		.ReDraw = False 

		iLngStartRow = .MaxRows + 1            '��: ��������� MaxRows 
		iLngLoopCnt = Ubound(pvArrRet, 1)           '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� 

		For iLngCnt = 0 to iLngLoopCnt - 1
			.MaxRows = .MaxRows + 1
			.Row = .MaxRows

			.Col = 0				:	.Text = ggoSpread.InsertFlag
			.Col = C_ItemCd			:	.text = pvArrRet(iLngCnt, C_REF2_ITEM_CD)			'ǰ�� 
			.Col = C_ItemNm			:	.text = pvArrRet(iLngCnt, C_REF2_ITEM_NM)			'ǰ��� 
			.Col = C_Spec			:	.text = pvArrRet(iLngCnt, C_REF2_SPEC)			'�԰� 
			.Col = C_TrackingNo		:	.text = pvArrRet(iLngCnt, C_REF2_TRACKING_NO)			'Tracking No
			.Col = C_DnUnit			:	.text = pvArrRet(iLngCnt, C_REF2_GI_UNIT)			'���� 
			.Col = C_DnQty			:	.text = pvArrRet(iLngCnt, C_REF2_REQ_QTY) -	pvArrRet(iLngCnt, C_REF2_GI_QTY)		'�������� 
			.CellTag = .text
			.Col = C_DnBonusQty		:	.text = pvArrRet(iLngCnt, C_REF2_REQ_BONUS_QTY) -  pvArrRet(iLngCnt, C_REF2_GI_BONUS_QTY)			'���������� 
			.Col = C_PickQty		:	.text = pvArrRet(iLngCnt, C_REF2_REQ_QTY) -	pvArrRet(iLngCnt, C_REF2_GI_QTY)			'Picking���� 
			.Col = C_PickBonusQty	:	.text = pvArrRet(iLngCnt, C_REF2_REQ_BONUS_QTY) -  pvArrRet(iLngCnt, C_REF2_GI_BONUS_QTY)			'Picking������ 
			.Col = C_LotNo			:	.Text = pvArrRet(iLngCnt, C_REF2_LOT_NO)
			.Col = C_LotSeq			:	.Text = pvArrRet(iLngCnt, C_REF2_LOT_SEQ)
			.Col = C_OnStkQty		:	.text = pvArrRet(iLngCnt, C_REF2_GOOD_ON_HAND_QTY)			'��� 
			.Col = C_BasicUnit		:	.text = pvArrRet(iLngCnt, C_REF2_BASIC_UNIT)			'������ 
			.Col = C_CartonNo		:	.Text = pvArrRet(iLngCnt, C_REF2_CARTON_NO)

			' �� �ݾ׵�..
'			.Col = C_GiAmt			:	.Text = pvArrRet(iLngCnt, C_REF2_GI_AMT)
'			.Col = C_GiAmtLoc		:	.Text = pvArrRet(iLngCnt, C_REF2_GI_AMT_LOC)
'			.Col = C_DepositAmt		:	.Text = pvArrRet(iLngCnt, C_REF2_DEPOSIT_AMT)
'			.Col = C_VatAmt			:	.Text = pvArrRet(iLngCnt, C_REF2_VAT_AMT)
'			.Col = C_VatAmtLoc		:	.Text = pvArrRet(iLngCnt, C_REF2_VAT_AMT_LOC)

'C_QMItemFlag
			.Col = C_QmFlag			:	.Text = pvArrRet(iLngCnt, C_REF2_QM_FLAG)

			.Col = C_PlantCd		:	.text = pvArrRet(iLngCnt, C_REF2_PLANT_CD)			'�����ڵ� 
			.Col = C_SlCd			:	.text = pvArrRet(iLngCnt, C_REF2_SL_CD)			'â���ڵ� 
			.Col = C_TolMoreQty		:	.text = pvArrRet(iLngCnt, C_REF2_TOL_MORE_QTY)			'��������뷮(+)
			.Col = C_TolLessQty		:	.text = pvArrRet(iLngCnt, C_REF2_TOL_LESS_QTY)			'��������뷮(-)
			.Col = C_CIQty			:	.Text = 0
			.Col = C_SoNo			:	.text = pvArrRet(iLngCnt, C_REF2_SO_NO)			'���ֹ�ȣ 
			.Col = C_SoSeq			:	.text = pvArrRet(iLngCnt, C_REF2_SO_SEQ)			'���ּ��� 
			.Col = C_SoSchdNo		:	.text = pvArrRet(iLngCnt, C_REF2_SO_SCHD_NO)			'������������ 
			.Col = C_LcNo			:	.text = pvArrRet(iLngCnt, C_REF2_LC_NO)			'L/C��ȣ 
			.Col = C_LcSeq			:	.text = pvArrRet(iLngCnt, C_REF2_LC_SEQ)			'L/C���� 
			.Col = C_RetType		:	.text = pvArrRet(iLngCnt, C_REF2_RET_TYPE)     		'��ǰ���� 
			.Col = C_RetTypeNm		:	.text = pvArrRet(iLngCnt, C_REF2_RET_TYPE_NM)     		'��ǰ������ 
			.Col = C_Remark			:	.text = pvArrRet(iLngCnt, C_REF2_REMARK)			'��� 
''''			.Col = C_LotReqmtFlag	:	.text = pvArrRet(iLngCnt, 25)			' ��ǰ���� 
			.Col = C_LotFlag		:	.text = pvArrRet(iLngCnt, C_REF2_LOT_FLG)			'Lot �������� 
			.Col = C_DnSeq			:	.Text = 0

			.Col = C_RelBillNo		:	.text = pvArrRet(iLngCnt, C_REF2_REL_BILL_NO)			
			.Col = C_RelBillCnt		:	.text = pvArrRet(iLngCnt, C_REF2_REL_BILL_CNT)			

			.Col = C_DnReqNo		:	.text = pvArrRet(iLngCnt, C_REF2_DN_REQ_NO)			'���Ͽ�û��ȣ 
			.Col = C_DnReqSeq		:	.text = pvArrRet(iLngCnt, C_REF2_DN_REQ_SEQ)			'���Ͽ�û���� 

		Next

		Call SetSpreadColor(iLngStartRow, .MaxRows)

		' Focus ó�� 
		Call SubSetErrPos(iLngStartRow)

		.ReDraw = True    

	End With

	lgBlnFlgChgValue = True
End Function
 
'=====================================================
Sub SetQuerySpreadColor(ByVal pvRow)
	On Error Resume Next
	Dim i, iMaxRows
  
	iMaxRows = frm1.vspdData.MaxRows
    With ggoSpread  
  
		frm1.vspdData.ReDraw = False

		'--- ���ó���� �Ǿ������� Ȯ���Ѵ�.
		If Trim(frm1.txtGINo.value) = "" Then
			'--- ���ó���� ���� ���� ��� 
			.SSSetProtected C_ItemCd, pvRow, iMaxRows
			.SSSetProtected C_ItemNm, pvRow, iMaxRows
			.SSSetProtected C_Spec, pvRow, iMaxRows
			.SSSetProtected C_TrackingNo, pvRow, iMaxRows        
			.SSSetProtected C_DnQty, pvRow, iMaxRows
			.SSSetProtected C_DnUnit, pvRow, iMaxRows
			.SSSetProtected C_OnStkQty, pvRow, iMaxRows
			.SSSetProtected C_BasicUnit, pvRow, iMaxRows
			.SSSetProtected C_DnBonusQty, pvRow, iMaxRows

			.SSSetRequired C_PickQty, pvRow, iMaxRows

'			.SSSetProtected C_PlantCd, pvRow, iMaxRows
			.SSSetProtected C_CIQty, pvRow, iMaxRows
			.SSSetProtected C_SoNo, pvRow, iMaxRows
			.SSSetProtected C_SoSeq, pvRow, iMaxRows
			.SSSetProtected C_SoSchdNo, pvRow, iMaxRows
			.SSSetProtected C_LcNo, pvRow, iMaxRows
			.SSSetProtected C_LcSeq, pvRow, iMaxRows
			.SSSetProtected C_GiAmt, pvRow, iMaxRows
			.SSSetProtected C_GiAmtLoc, pvRow, iMaxRows
			.SSSetProtected C_DepositAmt, pvRow, iMaxRows
			.SSSetProtected C_VatAmt, pvRow, iMaxRows
			.SSSetProtected C_VatAmtLoc, pvRow, iMaxRows
			.SSSetProtected C_QMItemFlag, pvRow, iMaxRows
			.SSSetProtected C_QmFlag, pvRow, iMaxRows
   
		   .SSSetProtected C_TolMoreQty, pvRow, iMaxRows
		   .SSSetProtected C_TolLessQty, pvRow, iMaxRows
		   .SSSetProtected C_RetType, pvRow, iMaxRows
		   .SSSetProtected C_RetTypeNm, pvRow, iMaxRows

			.SSSetProtected C_DnReqNo, pvRow, iMaxRows
			.SSSetProtected C_DnReqSeq, pvRow, iMaxRows
			.SSSetProtected C_CUST_LOT_NO, pvRow, iMaxRows
			.SSSetProtected C_RCPT_LOT_NO, pvRow, iMaxRows

			If frm1.vspdData.MaxRows > 0 Then
				frm1.btnPosting.disabled = False
				frm1.btnPostCancel.disabled = True
			Else
				frm1.btnPosting.disabled = True
				frm1.btnPostCancel.disabled = True
			End If

			Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "D")
		   
		   '====================================================================
		   ' 02.06 SMJ
		   ' ��ǰ�� ��� lot no, lot seq���� ���ϵ��� 
		   '====================================================================
'		   If Trim(frm1.txtHRetFlag.value) = "Y" Then   
'				frm1.vspdData.Col = C_RetType : frm1.vspdData.ColHidden = False
'				frm1.vspdData.Col = C_RetTypeNm : frm1.vspdData.ColHidden = False
'				.SSSetProtected C_LotNo, pvRow, iMaxRows
'				.SSSetProtected C_LotSeq, pvRow,iMaxRows
'				.SpreadLock C_LotNoPopup, pvRow, C_LotNoPopup, iMaxRows
'		   Else
				frm1.vspdData.Col = C_RetType : frm1.vspdData.ColHidden = True
				frm1.vspdData.Col = C_RetTypeNm : frm1.vspdData.ColHidden = True
'		   End If

			' Picking ������ ��ϵ� ��� â�� ������ �� ����.
			For i = pvRow to iMaxRows
				frm1.vspdData.Row = i
				frm1.vspdData.Col = C_PickQty
				If UNICDbl(frm1.vspdData.Text)  > 0 Then
					.SSSetProtected C_SlCd, i, i
					.SSSetProtected C_SlCdPopup, i, i
				Else
					.SSSetRequired  C_SlCd, i, i
				End If
				
'			   If Trim(frm1.txtHRetFlag.value) <> "Y" Then
					' Lot ���� ǰ�� ��� Lot ���� ���� ���� 
					frm1.vspdData.Col = C_LotFlag
					If frm1.vspdData.Text = "Y" Then
						.SpreadUnLock C_LotNo, i, C_LotNo, i
						.SpreadUnLock C_LotSeq, i, C_LotSeq, i
						.SSSetRequired C_LotNo, i, i
						.SSSetRequired C_LotSeq, i, i
						.SpreadUnLock C_LotNoPopup, i, C_LotNoPopup, i
					Else
						.SpreadLock C_LotNo, i, C_LotNo, i
						.SpreadLock C_LotSeq, i, C_LotSeq, i
						.SSSetProtected C_LotNo, i, i
						.SSSetProtected C_LotSeq, i, i
						.SpreadLock C_LotNoPopup, i, C_LotNoPopup, i
					End If
'				End If
			Next
		Else
			'--- ���ó���� �� ��� 
			For i = 1 To frm1.vspdData.MaxCols
				.SSSetProtected i, pvRow, iMaxRows
			Next 

			If frm1.vspdData.MaxRows > 0 Then
				frm1.btnPosting.disabled = True
				frm1.btnPostCancel.disabled = False
			Else
				frm1.btnPosting.disabled = True
				frm1.btnPostCancel.disabled = True
			End If

			Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "Q")
			Call ggoOper.SetReqAttr(frm1.chkArFlag, "Q")
			Call ggoOper.SetReqAttr(frm1.chkVatFlag, "Q")
			Call ggoOper.SetReqAttr(frm1.txtInvMgr, "Q")

		End if
 
		frm1.vspdData.ReDraw = True
    
    End With

End Sub

'=================================================================
Function CookiePage(Byval Kubun)

' On Error Resume Next

 Const CookieSplit = 4877      'Cookie Split String : CookiePage Function Use
 
 Dim strTemp, arrVal

 If Kubun = 1 Then

  WriteCookie CookieSplit , frm1.txtConDnNo.value

 ElseIf Kubun = 0 Then

  strTemp = ReadCookie(CookieSplit)

  If strTemp = "" then Exit Function
   
  arrVal = Split(strTemp, parent.gRowSep)

  If arrVal(0) = "" Then Exit Function
  
  frm1.txtConDnNo.value =  arrVal(0)

  If Err.number <> 0 Then
   Err.Clear
   WriteCookie CookieSplit , ""
   Exit Function 
  End If

  Call MainQuery()

  frm1.txtHRefRoot.value =  arrVal(1)

  WriteCookie CookieSplit , ""
  
 End If

End Function

'========================================
Function JumpChgCheck()

 Dim IntRetCD

 '************ ��Ƽ�� ��� **************
 ggoSpread.Source = frm1.vspdData 
 If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
  If IntRetCD = vbNo Then
   Exit Function
  End If
 End If

 Call CookiePage(1)
 Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'=================================================================
Function BtnSpreadCheck()
	Dim IntRetCD
	Dim iCnt, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim rtn 
	 
	BtnSpreadCheck = False

	If Trim(frm1.txtActualGIDt.Text) = "" Then
		MsgBox "����������� �Է��ϼ���", vbInformation, parent.gLogoName
		Call SetFocusToDocument("M")	
		frm1.txtActualGIDt.focus
		Exit Function
	End If

	'==================================================
	' 2002.2.4 SMJ
	' ����������� �����Ϻ��� �۰��Էµǵ��� ���� 
	'==================================================
	If UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , parent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate , parent.gDateFormat , "") Then  
		IntRetCD = DisplayMsgBox("970024", "X", frm1.txtActualGIDt.ALT, "������") 
		Call SetFocusToDocument("M")	
		frm1.txtActualGIDt.focus
		Exit Function
	End If
	
'	rtn = CommonQueryRs(" sh.so_no ", " s_so_hdr sh, s_dn_dtl dd ", " sh.so_no = dd.so_no and dd.dn_no = '" & frm1.txtConDnNo.value & "' and sh.so_dt > '" & UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , gDateFormat , "") & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'	If rtn = True Then
		
'		iCnt = Split(lgF0, Chr(11))	

'		If iCnt(0) <> "" Then
'			IntRetCD = DisplayMsgBox("970023", "X", frm1.txtActualGIDt.ALT, "���ֹ�ȣ : " & iCnt(0) & " ������")	
'			Exit Function
'		End If
'			
'		If Err.number <> 0 Then
'			MsgBox Err.description 
'			Err.Clear 
'			Exit Function
'		End If
'	End If			

  
	ggoSpread.Source = frm1.vspdData

	'������ ������ ���� ���� ���� üũ��, YES�̸� �۾����࿩�� üũ ���Ѵ� 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	'������ ������ �۾����࿩�� üũ 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function

'=================================================================
Function CheckCreditlimitSvr()

    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		  Exit Function
	End If

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=ChkGiCreditLimit"
    strVal = strVal & "&txtConDnNo=" & Trim(frm1.txtConDnNo.value)   
    
	Call RunMyBizASP(MyBizASP, strVal)          
 
End Function

'=================================================================
Function JungBokMsg(strJungBok1,strJungBok2,strID1,strID2)

 Dim strJugBokMsg

 If Len(Trim(strJungBok1)) Then strJungBok1 = strID1 & Chr(13) & String(30,"=") & strJungBok1
 If Len(Trim(strJungBok2)) Then strJungBok2 = strID2 & Chr(13) & String(30,"=") & strJungBok2

 If Len(Trim(strJungBok1)) Then strJugBokMsg = strJungBok1 & Chr(13) & Chr(13)
 If Len(Trim(strJungBok2)) Then strJugBokMsg = strJugBokMsg & strJungBok2 & Chr(13) & Chr(13)

 If Len(Trim(strJugBokMsg)) Then
  strJugBokMsg = strJugBokMsg & "�̹� ������ ��ȣ�� ������ �����մϴ�"
  MsgBox strJugBokMsg, vbInformation, parent.gLogoName
 End If

End Function

'=================================================================
Function CheckLotNoLotFlag()

	CheckLotNoLotFlag = False

	With frm1

		Dim lRow
 
		For lRow = 1 to .vspdData.MaxRows

			.vspdData.Row = lRow : .vspdData.Col = 0
			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Row = lRow : .vspdData.Col = C_LotFlag
					If UCase(Trim(.vspdData.Text)) = "Y" Then
						.vspdData.Col = C_LotNo
						If Trim(.vspdData.Text) = "*" Then
							Call DisplayMsgBox("204230", "X", lRow & "��:", "X")
							Exit Function
						End If
					End If
			End Select
		Next

	End With

	CheckLotNoLotFlag = True

End Function

'=====================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
						
			C_ItemCd	    = iCurColumnPos(1)    
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_TrackingNo	= iCurColumnPos(4)
			C_DnUnit		= iCurColumnPos(5)
			C_DnQty			= iCurColumnPos(6)
			C_DnBonusQty	= iCurColumnPos(7)
			C_PickQty		= iCurColumnPos(8)
			C_PickBonusQty  = iCurColumnPos(9)
			C_LotNo			= iCurColumnPos(10)
			C_LotNoPopup	= iCurColumnPos(11)
			C_LotSeq		= iCurColumnPos(12)
			C_OnStkQty		= iCurColumnPos(13)
			C_BasicUnit		= iCurColumnPos(14)
			C_CartonNo		= iCurColumnPos(15)

			C_GiAmt			= iCurColumnPos(16)
			C_GiAmtLoc		= iCurColumnPos(17)
			C_DepositAmt	= iCurColumnPos(18)
			C_VatAmt		= iCurColumnPos(19)
			C_VatAmtLoc		= iCurColumnPos(20)

			C_QMItemFlag	= iCurColumnPos(21)
			C_QmFlag		= iCurColumnPos(22)
			C_QmNoPopup		= iCurColumnPos(23)

			C_PlantCd		= iCurColumnPos(24)
			C_PlantPopup	= iCurColumnPos(25)
			C_SlCd			= iCurColumnPos(26)
			C_SlCdPopup		= iCurColumnPos(27)
			C_TolMoreQty	= iCurColumnPos(28)
			C_TolLessQty	= iCurColumnPos(29)
			C_CIQty			= iCurColumnPos(30)
			C_SoNo			= iCurColumnPos(31)
			C_SoSeq			= iCurColumnPos(32)
			C_SoSchdNo		= iCurColumnPos(33)
			C_LcNo			= iCurColumnPos(34)
			C_LcSeq			= iCurColumnPos(35)
			C_RetType		= iCurColumnPos(36)
			C_RetTypeNm		= iCurColumnPos(37)
			C_Remark		= iCurColumnPos(38)
			C_LotReqmtFlag  = iCurColumnPos(39)
			C_LotFlag		= iCurColumnPos(40)
			C_DnSeq			= iCurColumnPos(41)
			C_RelBillNo     = iCurColumnPos(42)
			C_RelBillCnt    = iCurColumnPos(43)

			C_DnReqNo       = iCurColumnPos(44)
			C_DnReqSeq      = iCurColumnPos(45)
			C_RCPT_LOT_NO		= iCurColumnPos(46)
			C_CUST_LOT_NO		= iCurColumnPos(47)

			C_OUT_NO				= iCurColumnPos(48)
			C_TRANS_TIME		= iCurColumnPos(49)
			C_OUT_TYPE_SUB	= iCurColumnPos(50)
			C_CREATE_TYPE		= iCurColumnPos(51)
			C_REF_GUBUN			= iCurColumnPos(52)
			
			'2008-06-16 7:55���� :: hanc
			C_pgm_name			= iCurColumnPos(53)
			C_pgm_price			= iCurColumnPos(54)


			
    End Select    
End Sub

'========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow

           If Not Frm1.vspdData.ColHidden Then
			  Call SetActiveCell(frm1.vspdData, iDx, iRow,"M","X","X")
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'=========================================
Sub Form_Load() 
    Call LoadInfTB19029              '��: Load table , B_numeric_format    
    Call FormatField()
    Call LockFieldInit("L")
    Call InitSpreadSheet
	Call SetDefaultVal 
	Call InitVariables              

    Call SetToolbar("11000000000011")          '��: ��ư ���� ����    
	Call CookiePage(0)
End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	If Row <= 0 Then Exit Sub

	Dim strPlantCd, strSICd

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		.Row = Row
		
		Select Case Col
			Case C_PlantPopup
				.Col = Col - 1
				Call OpenDnDtl(.Text, 1)

			Case C_SlCdPopup
				.Col = Col - 1		:	strSICd = .Text
				.Col = C_PlantCd	:	strPlantCd = .Text

				Call OpenDnDtl(strSICd & parent.gColSep & strPlantCd, 2)

			Case C_LotNoPopup
				Call OpenLotNoPopup(Row)
				
			Case C_QmNoPopup
				Call OpenQMDtlRef(Row)
		End Select

		Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
		
	End With

End Sub

'=========================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	If lgIntFlgMode = parent.OPMD_UMODE Then
		If Len(Trim(frm1.txtGINo.value)) Then
			Call SetPopupMenuItemInf("0000111111")
		Else
			Call SetPopupMenuItemInf("0101111111")
		End If
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	

End Sub

'=========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'=========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'=========================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

	Select Case Col
	
		Case C_PickQty       '2008-05-08 5:34���� :: hanc
		    Call SumPicking()

    End Select
End Sub

Function SumPicking()

    Dim iLngRow, iPickingQty
    
    iPickingQty =   0
    
	With frm1.vspdData
		For iLngRow = 1 To .MaxRows
			.Row = iLngRow
			.Col = C_PickQty
			
			iPickingQty =   CDbl(iPickingQty) + UNIConvNum(.Text, 0)

		Next
	End With
	
	frm1.txtSumPicking.value    =   CStr(iPickingQty)
	
End Function 


'=========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then       '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess Then Exit Sub
	   
			Call DisableToolBar(parent.TBC_QUERY)
			Call DBQuery
		End If
	End if    
End Sub


'=============================================
' 2005.11.10 SMJ
'=============================================
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
    
   	
	Case "[���ֹ�ȣ]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow		

		if 	frm1.vspddata.value <> "" then
	
				
					pvKeyVal =   frm1.vspddata.value
					
									
					pvSingle =   ""
				
					pvFB_fg = "B"
					pvSelmvid = "SO_NO"
	
						Call Jump_Pgm (	pvSelmvid, _
										pvFB_fg, _
										pvSingle,  _
										pvKeyVal)
										
										
										
	End if 											
		 
	End select
End Function


'=========================================
Sub btnPosting_OnClick()
	Dim IntRetCD 
	 
	If BtnSpreadCheck = False Then Exit Sub
	  
	Call CheckCreditlimitSvr
End Sub

'=========================================
Sub btnPostCancel_OnClick()

	If BtnSpreadCheck = False Then Exit Sub
	Call BatchButton(3)

End Sub

'=========================================
Function BatchButton(ByVal iKubun)

    Err.Clear                                                               

	Select Case iKubun 
		Case 2
			frm1.txtBatch.value = "Posting"
		Case 3
			frm1.txtBatch.value = "PostCancel"
			If LayerShowHide(1) = False Then
				Exit Function
			End If
		Case Else
			Exit Function
	End Select

	frm1.txtARFlag.value = ""
	frm1.txtVatFlag.value = ""
	If frm1.chkARflag.checked = True Then frm1.txtARFlag.value = "Y"
	If frm1.chkVatflag.checked = True Then frm1.txtVatFlag.value = "Y"
	    
	Dim strPostVal
	strPostVal = BIZ_PGM_ID & "?txtMode=" & "ARPOST"         
	strPostVal = strPostVal & "&txtHDnNo=" & Trim(frm1.txtHDnNo.value)      
	strPostVal = strPostVal & "&txtActualGIDt=" & Trim(frm1.txtActualGIDt.Text)
	strPostVal = strPostVal & "&txtARFlag=" & Trim(frm1.txtARFlag.value)
	strPostVal = strPostVal & "&txtVatFlag=" & Trim(frm1.txtVatFlag.value)
	strPostVal = strPostVal & "&txtInvMgr=" & Trim(frm1.txtInvMgr.value)
	strPostVal = strPostVal & "&txtGINo=" & Trim(frm1.txtGINo.value)

	Call RunMyBizASP(MyBizASP, strPostVal)             
End Function

'=========================================
Sub txtActualGIDt_Change
	' ����������� �����Ϻ��� �۰��Էµǵ��� ���� 
'	If UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , parent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate , parent.gDateFormat , "") Then
'		Call DisplayMsgBox("970024", "X", frm1.txtActualGIDt.ALT, "������")
'		Call SetFocusToDocument("M")	
'        frm1.txtActualGIDt.Focus
'		Exit Sub
'	End If

	With frm1
		If Trim(frm1.txtActualGIDt.text) <> "" Then
			Call ggoOper.SetReqAttr(.txtInvMgr, "D")	
		Else
			Call ggoOper.SetReqAttr(.txtInvMgr, "Q")	
		End If
		
		' ������ �ʿ� ���� ��쳪, ���ǿ� ���ؼ��� ���ó���� ���ÿ� �����ڷḦ ������ �� �� ����.
		If Trim(frm1.txtActualGIDt.text) <> "" And Trim(.txtRetBillFlag.value) = "Y" And Trim(.txtExportFlag.value) = "N" Then
			Call ggoOper.SetReqAttr(.chkVatFlag, "D")
			Call ggoOper.SetReqAttr(.chkARflag, "D")	
		Else
			Call ggoOper.SetReqAttr(.chkVatFlag, "Q")
			Call ggoOper.SetReqAttr(.chkARflag, "Q")
		End If
	End With

	lgBlnFlgChgValue = True
End Sub

'=========================================
Sub txtActualGIDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtActualGIDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtActualGIDt.Focus
	End If
End Sub

'=======================================================
'   Event Name : chkTaxNo_OnPropertyChange
'   Event Desc : ���ݰ�꼭 �ڵ����� ���ο� ���� �����Է��׸� Change
'=======================================================
Sub chkArFlag_OnClick()

	On Error Resume Next

	Select Case frm1.chkArFlag.checked
	Case True
		lblArFlag.disabled = False
	Case False
		lblArFlag.disabled = True
		lblVatFlag.disabled = True
		frm1.chkVatFlag.checked = False
	End Select

	lgBlnFlgChgValue = True

	If Err.number <> 0 Then Err.Clear

End Sub

'=====================================================
Sub chkVatFlag_OnClick()

	On Error Resume Next

	Select Case frm1.chkVatFlag.checked
		Case True
			lblArFlag.disabled = False
			lblVatFlag.disabled = False
			frm1.chkARflag.checked = True  
		Case False
			lblArFlag.disabled = True
			lblVatFlag.disabled = True
			frm1.chkARflag.checked = False
	End Select

	lgBlnFlgChgValue = True

	If Err.number <> 0 Then Err.Clear
 
End Sub

'=====================================================
Function FncQuery() 
    Dim IntRetCD 
    on error resume next
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkField(Document, "1") Then Exit Function

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")              
    Call InitVariables               

    Call DbQuery

    FncQuery = True                
        
End Function

'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

    Call ggoOper.ClearField(Document, "A")
    Call LockFieldInit("N")                                       '��: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables               

    Call SetToolbar("11000000000011")          '��: ��ư ���� ���� 
    
    FncNew = True                

End Function

'=====================================================
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                                   
    
    FncDelete = False              
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If
    
    Call ggoOper.ClearField(Document, "A")                                   '��: Clear Condition Field
    FncDelete = True                                                        
    
End Function

'=====================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then		
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

	If ggoSpread.SSDefaultCheck = False Then
       Exit Function
    End If


 '--- [2002-01-08] : ��ǰ�� ���� Skip ---
'	If Trim(frm1.txtHRetFlag.value) <> "Y" Then
	 '[2002-01-08] : Lot�������ΰ� "Y"�ϰ��, LotNo�� "*"�� �ԷµǸ� �ȵȴ� ///
		If CheckLotNoLotFlag = False Then Exit Function
'	End If

    CAll DbSave

    FncSave = True                                                          
    
End Function

'=====================================================
Function FncCancel() 
 If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo  
End Function

'=====================================================
Function FncDeleteRow() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    
	lDelRows = ggoSpread.DeleteRow
 
    lgBlnFlgChgValue = True
    
    End With
    
End Function

'=====================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'=====================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'=====================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=====================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(1)
End Sub

'=====================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False

 ggoSpread.Source = frm1.vspdData 
 If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
	'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
	If IntRetCD = vbNo Then
		Exit Function
	End If
 End If

 FncExit = True
End Function

'=====================================================
Function DbDelete() 
    On Error Resume Next                                                    
End Function

'=====================================================
Function DbDeleteOk()              
    On Error Resume Next                                                    
End Function

'=====================================================
Function DbQuery() 

    Err.Clear                                                               

    DbQuery = False                                                         

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    Dim iStrVal

    If lgIntFlgMode = parent.OPMD_UMODE Then
		iStrVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         
		iStrVal = iStrVal & "&txtConDnNo=" & Trim(frm1.txtHDnNo.value)     
		iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		iStrVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001         
		iStrVal = iStrVal & "&txtConDnNo=" & Trim(frm1.txtConDnNo.value)     
		iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If
    
	iStrVal = iStrVal & "&txtLastRow=" & frm1.vspdData.MaxRows
	
	lgLngStartRow = frm1.vspdData.MaxRows + 1

	Call RunMyBizASP(MyBizASP, iStrVal)            
  
    DbQuery = True                 

End Function

'=====================================================
Function DbQueryOk()
on error resume next
    lgIntFlgMode = parent.OPMD_UMODE

	With frm1
		' ������ �������� �ʴ� ��� 
		If .vspdData.MaxRows = 0 Then
			.btnPosting.disabled = True
			.btnPostCancel.disabled = True
			Call ggoOper.SetReqAttr(.txtActualGIDt, "Q")
			frm1.txtConDnNo.focus
		Else
			frm1.vspdData.Focus()
			Call SetQuerySpreadColor(lgLngStartRow)
		End If

		' Scroll ��ȸ�ô� �������� �ʴ´�.
		If lgLngStartRow = 1 Then
			'���/�԰� ó�� ��ǰ���ο� ���� 
			If UCase(.txtHRetFlag.value) = "Y" Then
				.btnPosting.value = "�԰�ó��"
				.btnPostCancel.value = "�԰�ó�����"
			Else
				.btnPosting.value = "���ó��"
				.btnPostCancel.value = "���ó�����"
			End If

			' ���ó���� ��� 
			If Len(Trim(frm1.txtGINo.value)) Then
				Call SetToolbar("11100000000111")          '��: ��ư ���� ���� 
			Else
				Call SetToolbar("11101011000111")          '��: ��ư ���� ���� 
			End If

			lgBlnFlgChgValue = False
 		End If
	End With
	
    Call SumPicking()   '2008-06-11 10:47���� :: hanc

End Function

' ���������� �������� ���� ��� 
'=====================================================
Function DbQueryNotFound()
	Call SetDefaultVal
	Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "Q")
	Call SetToolbar("11000000000011")
	frm1.txtConDnNo.focus
End Function

'=====================================================
Function DbSave() 
	On Error Resume Next
	
    Err.Clear                
 
    Dim iLngRow, iLngRowsIns, iLngRowsUpd, iLngRowsDel
	Dim iArrData, iArrRowsIns, iArrRowsUpd, iArrRowsDel
 
    DbSave = False                                                          

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	iLngRowsIns = -1
	iLngRowsUpd = -1
	iLngRowsDel = -1
	
	Redim iArrRowsIns(0)
	Redim iArrRowsUpd(0)
	Redim iArrRowsDel(0)
	iArrRowsIns(0) = ""
	iArrRowsUpd(0) = ""
	iArrRowsDel(0) = ""

	Redim iArrdata(59)
	
	With frm1.vspdData
		For iLngRow = 1 To .MaxRows
			.Row = iLngRow
			.Col = 0

			'������ ��� 
			If .Text = ggoSpread.DeleteFlag then
				iLngRowsDel = iLngRowsDel + 1
				' Row ��ȣ, ���ϼ��� 
				.Col = C_DnSeq
				Redim Preserve iArrRowsDel(iLngRowsDel)
				iArrRowsDel(iLngRowsDel) = CStr(iLngRow) & parent.gColSep & Trim(.Text)
				
			' �Է�, ������ ��� 
			Elseif .Text <> "" Then
				iArrData(0) = iLngRow				' Row��ȣ 
				.Col = C_DnSeq			:	iArrData(1) = UNIConvNum(.Text, 0)	' ���ϼ��� 
				.Col = C_ItemCd			:	iArrData(2) = Trim(.Text)
				.Col = C_DnQty			:	iArrData(3) = UNIConvNum(.Text, 0)	' ����û���� 
				.Col = C_DnUnit			:	iArrData(4) = Trim(.Text)			' ���� 
				.Col = C_DnBonusQty		:	iArrData(5) = UNIConvNum(.Text, 0)	' ����û������ 
				.Col = C_PickQty		:	iArrData(6) = UNIConvNum(.Text, 0)	' Picking���� 
				.Col = C_PickBonusQty	:	iArrData(7) = UNIConvNum(.Text, 0)	' Picking������ 
				.Col = C_PlantCd		:	iArrData(8) = Trim(.Text)			' ���� 
				.Col = C_SlCd			:	iArrData(9) = Trim(.Text)			' â�� 
				.Col = C_TolMoreQty		:	iArrData(10) = UNIConvNum(.Text, 0)	' ��������뷮(+)
				.Col = C_TolLessQty		:	iArrData(11) = UNIConvNum(.Text, 0)	' ��������뷮(-)
				.Col = C_LotNo			:	iArrData(12) = Trim(.Text)			' LOT No
				.Col = C_LotSeq			:	iArrData(13) = UNIConvNum(.Text, 0) ' LOT No ���� 
				.Col = C_CIQty			:	iArrData(14) = UNIConvNum(.Text, 0)	' ����� 
				.Col = C_SoNo			:	iArrData(15) = Trim(.Text)			' ���ֹ�ȣ 
				.Col = C_SoSeq			:	iArrData(16) = UNIConvNum(.Text, 0)	' ���ּ��� 
				.Col = C_SoSchdNo		:	iArrData(17) = UNIConvNum(.Text, 0)	' ��ǰ���� 
				.Col = C_LcNo			:	iArrData(18) = Trim(.Text)			' L/C��ȣ 
				.Col = C_LcSeq			:	iArrData(19) = UNIConvNum(.Text, 0)	' L/C���� 
				.Col = C_Remark			:	iArrData(20) = Trim(.Text)			' ��� 
				.Col = C_QmFlag			:	iArrData(21) = Trim(.Text)			' �˻籸�� 

				iArrData(22) = "0"			' ext1_qty
				iArrData(23) = "0"			' ext1_amt
				iArrData(24) = GetSpreadText(frm1.vspdData,C_CUST_LOT_NO,iLngRow,"X","X")		' ext1_cd
				iArrData(25) = "0"			' ext2_qty
				iArrData(26) = "0"			' ext2_amt
				iArrData(27) = GetSpreadText(frm1.vspdData,C_RCPT_LOT_NO,iLngRow,"X","X")			' ext2_cd
				iArrData(28) = "0"			' ext3_qty
				iArrData(29) = "0"			' ext3_amt
				iArrData(30) = ""			' ext3_cd
				.Col = C_CartonNo		:	iArrData(31) = Trim(.Text)				' Carton ��ȣ 

				.Col = C_DnReqNo			:	iArrData(51) = Trim(.Text)			' ���Ͽ�û��ȣ 
				.Col = C_DnReqSeq			:	iArrData(52) = UNIConvNum(.Text, 0)	' ���Ͽ�û���� 

				.Col = C_OUT_NO					:	iArrData(53) = Trim(.Text)
				.Col = C_TRANS_TIME			:	iArrData(54) = Trim(.Text)
				.Col = C_OUT_TYPE_SUB		:	iArrData(55) = Trim(.Text)
				.Col = C_CREATE_TYPE		:	iArrData(56) = Trim(.Text)
				.Col = C_REF_GUBUN			:	iArrData(57) = Trim(.Text)

				.Col = C_pgm_name			:	iArrData(58) = Trim(.Text)      '2008-06-16 8:02���� :: hanc
				.Col = C_pgm_price			:	iArrData(59) = Trim(.Text)      '2008-06-16 8:02���� :: hanc
				
				.Col = 0
				' �Է� 
				If .Text = ggoSpread.InsertFlag then
					iLngRowsIns = iLngRowsIns + 1
					Redim Preserve iArrRowsIns(iLngRowsIns)
					iArrRowsIns(iLngRowsIns) = Join(iArrData, parent.gColSep)
				' ���� 
				ElseIf .Text = ggoSpread.UpdateFlag then

					' �˻��Ƿ������� ������ ����� ���� (2006-06-01 ������ ����)	
					Dim iCnt, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 , a
					Dim rtn

					rtn = CommonQueryRs("TOP 1 document_no", " Q_INSPECTION_REQUEST (nolock)", " document_no = '" & frm1.txtConDnNo.value & "' and document_seq_no = '" & iArrData(1) & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				        If rtn = True Then
					   Call DisplayMsgBox("223703","X" , "X","X")   '�˻��Ƿ� ������ �ڷᰡ �̹� �����մϴ�.
					   Call LayerShowHide(0)
				 	   Exit Function
					END if

					iLngRowsUpd = iLngRowsUpd + 1
					Redim Preserve iArrRowsUpd(iLngRowsUpd)
					iArrRowsUpd(iLngRowsUpd) = Join(iArrData, parent.gColSep)
				End if
			End If
		Next
	End With
	
	With frm1
		.txtMode.value = parent.UID_M0002
		If iLngRowsIns >= 0 Then .txtSpreadIns.value = Join(iArrRowsIns, parent.gRowSep) & parent.gRowSep
		If iLngRowsUpd >= 0 Then .txtSpreadUpd.value = Join(iArrRowsUpd, parent.gRowSep) & parent.gRowSep
		If iLngRowsDel >= 0 Then .txtSpreadDel.value = Join(iArrRowsDel, parent.gRowSep) & parent.gRowSep
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
 
    DbSave = True                                                           
    
End Function

'=====================================================
Function DbSaveOk()

    Call InitVariables
	frm1.txtConDnNo.value = frm1.txtHDnNo.value
	Call ggoOper.ClearField(Document, "2")
    Call MainQuery()

	frm1.txtBatch.value = ""

End Function

Function GetIssueFromMESOk()
		
      frm1.vspdData.Row = frm1.vspdData.MaxRows
      frm1.vspdData.Col = 0
      frm1.vspdData.Text = ggoSpread.InsertFlag
				
			SetSpreadColor frm1.vspdData.MaxRows, frm1.vspdData.MaxRows
			'Call vspdData_Change(C_BillQty,frm1.vspdData.MaxRows)
			
End Function


'==============================================================================================================================
' ������ �߰� (2006-04-27) 
'==============================================================================================================================
Function OpenGLRef() 

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
		
	Dim iCnt, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 , a
	Dim rtn

'	rtn = CommonQueryRs("document_year", "i_goods_movement_header", " item_document_no   = '" & frm1.txtGINo.value & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	rtn = CommonQueryRs("TOP 1 A.document_year", "i_goods_movement_header A , i_goods_movement_detail B ", " a.item_document_no = b.item_document_no and a.item_document_no   = '" & frm1.txtGINo.value & "' and B.dn_no   = '" & frm1.txtConDnNo.Value & "'" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        If rtn = True Then
		
		iCnt = replace(lgF0, Chr(11),"")	

		If iCnt  <> "" Then

		   rtn = CommonQueryRs("temp_gl_no, gl_no", "a_batch", " ref_no    = '" & frm1.txtGINo.value & "-" & iCnt & "'"  , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

'		   MSGBOX frm1.txtGINo.value & "-" & iCnt

		   if replace(lgF0, Chr(11),"") <> ""  Then 
			a = "T"
			arrParam(0) = replace(lgF0, Chr(11),"")
		   End if

		   if replace(lgF1, Chr(11),"") <> ""  Then 
			a = "A"
			arrParam(0) = replace(lgF1, Chr(11),"")
		   End if

		End if
	ELSE
		a = "B"
	End If


	arrParam(1) = ""

   If a = "A" Then               'ȸ����ǥ�˾� 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif a = "T" Then          '������ǥ�˾� 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif a= "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	lblnWinEvent = False
	
End Function

'==============================================================================================================================
' ������ �߰� (2006-04-27) END
'==============================================================================================================================

