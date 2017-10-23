
Option Explicit                               

Const BIZ_PGM_ID = "s4512mb1.asp"            '��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID = "s4511ma1"            '��: JUMP�� �����Ͻ� ���� ASP�� 

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
''''    Call FormatDATEField(frm1.txtActualGIDt)
End Sub

'=========================================
Sub LockFieldInit(ByVal pvFlag)
    With frm1
        ' ��¥ OCX
''''        Call LockObjectField(.txtActualGIDt, "P")

        If pvFlag = "N" Then
''''			Call LockHTMLField(.txtInvMgr, "P")	
''''			Call LockHTMLField(.chkArFlag, "P")	
''''			Call LockHTMLField(.chkVatFlag, "P")	
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
''''	frm1.btnPosting.disabled = True
''''	frm1.btnPostCancel.disabled = True
''''	frm1.btnPosting.value = "���ó��"
''''	frm1.btnPostCancel.value = "���ó�����"  
	 
	lgBlnFlgChgValue = False
''''	frm1.chkARflag.checked = False
''''	frm1.chkVatFlag.checked = False
	Call chkVatFlag_OnClick()
End Sub

'=========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	With ggoSpread

		.Source = frm1.vspdData
		.Spreadinit "V20030902",,parent.gAllowDragDropSpread    
		frm1.vspdData.ReDraw = false
		
		frm1.vspdData.MaxCols = C_RelBillCnt + 1            '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		frm1.vspdData.Col = frm1.vspdData.MaxCols               '��: ������Ʈ�� ��� Hidden Column
		frm1.vspdData.ColHidden = True
 
		frm1.vspdData.MaxRows = 0

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("7","5","0")

		.SSSetFloat C_DnSeq,"���Ͽ�û����" ,15,"7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"  
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
		
'''' Picking
		Call .SSSetColHidden(C_PickQty,C_PickBonusQty,True)

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
		.SSSetRequired  C_DnQty, pvStartRow, pvEndRow
		.SSSetProtected C_DnUnit, pvStartRow, pvEndRow
		.SSSetRequired  C_DnBonusQty, pvStartRow, pvEndRow
		.SSSetProtected C_OnStkQty, pvStartRow, pvEndRow
		.SSSetProtected C_BasicUnit, pvStartRow, pvEndRow

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
			  
		' ��ǰ�� ���� Lot ��ȣ�� ������ �� ���� 
		If Trim(frm1.txtHRetFlag.value) = "Y" Then   
			frm1.vspdData.Col = C_RetType	: frm1.vspdData.ColHidden = False
			frm1.vspdData.Col = C_RetTypeNm	: frm1.vspdData.ColHidden = False
			.SSSetProtected C_LotNo, pvStartRow, pvEndRow
			.SSSetProtected C_LotSeq, pvStartRow, pvEndRow
			.SpreadLock C_LotNoPopup, pvStartRow, C_LotNoPopup, pvEndRow
		Else
			frm1.vspdData.Col = C_RetType	: frm1.vspdData.ColHidden = True
			frm1.vspdData.Col = C_RetTypeNm	: frm1.vspdData.ColHidden = True

			For iRow = pvStartRow To pvEndRow
				frm1.vspdData.Row = iRow	:	frm1.vspdData.Col = C_LotFlag
				' Lot ���� ǰ�� ��� Lot ���� ���� ���� 
				If frm1.vspdData.Text = "Y" Then
					.SpreadUnLock C_LotNo, iRow, C_LotNo, iRow
					.SpreadUnLock C_LotSeq, iRow, C_LotSeq, iRow
'					.SSSetRequired C_LotNo, iRow, iRow
'					.SSSetRequired C_LotSeq, iRow, iRow
					.SpreadUnLock C_LotNoPopup, iRow, C_LotNoPopup, iRow
				Else
					.SpreadLock C_LotNo, iRow, C_LotNo, iRow
					.SpreadLock C_LotSeq, iRow, C_LotSeq, iRow
					.SSSetProtected C_LotNo, iRow, iRow
					.SSSetProtected C_LotSeq, iRow, iRow
					.SpreadLock C_LotNoPopup, iRow, C_LotNoPopup, iRow
				End If
			Next
		End If


'''''''''''' Picking
		.SSSetProtected C_PickQty, pvStartRow, pvEndRow
		.SSSetProtected C_PickBonusQty, pvStartRow, pvEndRow

	End With
End Sub

'========================================
Function OpenConDnDtl()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
	   
	IsOpenPop = True

	iCalledAspName = AskPRAspName("S4511PA1")
			
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4511PA1", "x")
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
''''Sub OpenInvMgrPopUp()
''''
''''	Dim iArrRet
''''	Dim iArrParam(5), iArrField(6), iArrHeader(6)

''''	If IsOpenPop Then Exit Sub

''''	With frm1
''''''''		If .txtInvMgr.readOnly Then	Exit Sub
''''
''''		IsOpenPop = True

''''		iArrParam(1) = "dbo.B_MINOR"
''''		iArrParam(2) = ""
''''		iArrParam(3) = ""											
''''		iArrParam(4) = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & ""
''''				
''''		iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
''''		iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
''''							
''''		iArrHeader(0) = ""
''''		iArrHeader(1) = ""
''''
''''''''		.txtInvMgr.focus
''''	End With
''''	
''''	iArrParam(0) = iArrHeader(0)							' �˾� Title
''''	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī 

''''	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
''''		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
''''
''''	IsOpenPop = False

''''	If iArrRet(0) <> "" Then
''''		frm1.txtInvMgr.value = iArrRet(0)
''''		frm1.txtInvMgrNm.value = iArrRet(1)
''''	End If	
''''End Sub

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

'''' Picking
''''	.vspdData.Col = C_PickQty : lsPickQty = UNICDbl(Trim(.vspdData.text))
''''    .vspdData.Col = C_PickBonusQty : lsPickBonusQty = UNICDbl(Trim(.vspdData.text))
	.vspdData.Col = C_PickQty : lsPickQty = 0
    .vspdData.Col = C_PickBonusQty : lsPickBonusQty = 0

'    lsTotDnQty = @@@UNICDbl(lsDnQty) + @@@UNICDbl(lsDnBonusQty)
    lsTotPickQty = UNIFormatNumber(UNICDbl(lsPickQty) + UNICDbl(lsPickBonusQty), ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)

    lsAvlQty = UNICDbl(arrRet(5))
    
    If lsAvlQty >= uniCDbl(lsTotPickQty) Then
     '.vspdData.Col = C_PickQty : .vspdData.Text = lsPickQty
     '.vspdData.Col = C_PickBonusQty : .vspdData.Text = lsPickBonusQty
    ElseIf lsAvlQty < uniCDbl(lsTotPickQty) Then
     If lsAvlQty <= lsPickQty Then
'''' Picking
''''  .vspdData.Col = C_PickQty : .vspdData.Text = lsAvlQty
''''  .vspdData.Col = C_PickBonusQty : .vspdData.Text = 0
	  .vspdData.Col = C_PickQty : .vspdData.Text = 0
      .vspdData.Col = C_PickBonusQty : .vspdData.Text = 0
     ElseIf lsAvlQty > lsPickQty Then
'''' Picking
''''  .vspdData.Col = C_PickQty : .vspdData.Text = lsPickQty
''''  .vspdData.Col = C_PickBonusQty : .vspdData.Text = UNIFormatNumber(UNICDbl(lsAvlQty) - UNICDbl(lsPickQty),  ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
	  .vspdData.Col = C_PickQty : .vspdData.Text = 0
      .vspdData.Col = C_PickBonusQty : .vspdData.Text = 0
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

	on error Resume Next

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

	on error Resume Next

	If Trim(frm1.txtPlannedGIDt.value) = "" Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

''''	If Len(Trim(frm1.txtGINo.value)) Then
''''		Msgbox "���ó���� ǰ���� ���ֳ����� ���� �� �� �����ϴ�",vbInformation, parent.gLogoName
''''		Exit Function
''''	End If

	iCalledAspName = AskPRAspName("S4512AA1")
					
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4512AA1", "x")
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
				Trim(frm1.txtPlantCd.Value)

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strParam), _
	"dialogWidth=850px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0, 0) = "" Then
		If Err.number <> 0 Then Err.Clear 
		Exit Function
	Else
		Call SetSODtlRef(arrRet)
	End If 
End Function
 
'========================================
Function OpenQMDtlRef(Row)
	Dim iCalledAspName
	Dim strRet
	Dim arrValue(2)
	Dim ItemCode
	Dim DnSeq

	on error Resume Next

	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Function

''''	If Len(Trim(frm1.txtGINo.value)) Then
''''		Exit Function
''''	End If

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
on error Resume Next

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
'''' Picking
''''			.Col = C_PickQty	:		.text = pvArrRet(iLngCnt, 3)			'Picking���� 
''''			.Col = C_PickBonusQty	:	.text = pvArrRet(iLngCnt, 4)			'Picking������ 
			.Col = C_PickQty	:		.text = 0			'Picking���� 
			.Col = C_PickBonusQty	:	.text = 0			'Picking������ 
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
 
'=====================================================
Sub SetQuerySpreadColor(ByVal pvRow)
	on error Resume Next
	Dim i, iMaxRows
  
	iMaxRows = frm1.vspdData.MaxRows
    With ggoSpread  
  
		frm1.vspdData.ReDraw = False

		'--- ���ó���� �Ǿ������� Ȯ���Ѵ�.
''''		If Trim(frm1.txtGINo.value) = "" Then
		If True Then
			'--- ���ó���� ���� ���� ��� 
			.SSSetProtected C_ItemCd, pvRow, iMaxRows
			.SSSetProtected C_ItemNm, pvRow, iMaxRows
			.SSSetProtected C_Spec, pvRow, iMaxRows
			.SSSetProtected C_TrackingNo, pvRow, iMaxRows        
			.SSSetRequired  C_DnQty, pvRow, iMaxRows
			.SSSetProtected C_DnUnit, pvRow, iMaxRows
			.SSSetProtected C_OnStkQty, pvRow, iMaxRows
			.SSSetProtected C_BasicUnit, pvRow, iMaxRows
			.SSSetRequired  C_DnBonusQty, pvRow, iMaxRows
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

''''			If frm1.vspdData.MaxRows > 0 Then
''''				frm1.btnPosting.disabled = False
''''				frm1.btnPostCancel.disabled = True
''''			Else
''''				frm1.btnPosting.disabled = True
''''				frm1.btnPostCancel.disabled = True
''''			End If

''''			Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "D")
		   
		   '====================================================================
		   ' 02.06 SMJ
		   ' ��ǰ�� ��� lot no, lot seq���� ���ϵ��� 
		   '====================================================================
		   If Trim(frm1.txtHRetFlag.value) = "Y" Then   
				frm1.vspdData.Col = C_RetType : frm1.vspdData.ColHidden = False
				frm1.vspdData.Col = C_RetTypeNm : frm1.vspdData.ColHidden = False
				.SSSetProtected C_LotNo, pvRow, iMaxRows
				.SSSetProtected C_LotSeq, pvRow,iMaxRows
				.SpreadLock C_LotNoPopup, pvRow, C_LotNoPopup, iMaxRows
		   Else
				frm1.vspdData.Col = C_RetType : frm1.vspdData.ColHidden = True
				frm1.vspdData.Col = C_RetTypeNm : frm1.vspdData.ColHidden = True
		   End If

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
				
			   If Trim(frm1.txtHRetFlag.value) <> "Y" Then
					' Lot ���� ǰ�� ��� Lot ���� ���� ���� 
					frm1.vspdData.Col = C_LotFlag
					If frm1.vspdData.Text = "Y" Then
						.SpreadUnLock C_LotNo, i, C_LotNo, i
						.SpreadUnLock C_LotSeq, i, C_LotSeq, i
'						.SSSetRequired C_LotNo, i, i
'						.SSSetRequired C_LotSeq, i, i
						.SpreadUnLock C_LotNoPopup, i, C_LotNoPopup, i
					Else
						.SpreadLock C_LotNo, i, C_LotNo, i
						.SpreadLock C_LotSeq, i, C_LotSeq, i
						.SSSetProtected C_LotNo, i, i
						.SSSetProtected C_LotSeq, i, i
						.SpreadLock C_LotNoPopup, i, C_LotNoPopup, i
					End If
				End If
			Next
		Else
			'--- ���ó���� �� ��� 
			For i = 1 To frm1.vspdData.MaxCols
				.SSSetProtected i, pvRow, iMaxRows
			Next 

''''			If frm1.vspdData.MaxRows > 0 Then
''''				frm1.btnPosting.disabled = True
''''				frm1.btnPostCancel.disabled = False
''''			Else
''''				frm1.btnPosting.disabled = True
''''				frm1.btnPostCancel.disabled = True
''''			End If

''''			Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "Q")
''''			Call ggoOper.SetReqAttr(frm1.chkArFlag, "Q")
''''			Call ggoOper.SetReqAttr(frm1.chkVatFlag, "Q")
''''			Call ggoOper.SetReqAttr(frm1.txtInvMgr, "Q")

		End if
 
		frm1.vspdData.ReDraw = True

'''''''''''Picking
		.SSSetProtected C_PickQty, pvRow, iMaxRows
		.SSSetProtected C_PickBonusQty, pvRow, iMaxRows

    End With

End Sub

'=================================================================
Function CookiePage(Byval Kubun)

 on error Resume Next

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

''''	If Trim(frm1.txtActualGIDt.Text) = "" Then
''''		MsgBox "����������� �Է��ϼ���", vbInformation, parent.gLogoName
''''		Call SetFocusToDocument("M")	
''''		frm1.txtActualGIDt.focus
''''		Exit Function
''''	End If

	'==================================================
	' 2002.2.4 SMJ
	' ����������� �����Ϻ��� �۰��Էµǵ��� ���� 
	'==================================================
''''	If UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , parent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate , parent.gDateFormat , "") Then  
''''		IntRetCD = DisplayMsgBox("970024", "X", frm1.txtActualGIDt.ALT, "������") 
''''		Call SetFocusToDocument("M")	
''''		frm1.txtActualGIDt.focus
''''		Exit Function
''''	End If
	
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
''''		If Len(Trim(frm1.txtGINo.value)) Then
''''			Call SetPopupMenuItemInf("0000111111")
''''		Else
			Call SetPopupMenuItemInf("0101111111")
''''		End If
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

End Sub

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
''''Sub btnPosting_OnClick()
''''	Dim IntRetCD 
''''	 
''''	If BtnSpreadCheck = False Then Exit Sub
''''	  
''''	Call CheckCreditlimitSvr
''''End Sub

'=========================================
''''Sub btnPostCancel_OnClick()

''''	If BtnSpreadCheck = False Then Exit Sub
''''	Call BatchButton(3)

''''End Sub

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
''''	If frm1.chkARflag.checked = True Then frm1.txtARFlag.value = "Y"
''''	If frm1.chkVatflag.checked = True Then frm1.txtVatFlag.value = "Y"
	    
	Dim strPostVal
	strPostVal = BIZ_PGM_ID & "?txtMode=" & "ARPOST"         
	strPostVal = strPostVal & "&txtHDnNo=" & Trim(frm1.txtHDnNo.value)      
''''	strPostVal = strPostVal & "&txtActualGIDt=" & Trim(frm1.txtActualGIDt.Text)
	strPostVal = strPostVal & "&txtARFlag=" & Trim(frm1.txtARFlag.value)
	strPostVal = strPostVal & "&txtVatFlag=" & Trim(frm1.txtVatFlag.value)
''''	strPostVal = strPostVal & "&txtInvMgr=" & Trim(frm1.txtInvMgr.value)
''''	strPostVal = strPostVal & "&txtGINo=" & Trim(frm1.txtGINo.value)

	Call RunMyBizASP(MyBizASP, strPostVal)             
End Function

'=========================================
''''Sub txtActualGIDt_Change
''''	' ����������� �����Ϻ��� �۰��Էµǵ��� ���� 
'''''	If UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , parent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate , parent.gDateFormat , "") Then
'''''		Call DisplayMsgBox("970024", "X", frm1.txtActualGIDt.ALT, "������")
'''''		Call SetFocusToDocument("M")	
'''''        frm1.txtActualGIDt.Focus
'''''		Exit Sub
'''''	End If
''''
''''	With frm1
''''		If Trim(frm1.txtActualGIDt.text) <> "" Then
''''			Call ggoOper.SetReqAttr(.txtInvMgr, "D")	
''''		Else
''''			Call ggoOper.SetReqAttr(.txtInvMgr, "Q")	
''''		End If
''''		
''''		' ������ �ʿ� ���� ��쳪, ���ǿ� ���ؼ��� ���ó���� ���ÿ� �����ڷḦ ������ �� �� ����.
''''		If Trim(frm1.txtActualGIDt.text) <> "" And Trim(.txtRetBillFlag.value) = "Y" And Trim(.txtExportFlag.value) = "N" Then
''''			Call ggoOper.SetReqAttr(.chkVatFlag, "D")
''''			Call ggoOper.SetReqAttr(.chkARflag, "D")	
''''		Else
''''			Call ggoOper.SetReqAttr(.chkVatFlag, "Q")
''''			Call ggoOper.SetReqAttr(.chkARflag, "Q")
''''		End If
''''	End With
''''
''''	lgBlnFlgChgValue = True
''''End Sub

'=========================================
''''Sub txtActualGIDt_DblClick(Button)
''''	If Button = 1 Then
''''		frm1.txtActualGIDt.Action = 7
''''		Call SetFocusToDocument("M")	
''''        frm1.txtActualGIDt.Focus
''''	End If
''''End Sub

'=======================================================
'   Event Name : chkTaxNo_OnPropertyChange
'   Event Desc : ���ݰ�꼭 �ڵ����� ���ο� ���� �����Է��׸� Change
'=======================================================
''''Sub chkArFlag_OnClick()
''''
''''	on error Resume Next
''''
''''	Select Case frm1.chkArFlag.checked
''''	Case True
''''		lblArFlag.disabled = False
''''	Case False
''''		lblArFlag.disabled = True
''''		lblVatFlag.disabled = True
''''		frm1.chkVatFlag.checked = False
''''	End Select
''''
''''	lgBlnFlgChgValue = True
''''
''''	If Err.number <> 0 Then Err.Clear
''''
''''End Sub

'=====================================================
Sub chkVatFlag_OnClick()

	on error Resume Next
'	Select Case frm1.chkVatFlag.checked
'		Case True
'			lblArFlag.disabled = False
'			lblVatFlag.disabled = False
''''			frm1.chkARflag.checked = True  
'		Case False
'			lblArFlag.disabled = True
'			lblVatFlag.disabled = True
''''			frm1.chkARflag.checked = False
'	End Select

'	lgBlnFlgChgValue = True

'	If Err.number <> 0 Then Err.Clear
 
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
	If Trim(frm1.txtHRetFlag.value) <> "Y" Then
	 '[2002-01-08] : Lot�������ΰ� "Y"�ϰ��, LotNo�� "*"�� �ԷµǸ� �ȵȴ� ///
		If CheckLotNoLotFlag = False Then Exit Function
	End If

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
    on error Resume Next                                                    
End Function

'=====================================================
Function DbDeleteOk()              
    on error Resume Next                                                    
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
''''			.btnPosting.disabled = True
''''			.btnPostCancel.disabled = True
''''			Call ggoOper.SetReqAttr(.txtActualGIDt, "Q")
			frm1.txtConDnNo.focus
		Else
			frm1.vspdData.Focus()
			Call SetQuerySpreadColor(lgLngStartRow)
		End If

		' Scroll ��ȸ�ô� �������� �ʴ´�.
		If lgLngStartRow = 1 Then
			'���/�԰� ó�� ��ǰ���ο� ���� 
''''			If UCase(.txtHRetFlag.value) = "Y" Then
''''				.btnPosting.value = "�԰�ó��"
''''				.btnPostCancel.value = "�԰�ó�����"
''''			Else
''''				.btnPosting.value = "���ó��"
''''				.btnPostCancel.value = "���ó�����"
''''			End If

			' ���ó���� ��� 
''''			If Len(Trim(frm1.txtGINo.value)) Then
''''				Call SetToolbar("11100000000111")          '��: ��ư ���� ���� 
''''			Else
				Call SetToolbar("11101011000111")          '��: ��ư ���� ���� 
''''			End If

			lgBlnFlgChgValue = False
 		End If
	End With
End Function

' ���������� �������� ���� ��� 
'=====================================================
Function DbQueryNotFound()
	Call SetDefaultVal
''''	Call ggoOper.SetReqAttr(frm1.txtActualGIDt, "Q")
	Call SetToolbar("11000000000011")
	frm1.txtConDnNo.focus
End Function

'=====================================================
Function DbSave() 
	on error Resume Next
	
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
	
	Redim iArrdata(31)
	
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
'''' Picking
''''			.Col = C_PickQty		:	iArrData(6) = UNIConvNum(.Text, 0)	' Picking���� 
''''			.Col = C_PickBonusQty	:	iArrData(7) = UNIConvNum(.Text, 0)	' Picking������ 
				.Col = C_PickQty		:	iArrData(6) = 0	' Picking���� 
				.Col = C_PickBonusQty	:	iArrData(7) = 0	' Picking������ 
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
				iArrData(24) = ""			' ext1_cd
				iArrData(25) = "0"			' ext2_qty
				iArrData(26) = "0"			' ext2_amt
				iArrData(27) = ""			' ext2_cd
				iArrData(28) = "0"			' ext3_qty
				iArrData(29) = "0"			' ext3_amt
				iArrData(30) = ""			' ext3_cd
				.Col = C_CartonNo		:	iArrData(31) = Trim(.Text)				' Carton ��ȣ 
				
				.Col = 0
				' �Է� 
				If .Text = ggoSpread.InsertFlag then
					iLngRowsIns = iLngRowsIns + 1
					Redim Preserve iArrRowsIns(iLngRowsIns)
					iArrRowsIns(iLngRowsIns) = Join(iArrData, parent.gColSep)
				' ���� 
				ElseIf .Text = ggoSpread.UpdateFlag then
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
