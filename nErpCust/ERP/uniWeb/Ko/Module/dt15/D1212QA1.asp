<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : D1211MA1
'*  4. Program Name         : ���ڰ�꼭 ����(����) - ��ť�� 
'*  5. Program Desc         : ���ڰ�꼭�� ���Ͽ� ���� �Ǵ� ��������ϴ� ��� 
'*  6. Component List       :  
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 2009/10/31
'*  9. Modifier (First)     : Lee MIn Hyung
'* 10. Modifier (Last)      : Chon, Jaehyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBSCRIPT">

Option Explicit  

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate


iDBSYSDate = "<%=GetSvrDate%>"

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "D1212QB1.asp"
Const BIZ_PGM_ID2 = "D1212QB2.asp"
Const BIZ_PGM_ID3 = "D1212QB3.asp"

'==========================================  1.2.1 Global ��� ����  ======================================
'=                       4.2 Constant variables 
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"

'==========================================================================================================

'add header datatable column
Dim	C1_send_check
Dim	C1_proc_flag_nm 
Dim	C1_is_send_nm 
Dim	C1_inv_no 
Dim	C1_pub_date 
Dim	C1_sup_reg_num 
Dim	C1_sup_cmp_name 
Dim	C1_total_amt 
Dim	C1_sup_tot_amt 
Dim	C1_sur_tax 
Dim	C1_apply_system 
Dim	C1_apply_system_pop
Dim	C1_snd_mgm_no 
Dim	C1_disuse_reason 
Dim	C1_legacy_pk 
Dim	C1_sale_no 
Dim	C1_remark 
Dim	C1_remark2 
Dim	C1_remark3 
Dim	C1_proc_flag 
Dim	C1_is_send 
Dim	C1_proc_date 
Dim	C1_inv_type2 
Dim	C1_old_apply_system


'add detail datatable column
Dim	C2_item
Dim	C2_item_std
Dim	C2_item_prc
Dim	C2_item_qty
Dim	C2_item_date
Dim	C2_item_amt
Dim	C2_item_tax
Dim	C2_item_memo
Dim C2_inv_no

Dim docStatusMake   '�ۼ� 
Dim docStatusIssue 	'���� 
Dim docStatusOpen	'���� 
Dim docStatusReject  	' �ݷ� 
Dim docStatusSalesApprove ' ������� 
Dim docStatusRequestDisuse  ' ����û 
Dim docStatusCancelDisuse 	' ������ 
Dim docStatusDisuse  	' ��� 
Dim docStatusDelete  	' ����	
Dim docStatusApprove	'���� 

Dim sendStatusFail


docStatusMake = "T0" '�ۼ� 
docStatusIssue = "10"	'���� 
docStatusOpen = "50"	'���� 
docStatusReject = "60"	' �ݷ� 
docStatusSalesApprove = "70" ' ������� 
docStatusApprove = "80" ' ���� 
docStatusRequestDisuse = "81"	' ����û 
docStatusCancelDisuse = "82"	' ������ 
docStatusDisuse = "90"	' ��� 
docStatusDelete = "D0"	' ����	

sendStatusFail = "2" '���� 

Dim lgStrPrevKeyTempGlNo
Dim lgStrPrevKeyTempGlDt
Dim lgQueryFlag					' �ű���ȸ �� �߰���ȸ ���� Flag
Dim lgGridPoupMenu              ' Grid Popup Menu Setting
Dim lgAllSelect

Dim lgIsOpenPop
Dim IsOpenPop       
Dim lgPageNo_B
Dim lgSortKey_B
Dim lgOldRow1
Dim lgOldRow, lgRow
Dim lgSortKey1, lgSortKey2
Dim IntRetCD

Dim lblnWinEvent
Dim gblnWinEvent				

Const C_MaxKey = 3
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'========================================================================================================= 
'========================================================================================================= 
Sub Form_Load()
   Call LoadInfTB19029

   With frm1
      Call FormatDATEField(.txtIssuedFromDt)
      Call LockObjectField(.txtIssuedFromDt,"R")
      Call FormatDATEField(.txtIssuedToDt)
      Call LockObjectField(.txtIssuedToDt, "R")
      Call InitSpreadSheet()
      Call InitSpreadSheet2

      Call SetDefaultVal
      Call InitVariables
      Call InitComboBox()
 
      Call SetToolbar("110000000000111")										'��: ��ư ���� ����    	
 
      .txtSupplierCd.focus

   End With		
End Sub

'========================================================================================================= 
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
   '�ڷ�����(Data Type)
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTaxDocumentType, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTransmitStatus, lgF0, lgF1, Chr(11))
	


End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
	C1_send_check	=	1
	C1_proc_flag_nm 	=	2
	C1_is_send_nm 	=	3
	C1_inv_no 	=	4
	C1_pub_date 	=	5
	C1_sup_reg_num 	=	6
	C1_sup_cmp_name 	=	7
	C1_total_amt 	=	8
	C1_sup_tot_amt 	=	9
	C1_sur_tax 	=	10
	C1_apply_system 	=	11
	C1_apply_system_pop 	=	12
	C1_snd_mgm_no 	=	13
	C1_disuse_reason 	=	14
	C1_legacy_pk 	=	15
	C1_sale_no 	=	16
	C1_remark 	=	17
	C1_remark2 	=	18
	C1_remark3 	=	19
	C1_proc_flag 	=	20
	C1_is_send 	=	21
	C1_proc_date 	=	22
	C1_inv_type2 	=	23
	C1_old_apply_system	=	24

End Sub

Sub InitSpreadPosVariables2()
	'add tab1 detail datatable column
	C2_item	=	1
	C2_item_std	=	2
	C2_item_prc	=	3
	C2_item_qty	=	4
	C2_item_date	=	5
	C2_item_amt	=	6
	C2_item_tax	=	7
	C2_item_memo	=	8
	C2_inv_no = 9

 
End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
   
    lgStrPrevKeyTempGlNo = ""							'initializes Previous Key
    lgLngCurRows = 0									   'initializes Deleted Rows Count
    
    lgPageNo_B		= ""                          'initializes Previous Key for spreadsheet #2    
    lgSortKey_B	= "1"
    
    lgOldRow = 0
    lgRow = 0

    lblnWinEvent = False
	gblnWinEvent = False
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	'������ ���ڴ� ������ ���ڸ� ��ȸ�Ѵ�.
    Dim EndDate
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

	'������ ���ڴ� ���� ~ ���� �̴�.
	frm1.txtIssuedFromDt.text  = EndDate
	frm1.txtIssuedToDt.text    = EndDate
	frm1.txtSupplierCd.focus
	'lgGridPoupMenu          = GRID_POPUP_MENU_PRT
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData	
		.MaxCols = C1_old_apply_system + 1								'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols												'��: ����� �� Hidden Column
		.ColHidden = True
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20090708",, parent.gAllowDragDropSpread

		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetCheck	C1_send_check	,"����",	4,  -10, "", True, -1
		ggoSpread.SSSetEdit	C1_proc_flag_nm 	,"��꼭����",	10
		ggoSpread.SSSetEdit	C1_is_send_nm 		,"�۽Ż���",	8
		ggoSpread.SSSetEdit	C1_inv_no 			,"��꼭��ȣ",	15
		ggoSpread.SSSetDate	C1_pub_date 		,"������",	10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit	C1_sup_reg_num		,"�ŷ�ó ����ڹ�ȣ",	15
		ggoSpread.SSSetEdit	C1_sup_cmp_name 	,"�ŷ�ó ����ڸ�",	20
		ggoSpread.SSSetFloat	C1_total_amt 	,"�հ�ݾ�",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_sup_tot_amt 	,"���ް���",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_sur_tax 		,"VAT �ݾ�",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit	C1_apply_system 	,"�ý��� �����ȣ",	15
		ggoSpread.SSSetButton	C1_apply_system_pop
		ggoSpread.SSSetEdit	C1_snd_mgm_no 		,"�߽��� ������ȣ",	15
		ggoSpread.SSSetEdit	C1_disuse_reason 	,"�ݷ�/������",	15, ,,200
		ggoSpread.SSSetEdit	C1_legacy_pk 		,"legacy pk",	15
		ggoSpread.SSSetEdit	C1_sale_no			,"������ȣ",	15
		ggoSpread.SSSetEdit	C1_remark 			,"���1",	15
		ggoSpread.SSSetEdit	C1_remark2 			,"���2",	15
		ggoSpread.SSSetEdit	C1_remark3 			,"���3",	15
		ggoSpread.SSSetEdit	C1_proc_flag 		,"��꼭����",	10
		ggoSpread.SSSetEdit	C1_is_send 			,"�۽Ż���",	10
		ggoSpread.SSSetEdit	C1_proc_date 		,"ó������",	15
		ggoSpread.SSSetEdit	C1_inv_type2 		,"��꼭����",	10
		ggoSpread.SSSetEdit	C1_old_apply_system	,"�ý��� ���Թ�ȣ",	15
		

		Call ggoSpread.MakePairsColumn(C1_apply_system, C1_apply_system_pop, "1")

        Call ggoSpread.SSSetColHidden(C1_send_check, C1_send_check, True)
        Call ggoSpread.SSSetColHidden(C1_proc_flag, C1_old_apply_system, True)


		.ReDraw = True
	End With

	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()
	Call initSpreadPosVariables2()
	With frm1.vspdData2	
		.MaxCols = C2_inv_no + 1								'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols												'��: ����� �� Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		.ReDraw = False 
		ggoSpread.Spreadinit "V20090708",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos2("A")
		ggoSpread.SSSetEdit		C2_item			,"ǰ��",	15
		ggoSpread.SSSetEdit		C2_item_std		,"�԰�",	15
		ggoSpread.SSSetFloat	C2_item_prc		,"�ܰ�",	18, parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C2_item_qty		,"����",	18, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetDate		C2_item_date	,"������",	10, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C2_item_amt		,"���ް���",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C2_item_tax		,"VAT�ݾ�",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit		C2_item_memo	,"���",	15
		ggoSpread.SSSetEdit		C2_inv_no		,"��꼭��ȣ",	15
		Call ggoSpread.SSSetColHidden(C2_inv_no, C2_inv_no, True)

		.ReDraw = True
	End With	
	Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
        ggoSpread.Source = .vspdData
        .vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()	
        .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadLock_B()
    With frm1
        .vspdData2.ReDraw = False
        ggoSpread.Source = .vspdData2
        ggoSpread.SpreadLockWithOddEvenRowColor()
        .vspdData2.ReDraw = True
    End With
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C1_send_check	=	iCurColumnPos(1)
			C1_proc_flag_nm 	=	iCurColumnPos(2)
			C1_is_send_nm 	=	iCurColumnPos(3)
			C1_inv_no 	=	iCurColumnPos(4)
			C1_pub_date 	=	iCurColumnPos(5)
			C1_sup_reg_num 	=	iCurColumnPos(6)
			C1_sup_cmp_name 	=	iCurColumnPos(7)
			C1_total_amt 	=	iCurColumnPos(8)
			C1_sup_tot_amt 	=	iCurColumnPos(9)
			C1_sur_tax 	=	iCurColumnPos(10)
			C1_apply_system 	=	iCurColumnPos(11)
			C1_apply_system_pop 	=	iCurColumnPos(12)
			C1_snd_mgm_no 	=	iCurColumnPos(13)
			C1_disuse_reason 	=	iCurColumnPos(14)
			C1_legacy_pk 	=	iCurColumnPos(15)
			C1_sale_no 	=	iCurColumnPos(16)
			C1_remark 	=	iCurColumnPos(17)
			C1_remark2 	=	iCurColumnPos(18)
			C1_remark3 	=	iCurColumnPos(19)
			C1_proc_flag 	=	iCurColumnPos(20)
			C1_is_send 	=	iCurColumnPos(22)
			C1_proc_date 	=	iCurColumnPos(23)
			C1_inv_type2 	=	iCurColumnPos(24)
			C1_old_apply_system	=	iCurColumnPos(24)

	End Select    
End Sub

'========================================================================================
Sub GetSpreadColumnPos2(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C2_item	=	iCurColumnPos(1)
			C2_item_std	=	iCurColumnPos(2)
			C2_item_prc	=	iCurColumnPos(3)
			C2_item_qty	=	iCurColumnPos(4)
			C2_item_date	=	iCurColumnPos(5)
			C2_item_amt	=	iCurColumnPos(6)
			C2_item_tax	=	iCurColumnPos(7)
			C2_item_memo	=	iCurColumnPos(8)
			C2_inv_no	=	iCurColumnPos(9)

	End Select    
End Sub

'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ŷ�ó"					
	arrParam(1) = "b_biz_partner"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('C', 'CS')"	
	arrParam(5) = "�ŷ�ó"						

	arrField(0) = "bp_cd"					
	arrField(1) = "bp_nm"	
	arrField(2) = "bp_rgst_no"				

	arrHeader(0) = "�ŷ�ó"				
	arrHeader(1) = "�ŷ�ó��"	
	arrHeader(2) = "����ڵ�Ϲ�ȣ"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
											  Array(arrParam, arrField, arrHeader), _
											  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus
	End If

	Set gActiveElement = document.activeElement 
End Function



'------------------------------------------  OpenHistoryRef()  -------------------------------------------------
'	Name : OpenHistoryRef()
'	Description : Altered Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenHistoryRef()
	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("D1211PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "D1211PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C1_inv_no
	
	If Trim(frm1.vspdData.Text) = "" Then
		IntRetCD = DisplayMsgBox("205903", parent.VB_INFORMATION, "X", "X")
		IsOpenPop = False
		Exit Function
	End If
                
	arrParam(0) = Trim(frm1.vspdData.Text)			'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False
	
End Function


'------------------------------------------  OpenBillRef()  -------------------------------------------------
'	Name : OpenBillRef()
'	Description : Altered Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBillRef()
	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName


	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	iCalledAspName = AskPRAspName("D1212PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "D1212PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C1_sale_no
	If Trim(frm1.vspdData.Text) = "" Then
		IntRetCD = DisplayMsgBox("205928", parent.VB_INFORMATION, "X", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	frm1.vspdData.Col = C1_inv_no
                
	arrParam(0) = Trim(frm1.vspdData.Text)			'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0)), _
		"dialogWidth=760px; dialogHeight=640px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

Function OpenIvNo()
	Dim iCalledAspName
	Dim strRet
	
	'If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s5311pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5311pa1", "x")
		gblnWinEvent = False
		exit Function
	end if
	
	strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C1_apply_system
		frm1.vspdData.Text = strRet

        Call vspdData_Change(C1_apply_system, frm1.vspdData.Row)  ' ������ �о�ٰ� �˷��� 
	End If	
End Function

Sub vspdData_Change(Col , Row)
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub
'========================================================================================
Function fnResend()
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt
	Dim strCheckProc

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, is_send_flag, proc_flag
	
	Dim objDTI 
	Dim arrayM, userDN, userInfo, userInfoSet
	
	Dim RetFlag
	
	Dim StrSaveFlag , StrMessageNo
	
	'��: Processing is NG
'	Call LayerShowHide(1)

	StrSaveFlag = "SD"

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDtlVal = ""
		iSelectCnt = 0
		lgAllSelect = False

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then
			
				
				
				.vspdData.Col = C1_is_send
				is_send_flag = Trim(.vspdData.text)
				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)
				

				If Not (is_send_flag = sendStatusFail _
							AND (proc_flag = docStatusOpen or proc_flag = docStatusSalesApprove or proc_flag = docStatusReject or proc_flag = docStatusCancelDisuse  or proc_flag = docStatusDisuse) ) Then
					
					.vspdData.Col = C1_inv_no
					
					Call DisplayMsgBox("205902","X", .vspdData.text, "X")   '�� �ٲ�κ� 
					Call LayerShowHide(0)
					Exit Function
					
				End If

				.vspdData.Col = C1_inv_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_proc_flag :	strDtlVal = strDtlVal & Trim(.vspdData.Text) & parent.gColSep

			End If
		Next
		
		If strVal = "" Then
			Call DisplayMsgBox("181216","X", .vspdData.text, "X")   '�� �ٲ�κ� 
			Call LayerShowHide(0)
			Exit Function
		End If
	
		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal
		.txtDtlSpread.value = strDtlVal

		.txtuserDN.value = userDN
		.txtuserInfo.value = userInfo
		.txtbtnFlag.value = "Resend"
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ��: �����Ͻ� ASP �� ���� 
	End With
End Function

Function fnSalesApproval()
	With frm1
		Call ChangeDocStatus(docStatusSalesApprove, .btnSalesApproval.title)
	End With
End Function

Function fnApprovalDisuse()
	With frm1
		Call ChangeDocStatus(docStatusDisuse, .btnApprovalDisuse.title)
	End With
End Function

Function fnCancelDisuse()
	 
	With frm1
		Call ChangeDocStatus(docStatusCancelDisuse, .btnCancelDisuse.title)
	End With

End Function

Function fnReject()
	 
	With frm1
		Call ChangeDocStatus(docStatusReject, .btnReject.title)
	End With

End Function

'========================================================================================
Function fnOpen()
	With frm1
		Call ChangeDocStatus(docStatusOpen, .btnOpen.title)
	End With
End Function

Sub ChangeDocStatus(pvChangeCode, pvMessageText)
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt, StrSaveFlag
	Dim BlnChkErr
	Dim strErrInvNo

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, is_send_flag, proc_flag
	
	Dim objDTI 
	Dim arrayM, userDN, userInfo, userInfoSet
	StrSaveFlag = "SD"
	BlnChkErr = False

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDtlVal = ""
		iSelectCnt = 0
		lgAllSelect = False

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then
			
				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)
				
				Select Case pvChangeCode
					Case docStatusOpen
						If not proc_flag =  docStatusIssue Then
							BlnChkErr = True
						End If	
					
					Case docStatusSalesApprove
						If Not proc_flag =  docStatusOpen Then
							BlnChkErr = True
						End If

					Case docStatusReject
						If Not proc_flag =  docStatusOpen Then
							BlnChkErr = True
						End If
                        
					Case docStatusDisuse
						If Not proc_flag = docStatusRequestDisuse Then
							BlnChkErr = True
						End If
						
					Case docStatusCancelDisuse
						If Not proc_flag = docStatusRequestDisuse Then
							BlnChkErr = True
						End If

				End Select	
				
				If BlnChkErr = True Then
					
					.vspdData.Col = C1_inv_no
					Call DisplayMsgBox("205906","X", Trim(.vspdData.text), pvMessageText)   '�� �ٲ�κ� 
					Call LayerShowHide(0)
					Exit Sub
				End if	
				
				
				If pvChangeCode = docStatusReject Then
					.vspdData.Col = C1_disuse_reason
					If Trim(.vspdData.text) = "" Then
						.vspdData.Col = C1_inv_no
						Call DisplayMsgBox("205907","X", pvMessageText & " : " & Trim(.vspdData.text), "X")   '�� �ٲ�κ� 
						Call LayerShowHide(0)
						Call SheetFocus(lRow,C1_disuse_reason)
						Exit Sub
					End If
					
				End If
				
				.vspdData.Col = C1_inv_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_disuse_reason 
				If pvChangeCode =docStatusReject Then
					strDtlVal = strDtlVal & Trim(.vspdData.Text) & parent.gColSep
				Else
					strDtlVal = strDtlVal & parent.gColSep	
				End If	
			End If
		Next
		
		If strVal = "" Then
			Call DisplayMsgBox("181216","X", .vspdData.text, "X")   '�� �ٲ�κ� 
			Call LayerShowHide(0)
			Exit Sub
		End If
	
		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal
		.txtDtlSpread.value = strDtlVal

		.txtuserDN.value = userDN
		.txtuserInfo.value = userInfo
		.txtbtnFlag.value = "ChangeDocStatus"
		.txtChangeStatus.value = pvChangeCode
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ��: �����Ͻ� ASP �� ���� 
	End With
End Sub 

'========================================================================================================= 
Sub txtIssuedFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssuedToDt.focus
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub txtIssuedToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtIssuedFromDt.focus
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub txtBizAreaCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		exit sub
	ElseIf KeyAscii = 13 Then 
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub txtBizAreaCd1_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		exit sub
	ElseIf KeyAscii = 13 Then 
		Call FncQuery
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	
	Call SetPopupMenuItemInf("0000111111")

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData

 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		
 		frm1.vspddata.Row = frm1.vspdData.ActiveRow
		frm1.vspddata.Col = C1_inv_no
    
		frm1.vspddata2.MaxRows = 0

		If DbQuery2 = False Then
			Call RestoreToolBar()
			Exit Sub
		End If

		lgOldRow = frm1.vspddata.Row
	Else
		If lgOldRow <> Row Then
 			'------ Developer Coding part (Start)
			frm1.vspddata.Row = Row
			frm1.vspddata.Col = C1_inv_no
			frm1.vspddata2.MaxRows = 0
			
			lgOldRow = Row
			
			If DbQuery2 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
	 		'------ Developer Coding part (End)
	 	End If	
 	End If
End Sub

Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C1_apply_system_pop Then
		    .Col = Col - 1
		    .Row = Row
			Call OpenIvNo()
		End If
    
	End With

	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	 
	If frm1.vspdData.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	 
	Call vspdData_Click(NewCol, NewRow)
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C1_send_check Or NewCol <= C1_send_check Then
        Cancel = True
        Exit Sub
    End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	'----------  Coding part  -------------------------------------------------------------   
	' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData

	End With
End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
Sub txtFromReqDt_DblClick(Button)
'    If Button = 1 Then
'        frm1.txtFromReqDt.Action = 7
'        Call SetFocusToDocument("M")
'    End If
End Sub

Sub txtFromReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub txttoReqDt_DblClick(Button)
'    If Button = 1 Then
'        frm1.txttoReqDt.Action = 7
'        Call SetFocusToDocument("M")
'        frm1.txttoReqDt.focus
'    End If
End Sub

'========================================================================================================= 
Sub txttoReqDt_Change(Button)
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================


'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True													'��: Indicates that value changed
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
		If lgStrPrevKeyTempGlNo <> "" Then                         
			'Call DbQuery("1",frm1.vspddata.row)
		End If
	End If		
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	'��: ������ üũ'
		If lgPageNo_B <> "" Then													'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           'Call DbQuery("2",frm1.vspddata.ActiveRow)
		End If
   End if
End Sub

'#########################################################################################################
'												4. Common Function�� 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False																		'��: Processing is NG

    Err.Clear																				'��: Protect system from crashing
	
    '-----------------------
    'Check previous data area
    '-----------------------
    With frm1
	    ggoSpread.Source = .vspdData
	    If  ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
	    	If IntRetCD = vbNo Then
		      	Exit Function
	    	End If
	    End If

		'-----------------------
	    'Check condition area
	    '-----------------------
		If Not chkFieldByCell(.txtIssuedFromDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.txtIssuedToDt, "A", "1") Then Exit Function

	   If CompareDateByFormat( .txtIssuedFromDt.text, _
										.txtIssuedToDt.text, _
										.txtIssuedFromDt.Alt, _
										.txtIssuedToDt.Alt, _
										"970025", _
										.txtIssuedFromDt.UserDefinedFormat, _
										parent.gComDateType, _
										True) = False Then		
			Exit Function
		End If

		If frm1.txtSupplierCd.value = "" Then
			frm1.txtSupplierNm.value = ""
		End If
		
		'-----------------------
		'Erase contents area
		'-----------------------
		'	    Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'��: Initializes local global variables

		
	End With

	If DBquery = False Then
		Call RestoreToolBar()
		Exit Function
	End If	
	
	FncQuery = True	
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 

	FncNew = False																	'��: Processing is NG

	Err.Clear																			'��: Protect system from crashing
	'On Error Resume Next															'��: Protect system from crashing

	'-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '�� �ٲ�κ�    
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")												'��: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	Call LockObjectField(.txtFromReqDt,"R")
	Call LockObjectField(.txtToReqDt,"R")      				    

	'Call ggoOper.LockField(Document, "N")												'��: Lock  Suitable  Field
	Call SetDefaultVal
	Call InitVariables																	'��: Initializes local global variables

	FncNew = True																		'��: Processing is OK
End Function


'========================================================================================
Function FncSave() 
	Dim IntRetCD 

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

	FncSave = False																		'��: Processing is NG

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then								  '��:match pointer
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")				  '��:There is no changed data.  
        Exit Function
    End If    
 
    IF ggoSpread.SSDefaultCheck = False Then								  '��: Check contents area
		Exit Function
    End If
   
	Call DbSave()  
'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    If Err.number = 0 Then	
       FncSave = True                                                             '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo																	'��: Protect system from crashing    
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()																'��: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)											'��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")								'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtStatusflag

	DbQuery = False

	With frm1

		If .rdoStatusflag1.checked = True Then
			txtStatusflag = ""
		ElseIf .rdoStatusflag2.checked = True Then
			txtStatusflag = .rdoStatusflag2.value
		ElseIf .rdoStatusflag3.checked = True Then
			txtStatusflag = .rdoStatusflag3.value
		End If
		

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&txtSupplierCd=" & Trim(.txtSupplierCd.value) & _
		                      "&cboTaxDocumentType=" & Trim(.cboTaxDocumentType.value) & _
		                      "&cboTransmitStatus=" & Trim(.cboTransmitStatus.value) & _
		                      "&rdoStatusflag=" & Trim(txtStatusflag) & _
		                      "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text) & _
		                      "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)
	
	End With
	
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================
Function DbQuery2() 
	DbQuery2 = False 

	Dim strVal                                                        			'��: Processing is NG
	Dim iTaxBillNo
	Dim strWhereFlag

	Call LayerShowHide(1)

	ggoSpread.Source = frm1.vspdData 
	
	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_inv_no
	iTaxBillNo = frm1.vspddata.Text

	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001 & _
								  "&txtTaxBillNo=" & Trim(iTaxBillNo) 
	
	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery2 = True                                                     
End Function

'========================================================================================
Function DbQueryOk()																		'��: ��ȸ ������ ������� 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode
	
	lgOldRow = 1
	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1
	
	With frm1	
														'��: This function lock the suitable field
		'Call LayerShowHide(0)
		Call SetToolbar("110010000001111")																'��: ��ư ���� ���� 
    
	
		If .vspdData.MaxRows > 0 Then
			
			If Dbquery2 = False Then
				Call RestoreToolbar()
				Exit Function
			End If	
			
			Call SetActiveCell(frm1.vspdData,2,1,"M","X","X")
			Set gActiveElement = document.activeElement
		End If
	End With
End Function

'======================================================================================================
Function SetGridFocus()
	with frm1
		.vspdData.Row = 1
		.vspdData.Col = 1
		.vspdData.Action = 1
	end with 
End Function 

'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt

    DbSave = False              <% '��: Processing is OK %>

    On Error Resume Next            <% '��: Protect system from crashing %>


    If   LayerShowHide(1) = False Then
        Exit Function 
    End If

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDtlVal = ""
		iSelectCnt = 0
		lgAllSelect = False

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0

            Select Case .vspdData.Text
                Case ggoSpread.UpdateFlag

                .vspdData.Col = C1_apply_system
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep  

                .vspdData.Col = C1_old_apply_system
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep  

                .vspdData.Col = C1_inv_no
               strVal = strVal & Trim(.vspdData.Text) & parent.gColSep  

                strVal = strVal & lRow & parent.gRowSep

                lGrpCnt = lGrpCnt + 1
            End Select
        Next
	

		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal

		.txtbtnFlag.value = "Save"
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ��: �����Ͻ� ASP �� ���� 
	End With 
End Function

'========================================================================================
Function SaveResult()
	Call ExecMyBizASP(frm1, BIZ_PGM_ID4)			' ��: �����Ͻ� ASP �� ���� 
End Function

'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 

    lgLngCurRows = 0                            'initializes Deleted Rows Count

	ggoSpread.source = frm1.vspdData
    frm1.vspdData.MaxRows = 0
	
	Call MainQuery
	
End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus2(lRow, lCol)
	Dim pvCnt, pvInputCnt
	
	pvInputCnt = 0
	
	For pvCnt = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = pvCnt
		frm1.vspdData.Col = C1_send_check 
		If frm1.vspdData.text = "1" Then
			If lRow = pvInputCnt Then
				Exit For
			End If
			pvInputCnt = pvInputCnt + 1	
		End if
		
	Next
	frm1.vspdData.focus
	frm1.vspdData.Row = pvCnt
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedFromDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtIssuedFromDt.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtIssuedToDt.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### %>
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���⿪������ȸ</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
									 </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:OpenBillRef()">�ŷ�����</A> | <A href="vbscript:OpenHistoryRef()">�̷���ȸ</A></TD>
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
											<TD CLASS="TD5"NOWRAP>������</TD>
											<TD CLASS="TD6"NOWRAP>
												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT> ~
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT>
 											</TD>
 											<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="�ŷ�ó" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
										</TR>
										<TR>
											<TD CLASS="TD5"NOWRAP>��꼭����</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboTaxDocumentType" ALT="��꼭����" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
											<TD CLASS="TD5"NOWRAP>�۽Ż���</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboTransmitStatus" ALT="�۽Ż���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�ý��۹ݿ�����</TD>
											<TD CLASS=TD6 NOWRAP>
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag1" value="%" tag = "11X" checked><label for="rdoCfmAll">��ü</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag2" value="R" tag = "11X"><label for="rdoCfmreceipt">�ݿ�</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag3" value="D" tag = "11X"><label for="rdoCfmdemand">�̹ݿ�</label>
											</TD>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>
										
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=* valign=top>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR HEIGHT="60%">
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData ID = "A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
									</TR>
									<TR HEIGHT="40%">
										<TD WIDTH="100%" colspan="4"><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vspdData2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
			<TR>
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<TEXTAREA CLASS="hidden" NAME="txtDtlSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserDN" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserInfo" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtbtnFlag" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtChangeStatus" tag="24" TABINDEX="-1">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=280 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
	<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
		<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="date">	
	</Form>
</BODY>
</HTML>
