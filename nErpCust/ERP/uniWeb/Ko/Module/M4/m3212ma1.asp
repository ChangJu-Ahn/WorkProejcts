
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Detail ��� ASP													*
'*  6. Component List       :																			*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2003/05/19																*
'*  9. Modifier (First)     : Sun-jung Lee
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              :																			*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc ����  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<!--
'============================================  1.1.2 ���� Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 

<Script Language="VBS">
Option Explicit
	

Const BIZ_PGM_QRY_ID = "m3212mb1.asp"			<% '��: �����Ͻ� ���� ASP�� %>
Const BIZ_PGM_SAVE_ID = "m3212mb2.asp"			<% '��: �����Ͻ� ���� ASP�� %>
Const LC_HEADER_ENTRY_ID = "m3211ma1"			<% '��: �����Ͻ� ���� ASP�� %>
		
<!--
'============================================  1.2.2 Global ���� ����  ==================================
-->
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim gblnWinEvent		
	
Dim C_LcSeq 		
Dim C_ItemCd 			
Dim C_ItemNm 
Dim C_Spec			
Dim C_Unit	 		
Dim C_LcQty 		
Dim C_Price 		
Dim C_DocAmt	 	
Dim C_PORemainQty 	
Dim C_HsCd 			
Dim C_HsNm 			
Dim C_PoNo 			
Dim C_PoSeq 		
Dim C_OverTolerance 
Dim C_UnderTolerance
Dim C_BlQty			
Dim C_ChgFlg 		
Dim C_TrackingNo
'��ǰ��ݾװ���� ���� �߰�(2003.05)
Dim C_OrgDocAmt		'��ȭ�� ���� 
Dim C_OrgDocAmt1	'��ȸ�� �ʱⰪ ���� 

Dim lgTotalLcAmt	'ȭ�鿡 �������� �ʴ� ǰ����� �ݾ��հ�	2003.08

'������ ���(2003.04.08)
Dim C_LcQty_Ref							

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)


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
	frm1.vspdData.MaxRows = 0
		
End Function

'=======================================  initSpreadPosVariables()  ========================================
Sub InitSpreadPosVariables() 
	C_LcSeq 		 = 1
	C_ItemCd 		 = 2	
	C_ItemNm 		 = 3	
	C_Spec			 = 4
	C_Unit	 		 = 5
	C_LcQty 	     = 6	
	C_Price 		 = 7
	C_DocAmt	 	 = 8
	C_PORemainQty 	 = 9
	C_HsCd 			 = 10
	C_HsNm 			 = 11
	C_PoNo 			 = 12
	C_PoSeq 		 = 13
	C_OverTolerance  = 14
	C_UnderTolerance = 15
	C_BlQty			 = 16
	C_ChgFlg 		 = 17
	C_TrackingNo	 = 18
	C_OrgDocAmt		 = 19
	C_OrgDocAmt1	 = 20

End Sub
<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
Sub SetDefaultVal()
	Call SetToolbar("1110000000001111")
	frm1.txtOpenDt.text = EndDate
	frm1.txtDocAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	frm1.txtTotItemAmt.text= UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
	frm1.txtLcNo.Focus
	Set gActiveElement = document.activeElement
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
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
Sub InitSpreadSheet()
    With frm1
		Call InitSpreadPosVariables()

		ggoSpread.Source = .vspdData
		ggoSpread.SpreadInit "V20030530",,Parent.gAllowDragDropSpread
			
		.vspdData.ReDraw = False

		.vspdData.MaxCols = C_OrgDocAmt1 + 1
		.vspdData.MaxRows = 0
		.vspdData.Col = .vspdData.MaxCols:  .vspdData.ColHidden = True
			
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetEdit		C_LcSeq, "L/C����", 10, 2
		ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 18, 0
		ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 20, 0
		ggoSpread.SSSetEdit		C_Spec, "ǰ��԰�", 20, 0
		ggoSpread.SSSetEdit		C_Unit, "����", 10, 2
		SetSpreadFloatLocal		C_LcQty,  "L/C����", 15, 1, 3
		SetSpreadFloatLocal		C_Price, "�ܰ�", 15, 1, 4
		SetSpreadFloatLocal		C_DocAmt, "�ݾ�", 15, 1, 2
		SetSpreadFloatLocal		C_PORemainQty,  "�����ܷ�", 15, 1, 3
		ggoSpread.SSSetEdit		C_HsCd, "HS��ȣ", 20, 0
		ggoSpread.SSSetEdit		C_HsNm, "HS��", 20, 0
		ggoSpread.SSSetEdit		C_PoNo, "���ֹ�ȣ", 18, 0
		ggoSpread.SSSetEdit		C_PoSeq, "���ּ���", 10, 2
		SetSpreadFloatLocal		C_OverTolerance, "�����������(+)", 15, 1, 5
		SetSpreadFloatLocal		C_UnderTolerance, "�����������(-)", 15, 1, 5
		SetSpreadFloatLocal		C_BlQty, "BlQty",15, 1, 3
		ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 5, 0
		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.",  15,,,25,2
		SetSpreadFloatLocal		C_OrgDocAmt, "C_OrgDocAmt",15,1,2
		SetSpreadFloatLocal		C_OrgDocAmt1, "C_OrgDocAmt1",15,1,2

		Call ggoSpread.SSSetColHidden(C_BlQty, C_ChgFlg, True)
		Call ggoSpread.SSSetColHidden(C_OrgDocAmt, C_OrgDocAmt1, True)
		Call SetSpreadLock()
			
		.vspdData.ReDraw = True
	End With
End Sub

<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
			
		'.vspdData.ReDraw = False
		ggoSpread.SpreadLock -1,-1
	    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
'			ggoSpread.SpreadLock C_ItemCd,		-1,			C_ItemCd,		-1
'			ggoSpread.SpreadLock C_ItemNm,		-1,			C_ItemNm,		-1
'			ggoSpread.SpreadLock C_Unit,			-1,			C_Unit,			-1
		ggoSpread.SpreadUnLock C_LcQty,		-1,			C_LcQty,		-1
'			ggoSpread.SpreadLock C_PORemainQty,  -1,			C_PORemainQty,	-1
		ggoSpread.SpreadUnLock C_Price,		-1,			C_Price,		-1
		ggoSpread.SpreadUnLock C_DocAmt,		-1,			C_DocAmt,		-1			
'			ggoSpread.SpreadLock C_HsCd,			-1,			C_HsCd,			-1
'			ggoSpread.SpreadLock C_HsNm,			-1,			C_HsNm,			-1
'			ggoSpread.SpreadLock C_LcSeq,		-1,			C_LcSeq,		-1
'			ggoSpread.SpreadLock C_PoNo,			-1,			C_PoNo,			-1
'			ggoSpread.SpreadLock C_PoSeq,		-1,			C_PoSeq,		-1
			
		if .vspdData.MaxRows > 0 then
			.vspdData.Col = C_BlQty
				
			if .vspdData.Text > 0 then
				ggoSpread.SpreadLock C_OverTolerance,		-1,		C_OverTolerance,	-1
				ggoSpread.SpreadLock C_UnderTolerance,		-1,		C_UnderTolerance,   -1
			else
				ggoSpread.SpreadUnLock C_OverTolerance,		-1,		C_OverTolerance,	-1
				ggoSpread.SpreadUnLock C_UnderTolerance,		-1,		C_UnderTolerance,   -1
			end if
		else
				ggoSpread.SpreadUnLock C_OverTolerance,		-1,		C_OverTolerance,	-1
				ggoSpread.SpreadUnLock C_UnderTolerance,		-1,		C_UnderTolerance,   -1
		end if
			
'			ggoSpread.SpreadLock C_ChgFlg,		-1,			C_ChgFlg,		-1
		
		'.vspdData.ReDraw = True
	End With
End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		.ReDraw = False

	    ggoSpread.SSSetProtected frm1.vspddata.maxcols, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_LcQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PORemainQty, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	 C_Price, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired	 C_DocAmt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_HsCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_HsNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LcSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoSeq, pvStartRow, pvEndRow
		.Row = pvEndRow
		.Col = C_BlQty
		if .Text > 0 then
			ggoSpread.SSSetProtected C_OverTolerance, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_UnderTolerance, pvStartRow, pvEndRow
		end if
		ggoSpread.SSSetProtected C_ChgFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow
	
		.ReDraw = True
	End With
End Sub
'===================================  GetSpreadColumnPos()  ======================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_LcSeq 		 = iCurColumnPos(1)
			C_ItemCd 		 = iCurColumnPos(2)	
			C_ItemNm 		 = iCurColumnPos(3)	
			C_Spec			 = iCurColumnPos(4)	
			C_Unit	 		 = iCurColumnPos(5)
			C_LcQty 	     = iCurColumnPos(6)	
			C_Price 		 = iCurColumnPos(7)
			C_DocAmt	 	 = iCurColumnPos(8)
			C_PORemainQty 	 = iCurColumnPos(9)
			C_HsCd 			 = iCurColumnPos(10)
			C_HsNm 			 = iCurColumnPos(11)
			C_PoNo 			 = iCurColumnPos(12)
			C_PoSeq 		 = iCurColumnPos(13)
			C_OverTolerance  = iCurColumnPos(14)
			C_UnderTolerance = iCurColumnPos(15)
			C_BlQty			 = iCurColumnPos(16)
			C_ChgFlg 		 = iCurColumnPos(17)
			C_TrackingNo	 = iCurColumnPos(18)	
			C_OrgDocAmt		 = iCurColumnPos(19)
			C_OrgDocAmt1	 = iCurColumnPos(20)
	End Select
End Sub

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCNoPop()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLCNoPop()
	Dim strRet,IntRetCD
	Dim iCalledAspName
		
	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog("m3211pa1.asp", Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtLCNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtLCNo.value = strRet
		frm1.txtLCNo.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPODtlRef()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenPODtlRef()
	Dim arrRet
	Dim arrParam(10)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode <> Parent.OPMD_UMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End if 

	if Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
		Call DisplayMsgBox("173421", "X", "X", "X")
		Exit Function
	End if 

	arrParam(0) = UCase(Trim(frm1.hdnPoNo.value))
	arrParam(1) = UCase(Trim(frm1.hdnPayMethCd.Value))
	arrParam(2) = Trim(frm1.hdnPayMethNm.Value)
	arrParam(3) = UCase(Trim(frm1.hdnIncotermsCd.Value))
	arrParam(4) = Trim(frm1.hdnIncotermsNm.Value)
	arrParam(5) = UCase(Trim(frm1.txtCurrency.Value))
	arrParam(6) = UCase(Trim(frm1.txtBeneficiary.Value))
	arrParam(7) = Trim(frm1.txtBeneficiaryNm.Value)
	arrParam(8) = UCase(Trim(frm1.hdnGrpCd.Value))
	arrParam(9) = Trim(frm1.hdnGrpNm.Value)
	arrParam(10)= "M"
		
	If gblnWinEvent = True Then Exit Function
		
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("M3112RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False
		
	If arrRet(0, 0) = "" Then
		frm1.txtLCNo.focus
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
	Dim TempRow, I, j, intEndRow, Row1
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt
	Dim strtemp1,strtemp2
	Dim strMessage
	Dim intInsertRowsCount,intStartRowNo
		
	Const C_REF_ItemCd			= 0
	Const C_REF_ItemNm			= 1
	Const C_REF_LcQty			= 2
	Const C_REF_ItemSpec		= 3
	Const C_REF_Unit			= 4
	Const C_REF_Price			= 5
	Const C_REF_DocAmt			= 6
	Const C_REF_PoNo			= 7
	Const C_REF_PoSeq			= 8
	Const C_REF_HsCd			= 9
	Const C_REF_HsNm			= 10
	Const C_REF_OverTolerance	= 11
	Const C_REF_UnderTolerance	= 12
	Const C_REF_TrackingNo		= 13
		
	With frm1 
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	

		TempRow = .vspdData.MaxRows		
		
		intStartRowNo=.vspdData.MaxRows	+ 1
		
		intInsertRowsCount = 0 '�ߺ� �ȵɶ��� MAXROW�� 1�� �߰��ϱ� ���Ѻ��� 
		intLoopCnt = Ubound(arrRet, 1)	
		'�ߺ��� ��û�������� MAXROW����� ���� ���� 200308	
		For intCnt = 1 to intLoopCnt + 1
			blnEqualFlg = False
				
			If TempRow <> 0 Then
				For j = 1 To TempRow
					.vspdData.Row = j
					.vspdData.Col = C_PoNo
					strtemp1 = .vspdData.text
					.vspdData.Col = C_PoSeq
					strtemp2 = .vspdData.text
					If strtemp1 = arrRet(intCnt - 1, C_REF_PoNo) and strtemp2 = arrRet(intCnt - 1, C_REF_PoSeq)Then
						strMessage = strMessage & strtemp1 & "-" & strtemp2 & ";"
						blnEqualFlg = True
						intInsertRowsCount = 0		'�ߺ��ɶ� MAXROW�� ������Ű�� ����.					
						Exit For
					Else 
						intInsertRowsCount =  1
					End If
				Next
			Else 
				intInsertRowsCount =  1
			End If

			If blnEqualFlg = False Then
				.vspdData.MaxRows = CLng(TempRow) + CLng(intInsertRowsCount) '����MAXROW���� �ߺ����� ������ 1���� 
				.vspdData.Row = CLng(TempRow) + CLng(intInsertRowsCount)
				
				TempRow = CLng(TempRow) + CLng(intInsertRowsCount) '���� MAXROW���� ���̽��� �� TempRow �� ������Ŵ.
				Row1 = CLng(TempRow) + CLng(intInsertRowsCount)				
				
				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_REF_ItemCd))
				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_REF_ItemNm))
				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_REF_ItemSpec))
				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_REF_Unit))
				Call .vspdData.SetText(C_LcQty,	Row1, arrRet(intCnt - 1, C_REF_LcQty))
				Call .vspdData.SetText(C_Price,	Row1, arrRet(intCnt - 1, C_REF_Price))
				Call .vspdData.SetText(C_DocAmt,	Row1, arrRet(intCnt - 1, C_REF_DocAmt))
				Call .vspdData.SetText(C_PORemainQty,	Row1, arrRet(intCnt - 1, C_REF_LcQty))
				Call .vspdData.SetText(C_HsCd,	Row1, arrRet(intCnt - 1, C_REF_HsCd))
				Call .vspdData.SetText(C_HsNm,	Row1, arrRet(intCnt - 1, C_REF_HsNm))
				Call .vspdData.SetText(C_LcSeq,	Row1, "")
				Call .vspdData.SetText(C_PoNo,	Row1, arrRet(intCnt - 1, C_REF_PoNo))
				Call .vspdData.SetText(C_PoSeq,	Row1, arrRet(intCnt - 1, C_REF_PoSeq))
				'Tolerance Format ���� ����(2003.06.13)
				Call .vspdData.SetText(C_OverTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_REF_OverTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_UnderTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_REF_UnderTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_BlQty,	Row1, 0)
				'Tracking No.�߰�(2003.07.11)
				Call .vspdData.SetText(C_TrackingNo,	Row1, arrRet(intCnt - 1, C_REF_TrackingNo))
					
				Call vspdData_Change(C_LcQty_Ref, .vspdData.Row)		

				'SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
			End If
		Next
		'ȭ�鼺�ɰ���(2003.04.08)-Lee Eun Hee
		intEndRow = .vspdData.MaxRows
		Call SetSpreadColor(intStartRowNo,intEndRow)	
		'Call TotalSum
		call HSumAmtNewCalc
			
		if strMessage<>"" then
			Call DisplayMsgBox("17a005", "X",strmessage,"���ֹ�ȣ" & "," & "���ּ���")
			.vspdData.ReDraw = True
			Exit Function
		End if
			
		.vspdData.ReDraw = True

	End With
End Function
	

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	ggoOper.FormatFieldByObjectOfCur frm1.txtTotItemAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

With frm1

	ggoSpread.Source = frm1.vspdData
	'�ܰ� 
	ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
	'�ݾ� 
	ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt1,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"

End With

End Sub
'==========================================================================================
'   Event Name : changeTag
'==========================================================================================
Sub changeTag()

	frm1.vspdData.Redraw = False

	If Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" Then
		Call ggoSpread.SpreadLock(-1,-1)
		Call SetToolbar("1110000000001111")
	Else
		Call SetSpreadLock()
		Call SetSpreadColor(1, frm1.vspdData.MaxRows)
			
		Call SetToolbar("11101011000111")
	End If

	frm1.vspdData.Redraw = True

End Sub
<!--
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
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
<!--
'=============================================  2.5.1 LoadLCHdr()  ======================================
'=	Event Name : LoadLCHdr																				=
'========================================================================================================
-->
Function LoadLCHdr()
	Dim strDtlOpenParam
	Dim IntRetCD
		
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      <%'Check if there is retrived data%>
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
		
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	WriteCookie "LCNo", UCase(Trim(frm1.txtLCNo.value))

	PgmJump(LC_HEADER_ENTRY_ID)

End Function
	
<!--
'============================================  2.5.1 OpenCookie()  ======================================
'=	Name : OpenCookie()																					=
'=	Description : Open L/C Header ȭ�����κ��� �Ѱܹ��� parameter setting(Cookie ���)				=
'========================================================================================================
-->
Function OpenCookie()
		
	frm1.txtLCNo.Value = ReadCookie("LCNo")		
	frm1.hdnQueryType.Value = "autoQuery"
		
	WriteCookie "LCNo", ""
		
	If UCase(Trim(frm1.txtLCNo.value)) <> "" Then
		Call dbQuery()
	End If
		
End Function

<!--
'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header ȭ�����κ��� �Ѱܹ��� parameter setting(Cookie ���)				=
'========================================================================================================
-->
Sub TotalSum()
	Dim SumTotal, lRow
		
	SumTotal = UNICDbl(frm1.txtTotItemAmt.Text)
	ggoSpread.source = frm1.vspdData
		
	For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text = ggoSpread.InsertFlag then
			frm1.vspdData.Col = C_DocAmt
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	Next
	frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(CStr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

End Sub
'########################################################################################
'============================================  2.5.1 TotalSumNew()  ======================================
'=	Name : TotalSumNew()																					=
'=	Description : Master L/C Header ȭ�����κ��� �Ѱܹ��� parameter setting(Cookie ���)				=
'========================================================================================================
Sub TotalSumNew(ByVal row)
		
    Dim SumTotal, lRow, tmpGrossAmt

	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtTotItemAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_DocAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = C_OrgDocAmt							
	SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))

        
    frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub
'######################################################################################
'==========================================   HSumAmtNewCalc()  ===============================
'	Name : HSumAmtNewCalc()
'	Description : detail �ݾ��� ���Ҷ� ��ȸ�� �Ѿ׺��� Event �ռ� 200308
'==============================================================================================
Function HSumAmtNewCalc()

	Dim iIndex
	Dim SumLcAmt
	Dim LcAmt
	
	SumLcAmt = lgTotalLcAmt
				
	With frm1.vspdData
	
		If .Maxrows >= 1 then 
			For iIndex = 1 to .Maxrows
				.Row = iIndex
				.Col = 0
				If Trim(.text) <> ggoSpread.DeleteFlag then 	
					'��LC�ݾ� 
					.Col = C_DocAmt
					LcAmt	=	 unicdbl(.text)				
											
					SumLcAmt = SumLcAmt + LcAmt
					
				End if
			Next
		Else
			SumLcAmt = 0		
		End if
			
	End with						
	frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(SumLcAmt,frm1.txtCurrency.value,parent.ggAmtOfMoneyNo,"X" , "X")	
	
End Function
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
	
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")		
	Call InitSpreadSheet						
		
	Call SetDefaultVal
	Call InitVariables
	Call OpenCookie()
	'Call SetToolbar("1110110100101111")
		
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	    
End Sub
'===================================  vspdData_Click()  ========================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
   'If Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		If frm1.vspddata.maxRows > 0 Then
			If Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
				Call SetPopupMenuItemInf("0000111111")
			Else
				Call SetPopupMenuItemInf("0101111111")
			End If
		Else 	
			Call SetPopupMenuItemInf("0001111111")
		End If
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
'===================================  vspdData_DblClick()  ========================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
If Row <= 0 Then
	Exit Sub
End If
If frm1.vspddata.MaxRows=0 Then	Exit Sub
	
End Sub
'===================================  vspdData_ColWidthChange()  ========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
ggoSpread.Source = frm1.vspdData
Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'===================================  vspdData_MouseDown()  ========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

If Button = 2 And gMouseClickStatus = "SPC" Then
  gMouseClickStatus = "SPCR"
End If
End Sub    
'===================================  vspdData_ScriptDragDropBlock()  ========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
	Call GetSpreadColumnPos("A")
End Sub

'===================================  FncSplitColumn()  ========================================
Function FncSplitColumn()
    
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Function
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function
<!--
'=========================================  3.2.1 btnLCNoOnClick()  ====================================
-->
Sub btnLCNoOnClick()
	Call OpenLCNoPop()
End Sub	
<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim Qty, Price, DocAmt
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
		
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	Select Case col
	Case C_LcQty, C_Price, C_LcQty_Ref
		frm1.vspdData.Col = C_LcQty
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
			
		DocAmt = Qty * Price
		frm1.vspdData.Col = C_DocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(DocAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
		If col <> C_LcQty_Ref Then
			'Call TotalSumNew(Row)					'��ǰ��ݾ��հ� 
			call HSumAmtNewCalc
		End If
		'�ѱݾװ���� ���� �ʿ�(2003.05)
		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt

	Case  C_DocAmt
		'Call TotalSumNew(Row)					'��ǰ��ݾ��հ� 
		call HSumAmtNewCalc
		'�ѱݾװ���� ���� �ʿ�(2003.05)
		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt
	end select
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
		
	If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    	
	Call InitVariables							

	If Not chkField(Document, "1") Then					
		Exit Function
	End If
	
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
		
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")		
	Call ggoOper.LockField(Document, "N")		
	Call SetDefaultVal
	Call InitVariables							
		
	FncNew = True		
	Set gActiveElement = document.activeElement						
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
		
	ggoSpread.Source = frm1.vspdData
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then			
		Call DisplayMsgBox("900002", "X", "X", "X")
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
		
	ggoSpread.Source = frm1.vspdData  
    
	If ggoSpread.SSCheckChange = False  Then 
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")  
	    Exit Function
	End If

	ggoSpread.Source = frm1.vspdData                 
	If Not ggoSpread.SSDefaultCheck Then     
	   Exit Function
	End If
		
	If DbSave = False Then Exit Function
		
	If frm1.txtHLCNo.value <> frm1.txtLcNo.value then
		frm1.txtLcNo.value =	frm1.txtHLCNo.value
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
		
	lgIntFlgMode = Parent.OPMD_CMODE											

	frm1.vspdData.ReDraw = False

	if frm1.vspdData.Maxrows < 1	then exit function
		
	ggoSpread.Source = frm1.vspdData	
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
	
	
	'�ѱݾװ�����(2003.05.28) -> ���� 200308
	'---------------------------------------------
    'SumTotal = UNICDbl(frm1.txtTotItemAmt.Text)
	'Row = frm1.vspdData.SelBlockRow
		
	'frm1.vspdData.Row = Row
	'frm1.vspdData.Col = C_DocAmt
	'tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	'frm1.vspdData.Col = C_OrgDocAmt1
	'orgtmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	'frm1.vspdData.Col = 0
	'CUDflag = frm1.vspdData.Text
				
    'If CUDflag = ggoSpread.UpdateFlag Then
    '    SumTotal = SumTotal + (orgtmpGrossAmt - tmpGrossAmt )
    'ElseIf CUDflag = ggoSpread.InsertFlag  Then
    '    SumTotal = SumTotal - tmpGrossAmt
    'End If

	'frm1.txtTotItemAmt.Text = SumTotal
	'--------------------------------------------
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo	
	
	Set gActiveElement = document.activeElement	
	
	call HSumAmtNewCalc	
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow()
    Dim IntRetCD
    Dim imRow
		    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
		    
    FncInsertRow = False                                                         '��: Processing is NG
	imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then Exit Function
		    
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
		        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With

	If Err.number = 0 Then FncInsertRow = True                                                          '��: Processing is OK
		    
    Set gActiveElement = document.ActiveElement   
        
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
			
	End With
	Set gActiveElement = document.activeElement
	
	call HSumAmtNewCalc
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
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgPrevNo = "" Then					
		Call DisplayMsgBox("900011", "X", "X", "X")
	End If
	Set gActiveElement = document.activeElement	
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
		
	ggoSpread.Source = frm1.vspdData
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then				
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	ElseIf lgNextNo = "" Then						
		Call DisplayMsgBox("900012", "X", "X", "X")
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
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	    
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(1, frm1.vspdData.MaxRows)
End Sub
<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
			
	Dim IntRetCD
		
	FncExit = False
		
    ggoSpread.Source = frm1.vspdData
	    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")  
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

	if LayerShowHide(1) =false then
	    exit Function
	end if
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		'����(2003.06.10)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		'����(2003.06.10)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
	End If
		
	strVal = strVal & "&txtQueryType=" & Trim(frm1.hdnQueryType.value)
		
    frm1.hdnmaxrow.value = frm1.vspdData.MaxRows
		
	Call RunMyBizASP(MyBizASP, strVal)								
	
	DbQuery = True
		
End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave() 

	Dim lRow
	Dim strVal, strDel
	Dim ColSep, RowSep

	Dim strUnit,strLcQty,strPrice,strDocAmt,strLocAmt,strHsCd,strLcSeq,strPoNo,strPoSeq,strMvmtNo
	Dim strOver,strUnder,strReQty,strBlQty,strTrackingNo
		                         
	
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
		
	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ʱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '�ʱ� ������ ����[����]
    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	if LayerShowHide(1) =false then
	    exit Function
	end if

	With frm1
		.txtMode.value = Parent.UID_M0002
    
		strVal = ""
		strDel = ""
			
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
					
			Select Case .vspdData.Text					'2003.04.21
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag	 'insert/update flg ��ħ.
					if .vspdData.Text=ggoSpread.InsertFlag then
						strVal = "C" & ColSep	
					Else
						strVal = "U" & ColSep
					End if      	
										
					.vspdData.Col = C_LcQty
					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
						Call DisplayMsgBox("970021", "X","L/C����", "X")
						Call SetActiveCell(frm1.vspdData,C_LcQty,lRow,"M","X","X")
						Call LayerShowHide(0)
						Exit Function
					End if
						
					.vspdData.Col = C_Price
					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
						Call DisplayMsgBox("970021", "X","�ܰ�", "X")
						Call SetActiveCell(frm1.vspdData,C_Price,lRow,"M","X","X")
						Call LayerShowHide(0)
						Exit Function
					End if
							
					.vspdData.Col = C_DocAmt				
					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
						Call DisplayMsgBox("970021", "X","�ݾ�", "X")
						Call SetActiveCell(frm1.vspdData,C_DocAmt,lRow,"M","X","X")
						Call LayerShowHide(0)
						Exit Function
					End if

					.vspdData.Col = C_Unit		
					strUnit = Trim(.vspdData.Text)					

					.vspdData.Col = C_LCQty								
					strLcQty = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_Price								
					strPrice = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_DocAmt							
					strDocAmt = UNIConvNum(Trim(.vspdData.Text), 0)
						
					'local amount
					strLocAmt = 0

					.vspdData.Col = C_HsCd								
					strHsCd =  Trim(.vspdData.Text)

					.vspdData.Col = C_LcSeq								
					strLcSeq =  Trim(.vspdData.Text)

					.vspdData.Col = C_PoNo								
					strPoNo = Trim(.vspdData.Text) 

					.vspdData.Col = C_PoSeq								
					strPoSeq = Trim(.vspdData.Text)
					
					'mvmt no				
					strMvmtNo = ""
						
					.vspdData.Col = C_OverTolerance						
					strOver = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_UnderTolerance					
					strUnder= UNIConvNum(Trim(.vspdData.Text), 0)
						
					'receipt qty
					strReQty = 0
						
					.vspdData.Col = C_BlQty					
					strBlQty = UNIConvNum(Trim(.vspdData.Text), 0)
						
					'tracking_no 2003-04 �߰� 
					.vspdData.Col = C_TrackingNo					
					strTrackingNo = Trim(.vspdData.Text)	

					strVal = strVal & strUnit & ColSep & strLcQty & ColSep & strPrice & ColSep & strDocAmt & ColSep & strLocAmt & ColSep & strHsCd & ColSep & strLcSeq & ColSep & strPoNo & ColSep & strPoSeq & ColSep & _   
							strMvmtNo & ColSep & strOver & ColSep & strUnder & ColSep & strReQty & ColSep & strBlQty & ColSep & strTrackingNo & ColSep &lRow & RowSep
													                    
				Case ggoSpread.DeleteFlag	
					strDel = "D" & ColSep
						
					.vspdData.Col = C_Unit		
					strUnit = Trim(.vspdData.Text)					

					.vspdData.Col = C_LCQty								
					strLcQty = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_Price								
					strPrice = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_DocAmt							
					strDocAmt = UNIConvNum(Trim(.vspdData.Text), 0)
						
					'local amount
					strLocAmt = 0

					.vspdData.Col = C_HsCd								
					strHsCd =  Trim(.vspdData.Text)

					.vspdData.Col = C_LcSeq								
					strLcSeq =  Trim(.vspdData.Text)

					.vspdData.Col = C_PoNo								
					strPoNo = Trim(.vspdData.Text) 

					.vspdData.Col = C_PoSeq								
					strPoSeq = Trim(.vspdData.Text)
						
					strMvmtNo = ""
						
					.vspdData.Col = C_OverTolerance						
					strOver = UNIConvNum(Trim(.vspdData.Text), 0)

					.vspdData.Col = C_UnderTolerance					
					strUnder= UNIConvNum(Trim(.vspdData.Text), 0)
						
					'receipt qty
					strReQty = 0
						
					.vspdData.Col = C_BlQty					
					strBlQty = UNIConvNum(Trim(.vspdData.Text), 0)
						
					'tracking_no 2003-04 �߰� 
					.vspdData.Col = C_TrackingNo					
					strTrackingNo = Trim(.vspdData.Text)	
								
					strDel = strDel & strUnit & ColSep & strLcQty & ColSep & strPrice & ColSep & strDocAmt & ColSep & strLocAmt & ColSep & strHsCd & ColSep & strLcSeq & ColSep & strPoNo & ColSep & strPoSeq & ColSep & _   
							strMvmtNo & ColSep & strOver & ColSep & strUnder & ColSep & strReQty & ColSep & strBlQty & ColSep & strTrackingNo & ColSep & lRow & RowSep
									
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
		
	lgIntFlgMode = Parent.OPMD_UMODE						

	lgBlnFlgChgValue = False
	
	'Call TotalSum									

	Call ggoOper.LockField(Document, "Q")			
	
	Call RemovedivTextArea
			
	if frm1.vspdData.MaxRows < 1 then
		Call SetToolbar("11101001000111")
		frm1.txtLcNo.Focus
	else
		Call changeTag()
		frm1.vspdData.focus
	end if
	
	Dim iIndex
	Dim SumLcAmt
	Dim LcAmt
						
	With frm1.vspdData
	
		If .Maxrows >= 1 then 
			For iIndex = 1 to .Maxrows
				.Row = iIndex
				.Col = 0
				If Trim(.text) <> ggoSpread.DeleteFlag then 	
					'��LC�ݾ� 
					.Col = C_DocAmt
					LcAmt	=	 unicdbl(.text)				
											
					SumLcAmt = SumLcAmt + LcAmt
					
				End if
			Next
		Else
			SumLcAmt = 0		
		End if
			
	End with							
	'ȭ�鿡 �Ⱥ��̴� �ݾ��� ����� ���رݾ��� ��.
	lgTotalLcAmt	= unicdbl(frm1.txtTotItemAmt.Text) - SumLcAmt	
			
End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()									
	Call InitVariables
	Call MainQuery()
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()								
	'Call FncNew()
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C ��������</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenPODtlRef()">���ֳ�������</A></TD>
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
										<TD CLASS="TD5" NOWRAP>L/C ������ȣ</TD>
										<TD CLASS="TD6"><INPUT NAME="txtLCNo" ALT="L/C ������ȣ" TYPE="Text" SIZE=30 MAXLENGTH=18  TAG="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" align=top TYPE="BUTTON" onclick="vbscript:btnLCNoOnClick()"></TD>
										<TD CLASS="TD6"></TD>
										<TD CLASS="TD6"></TD>
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
									<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C��ȣ" TYPE=TEXT MAXLENGTH=35 SIZE=28  TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3212ma1_fpDateTime1_txtOpenDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ѱ����ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="��ȭ">&nbsp;&nbsp;
												</TD>
												<TD NOWRAP>
													<script language =javascript src='./js/m3212ma1_fpDoubleSingle5_txtDocAmt.js'></script></TD>
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>��ǰ��ݾ�</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3212ma1_fpDoubleSingle5_txtTotItemAmt.js'></script></TD>
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/m3212ma1_I538522344_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<tr>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</tr>
		<tr HEIGHT="20">
			<td WIDTH="100%">
				<table <%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadLCHdr()">L/C���</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
