<%@ LANGUAGE="VBSCRIPT"%>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212ma2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Detail 등록 ASP											*
'*  6. Component List       :												*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2003/05/19																*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/23 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<!--
'============================================  1.1.2 공통 Include  ======================================
-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 

<Script Language="VBS">
Option Explicit					<% '☜: indicates that All variables must be declared in advance %>
	
<!--
'============================================  1.2.1 Global 상수 선언  ==================================
-->
 Const BIZ_PGM_QRY_ID = "m3212mb5_ko441.asp"	
 Const BIZ_PGM_SAVE_ID = "m3212mb6_ko441.asp"	
 Const LC_HEADER_ENTRY_ID = "m3211ma2"
 Const BIZ_PGM_CAL_AMT_ID = "m3211mb10_ko441.asp"
	
<!--
'============================================  1.2.2 Global 변수 선언  ==================================
-->
 Dim C_ItemCd		
 Dim C_ItemNm
 Dim C_Spec		
 Dim C_Unit			
 Dim C_LCQty			
 Dim C_Price			
 Dim C_DocAmt		
 Dim C_LocAmt		
 Dim C_PORemainQty	
 Dim C_HsCd			
 Dim C_PopUp			
 Dim C_HsNm			
 Dim C_LcSeq			
 Dim C_PoNo			
 Dim C_PoSeq	
 Dim C_MvmtRcptNo		
 Dim C_MvmtNo		
 Dim C_OverTolerance	
 Dim C_UnderTolerance
 Dim C_AfterLCFlg
 Dim C_TrackingNo
 '총품목금액계산을 위해 추가(2003.05)
 Dim C_OrgLocAmt		'변화값 저장 
 Dim C_OrgLocAmt1		'조회후 초기값 저장 
 
 Dim gblnWinEvent	
	
 Dim lgPOorGRFlag
 '참조시 사용(2003.04.08)
 Dim C_LCQty_Ref
	
<!-- #Include file="../../inc/lgvariables.inc" -->

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

 lgPOorGRFlag = ""		
		
End Function
'========================================================================================================
' Name : initSpreadPosVariables()	
'========================================================================================================
 Sub InitSpreadPosVariables() 
 	 C_ItemCd			= 1	
 	 C_ItemNm			= 2
 	 C_Spec				= 3
 	 C_Unit				= 4
 	 C_LCQty			= 5
 	 C_Price			= 6
 	 C_DocAmt			= 7
 	 C_LocAmt			= 8
 	 C_PORemainQty		= 9
 	 C_HsCd				= 10
 	 C_HsNm				= 11
 	 C_LcSeq			= 12
 	 C_PoNo				= 13
 	 C_PoSeq			= 14
 	 C_MvmtRcptNo		= 15
 	 C_MvmtNo			= 16
 	 C_OverTolerance	= 17
 	 C_UnderTolerance	= 18
 	 C_AfterLCFlg		= 19
 	 C_TrackingNo		= 20
 	 C_OrgLocAmt        = 21
 	 C_OrgLocAmt1       = 22
 End Sub

<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
 Sub SetDefaultVal()
 	frm1.txtDocAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 	frm1.txtTotItemAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 	'툴바수정(2003.05.30)
 	Call SetToolbar("1110000000001111")
 	frm1.txtLcNo.focus
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
 		ggoSpread.SpreadInit "V20030530", , Parent.gAllowDragDropSpread
			
 		.vspdData.ReDraw = False

 		.vspdData.MaxCols = C_OrgLocAmt1 + 1
 		.vspdData.MaxRows = 0
 		'.vspdData.Col = C_ChgFlg:    .vspdData.ColHidden = True
			
 		Call GetSpreadColumnPos("A")
			
 		ggoSpread.SSSetEdit		C_ItemCd, "품목", 18, 0
 		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 20, 0
 		ggoSpread.SSSetEdit		C_Spec, "품목규격", 20, 0
 		ggoSpread.SSSetEdit		C_Unit, "단위", 10, 2
 		SetSpreadFloatLocal		C_LCQty,  "L/C수량", 15, 1, 3
 		SetSpreadFloatLocal		C_Price, "단가", 15, 1, 4
 		SetSpreadFloatLocal		C_DocAmt, "금액", 15, 1, 2
 		SetSpreadFloatLocal		C_LocAmt, "원화금액", 15, 1, 2
 		SetSpreadFloatLocal		C_PORemainQty,  "발주잔량", 15, 1, 3
 		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 20, 0
 		ggoSpread.SSSetEdit		C_HsNm, "HS명", 20, 0
 		ggoSpread.SSSetEdit		C_LcSeq, "L/C순번", 10, 2
 		ggoSpread.SSSetEdit		C_PoNo, "발주번호", 18, 0
 		ggoSpread.SSSetEdit		C_PoSeq, "발주순번", 10, 2
 		ggoSpread.SSSetEdit		C_MvmtRcptNo, "입고번호", 18, 0
 		ggoSpread.SSSetEdit		C_MvmtNo, "MVMT_NO",18 , 0
 		SetSpreadFloatLocal		C_OverTolerance, "과부족허용율(+)", 15, 1, 5
 		SetSpreadFloatLocal		C_UnderTolerance, "과부족허용율(-)", 15, 1, 5
 		ggoSpread.SSSetEdit		C_AfterLCFlg, "",5,0
 		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.",  15,,,25,2
		SetSpreadFloatLocal		C_OrgLocAmt, "C_OrgLocAmt", 15, 1, 2
		SetSpreadFloatLocal		C_OrgLocAmt1, "C_OrgLocAmt1", 15, 1, 2
		
 		Call ggoSpread.SSSetColHidden(C_MvmtNo, C_MvmtNo, True)
 		Call ggoSpread.SSSetColHidden(C_AfterLCFlg, C_AfterLCFlg, True)
 		Call ggoSpread.SSSetColHidden(C_OrgLocAmt, C_OrgLocAmt1, True)
 		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
			
 		Call SetSpreadLock()

 		.vspdData.ReDraw = True
 	End With
 End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
 Sub GetSpreadColumnPos(ByVal pvSpdNo)
     Dim iCurColumnPos
	    
     Select Case UCase(pvSpdNo)
        Case "A"
             ggoSpread.Source = frm1.vspdData
             Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
                C_ItemCd			= iCurColumnPos(1)	
 				C_ItemNm			= iCurColumnPos(2)
 				C_Spec				= iCurColumnPos(3)
 				C_Unit				= iCurColumnPos(4)
 				C_LCQty				= iCurColumnPos(5)
 				C_Price				= iCurColumnPos(6)
 				C_DocAmt			= iCurColumnPos(7)
 				C_LocAmt			= iCurColumnPos(8)
 				C_PORemainQty		= iCurColumnPos(9)
 				C_HsCd				= iCurColumnPos(10)
 				C_HsNm				= iCurColumnPos(11)
 				C_LcSeq				= iCurColumnPos(12)
 				C_PoNo				= iCurColumnPos(13)
 				C_PoSeq				= iCurColumnPos(14)
 				C_MvmtRcptNo		= iCurColumnPos(15)
 				C_MvmtNo			= iCurColumnPos(16)
 				C_OverTolerance		= iCurColumnPos(17)
 				C_UnderTolerance	= iCurColumnPos(18)
 				C_AfterLCFlg		= iCurColumnPos(19)
 				C_TrackingNo		= iCurColumnPos(20)
 				C_OrgLocAmt			= iCurColumnPos(21)
 				C_OrgLocAmt1		= iCurColumnPos(22)
 	End Select
 End Sub


<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
 Sub SetSpreadLock()
 	Dim rowCount
     With frm1
 		ggoSpread.Source = .vspdData
			
 		'.vspdData.ReDraw = False
			
 	    ggoSpread.SpreadLock frm1.vspddata.maxcols, -1
 		ggoSpread.SpreadLock C_ItemCd,				-1,				C_ItemCd,			-1
 		ggoSpread.SpreadLock C_ItemNm,				-1,				C_ItemNm,			-1
 		ggoSpread.SpreadLock C_Spec,				-1,				C_Spec,				-1
 		ggoSpread.SpreadLock C_Unit,				-1,				C_Unit,				-1
 		ggoSpread.SpreadUnLock C_LCQty,				-1,				C_LCQty,			-1
 		'ggoSpread.SpreadUnLock C_Price, lRow, -1
 		ggoSpread.SpreadUnLock C_Price,				-1,				C_Price,			-1
 		ggoSpread.SpreadUnLock C_DocAmt,			-1,				C_DocAmt,			-1
 		ggoSpread.SpreadUnLock C_LocAmt,			-1,				C_LocAmt,			-1
 		ggoSpread.SpreadLock C_PORemainQty,			-1,				C_PORemainQty,		-1
 		ggoSpread.SpreadUnLock C_HsCd,				-1,				C_HsCd,				-1
 		ggoSpread.SpreadLock C_HsNm,				-1,				C_HsNm,				-1
 		ggoSpread.SpreadLock C_LcSeq,				-1,				C_LcSeq,			-1
 		ggoSpread.SpreadLock C_PoNo,				-1,				C_PoNo,				-1
 		ggoSpread.SpreadLock C_PoSeq,				-1,				C_PoSeq,			-1
 		ggoSpread.SpreadLock C_MvmtRcptNo,				-1,			C_MvmtRcptNo,			-1
 		ggoSpread.SpreadUnLock	C_OverTolerance,		-1,				C_OverTolerance,	-1
 		ggoSpread.SpreadUnLock	C_UnderTolerance,		-1,				C_UnderTolerance,	-1
 		ggoSpread.SpreadLock	C_MvmtNo,				-1,				C_MvmtNo,			-1
 		ggoSpread.SpreadLocK	C_TrackingNo			-1,				C_TrackingNo		-1

 	End With
 End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
 Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
 	Dim rowCount1
 	ggoSpread.Source = frm1.vspdData

     With frm1.vspdData
	    
 	    ggoSpread.SSSetProtected frm1.vspddata.maxcols, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemCd,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_ItemNm,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_Spec,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_Unit,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetRequired  C_LCQty,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_PORemainQty,		pvStartRow,			pvEndRow
 		'ggoSpread.SSSetRequired C_Price, lRow, lRow
 		ggoSpread.SSSetRequired C_Price,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetRequired C_DocAmt,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetRequired C_LocAmt,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_HsCd,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_HsNm,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_LcSeq,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_PoNo,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_PoSeq,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_PoSeq,			pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_MvmtRcptNo,		pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_OverTolerance,	pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_UnderTolerance,	pvStartRow,			pvEndRow
 		ggoSpread.SSSetProtected C_TrackingNo,		pvStartRow,			pvEndRow

 	End With
 End Sub

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCNoPop()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenLCNoPop()																				+
'+	Description : Master L/C No PopUp Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenLCNoPop()
 	Dim strRet,IntRetCD
 	Dim iCalledAspName
		
 	If gblnWinEvent = True Or UCase(frm1.txtLCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
 	gblnWinEvent = True
		
 	iCalledAspName = AskPRAspName("M3211PA2_KO441")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211PA2_KO441", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	strRet = window.showModalDialog("m3211pa2.asp", Array(window.parent), _
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
'+	Name : OpenPODtlRef()																				+
'+	Description : S/O Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenPODtlRef()
 	Dim arrRet
 	Dim arrParam(8)
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	if lgIntFlgMode <> Parent.OPMD_UMODE then
 		Call DisplayMsgBox("900002", "X", "X", "X")	
 		Exit Function
 	End If
		
 	if Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
 		Call DisplayMsgBox("173421", "X", "X", "X")
 		Exit Function
 	End if 

 	If lgPOorGRFlag = "GR" then 
 		Call DisplayMsgBox("173531", "X", "X", "X")
 		exit Function
 	End if

 	arrParam(0) = Trim(frm1.txtHPurGrp.value)					
 	arrParam(1) = Trim(frm1.txtHPurGrpNm.value)	
 	arrParam(2) = Trim(frm1.txtBeneficiary.value)			
 	arrParam(3) = Trim(frm1.txtBeneficiaryNm.value)
 	arrParam(4) = Trim(frm1.txtCurrency.value) 									
 	arrParam(5) = Trim(frm1.txtHPayTerms.value)	
 	arrParam(6) = Trim(frm1.txtHPayTermsNm.value)	
 	arrParam(7) = Trim(frm1.txtHPONo.value)	
		
 	If gblnWinEvent = True Then Exit Function
		
 	gblnWinEvent = True

 	iCalledAspName = AskPRAspName("M3112RA3")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA3", "X")
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
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenGnDtlRef()  +++++++++++++++++++++++++++++++++++++++
'+	Name : OpenGnDtlRef()																				+
'+	Description : S/O Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenGnDtlRef()
 	Dim arrRet
 	Dim strPONo
 	Dim arrParam(8)
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
 		Call DisplayMsgBox("900002", "X", "X", "X")	
 		Exit Function
 	End If
		
 	if Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
 		Call DisplayMsgBox("173421", "X", "X", "X")
 		Exit Function
 	End if 
		
 	If lgPOorGRFlag = "PO" then 
 		Call DisplayMsgBox("173531", "X", "X", "X")
 		exit Function
 	End if

 	arrParam(0) = Trim(frm1.txtBeneficiary.value)					
 	arrParam(1) = Trim(frm1.txtBeneficiaryNm.value)	
 	arrParam(2) = Trim(frm1.txtHPayTerms.value)			
 	arrParam(3) = Trim(frm1.txtHPayTermsNm.value)	
 	arrParam(4) = Trim(frm1.txtHPurGrp.value)	
 	arrParam(5) = Trim(frm1.txtHPurGrpNm.value)		
 	arrParam(6) = Trim(frm1.txtCurrency.value)
 	arrParam(7) = Trim(frm1.txtHPONo.value) 	
		
 	If gblnWinEvent = True Then Exit Function
		
 	gblnWinEvent = True

 	iCalledAspName = AskPRAspName("M4111RA4")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111RA4", "X")
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
 		Call SetGNDtlRef(arrRet)
 	End If	
 End Function	
<!--
'------------------------------------------  OpenHS()  -------------------------------------------------
-->
Function OpenHS()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function

 gblnWinEvent = True

 arrParam(0) = "HS부호"	
 arrParam(1) = "B_HS_code"
 frm1.vspdData.Col=C_HSCd
 frm1.vspdData.Row=frm1.vspdData.ActiveRow 
 arrParam(2) = Trim(frm1.vspdData.text)
 arrParam(3) = ""
 arrParam(4) = ""			
 arrParam(5) = "HS부호"			
	
 arrField(0) = "HS_CD"	
 arrField(1) = "HS_NM"	
    
 arrHeader(0) = "HS부호"		
 arrHeader(1) = "HS명"		
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
 gblnWinEvent = False
	
 If arrRet(0) = "" Then
 	Exit Function
 Else
 	frm1.vspdData.Col = C_HSCd
 	frm1.vspdData.Text = arrRet(0)
 	frm1.vspdData.Col = C_HSNm
 	frm1.vspdData.Text = arrRet(1)
 	Call vspdData_Change(C_HSCd, frm1.vspdData.Row)		
	
	lgBlnFlgChgValue = True
 End If	
	
End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetPODtlRef()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : SetPODtlRef()																				+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function SetPODtlRef(arrRet)
 	Dim intRtnCnt, strData
 	Dim TempRow, I, j, intEndRow, Row1
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim strMessage
 	Dim	dblPrice, dblQty, dblAmt

 	Const C_Ref_ItemCd			= 0
 	Const C_Ref_ItemNm			= 1
 	Const C_Ref_PORemainQty		= 2
 	Const C_Ref_Spec			= 3
 	Const C_Ref_Unit			= 4
 	Const C_Ref_Price			= 5
 	Const C_Ref_DocAmt			= 6
 	Const C_Ref_PoNo			= 7
 	Const C_Ref_PoSeq			= 8
 	Const C_Ref_HsCd			= 9
 	Const C_Ref_OverTolerance	= 10
 	Const C_Ref_UnderTolerance	= 11
 	Const C_Ref_TrackingNo		= 12
			
 	lgPOorGRFlag = "PO" 

 	With frm1 
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows			
 		intLoopCnt = Ubound(arrRet, 1)		
			
 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_PoSeq
 					If .vspdData.Text = arrRet(intCnt - 1, C_Ref_PoSeq) Then
 					    .vspdData.Row = j
 						.vspdData.Col = C_PoNo
 						If .vspdData.Text = arrRet(intCnt - 1, C_Ref_PoNo) Then
 							strMessage = arrRet(intCnt - 1, C_Ref_PoNo) & "-" & arrRet(intCnt - 1, C_Ref_PoSeq)
 							blnEqualFlg = True
 							Exit For
 						End If
 					End If
 				Next
 			End If

 			If blnEqualFlg = False Then					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = .vspdData.MaxRows	
 				Row1 = .vspdData.Row
					
 				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
 				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
 				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
 				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
 				Call .vspdData.SetText(C_PORemainQty,	Row1, arrRet(intCnt - 1, C_Ref_PORemainQty))
 				Call .vspdData.SetText(C_LCQty,	Row1, arrRet(intCnt - 1, C_Ref_PORemainQty))
 				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
 				Call .vspdData.SetText(C_Price,	Row1, arrRet(intCnt - 1, C_Ref_Price))
 				Call .vspdData.SetText(C_DocAmt,	Row1, arrRet(intCnt - 1, C_Ref_DocAmt))
 				Call .vspdData.SetText(C_PoNo,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
 				Call .vspdData.SetText(C_PoSeq,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
 				Call .vspdData.SetText(C_HsCd,	Row1, arrRet(intCnt - 1, C_Ref_HsCd))
 				'Tolerance Format 오류 수정(2003.06.13)
				Call .vspdData.SetText(C_OverTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_OverTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_UnderTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_UnderTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
 				'Tracking No.추가(2003.07.11)
 				Call .vspdData.SetText(C_TrackingNo,	Row1, arrRet(intCnt - 1, C_REF_TrackingNo))
 				.vspdData.Col = C_LocAmt											
 				.vspdData.text = UNIConvNumPCToCompanyByCurrency(CStr(UNICDbl(.txtHXchRate.value) * Parent.UNICDbl(arrRet(intCnt - 1, C_Ref_DocAmt))),Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
										
 				Call vspdData_Change(C_LCQty_Ref, .vspdData.Row)	
					
'					ChangeSpreadColor CLng(TempRow) + CLng(intCnt)
 				'SetSpreadColor CLng(TempRow) + CLng(intCnt),  CLng(TempRow) + CLng(intCnt)
 			End If
 		Next
 		'툴바수정(2003.05.30)
 		Call SetToolbar("11101011000000")	
			
 		intEndRow = .vspdData.MaxRows
 		Call SetSpreadColor(TempRow+1,intEndRow)	
 		Call TotalSum()
			
 		if strMessage<>"" then

 			Call DisplayMsgBox("17a005", "X",strMessage,"발주번호, 발주순번")
 			.vspdData.ReDraw = True
 			Exit Function
 		End if		
 		.vspdData.ReDraw = True

 	End With
 End Function
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetGNDtlRef()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : SetGNDtlRef()																				+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function SetGNDtlRef(arrRet)
 	Dim intRtnCnt, strData
 	Dim TempRow, I, j, intEndRow, Row1
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim strMessage
 	Dim	dblPrice, dblQty, dblAmt

 	Const C_Ref_MvmtRcptNo		= 0
 	Const C_Ref_ItemCd			= 1
 	Const C_Ref_ItemNm			= 2
 	Const C_Ref_Spec			= 3
 	Const C_Ref_Unit			= 4
 	Const C_Ref_PORemainQty		= 5
 	Const C_Ref_LCQTY			= 6
 	Const C_Ref_MvmtDt			= 7
 	Const C_Ref_Price			= 8
 	Const C_Ref_DocAmt			= 9
 	Const C_Ref_PoNo			= 10
 	Const C_Ref_PoSeq			= 11
 	Const C_Ref_OverTolerance	= 12 
 	Const C_Ref_UnderTolerance	= 13
 	Const C_Ref_MvmtNo			= 14
 	Const C_Ref_TrackingNo		= 15
		
 	lgPOorGRFlag = "GR" 
		
 	With frm1 
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows					
 		intLoopCnt = Ubound(arrRet, 1)				
			
 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_MvmtNo

 					If .vspdData.Text = arrRet(intCnt - 1, C_Ref_MvmtNo) Then
 						strMessage = arrRet(intCnt - 1, C_Ref_MvmtNo) 
 						blnEqualFlg = True
 						Exit For
 					End If
					
 				Next
 			End If

 			If blnEqualFlg = False Then
					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = .vspdData.MaxRows
 				Row1 = .vspdData.Row
					
 				Call .vspdData.SetText(0				,	Row1, ggoSpread.InsertFlag)
 				Call .vspdData.SetText(C_MvmtRcptNo	,	Row1, arrRet(intCnt - 1, C_Ref_MvmtRcptNo))
 				Call .vspdData.SetText(C_ItemCd		,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
 				Call .vspdData.SetText(C_ItemNm		,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
 				Call .vspdData.SetText(C_Spec		,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
 				Call .vspdData.SetText(C_PORemainQty	,	Row1, arrRet(intCnt - 1, C_Ref_PORemainQty))
 				Call .vspdData.SetText(C_LCQty		,	Row1, arrRet(intCnt - 1, C_Ref_LCQTY))
 				Call .vspdData.SetText(C_Unit		,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
 				Call .vspdData.SetText(C_Price		,	Row1, arrRet(intCnt - 1, C_Ref_Price))
 				Call .vspdData.SetText(C_DocAmt		,	Row1, arrRet(intCnt - 1, C_Ref_DocAmt))
 				Call .vspdData.SetText(C_PoNo		,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
 				Call .vspdData.SetText(C_PoSeq		,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
 				Call .vspdData.SetText(C_OverTolerance,	Row1, arrRet(intCnt - 1, C_Ref_OverTolerance))
 				Call .vspdData.SetText(C_UnderTolerance,	Row1, arrRet(intCnt - 1, C_Ref_UnderTolerance))
 				Call .vspdData.SetText(C_MvmtNo		,	Row1, arrRet(intCnt - 1, C_Ref_MvmtNo))
 				'Tracking No.추가(2003.07.11)
 				Call .vspdData.SetText(C_TrackingNo,	Row1, arrRet(intCnt - 1, C_REF_TrackingNo))
					
 				Call vspdData_Change(C_LCQty_Ref, .vspdData.Row)	
					
 				'SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
 				'입고참조시에도 수량 잠그지 않음. 2003.03
 				'ggoSpread.SpreadLock C_LCQty, .vspdData.MaxRows,C_LCQty, .vspdData.MaxRows  
 			End If
 		Next
 		'화면성능개선(2003.04.08)-Lee Eun Hee
 		intEndRow = .vspdData.MaxRows
 		Call SetSpreadColor(TempRow+1,intEndRow)
 		Call TotalSum()
 		Call SetToolbar("11101011000000")
		
 		if strMessage<>"" then

 			Call DisplayMsgBox("17a005", "X",strMessage,"입고번호")
 			.vspdData.ReDraw = True
 			Exit Function
 		End if
 		.vspdData.ReDraw = True

 	End With
 End Function		
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

 With frm1
 	ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
 	'ggoOper.FormatFieldByObjectOfCur .txtTotItemAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
 End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

 With frm1

 	ggoSpread.Source = frm1.vspdData
 	'단가 
 	ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	'금액 
 	ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1,	.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_LocAmt,-1,	Parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgLocAmt,-1, Parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgLocAmt1,-1,Parent.gCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

 End With

End Sub
'==========================================================================================
'   Event Name : changeTag
'==========================================================================================
sub changeTag()
'툴바수정(2003.05.30)	
 frm1.vspdData.Redraw = false
 if Trim(frm1.txtLCAmendSeq.value) <> "" And Trim(frm1.txtLCAmendSeq.value) <> "0" then
 	Call ggoSpread.SpreadLock(-1,-1)
 	Call SetToolbar("1110000000000111")
 else
 	Call SetSpreadLock()
 	Call SetSpreadColor(-1, -1)
 	Call SetToolbar("11101011000101")
 end if

 frm1.vspdData.Redraw = true

end sub
<!--
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
-->
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                 ByVal dColWidth , ByVal HAlign , _
                 ByVal iFlag )
	        
Select Case iFlag
     Case 2                                                              '금액 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
     Case 3                                                              '수량 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
     Case 4                                                              '단가 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
     Case 5                                                              '환율 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
 End Select
         
End Sub
<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
 Function OpenCookie()
 		frm1.hdnQueryType.Value = "autoQuery"
 End Function
<!--
'===========================================  2.5.1 TotalSum()  =========================================
'=	Event Name : TotalSum																				=
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
			frm1.vspdData.Col = C_LocAmt
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	Next
	frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(CStr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

End Sub
'########################################################################################
'============================================  2.5.1 TotalSumNew()  ======================================
'=	Name : TotalSumNew()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
Sub TotalSumNew(ByVal row)
		
    Dim SumTotal, lRow, tmpGrossAmt

	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtTotItemAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_LocAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = C_OrgLocAmt							
	SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))

        
    frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub
'######################################################################################

<!--
'==========================================  2.2.5 ChangeSpreadColor()  =================================
-->
Sub ChangeSpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
	    
		.Redraw = False

		ggoSpread.SSSetProtected C_ItemCd, lRow, lRow
		ggoSpread.SSSetProtected C_ItemNm, lRow, lRow
		ggoSpread.SSSetProtected C_Unit, lRow, lRow
		ggoSpread.SSSetRequired  C_LCQty, lRow, lRow
		ggoSpread.SSSetProtected C_PORemainQty, lRow, lRow
		'ggoSpread.SSSetRequired C_Price, lRow, lRow
		ggoSpread.SSSetRequired C_Price, lRow, lRow
		ggoSpread.SSSetRequired C_DocAmt, lRow, lRow
		ggoSpread.SSSetRequired C_LocAmt, lRow, lRow
		ggoSpread.SSSetProtected C_HsCd, lRow, lRow
		ggoSpread.SSSetProtected C_HsNm, lRow, lRow
		ggoSpread.SSSetProtected C_LcSeq, lRow, lRow
		ggoSpread.SSSetRequired C_PoNo, lRow, lRow
		ggoSpread.SSSetRequired C_PoSeq, lRow, lRow
		ggoSpread.SSSetProtected C_MvmtNo, lRow, lRow
		ggoSpread.SSSetRequired C_OverTolerance, lRow, lRow
		ggoSpread.SSSetRequired C_UnderTolerance, lRow, lRow
			
		.ReDraw = True
	End With
End Sub			
<!--
'===========================================  2.5.1 CookiePage()  =======================================
'=	Event Name : CookiePage																				=
'========================================================================================================
-->
 Function CookiePage(ByVal Kubun)

 	On Error Resume Next

 	Const CookieSplit = 4877					
 	Dim strTemp, arrVal
 	Dim IntRetCD

 	If Kubun = 1 Then

 	    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
 	        Call DisplayMsgBox("900002", "X", "X", "X")
 	        Exit Function
 	    End If
			
 	    If lgBlnFlgChgValue = True Then
 			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
 			If IntRetCD = vbNo Then
 				Exit Function
 			End If
 	    End If
	    
 		WriteCookie CookieSplit , frm1.txtLCNo.value
			
 		Call PgmJump(LC_HEADER_ENTRY_ID)

 	ElseIf Kubun = 0 Then

 		strTemp = ReadCookie(CookieSplit)
				
 		If strTemp = "" then Exit Function
				
 		frm1.txtLCNo.value =  strTemp
			
 		If Err.number <> 0 Then
 			Err.Clear
 			WriteCookie CookieSplit , ""
 			Exit Function 
 		End If
			
 		Call MainQuery()
						
 		WriteCookie CookieSplit , ""
			
 	End If

 End Function	
<!--
'========================================  2.5.1 SettxtTotItemAmt()  ====================================
'=	Event Name : SettxtTotItemAmt																		=
'========================================================================================================
-->
 Sub SettxtTotItemAmt()	
 	frm1.txtTotItemAmt.text = UNIFormatNumber(UNICDbl(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
'		frm1.txtTotItemAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 End Sub	
<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
 Sub Form_Load()
	
 	Call LoadInfTB19029							
 	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.LockField(Document, "N")		
 	Call InitSpreadSheet						
 	Call InitVariables
 	Call OpenCookie()
 	Call SetDefaultVal
 	Call CookiePage(0)	

 End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
 Sub Form_QueryUnload(Cancel, UnloadMode)
	   
 End Sub
	
'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
 Set gActiveSpdSheet = frm1.vspdData
	   
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

'========================================  3.3.2 vspdData_ColWidthChange()  ==================================
 Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
 	ggoSpread.Source = frm1.vspdData
		
 	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
 End Sub
	
'========================================  3.3.2 vspdData_DblClick()  ==================================
 Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then
 		Exit Sub
 	End If
 	If frm1.vspddata.MaxRows=0 Then
 		Exit Sub
 	End If
	
 End Sub
	
'========================= vspdData_MouseDown() ==========================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SPC" Then
    gMouseClickStatus = "SPCR"
 End If
End Sub    

'========================= vspdData_ScriptDragDropBlock() ==================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

 ggoSpread.Source = frm1.vspdData
 Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
 Call GetSpreadColumnPos("A")
End Sub

'======================== FncSplitColumn() ========================================================
Function FncSplitColumn()
    
  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
    Exit Function
 End If

 ggoSpread.Source = gActiveSpdSheet
 ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

'======================= PopSaveSpreadColumnInf() ==================================================
Sub PopSaveSpreadColumnInf()
 ggoSpread.Source = gActiveSpdSheet
 Call ggoSpread.SaveSpreadColumnInf()
End Sub

'======================= PopRestoreSpreadColumnInf() ==================================================
Sub PopRestoreSpreadColumnInf()

 ggoSpread.Source = gActiveSpdSheet
    
 Call ggoSpread.RestoreSpreadInf()
 Call InitSpreadSheet()      
 Call ggoSpread.ReOrderingSpreadData()
 Call SetSpreadColor(1, frm1.vspdData.MaxRows)
End Sub

<!--
'=========================================  3.2.1 btnLCNoOnClick()  ====================================
-->
 Sub btnLCNoOnClick()
 	Call OpenLCNoPop()
 End Sub

<!--
'*********************************************  환율계산  **********************************************
'* Change Event 처리																		*
'********************************************************************************************************
 Sub TxtdblAmt(ByVal Row)

 	DIM RateAmt
        
     frm1.vspdData.Row = Row
 	frm1.vspddata.Col = C_DocAmt

 	IF Trim(frm1.hdnDiv.value) = "*" THEN
 		RateAmt = UNICDbl(frm1.vspdData.Text) * UNICDbl(frm1.txtHXchRate.value)

 		frm1.vspdData.Col = C_LocAmt

 		'frm1.vspdData.Text = UNIFormatNumber(cstr(RateAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(cstr(RateAmt), parent.gCurrency, Parent.ggAmtOfMoneyNo,Parent.gLocRndPolicyNo,"X")
			
 	ELSEIF Trim(frm1.hdnDiv.value) = "/" THEN
			
 		RateAmt = UNICDbl(frm1.vspdData.Text) / UNICDbl(frm1.txtHXchRate.value)

 		frm1.vspdData.Col = C_LocAmt

 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(cstr(RateAmt), parent.gCurrency, Parent.ggAmtOfMoneyNo,Parent.gLocRndPolicyNo,"X")
'			frm1.vspdData.Text = UNIFormatNumber(cstr(RateAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)

 	END IF

 End Sub	
-->
<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
 Sub vspdData_Change(ByVal Col, ByVal Row )
 	Dim dblQty
 	Dim dblPrice
 	Dim dblAmt
 	Dim dblLocAmt
 	Dim Todate
 	Dim strVal

 	ggoSpread.Source = frm1.vspdData
		
 	Select Case Col
 		Case C_LCQty, C_LCQty_Ref
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_LCQty

 			dblQty = frm1.vspdData.Text
				
 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_Price

 			dblPrice = frm1.vspdData.Text
				
 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_DocAmt
 			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(dblAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

 			If frm1.txtCurrency.value = Parent.gCurrency Then
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(dblAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 			Else														
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt

 				Call TxtdblAmt(Row)
 			End If				
 			If Col <> C_LCQty_Ref Then
 				Call TotalSumNew(Row)	
 			End If
 			
 			'총금액계산을 위해 필요(2003.05)
 			frm1.vspdData.Col = C_LocAmt
 			dblLocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgLocAmt		
			frm1.vspdData.Text = dblLocAmt
 		'Case C_Price, C_DocAmt	--- KFC2에서는 C_DocAmt 빠져있음 
				
 		Case C_Price
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col

 			dblPrice = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_LCQty

 			dblQty = frm1.vspdData.Text

 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)
				
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_DocAmt
 			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(dblAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

 			dblAmt = UNICDbl(frm1.vspdData.Text)

 			If frm1.txtCurrency.value = Parent.gCurrency Then
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(dblAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 			Else		
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt

 				Call TxtdblAmt(Row)			

 			End If
							
 			Call TotalSumNew(Row)
 			'총금액계산을 위해 필요(2003.05)
 			frm1.vspdData.Col = C_LocAmt
 			dblLocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgLocAmt		
			frm1.vspdData.Text = dblLocAmt
				
 		Case C_DocAmt
			
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col

 			dblAmt = UNICDbl(frm1.vspdData.Text)

 			If frm1.txtCurrency.value = Parent.gCurrency Then					

 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt
 				'frm1.vspdData.Text = UNIFormatNumber(dblAmt,Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(dblAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 				'frm1.vspdData.Text = UNIFormatNumber(UNICDbl(dblAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 			Else
								
					
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_LocAmt

 				Call TxtdblAmt(Row)
 			End If

 			Call TotalSumNew(Row)
 			'총금액계산을 위해 필요(2003.05)
 			frm1.vspdData.Col = C_LocAmt
 			dblLocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgLocAmt		
			frm1.vspdData.Text = dblLocAmt	
			
 		Case C_LocAmt
 			Call TotalSumNew(Row)
 			'총금액계산을 위해 필요(2003.05)
 			frm1.vspdData.Col = C_LocAmt
 			dblLocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgLocAmt		
			frm1.vspdData.Text = dblLocAmt	
				
 		Case Else

 	End Select
 	ggoSpread.UpdateRow Row

 	lgBlnFlgChgValue = True

 End Sub
<!--
'================ vspdData_ButtonClicked() ========================================================
-->
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 Dim strTemp
 Dim intPos1
	   
 With frm1.vspdData 
	
 ggoSpread.Source = frm1.vspdData
   
 If Row > 0 Then
     .Col = Col
     .Row = Row
 End If
    
 End With
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
'========================================  3.3.3 vspdData_TopLeftChange()  ==============================
-->
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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

 	If lgBlnFlgChgValue = True Then
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

 	If DbQuery = False Then Exit Function
		
 	frm1.hdnQueryType.Value = "Query"
		
 	FncQuery = True	
 	Set gActiveElement = document.activeElement										
 End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
 Function FncNew()
 	Dim IntRetCD 

 	FncNew = False                                          

 	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	Call ggoOper.ClearField(Document, "A")				
 	Call ggoOper.LockField(Document, "N")				
 	Call InitVariables									
 	'Call SetToolbar("1110101100011")		
		
 	Call SetDefaultVal

 	FncNew = True
 	Set gActiveElement = document.activeElement										

 End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
 Function FncDelete()
		
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
    
 	If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  
 	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")           
 	    Exit Function
 	End If

 	ggoSpread.Source = frm1.vspdData              
 	If Not ggoSpread.SSDefaultCheck         Then  
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
	
 	frm1.vspdData.ReDraw = False
 	if frm1.vspdData.Maxrows < 1	then exit function

 	'입고내역참조시에만 복사가능하게 변경 2003.03.
 	frm1.vspdData.Col = C_MvmtRcptNo
 	frm1.vspdData.row = frm1.vspdData.Activerow
		
 	if Trim(frm1.vspdData.text) = "" then 
 		Call DisplayMsgBox("173529", "X", "X", "X")
 		Exit Function
 	End If
		
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
 	'총금액계산수정(2003.05.28)
	'---------------------------------------------
    SumTotal = UNICDbl(frm1.txtTotItemAmt.Text)
	Row = frm1.vspdData.SelBlockRow
		
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_LocAmt
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = C_OrgLocAmt1
	orgtmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = 0
	CUDflag = frm1.vspdData.Text
				
    If CUDflag = ggoSpread.UpdateFlag Then
        SumTotal = SumTotal + (orgtmpGrossAmt - tmpGrossAmt )
    ElseIf CUDflag = ggoSpread.InsertFlag  Then
        SumTotal = SumTotal - tmpGrossAmt
    End If

	frm1.txtTotItemAmt.Text = SumTotal
	'--------------------------------------------
	
 	'수정(2003.04.29)-Lee Eun Hee
	If frm1.vspdData.Maxrows = 1 then lgPOorGRFlag = ""
 	ggoSpread.Source = frm1.vspdData
 	ggoSpread.EditUndo							
 	
 	Set gActiveElement = document.activeElement
 End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
 Function FncInsertRow()
 	Dim IntRetCD
     Dim imRow
		    
     On Error Resume Next                                                          '☜: If process fails
     Err.Clear                                                                     '☜: Clear error status
		    
     FncInsertRow = False                                                         '☜: Processing is NG
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

 	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK
		    
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

 		lgBlnFlgChgValue = True
 	End With
 	Set gActiveElement = document.activeElement
 End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
 Function FncPrint()
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
 End Function

<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
 Function FncPrev() 
	
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
 	Call parent.FncExport(Parent.C_SINGLEMULTI)
 	Set gActiveElement = document.activeElement
 End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
 Function FncFind() 
 	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
 	Set gActiveElement = document.activeElement
 End Function

<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
 Function FncExit()
 	Dim IntRetCD

 	FncExit = False

 	If lgBlnFlgChgValue = True Then
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
		
 	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
 	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)	
 	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
 	strVal = strVal & "&txtQueryType=" & Trim(frm1.hdnQueryType.value)
 	'수정(2003.06.10)
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		
 	Call RunMyBizASP(MyBizASP, strVal)							
	
 	DbQuery = True												
 End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
 Function DbSave() 

 	Dim lRow,ColSep,RowSep
 	Dim strVal,strDel
		
 	Dim strUnit,strLcQty,strPrice,strDocAmt,strLocAmt,strHsCd,strLcSeq,strPoNo,strPoSeq,strMvmtNo
 	Dim strOver,strUnder,strReQty,strBlQty,strTrackingNo
		                         
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size	
		
 	DbSave = False												
		
 	ColSep = Parent.gColSep																
 	RowSep = Parent.gRowSep			
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '초기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '초기 버퍼의 설정[삭제]
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
			
 	'On Error Resume Next
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

 			Select Case .vspdData.Text
				
 				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag	 'insert/update flg 합침.
 					if .vspdData.Text=ggoSpread.InsertFlag then
 						strVal = "C" & ColSep	
 					Else
 						strVal = "U" & ColSep
 					End if      
					
 					.vspdData.Col = C_LCQty
 					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 						Call DisplayMsgBox("970021", "X","L/C수량", "X")
 						Call SetActiveCell(frm1.vspdData,C_LCQty,lRow,"M","X","X")
 						Call LayerShowHide(0)
 						Exit Function
 					End if
		
 					.vspdData.Col = C_Price								
 					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 						Call DisplayMsgBox("970021", "X","단가", "X")
 						Call SetActiveCell(frm1.vspdData,C_Price,lRow,"M","X","X")
 						Call LayerShowHide(0)
 						Exit Function
 					End if
						
 					.vspdData.Col = C_DocAmt							
 					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 						Call DisplayMsgBox("970021", "X","금액", "X")
 						Call SetActiveCell(frm1.vspdData,C_DocAmt,lRow,"M","X","X")
 						Call LayerShowHide(0)
 						Exit Function
 					End if
						
 					.vspdData.Col = C_LocAmt							
 					if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 						Call DisplayMsgBox("970021", "X","원화금액", "X")
 						Call SetActiveCell(frm1.vspdData,C_LocAmt,lRow,"M","X","X")
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
						
 					.vspdData.Col = C_LocAmt
 					strLocAmt = UNIConvNum(Trim(.vspdData.Text), 0)
									
 					.vspdData.Col = C_HsCd								
 					strHsCd =  Trim(.vspdData.Text)

 					.vspdData.Col = C_LcSeq								
 					strLcSeq =  Trim(.vspdData.Text)


 					.vspdData.Col = C_PoNo								
 					strPoNo = Trim(.vspdData.Text) 


 					.vspdData.Col = C_PoSeq								
 					strPoSeq = Trim(.vspdData.Text)

						
 					.vspdData.Col = C_MvmtNo							
 					strMvmtNo = Trim(.vspdData.Text)

						
 					.vspdData.Col = C_OverTolerance						
 					strOver = UNIConvNum(Trim(.vspdData.Text), 0)

 					.vspdData.Col = C_UnderTolerance					
 					strUnder= UNIConvNum(Trim(.vspdData.Text), 0)
						
 					'receipt qty												
 					strReQty = 0

 					'bl qty
 					strBlQty = 0
						
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

 					'mvmt no				
 					strMvmtNo = ""
						
 					.vspdData.Col = C_OverTolerance						
 					strOver = UNIConvNum(Trim(.vspdData.Text), 0)

 					.vspdData.Col = C_UnderTolerance					
 					strUnder= UNIConvNum(Trim(.vspdData.Text), 0)

 					'receipt qty
 					strReQty = 0

 					'.vspdData.Col = C_BlQty					
 					strBlQty = 0
					
 					'tracking_no 2003-04 추가 
 					.vspdData.Col = C_TrackingNo					
 					strTrackingNo = Trim(.vspdData.Text)	
				
 					strDel = strDel & strUnit & ColSep & strLcQty & ColSep & strPrice & ColSep & strDocAmt & ColSep & strLocAmt & ColSep & strHsCd & ColSep & strLcSeq & ColSep & strPoNo & ColSep & strPoSeq & ColSep & _   
 							strMvmtNo & ColSep & strOver & ColSep & strUnder & ColSep & strReQty & ColSep & strBlQty & ColSep & strTrackingNo & ColSep & lRow & RowSep
						
 			End Select
			
			'=====================
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
				                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
				 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
				       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
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

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
				         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select  

			'=====================	
 		Next
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
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
 	Call ggoOper.LockField(Document, "Q")	
	
	Call RemovedivTextArea
		
 	if frm1.vspdData.MaxRows < 1 then
 		Call SetToolbar("11101001000111")
 		frm1.txtLcNo.focus
 	else
 		Call changeTag()
 		with frm1.vspdData
 			.Row = 1
 			.Col = C_MvmtRcptNo			
 			If Trim (.text) <> "" then 
 				lgPOorGRFlag = "GR"
 			Else
 				lgPOorGRFlag = "PO"
 			End if
 			.focus
 		end with			
 	end if
				
 	'Call TotalSum()	'--총금액 계산로직 변경으로 삭제(2003.05)
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
 	Call FncNew()
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>LOCAL L/C내역정보</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenPODtlRef">발주내역참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenGNDtlRef">입고내역참조</TD>
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
										<TD CLASS="TD5" NOWRAP>LOCAL L/C 관리번호</TD>
										<TD CLASS="TD6"><INPUT NAME="txtLcNo" ALT="LOCAL L/C 관리번호" TYPE="Text" Size=29 MAXLENGTH=18  TAG="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCNo" align=top TYPE="BUTTON" onclick="vbscript:btnLCNoOnClick()"></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=27  TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3212ma2_fpDateTime2_txtOpenDt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>수혜자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수혜자">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총개설금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="통화">&nbsp;</TD>
												<TD><script language =javascript src='./js/m3212ma2_fpDoubleSingle1_txtDocAmt.js'></script></TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>총품목원화금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><script language =javascript src='./js/m3212ma2_fpDoubleSingle2_txtTotItemAmt.js'></script></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/m3212ma2_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:CookiePage(1)">LOCAL L/C등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="SpdCount" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurGrp" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurGrpNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHApplicantNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTermsNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHMultiDiv" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
