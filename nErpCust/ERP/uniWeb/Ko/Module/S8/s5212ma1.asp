<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution													    *
'*  2. Function Name        :																			*
'*  3. Program ID           : S5212MA1    																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수출B/L 내역등록 ASP														*
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr		*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk													            *
'* 10. Modifier (Last)      : AHN TAE HEE																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*							  3. 2002/07/04 : VB CONVERSION												*
'*							  3. 2002/11/15 : UI성능 적용												*	
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                 '☜: indicates that All variables must be declared in advance
'========================================================================================================
 Const BIZ_PGM_QRY_ID = "s5212mb1.asp"  
 Const BIZ_PGM_SAVE_ID = "s5212mb1.asp"  
 Const BIZ_PGM_POSTING_ID = "s5212mb1.asp" 
 Const EXBL_HEADER_ENTRY_ID = "s5211ma1"  
 Const EXPORT_CHARGE_ENTRY_ID = "s6111ma1"  
 Const BIZ_BillCollect_JUMP_ID = "s5115ma1"
'========================================================================================================
 Dim C_ItemCd         
 Dim C_ItemNm   
 Dim C_Unit   
 Dim C_Qty    
 Dim C_Price   
 Dim C_DocAmt   
 Dim C_VatIncFlag  
 Dim C_VatType   
 Dim C_VatRate   
 Dim C_VatAmt   
 Dim C_LocAmt   
 Dim C_VatLocAmt  
 Dim C_GrossWeight  
 Dim C_GrossVolume  
 Dim C_NetWeight  
 Dim C_TrackingNo   
 Dim C_Plant   
 Dim C_HsCd   
 Dim C_CcNo   
 Dim C_CcSeq  
 Dim C_SoNo   
 Dim C_SoSeq   
 Dim C_LCNo   
 Dim C_LCSeq   
 Dim C_DNNo   
 Dim C_DNSeq   
 Dim C_BLSeq
 Dim C_Spec   
 Dim C_ChgFlg
    

 Const PostFlag = "PostFlag"
'========================================================================================================
 Dim lgBlnFlgChgValue  
 Dim lgIntGrpCount   
 Dim lgIntFlgMode   

 Dim lgSortKey
 Dim lgStrPrevKey
 Dim lgLngCurRows
 Dim gblnWinEvent   
 Dim IntRetCD
'========================================================================================================
 Dim IsOpenPop
'========================================================================================================
Sub initSpreadPosVariables()

  C_ItemCd  = 1      
  C_ItemNm  = 2
  C_Unit  = 3
  C_Qty   = 4
  C_Price  = 5
  C_DocAmt  = 6
  C_VatIncFlag  = 7
  C_VatType  = 8
  C_VatRate  = 9
  C_VatAmt  = 10
  C_LocAmt  = 11
  C_VatLocAmt = 12
  C_GrossWeight = 13
  C_GrossVolume = 14
  C_NetWeight = 15
  C_TrackingNo  = 16
  C_Plant  = 17 
  C_HsCd  = 18
  C_CcNo  = 19
  C_CcSeq  = 20
  C_SoNo  = 21
  C_SoSeq  = 22
  C_LCNo  = 23
  C_LCSeq  = 24
  C_DNNo  = 25
  C_DNSeq  = 26
  C_BLSeq  = 27
  C_Spec	= 28
  C_ChgFlg  = 29

End Sub  


 Function InitVariables()
  lgIntFlgMode = Parent.OPMD_CMODE    
  lgBlnFlgChgValue = False    
  lgIntGrpCount = 0      
  '---- Coding part--------------------------------------------------------------------
  lgStrPrevKey = ""      
  lgLngCurRows = 0       
  
  gblnWinEvent = False
  Call BtnDisabled(1)
 End Function
'=========================================================================================================
 Sub SetDefaultVal()
  With frm1
   .btnPosting.value = "{{확정}}"
   .btnPosting.disabled = True
   .btnGLView.disabled = True
   .btnPreRcptView.disabled = True
  End With

  lgBlnFlgChgValue = False
 End Sub
'========================================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
    <% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'==========================================================================================================
Sub InitSpreadSheet()
	 
	Call  initSpreadPosVariables()
	With frm1.vspdData
		  
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    
		.ReDraw = False
			   
		.MaxCols = C_ChgFlg
		.MaxRows = 0
			   
		Call GetSpreadColumnPos("A")
		   
		ggoSpread.SSSetEdit  C_BLSeq, "{{B/L순번}}", 10, 0 
		ggoSpread.SSSetEdit  C_ItemCd, "{{품목}}", 18, 0
		ggoSpread.SSSetEdit  C_ItemNm, "{{품목명}}", 25
		ggoSpread.SSSetEdit  C_Spec, "{{품목규격}}", 25
		ggoSpread.SSSetEdit  C_Unit, "{{단위}}", 10, 0
		ggoSpread.SSSetFloat C_Qty,"{{B/L수량}}" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_Price,"{{단가}}",15,Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_DocAmt,"{{금액}}",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_GrossWeight,"{{총중량}}" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_GrossVolume,"{{용적}}" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_NetWeight,"{{순중량}}" ,15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit  C_TrackingNo, "{{Tracking No}}", 18,,,25,2
		ggoSpread.SSSetEdit  C_Plant, "{{공장}}", 10, 0
		ggoSpread.SSSetEdit  C_HsCd, "{{HS부호}}", 20, 0
		'ggoSpread.SSSetEdit  C_BLSeq, "{{B/L순번}}", 10, 1
		ggoSpread.SSSetEdit  C_CcNo, "{{통관번호}}", 18, 0
		ggoSpread.SSSetEdit  C_CcSeq, "{{통관순번}}", 10, 1
		ggoSpread.SSSetEdit  C_SoNo, "{{수주번호}}", 18, 0
		ggoSpread.SSSetEdit  C_SoSeq, "{{수주순번}}", 10, 1
		ggoSpread.SSSetEdit  C_LCNo, "{{L/C번호}}", 18, 0
		ggoSpread.SSSetEdit  C_LcSeq, "{{L/C순번}}", 10, 1
		ggoSpread.SSSetEdit  C_DNNo, "{{출하번호}}", 18, 0
		ggoSpread.SSSetEdit  C_DNSeq, "{{출하순번}}", 10, 1

		ggoSpread.SSSetEdit  C_VatType, "{{VAT유형}}", 10
		ggoSpread.SSSetFloat C_VatRate,"{{VAT율}}",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit  C_VatIncFlag, "{{VAT포함구분}}", 1
		ggoSpread.SSSetFloat C_VatAmt,"{{VAT금액}}",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_LocAmt,"{{자국금액}}",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_VatLocAmt,"{{VAT자국금액}}",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit  C_ChgFlg, "Chgfg", 1, 2

		SetSpreadLock "", 0, -1, ""

		Call ggoSpread.SSSetColHidden(C_VatIncFlag,C_VatIncFlag,True)
	    Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
	    Call ggoSpread.SSSetColHidden(C_VatRate,C_VatRate,True)
	    Call ggoSpread.SSSetColHidden(C_VatAmt,C_VatAmt,True)
	    Call ggoSpread.SSSetColHidden(C_VatLocAmt,C_VatLocAmt, True)
	    Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg, True)
		
		.ReDraw = True
	End With
End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ItemCd       = iCurColumnPos(1)     
            C_ItemNm       = iCurColumnPos(2)
            C_Unit         = iCurColumnPos(3)
            C_Qty          = iCurColumnPos(4)
            C_Price        = iCurColumnPos(5)
            C_DocAmt       = iCurColumnPos(6)
            C_VatIncFlag   = iCurColumnPos(7)
            C_VatType      = iCurColumnPos(8)
            C_VatRate      = iCurColumnPos(9)
            C_VatAmt       = iCurColumnPos(10)
            C_LocAmt       = iCurColumnPos(11)
            C_VatLocAmt    = iCurColumnPos(12)
            C_GrossWeight  = iCurColumnPos(13)
            C_GrossVolume  = iCurColumnPos(14)
            C_NetWeight    = iCurColumnPos(15)
            C_TrackingNo   = iCurColumnPos(16)
            C_Plant        = iCurColumnPos(17)
            C_HsCd         = iCurColumnPos(18)
            C_CcNo         = iCurColumnPos(19)
            C_CcSeq        = iCurColumnPos(20)
            C_SoNo         = iCurColumnPos(21)
            C_SoSeq        = iCurColumnPos(22)
            C_LCNo         = iCurColumnPos(23)
            C_LCSeq        = iCurColumnPos(24)
            C_DNNo         = iCurColumnPos(25)
            C_DNSeq        = iCurColumnPos(26)
            C_BLSeq        = iCurColumnPos(27)
            C_Spec         = iCurColumnPos(28)
            C_ChgFlg       = iCurColumnPos(29)
							
    End Select    
End Sub
'===========================================================================================================
 Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
     With frm1
   ggoSpread.Source = .vspdData
   
   .vspdData.ReDraw = False
   
   ggoSpread.SpreadLock C_BLSeq, lRow, -1
   ggoSpread.SpreadLock C_ItemCd, lRow, -1
   ggoSpread.SpreadLock C_ItemNm, lRow, -1
   ggoSpread.SpreadLock C_Spec, lRow, -1
   ggoSpread.SpreadLock C_Unit, lRow, -1
   ggoSpread.SpreadUnLock C_Qty, lRow, -1
   ggoSpread.SSSetRequired C_Qty, lRow, -1
   ggoSpread.SSSetRequired C_DocAmt, lRow, -1
   '2004.02.10 SMJ
'   ggoSpread.SpreadLock C_LocAmt, lRow, -1
   ggoSpread.SpreadLock C_Price, lRow, -1
   ggoSpread.SpreadLock C_TrackingNo, lRow, -1
   ggoSpread.SpreadLock C_Plant, lRow, -1
   ggoSpread.SpreadLock C_HsCd, lRow, -1
   ggoSpread.SpreadLock C_CcNo, lRow, -1
   ggoSpread.SpreadLock C_CcSeq, lRow, -1
   ggoSpread.SpreadLock C_SoNo, lRow, -1
   ggoSpread.SpreadLock C_SoSeq, lRow, -1
   ggoSpread.SpreadLock C_LCNo, lRow, -1
   ggoSpread.SpreadLock C_LCSeq, lRow, -1
   ggoSpread.SpreadLock C_DNNo, lRow, -1
   ggoSpread.SpreadLock C_DnSeq, lRow, -1
   ggoSpread.SpreadLock C_VatIncflag, lRow, -1
   ggoSpread.SpreadLock C_VatType, lRow, -1
   ggoSpread.SpreadLock C_VatRate, lRow, -1
   ggoSpread.SpreadLock C_VatAmt, lRow, -1
   ggoSpread.SpreadLock C_VatLocAmt, lRow, -1
   
   .vspdData.ReDraw = True
  End With
 End Sub
'===========================================================================================================
 Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
  ggoSpread.Source = frm1.vspdData
     With frm1.vspdData
   .Redraw = False
   ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
   ggoSpread.SSSetRequired  C_Qty, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_Price, pvStartRow, pvEndRow
   ggoSpread.SSSetRequired  C_DocAmt, pvStartRow, pvEndRow
   '2004.02.10 SMJ
'   ggoSpread.SSSetProtected C_LocAmt, pvStartRow, pvEndRow
   
   ggoSpread.SSSetProtected C_VatIncFlag, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_VatType, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_VatRate, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_VatAmt, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_VatLocAmt, pvStartRow, pvEndRow

   ggoSpread.SSSetProtected C_TrackingNo, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_Plant, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_HsCd, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_BLSeq, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_CcNo, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_CcSeq, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_SoNo, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_SoSeq, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_LCNo, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_LCSeq, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_DNNo, pvStartRow, pvEndRow
   ggoSpread.SSSetProtected C_DNSeq, pvStartRow, pvEndRow

   .Col = 1
'   .Row = .ActiveRow
   .Action = 0
   .EditMode = True

   .ReDraw = True
  End With
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenEXBLNoPop()
	Dim iCalledAspName
	Dim strRet
	  
	If gblnWinEvent = True Or UCase(frm1.txtBLNo.className) = "PROTECTED" Then Exit Function
	  
	iCalledAspName = AskPRAspName("s5211pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5211pa1", "x")
		gblnWinEvent = False
		exit Function
	end if
	gblnWinEvent = True
	  
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	  
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetExBLNo(strRet)
		frm1.txtBLNo.focus
	End If 
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSODtRef()
	Dim iCalledAspName  
	Dim strRet
	Dim arrParam(14)

	If Trim(frm1.txtBLNo.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		frm1.txtBLNo.focus
		Exit Function
	End If  

	If RefCheckMessage("S") = False Then Exit Function
	  
	arrParam(0)  = Trim(frm1.txtSONo.value)     
	arrParam(1)  = Trim(frm1.txtApplicant.value) 
	arrParam(2)  = Trim(frm1.txtApplicantNm.value)     
	arrParam(3)  = Trim(frm1.txtSalesGroup.value) 
	arrParam(4)  = Trim(frm1.txtSalesGroupNm.value)   
	arrParam(5)  = Trim(frm1.txtPayTerms.value) 
	arrParam(6)  = Trim(frm1.txtPayTermsNm.value)    
	arrParam(7)  = Trim(frm1.txtCurrency.value) 
	arrParam(8)  = Trim(frm1.txtHBLIssueDt.value)
	arrParam(9)  = Trim(frm1.txtBillType.value) 
	arrParam(10) = Trim(frm1.txtIncoTerms.value) 

	iCalledAspName = AskPRAspName("s3112ra8")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3112ra8", "x")
		gblnWinEvent = False
		exit Function
	end if
	gblnWinEvent = True

	strRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	    
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetSODtlRef(strRet)
	End If 
End Function 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenLCDtRef()
	Dim iCalledAspName
	Dim strRet
	Dim arrParam(14)

	If Trim(frm1.txtBLNo.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		frm1.txtBLNo.focus
		Exit Function
	End If  

	If RefCheckMessage("L") = False Then Exit Function
	   
	arrParam(0)  = Trim(frm1.txtSONo.value)     
	arrParam(1)  = Trim(frm1.txtApplicant.value) 
	arrParam(2)  = Trim(frm1.txtApplicantNm.value)     
	arrParam(3)  = Trim(frm1.txtSalesGroup.value) 
	arrParam(4)  = Trim(frm1.txtSalesGroupNm.value)   
	arrParam(5)  = Trim(frm1.txtPayTerms.value) 
	arrParam(6)  = Trim(frm1.txtPayTermsNm.value)    
	arrParam(7)  = Trim(frm1.txtCurrency.value) 
	arrParam(8)  = Trim(frm1.txtHBLIssueDt.value)
	arrParam(9)  = Trim(frm1.txtBillType.value) 
	arrParam(10) = Trim(frm1.txtIncoTerms.value) 

	iCalledAspName = AskPRAspName("s3212ra8")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3212ra8", "x")
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetLCDtlRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCCDtlRef()
	Dim iCalledAspName
	Dim strRet
	Dim arrParam(15)
	  
	If Trim(frm1.txtBLNo.value) = "" Then  
		Call DisplayMsgBox("900002", "x", "x", "x")
		frm1.txtBLNo.focus
		Exit Function
	End If  
	  
	If RefCheckMessage("CM") = False Then Exit Function
	 
	arrParam(0)  = Trim(frm1.txtSONo.value)     
	arrParam(1)  = Trim(frm1.txtApplicant.value) 
	arrParam(2)  = Trim(frm1.txtApplicantNm.value)     
	arrParam(3)  = Trim(frm1.txtSalesGroup.value) 
	arrParam(4)  = Trim(frm1.txtSalesGroupNm.value)   
	arrParam(5)  = Trim(frm1.txtPayTerms.value) 
	arrParam(6)  = Trim(frm1.txtPayTermsNm.value)    
	arrParam(7)  = Trim(frm1.txtCurrency.value) 
	arrParam(8)  = Trim(frm1.txtHBLIssueDt.value)
	arrParam(9)  = Trim(frm1.txtBillType.value) 
	arrParam(10) = Trim(frm1.txtIncoTerms.value) 
	arrParam(11) = Trim(frm1.txtRefFlg.value)         

	iCalledAspName = AskPRAspName("s4212ra8")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s4212ra8", "x")
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value & "&txtRefFlag=" & frm1.txtRefFlg.value, Array(window.parent,arrParam), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	    
	If strRet(0, 0) <> "" Then
		Call SetCCDtlRef(strRet)
	End If 
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetExBLNo(strRet)
  frm1.txtBLNo.value = strRet(0)
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetSODtlRef(arrRet)
  Dim intRtnCnt, strData
  Dim TempRow, I, j
  Dim blnEqualFlg
  Dim intLoopCnt
  Dim intCnt
  Dim strtemp1, strtemp2, strMessage

  With frm1
   .vspdData.focus
   ggoSpread.Source = .vspdData
   .vspdData.ReDraw = False 

   TempRow = .vspdData.MaxRows  
   intLoopCnt = Ubound(arrRet, 1) 
   
   For intCnt = 1 to intLoopCnt
    blnEqualFlg = False

    If TempRow <> 0 Then
     For j = 1 To TempRow
      .vspdData.Row = j
      .vspdData.Col = C_SoNo
      strtemp1 = .vspdData.text
      
      If .vspdData.Text = arrRet(intCnt - 1, 1) Then
       .vspdData.Row = j
       .vspdData.Col = C_SoSeq
       strtemp2 = .vspdData.text
      
       If .vspdData.Text = arrRet(intCnt - 1, 8) Then
        strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
        blnEqualFlg = True
        Exit For
       End If
      End If
     Next
    End If

    If blnEqualFlg = False Then
     .vspdData.MaxRows = .vspdData.MaxRows + 1

     .vspdData.Row = .vspdData.MaxRows

     .vspdData.Col = 0
     .vspdData.Text = ggoSpread.InsertFlag

     .vspdData.Col = C_SoNo        
     .vspdData.text = arrRet(intCnt - 1, 1)
     .vspdData.Col = C_ItemCd       
     .vspdData.text = arrRet(intCnt - 1, 2)
     .vspdData.Col = C_ItemNm        
     .vspdData.text = arrRet(intCnt - 1, 3)
     .vspdData.Col = C_Qty         
     .vspdData.text = arrRet(intCnt - 1, 4)
     .vspdData.Col = C_Unit         
     .vspdData.text = arrRet(intCnt - 1, 5)
     .vspdData.Col = C_Price         
     .vspdData.text = arrRet(intCnt - 1, 6)
     .vspdData.Col = C_DocAmt        
     .vspdData.text = arrRet(intCnt - 1, 7)
     .vspdData.Col = C_SoSeq        
     .vspdData.text = arrRet(intCnt - 1, 8)
     .vspdData.Col = C_TrackingNo
     .vspdData.text = arrRet(intCnt - 1, 9)
     .vspdData.Col = C_Plant         
     .vspdData.text = arrRet(intCnt - 1, 10)
     .vspdData.Col = C_HsCd         
     .vspdData.text = arrRet(intCnt - 1, 12)
     .vspdData.Col = C_VatType         
     .vspdData.text = arrRet(intCnt - 1, 13)
     .vspdData.Col = C_VatRate         
     .vspdData.text = arrRet(intCnt - 1, 14)
     .vspdData.Col = C_VatIncFlag
     .vspdData.text = arrRet(intCnt - 1, 15)
     .vspdData.Col = C_Spec
     .vspdData.text = arrRet(intCnt - 1, 11)

     .vspdData.Col = C_ChgFlg        
     .vspdData.text = .vspdData.Row

     Call Calcamt(.vspdData.MaxRows)
     Call SetSpreadColor(.vspdData.MaxRows,.vspdData.MaxRows)
     lgBlnFlgChgValue = True
    End If
   Next

   If strMessage <> "" Then Call DisplayMsgBox("17a005", "X",strmessage,"{{수주번호}}" & "," & "{{수주순번}}")

   .vspdData.ReDraw = True
   
   .vspdData.Row = TempRow + 1
   .vspdData.Action = 0

   ' Header의 총금액 계산
   Call SumBlAmtMulti("DL")
  End With
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetLCDtlRef(arrRet)
  Dim intRtnCnt, strData
  Dim TempRow, I, j
  Dim blnEqualFlg
  Dim intLoopCnt
  Dim intCnt
  Dim dblAmt
  Dim strtemp1, strtemp2, strMessage

  With frm1
   .vspdData.focus
   ggoSpread.Source = .vspdData
   .vspdData.ReDraw = False 

   TempRow = .vspdData.MaxRows        
   intLoopCnt = Ubound(arrRet, 1)       
   
   For intCnt = 1 to intLoopCnt
    blnEqualFlg = False

    If TempRow <> 0 Then
     For j = 1 To TempRow
      .vspdData.Row = j
      .vspdData.Col = C_LCNo
      strtemp1 = .vspdData.text
      
      If .vspdData.Text = arrRet(intCnt - 1, 1) Then
       .vspdData.Row = j
       .vspdData.Col = C_LCSeq
       strtemp2 = .vspdData.text
      
       If .vspdData.Text = arrRet(intCnt - 1, 9) Then
        strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
        blnEqualFlg = True
        Exit For
       End If
      End If
     Next
    End If

    If blnEqualFlg = False Then
     .vspdData.MaxRows = .vspdData.MaxRows + 1
     .vspdData.Row = .vspdData.MaxRows


     .vspdData.Col = 0
     .vspdData.Text = ggoSpread.InsertFlag
     
     .vspdData.Col = C_LCNo         
     .vspdData.text = arrRet(intCnt - 1, 1)
     .vspdData.Col = C_ItemCd         
     .vspdData.text = arrRet(intCnt - 1, 3)
     .vspdData.Col = C_ItemNm         
     .vspdData.text = arrRet(intCnt - 1, 4)
     .vspdData.Col = C_Qty        
     .vspdData.text = arrRet(intCnt - 1, 5)
     .vspdData.Col = C_Unit        
     .vspdData.text = arrRet(intCnt - 1, 6)
     .vspdData.Col = C_Price        
     .vspdData.text = arrRet(intCnt - 1, 7)
     .vspdData.Col = C_DocAmt       
     .vspdData.text = arrRet(intCnt - 1, 8)
     .vspdData.Col = C_LCSeq          
     .vspdData.text = arrRet(intCnt - 1, 9)
     .vspdData.Col = C_SoNo        
     .vspdData.text = arrRet(intCnt - 1, 10)
     .vspdData.Col = C_SoSeq        
     .vspdData.text = arrRet(intCnt - 1, 11)
     .vspdData.Col = C_TrackingNo
     .vspdData.text = arrRet(intCnt - 1, 12)
     .vspdData.Col = C_Plant        
     .vspdData.text = arrRet(intCnt - 1, 13)
     .vspdData.Col = C_Spec        
     .vspdData.text = arrRet(intCnt - 1, 14)
     
     .vspdData.Col = C_HsCd        
     .vspdData.text = arrRet(intCnt - 1, 15)
     
     .vspdData.Col = C_VatType         
     .vspdData.text = arrRet(intCnt - 1, 16)
     .vspdData.Col = C_VatRate         
     .vspdData.text = arrRet(intCnt - 1, 17)
     .vspdData.Col = C_VatIncFlag
     .vspdData.text = arrRet(intCnt - 1, 18)

     .vspdData.Col = C_ChgFlg        
     .vspdData.text = .vspdData.Row

     Call Calcamt(.vspdData.MaxRows)
     Call SetSpreadColor(.vspdData.MaxRows,.vspdData.MaxRows)

     lgBlnFlgChgValue = True
    End If
   Next

   If strMessage <> "" Then Call DisplayMsgBox("17a005", "X",strmessage,"{{L/C번호}}" & "," & "{{L/C순번}}")

   .vspdData.ReDraw = True

   .vspdData.Row = TempRow + 1
   .vspdData.Action = 0

   ' Header의 총금액 계산
   Call SumBlAmtMulti("DL")
  End With
 End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 Function SetCCDtlRef(arrRet)
  Dim intRtnCnt, strData
  Dim TempRow, I, j
  Dim blnEqualFlg
  Dim intLoopCnt
  Dim intCnt
  Dim strCCNo
  Dim strCCSeq
  Dim dblAmt, dblLocAmt
  Dim strtemp1, strtemp2, strMessage

  With frm1
   .vspdData.focus
   ggoSpread.Source = .vspdData
   .vspdData.ReDraw = False 

   TempRow = .vspdData.MaxRows       
   intLoopCnt = Ubound(arrRet, 1)      
   
   For intCnt = 1 to intLoopCnt
    blnEqualFlg = False

    If TempRow <> 0 Then
     For j = 1 To TempRow
      .vspdData.Row = j
      .vspdData.Col = C_CcNo
      strtemp1 = .vspdData.text
      
      If .vspdData.Text = arrRet(intCnt - 1, 1) Then
       .vspdData.Row = j
       .vspdData.Col = C_CcSeq
       strtemp2 = .vspdData.text
        
       If .vspdData.Text = arrRet(intCnt - 1, 11) Then 
        strMessage = strMessage & strtemp1 & "-" & strtemp2 & vbCrlf
        blnEqualFlg = True
        Exit For
       End If
      End If
     Next
    End If

    If blnEqualFlg = False Then
     .vspdData.MaxRows = .vspdData.MaxRows + 1
     .vspdData.Row = .vspdData.MaxRows

     .vspdData.Col = 0
     .vspdData.Text = ggoSpread.InsertFlag

     .vspdData.Col = C_CcNo        
     .vspdData.text = arrRet(intCnt - 1, 1)
     .vspdData.Col = C_ItemCd        
     .vspdData.text = arrRet(intCnt - 1, 3)
     .vspdData.Col = C_ItemNm        
     .vspdData.text = arrRet(intCnt - 1, 4)
     .vspdData.Col = C_Qty         
     .vspdData.text = arrRet(intCnt - 1, 5)
     .vspdData.Col = C_Unit         
     .vspdData.text = arrRet(intCnt - 1, 6)
     .vspdData.Col = C_Price           
     .vspdData.text = arrRet(intCnt - 1, 7)
     .vspdData.Col = C_DocAmt          
     .vspdData.text = arrRet(intCnt - 1, 8)
     .vspdData.Col = C_NetWeight         
     .vspdData.text = arrRet(intCnt - 1, 9)
     .vspdData.Col = C_Plant          
     .vspdData.text = arrRet(intCnt - 1, 10)
     .vspdData.Col = C_CcSeq         
     .vspdData.text = arrRet(intCnt - 1, 11)
     .vspdData.Col = C_SoNo          
     .vspdData.text = arrRet(intCnt - 1, 12)
     .vspdData.Col = C_SoSeq          
     .vspdData.text = arrRet(intCnt - 1, 13)
     .vspdData.Col = C_TrackingNo
     .vspdData.text = arrRet(intCnt - 1, 14)
     .vspdData.Col = C_LCNo           
     .vspdData.text = arrRet(intCnt - 1, 15)
     .vspdData.Col = C_LCSeq          
     .vspdData.text = arrRet(intCnt - 1, 16)
     .vspdData.Col = C_DNNo          
     .vspdData.text = arrRet(intCnt - 1, 17)
     .vspdData.Col = C_DNSeq         
     .vspdData.text = arrRet(intCnt - 1, 18)
     .vspdData.Col = C_HsCd          
     .vspdData.text = arrRet(intCnt - 1, 20)
     .vspdData.Col = C_VatType         
     .vspdData.text = arrRet(intCnt - 1, 21)
     .vspdData.Col = C_VatRate         
     .vspdData.text = arrRet(intCnt - 1, 22)
     .vspdData.Col = C_VatIncFlag
     .vspdData.text = arrRet(intCnt - 1, 23)

     .vspdData.Col = C_Spec
     .vspdData.text = arrRet(intCnt - 1, 19)

     .vspdData.Col = C_ChgFlg         
     .vspdData.text = .vspdData.Row

     Call Calcamt(.vspdData.MaxRows)
     Call SetSpreadColor(.vspdData.MaxRows,.vspdData.MaxRows)

     lgBlnFlgChgValue = True
    End If
   Next

   If strMessage <> "" Then Call DisplayMsgBox("17a005", "X",strmessage,"{{통관번호}}" & "," & "{{통관순번}}")

   .vspdData.ReDraw = True

   .vspdData.Row = TempRow + 1
   .vspdData.Action = 0

   ' Header의 총금액 계산
   Call SumBlAmtMulti("DL")
  End With
 End Function
'================================== =====================================================
Sub SumBlAmtMulti(ByVal AmtFlag)

 Dim ldbSumDocAmt, ldbSumVatAmt, ldbDocAmt, ldbVatAmt, lRow
 Dim ldbSumLocAmt, ldbSumVatLocAmt, ldbLocAmt, ldbVatLocAmt
 
 ldbSumDocAmt = 0  '금액총계
 ldbSumVatAmt = 0  'vat금액총계

 ldbSumLocAmt = 0  '자국금액총계
 ldbSumVatLocAmt = 0  'vat자국금액총계
 
 ggoSpread.source = frm1.vspdData
 For lRow = 1 To frm1.vspdData.MaxRows 
  frm1.vspdData.Row = lRow
  frm1.vspdData.Col = 0
  If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
   frm1.vspdData.Col = C_DocAmt : ldbDocAmt = UNICDbl(frm1.vspdData.Text)
   frm1.vspdData.Col = C_VatAmt : ldbVatAmt = UNICDbl(frm1.vspdData.Text)
   frm1.vspdData.Col = C_LocAmt : ldbLocAmt = UNICDbl(frm1.vspdData.Text)
   frm1.vspdData.Col = C_VatLocAmt : ldbVatLocAmt = UNICDbl(frm1.vspdData.Text)

   '부가세포함여부
   frm1.vspdData.col = C_VatIncFlag
   If frm1.vspdData.Text = "1" Then
    ldbSumDocAmt = ldbSumDocAmt + ldbDocAmt
    ldbSumLocAmt = ldbSumLocAmt + ldbLocAmt
   Else
    ldbSumDocAmt = ldbSumDocAmt + ldbDocAmt - ldbVatAmt
    ldbSumLocAmt = ldbSumLocAmt + ldbLocAmt - ldbVatLocAmt
   End if
   
   ldbSumVatAmt = ldbSumVatAmt + ldbVatAmt
   ldbSumVatLocAmt = ldbSumVatLocAmt + ldbVatLocAmt
  End If
 Next

 Select Case AmtFlag
 Case "DL"
  frm1.txtDocAmt.Text = UNIFormatNumberByCurrecny(ldbSumDocAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
  frm1.txtVatAmt.value = uniFormatNumberByTax(ldbSumVatAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
  frm1.txtLocAmt.Text = UNIFormatNumber(ldbSumLocAmt, ggAmtOfMoney.DecPoint, -2, 0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
  frm1.txtVatLocAmt.value = uniFormatNumberByTax(ldbSumVatLocAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo)
 Case "D"
  frm1.txtDocAmt.Text = UNIFormatNumberByCurrecny(ldbSumDocAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
  frm1.txtVatAmt.value = uniFormatNumberByTax(ldbSumVatAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
 Case "L"
  frm1.txtLocAmt.Text = UNIFormatNumber(ldbSumLocAmt, ggAmtOfMoney.DecPoint, -2, 0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
  frm1.txtVatLocAmt.value = uniFormatNumberByTax(ldbSumVatLocAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo)
 End Select
 
End Sub
'================================== =====================================================
Sub CalcDocAmt(ByVal intCol, ByVal intRow)

 ggoSpread.source = frm1.vspdData
 frm1.vspdData.Row = intRow
 frm1.vspdData.Col = 0
 If frm1.vspdData.Text = ggoSpread.DeleteFlag Then Exit Sub

 Dim ldbQty, ldbPrice, ldbDocAmt, ldbVatAmt

 Select Case intCol
 Case C_Qty
  frm1.vspdData.Col = C_Qty : ldbQty = UNICDbl(frm1.vspdData.Text)
  frm1.vspdData.Col = C_Price : ldbPrice = UNICDbl(frm1.vspdData.Text)

  '수량 변경시 B/L금액 재계산
  ldbDocAmt = ldbQty * ldbPrice
  frm1.vspdData.Col = C_DocAmt : frm1.vspdData.Text = UNIFormatNumberByCurrecny(ldbDocAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
   
  ldbDocAmt = UNICDbl(frm1.vspdData.Text)

  '부가세금액 계산
  Call VatAmtOnVatCalsType(C_VATAmt,intRow,ldbDocAmt,"D")

 Case C_DocAmt
  frm1.vspdData.Col = C_DocAmt : ldbDocAmt = UNICDbl(frm1.vspdData.Text)

  Call VatAmtOnVatCalsType(C_VATAmt,intRow,ldbDocAmt,"D")

 End Select

 ' Document금액 변경시 Local Amount 및 Local Vat금액 재계산
 Call CalcLocAmt(intRow, ldbDocAmt)

 ggoSpread.source = frm1.vspdData
 frm1.vspdData.Col = 0
 If frm1.vspdData.Text = ggoSpread.DeleteFlag Then Exit Sub

 ' Head의 총금액 계산
 Call SumBlAmtMulti("DL")

End Sub
'================================== =====================================================
Sub CalcLocAmt(ByVal intRow, ByVal ldbDocAmt)

 Dim ldbLocAmt
 
 frm1.vspdData.Row = intRow

 Select Case Trim(frm1.txtXchgRateOp.value)
 Case "*"
  ldbLocAmt = ldbDocAmt * UNICDbl(Trim(frm1.txtXchgRate.Value)) 
 Case "/"
  ldbLocAmt = ldbDocAmt / UNICDbl(Trim(frm1.txtXchgRate.Value)) 
 Case "+"
  ldbLocAmt = ldbDocAmt + UNICDbl(Trim(frm1.txtXchgRate.Value)) 
 Case "-"
  ldbLocAmt = ldbDocAmt + UNICDbl(Trim(frm1.txtXchgRate.Value)) 
 Case Else
  ldbLocAmt = ldbDocAmt
 End Select

 frm1.vspdData.Col = C_LocAmt : frm1.vspdData.Text = UNIFormatNumber(ldbLocAmt, ggAmtOfMoney.DecPoint, -2, 0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)

 ' 부가세액 계산
 Call VatAmtOnVatCalsType(C_VatLocAmt,intRow,ldbLocAmt,"L")

End Sub

' 금액 변경시 부가세 금액, Loca금액, Local 부가세 금액 계산
Sub CalcAmt(ByVal pvIntRow)

 Dim ldbDocAmt
 
 frm1.vspdData.Row = pvIntRow
 frm1.vspdData.Col = C_DocAmt  : ldbDocAmt = UNICDbl(frm1.vspdData.Text)

 '부가세금액 계산
 Call VatAmtOnVatCalsType(C_VATAmt,pvIntRow,ldbDocAmt,"D")

 ' Local Amount, Local Vat 계산
 Call CalcLocAmt(pvIntRow, ldbDocAmt)
End Sub
'================================== =====================================================
Sub VatAmtOnVatCalsType(ByVal intCol, ByVal intRow, ByVal ldbAmt, ByVal strAmtFlag)
 Err.Clear
 
 On Error Resume next

 Dim ldbVATAmt, ldbVatRate
 Dim strVatIncFlag

 ' 부가세율(품목별로 부가세 관리)
 frm1.vspdData.Row = intRow
 frm1.vspdData.Col = C_VatRate : ldbVatRate = UNICDbl(frm1.vspdData.Text)
 
 ' 부가세포함여부 
 frm1.vspdData.Col = C_VatIncFlag : strVatIncFlag = Trim(frm1.vspdData.Text)

 With frm1

  If strVatIncFlag = "1" Then
   ldbVATAmt = ldbAmt * ldbVatRate * 0.01
  Else
   ldbVATAmt = ldbAmt * ldbVatRate * 0.01 / (1 + ldbVatRate * 0.01)
  End if

  If Err.number <> 0 Then
   Msgbox Err.Description, vbInformation, Parent.gLogoName
   Exit Sub
  End If

  Select Case strAmtFlag
  ' Document Amount
  Case "D"
   .vspdData.Row = intRow : .vspdData.Col = intCol
   .vspdData.Text = uniFormatNumberByTax(ldbVATAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo)
  'Local Amount
  Case "L"
   .vspdData.Row = intRow : .vspdData.Col = intCol
   .vspdData.Text = uniFormatNumberByTax(ldbVATAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo)
  End Select

 End With
 
End Sub
'========================================================================================================
Function CookiePage(ByVal Kubun)

On Error Resume Next

	Const CookieSplit = 4877  
	Dim strTemp, arrVal

	Select Case Kubun
	  
		Case 0
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			frm1.txtBLNo.value =  strTemp
		    
		Call DbQuery()
			WriteCookie CookieSplit , ""
		Case 1
			WriteCookie CookieSplit , frm1.txtHBLNo.value
		'경비등록
		Case 2 
			WriteCookie CookieSplit , "EB" & Parent.gRowSep & frm1.txtSalesGroup.value & Parent.gRowSep & frm1.txtSalesGroupNm.value & Parent.gRowSep & frm1.txtHBLNo.value 
			   
	End Select

End Function
'========================================================================================================
 Function LoadBLHdr()
  Dim strDtlOpenParam

  WriteCookie "txtBLNo", UCase(Trim(frm1.txtBLNo.value))
  
  PgmJump(EXBL_HEADER_ENTRY_ID)
 End Function
'========================================================================================================
 Function OpenCookie()
  frm1.txtBLNo.value = ReadCookie("txtBLNo")

  WriteCookie "txtBLNo", ""
 End Function
'========================================================================================================
 Function PostBL()
  If Trim(frm1.txtHBLNo.value) = "" Then
   Call DisplayMsgBox("900002", "x", "x", "x")
   'Call MsgBox("조회를 선행하십시오.", Parent.VB_INFORMATION)
   Exit Function
  End If

  Dim strVal

  Call LayerShowHide(1)
      
  strVal = BIZ_PGM_POSTING_ID & "?txtMode=" & PostFlag         <%'☜: 비지니스 처리 ASP의 상태 %>
  strVal = strVal & "&txtHBillNo=" & Trim(frm1.txtHBLNo.value)      <%'☜: 조회 조건 데이타%>
  strVal = strVal & "&txtInsrtUserId=" & Parent.gUsrID
  strVal = strVal & "&txtChangeOrgId=" & Parent.gChangeOrgId

  Call RunMyBizASP(MyBizASP, strVal)         
 End Function
'========================================================================================================
 Function PostingOk()
  frm1.rdoPostingflg1.checked = True
  lgBlnFlgChgValue = False
  Call ggoOper.ClearField(Document, "2")    
  Call FncQuery()
 End Function
'========================================================================================================
 Sub ReleaseGrid()
  ggoSpread.Source = frm1.vspdData
     With frm1.vspdData
   .Redraw = False

   ggoSpread.SpreadUnLock  C_Qty, -1, -1
   ggoSpread.SpreadUnLock  C_GrossWeight, -1, -1
   ggoSpread.SpreadUnLock  C_GrossVolume, -1, -1
   ggoSpread.SpreadUnLock  C_NetWeight, -1, -1
   ggoSpread.SpreadUnLock  C_DocAmt, -1, -1
   ggoSpread.SpreadUnLock  C_LocAmt, -1, -1
   ggoSpread.SSSetRequired  C_Qty, -1, -1
   'ggoSpread.SSSetRequired  C_GrossWeight, -1, -1
   'ggoSpread.SSSetRequired  C_GrossVolume, -1, -1
   'ggoSpread.SSSetRequired  C_NetWeight, -1, -1
   ggoSpread.SSSetRequired  C_DocAmt, -1, -1

   .ReDraw = True
  End With
 End Sub
'========================================================================================================
 Sub ProtectGrid()
  ggoSpread.Source = frm1.vspdData
     With frm1.vspdData
   .Redraw = False

   ggoSpread.SSSetProtected  C_Qty, -1, -1
   ggoSpread.SSSetProtected  C_DocAmt, -1, -1
   ggoSpread.SSSetProtected  C_LocAmt, -1, -1
   ggoSpread.SSSetProtected  C_GrossWeight, -1, -1
   ggoSpread.SSSetProtected  C_GrossVolume, -1, -1
   ggoSpread.SSSetProtected  C_NetWeight, -1, -1

   .ReDraw = True
  End With
  
 End Sub 
'============================================================================================================
Function BtnSpreadCheck()

 BtnSpreadCheck = False

 Dim Answer
 <% '변경이 있을떄 저장 여부 먼저 체크후, YES이면 작업진행여부 체크 안한다 %>
 If lgBlnFlgChgValue = True Then Answer = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")  <%'데이타가 변경되었습니다. 계속 하시겠습니까?%>
 If Answer = VBNO Then Exit Function

 <% '변경이 없을때 작업진행여부 체크 %>
 If lgBlnFlgChgValue = False Then Answer = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x") <% '작업을 수행하시겠습니까? %> 
 If Answer = VBNO Then Exit Function

 BtnSpreadCheck = True

End Function
'==================================================================================================== 
Function RefCheckMessage(strRefFlag)

 If frm1.rdoPostingflg1.checked = True Then
  Msgbox "{{이미 확정처리가 되어서 참조 할 수 없습니다}}",vbInformation, Parent.gLogoName
  Exit Function
 End If

 RefCheckMessage = False
 If strRefFlag = "CM" Then
  If Trim(frm1.txtRefFlg.value) <> "C" And Trim(frm1.txtRefFlg.value) <> "M" Then
   Select Case Trim(frm1.txtRefFlg.value)
   Case "L"
    Call DisplayMsgBox("209002", "X", "{{L/C}}", "{{L/C내역참조}}")
    Exit Function
   Case "S"
    Call DisplayMsgBox("209002", "X", "{{수주}}", "{{수주내역참조}}")
    Exit Function
   Case "C", "M"
    Call DisplayMsgBox("209002", "X", "{{통관}}", "{{통관내역참조}}")
    Exit Function
   End Select
  End If
 ElseIf strRefFlag <> Trim(frm1.txtRefFlg.value) Then
  Select Case Trim(frm1.txtRefFlg.value)
  Case "L"
   Call DisplayMsgBox("209002", "X", "{{L/C}}", "{{L/C내역참조}}")
   Exit Function
  Case "S"
   Call DisplayMsgBox("209002", "X", "{{수주}}", "{{수주내역참조}}")
   Exit Function
  Case "C", "M"
   Call DisplayMsgBox("209002", "X", "{{통관}}", "{{통관내역참조}}")
   Exit Function
  End Select
 End If

 RefCheckMessage = True

End Function
'========================================================================================================
 Function HideNonCCGrid()
  With frm1
   ggoSpread.Source = .vspdData 
   .vspdData.ReDraw = False

   Select Case frm1.txtRefFlg.value 
   Case "S"
		Call ggoSpread.SSSetColHidden(C_CcNo,C_CcSeq,True)
		Call ggoSpread.SSSetColHidden(C_SoNo,C_SoSeq,False)
		Call ggoSpread.SSSetColHidden(C_LCNo,C_DNSeq,True)
   Case "L"
		Call ggoSpread.SSSetColHidden(C_CcNo,C_CcSeq,True)
		Call ggoSpread.SSSetColHidden(C_SoNo,C_LcSeq,False)
		Call ggoSpread.SSSetColHidden(C_DNNo,C_DNSeq,True)
   Case "C"   
		Call ggoSpread.SSSetColHidden(C_CcNo,C_DNSeq,False)
   Case "M"   
		Call ggoSpread.SSSetColHidden(C_CcNo,C_CcSeq,False)
		Call ggoSpread.SSSetColHidden(C_SoNo,C_DNSeq,True)
   Case Else
  
   End Select
   .vspdData.ReDraw = True
  End With
 End Function
'=========================================================================== 
Function JumpChgCheck(Byval pvIntCookieFlag, Byval pvStrJumpFlag)

 Dim IntRetCD

 ggoSpread.Source = frm1.vspdData 
 If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
  If IntRetCD = vbNo Then
   Exit Function
  End If
 End If

 Call CookiePage(pvIntCookieFlag)
 Call PgmJump(pvStrJumpFlag)

End Function
'========================================================================================
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub
'====================================================================================================
 Sub CurFormatNumericOCX()

  With frm1
   '매출채권금액
   ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
   'VAT금액
  End With

 End Sub
'====================================================================================================
 Sub CurFormatNumSprSheet()
' 2008-01-24 박정순 추가
  If lgIntFlgMode = Parent.OPMD_UMODE Then
	EXIT SUB
  END IF 

  With frm1
   ggoSpread.Source = frm1.vspdData
   '단가
   ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
   '금액
   ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
   'Vat금액
   ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
  End With
 End Sub
'========================================================================================================
 Sub Form_Load()
  Call LoadInfTB19029             
  Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
  Call ggoOper.LockField(Document, "N")        
  Call LoadInfTB19029
  Call InitSpreadSheet            
  Call SetDefaultVal
  Call InitVariables

  Call SetToolbar("1110000000001111")   

  Call CookiePage(0) 
  frm1.txtBLNo.focus
  Set gActiveElement = document.activeElement 
 End Sub
'========================================================================================================
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub
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
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call HideNonCCGrid()
	If frm1.rdoPostingflg1.checked = True Then 
		ggoSpread.SpreadLock -1,-1,-1
	End If 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 	
End Sub
'========================================================================================================
 Sub btnBLNoOnClick()
  frm1.txtBLNo.focus
  Call OpenExBLNoPop()
 End Sub
'========================================================================================================
 Sub btnPosting_OnClick()
  If frm1.btnPosting.disabled <> True Then
   If BtnSpreadCheck = False Then Exit Sub
   Call PostBL()
  End If
 End Sub
'==========================================================================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
 	
	If Trim(frm1.txtGLNo.value) <> "" Then
		 arrParam(0) = Trim(frm1.txtGLNo.value) '회계전표번호
		 
		 if arrParam(0) = "" THEN Exit Sub
		 
		 iCalledAspName = AskPRAspName("a5120ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If

		 arrRet = window.showModalDialog(iCalledAspName , Array(window.parent,arrParam), _
		      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		      
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
	     arrParam(0) = Trim(frm1.txtTempGLNo.value) '결의전표번호
	     
	     if arrParam(0) = "" THEN Exit Sub
	     
	     iCalledAspName = AskPRAspName("a5130ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If
		 
	     arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else 
	     Call DisplayMsgBox("205154", "X", "X", "X")
	End If 
	     lblnWinEvent = False
End Sub
'==========================================================================================
Sub btnPreRcptView_OnClick()
 Dim iCalledAspName
 Dim arrRet
 Dim arrParam(4)
 
 If IsOpenPop = True Then Exit Sub

 IsOpenPop = True
 arrParam(0) = Trim(frm1.txtHBLIssueDt.value)  '발행일
 arrParam(1) = Trim(frm1.txtApplicant.value)   '수입자
 arrParam(2) = Trim(frm1.txtApplicantNm.value)  '수입자명
 arrParam(3) = Trim(frm1.txtCurrency.value)   '화폐
 arrParam(4) = ""         '선수금번호
iCalledAspName = AskPRAspName("s5111ra7")	
if Trim(iCalledAspName) = "" then
	IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111ra7", "x")
	IsOpenPop = False
	exit sub
end if
 
 arrRet = window.showModalDialog(iCalledAspName & "?txtFlag=BL&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent, arrParam), _
       "dialogWidth=860px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 IsOpenPop = False
End Sub
'========================================================================================================
 Sub vspdData_Change(ByVal Col, ByVal Row )
  Dim dblQty
  Dim dblPrice
  Dim dblAmt

  ggoSpread.Source = frm1.vspdData

  Select Case Col
   Case C_Qty, C_DocAmt
    Call CalcDocAmt(Col,Row)

   Case Else

  End Select

  ggoSpread.UpdateRow Row

  lgBlnFlgChgValue = True
 End Sub
'========================================================================================================
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
  With frm1.vspdData
   If Row >= NewRow Then
    Exit Sub
   End If

   If NewRow = .MaxRows Then
    If lgStrPrevKey <> "" Then       '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음
     DbQuery
    End If
   End If
  End With
 End Sub
'========================================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
  If OldLeft <> NewLeft Then Exit Sub

  If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then 
    If CheckRunningBizProcess Then Exit Sub   
    Call DisableToolBar(Parent.TBC_QUERY)
    Call DBQuery   
  End if     

 End Sub
'========================================================================================
 Sub vspdData_Click(ByVal Col, ByVal Row)
 
  Call SetPopupMenuItemInf("0000111111")
     	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		Exit Sub
	End If 
	
    If frm1.rdoPostingflg1.checked = True Then 
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0101111111")   
	End IF
    
 End Sub
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
    End If
	
End Sub
'========================================================================================
 Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
             gMouseClickStatus = "SPCR"
    End If
 End Sub    
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
 Function FncQuery()
  Dim IntRetCD

  FncQuery = False         

  Err.Clear           

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")  
'   IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  Call ggoOper.ClearField(Document, "2")
  Call ggoSpread.ClearSpreadData()  
  Call InitVariables       

  If Not chkField(Document, "1") Then   
   Exit Function
  End If

  Call DbQuery()     

  FncQuery = True     
 End Function
'========================================================================================================
 Function FncNew()
  Dim IntRetCD 

  FncNew = False                                       

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")
'   IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)

   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  Call ggoOper.ClearField(Document, "1")     
  Call ggoOper.ClearField(Document, "2")     
  Call ggoOper.LockField(Document, "N")     
  Call InitVariables          
  Call SetToolbar("1110000000001111")      
  Call SetDefaultVal

  FncNew = True           

 End Function
'========================================================================================================
 Function FncDelete()
  Dim IntRetCD

  FncDelete = False        
  
  If lgIntFlgMode <> Parent.OPMD_UMODE Then    
   Call DisplayMsgBox("900002", "x", "x", "x")
'   Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
   Exit Function
  End If

  IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "x", "x")

  If IntRetCD = vbNo Then
   Exit Function
  End If

  Call DbDelete          

  FncDelete = True         
 End Function
'========================================================================================================
 Function FncSave()
  Dim IntRetCD
  
  FncSave = False           
  
  Err.Clear            
  
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = False Then
      IntRetCD = DisplayMsgBox("900001", "x", "x", "x")             
      Exit Function
  End If  

  ggoSpread.Source = frm1.vspdData

  If Not chkField(Document, "2") Then 
   Exit Function
  End If

  If Not ggoSpread.SSDefaultCheck Then
   Exit Function
  End If
  
  Call DbSave         
  
  FncSave = True        
 End Function
'========================================================================================================
 Function FncCopy()
  frm1.vspdData.ReDraw = False

  ggoSpread.Source = frm1.vspdData 
  If frm1.vspdData.MaxRows < 1 Then Exit Function
  ggoSpread.CopyRow
  SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

  frm1.vspdData.ReDraw = True
 End Function
'========================================================================================================
 Function FncCancel() 
  ggoSpread.Source = frm1.vspdData
  If frm1.vspdData.MaxRows < 1 Then Exit Function
  ggoSpread.EditUndo
  ' Header의 총금액 계산
  Call SumBlAmtMulti("DL")    
 End Function
'========================================================================================================
 Function FncInsertRow()
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    imRow = Parent.AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
		
		lgBlnFlgChgValue = True
    End With
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
 End Function
'========================================================================================================
 Function FncDeleteRow()
  Dim lDelRows
  Dim iDelRowCnt, i
 
  With frm1.vspdData
   If .MaxRows = 0 Then
    Exit Function
   End If
 
   .focus
   ggoSpread.Source = frm1.vspdData

   lDelRows = ggoSpread.DeleteRow
   
   ' Header의 총금액 계산
   Call SumBlAmtMulti("DL")    
   
   lgBlnFlgChgValue = True
  End With
 End Function
'========================================================================================================
 Function FncPrint()
  Call parent.FncPrint()
 End Function
'========================================================================================================
 Function FncPrev() 
  On Error Resume Next 
 End Function
'========================================================================================================
 Function FncNext()
  On Error Resume Next       
 End Function
'========================================================================================================
 Function FncExcel() 
  Call parent.FncExport(Parent.C_SINGLEMULTI)
 End Function
'========================================================================================================
 Function FncFind() 
  Call parent.FncFind(Parent.C_SINGLEMULTI, False)
 End Function
'========================================================================================================
 Function FncExit()
  Dim IntRetCD

  FncExit = False

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")  

'   IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  FncExit = True
 End Function
'========================================================================================================
 Function DbQuery()
  Err.Clear       

  DbQuery = False      

  Dim strVal

  Call LayerShowHide(1)

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

  Call RunMyBizASP(MyBizASP, strVal)     
 
  DbQuery = True          
 End Function
'========================================================================================================
 Function DbSave() 
  Dim lRow
  Dim lGrpCnt
  Dim strVal, strDel
  Dim intInsrtCnt

  DbSave = False          
    
  Call LayerShowHide(1)

  With frm1
   .txtMode.value = Parent.UID_M0002
   .txtUpdtUserId.value = Parent.gUsrID
   .txtInsrtUserId.value = Parent.gUsrID

   lGrpCnt = 1

   strVal = ""
   strDel = ""
   intInsrtCnt = 1

   For lRow = 1 To .vspdData.MaxRows
    .vspdData.Row = lRow
    .vspdData.Col = 0

    Select Case .vspdData.Text
     Case ggoSpread.InsertFlag        
      strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep 
      
      'B/L순번
      .vspdData.Col = C_BLSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '품목
      .vspdData.Col = C_ItemCd       
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '수량
      .vspdData.Col = C_Qty        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '단위
      .vspdData.Col = C_Unit        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '단가
      .vspdData.Col = C_Price        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '금액
      .vspdData.Col = C_DocAmt       
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat타입
      .vspdData.Col = C_VatType
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      'vat율
      .vspdData.Col = C_VatRate
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat금액
      .vspdData.Col = C_VatAmt        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & "" & Parent.gColSep
      '출하번호
      .vspdData.Col = C_DnNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '출하순번
      .vspdData.Col = C_DnSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '수주번호      
      .vspdData.Col = C_SoNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '수주순번
      .vspdData.Col = C_SoSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'L/C번호
      .vspdData.Col = C_LCNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      'L/C순번
      .vspdData.Col = C_LCSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '공장
      .vspdData.Col = C_Plant        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '원화금액
      .vspdData.Col = C_LocAmt
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat원화금액
      .vspdData.Col = C_VatLocAmt        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat포함여부
      .vspdData.Col = C_VatIncFlag
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & UNIConvNum(0,0) & Parent.gColSep
      strVal = strVal & UNIConvNum(0,0) & Parent.gColSep
      '반품여부
      strVal = strVal & "N" & Parent.gColSep
      'C/C번호
      .vspdData.Col = C_CcNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      'C/C순번
      .vspdData.Col = C_CcSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '총중량
      .vspdData.Col = C_GrossWeight      
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '순중량
      .vspdData.Col = C_NetWeight       
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '용적
      .vspdData.Col = C_GrossVolume      
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'HS번호
      .vspdData.Col = C_HsCd        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & "" & Parent.gRowSep
      lGrpCnt = lGrpCnt + 1
      intInsrtCnt = intInsrtCnt + 1

     Case ggoSpread.UpdateFlag
      strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep 

      'B/L순번
      .vspdData.Col = C_BLSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '품목
      .vspdData.Col = C_ItemCd       
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '수량
      .vspdData.Col = C_Qty        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '단위
      .vspdData.Col = C_Unit        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '단가
      .vspdData.Col = C_Price        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '금액
      .vspdData.Col = C_DocAmt       
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat타입
      .vspdData.Col = C_VatType
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      'vat율
      .vspdData.Col = C_VatRate
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat금액
      .vspdData.Col = C_VatAmt        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & "" & Parent.gColSep
      '출하번호
      .vspdData.Col = C_DnNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '출하순번
      .vspdData.Col = C_DnSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '수주번호      
      .vspdData.Col = C_SoNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '수주순번
      .vspdData.Col = C_SoSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'L/C번호
      .vspdData.Col = C_LCNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      'L/C순번
      .vspdData.Col = C_LCSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '공장
      .vspdData.Col = C_Plant        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '원화금액
      .vspdData.Col = C_LocAmt
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat원화금액
      .vspdData.Col = C_VatLocAmt        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'vat포함여부
      .vspdData.Col = C_VatIncFlag
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & UNIConvNum(0,0) & Parent.gColSep
      strVal = strVal & UNIConvNum(0,0) & Parent.gColSep
      '반품여부
      strVal = strVal & "N" & Parent.gColSep
      'C/C번호
      .vspdData.Col = C_CcNo        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      'C/C순번
      .vspdData.Col = C_CcSeq        
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '총중량
      .vspdData.Col = C_GrossWeight      
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '순중량
      .vspdData.Col = C_NetWeight       
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      '용적
      .vspdData.Col = C_GrossVolume      
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
      'HS번호
      .vspdData.Col = C_HsCd        
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      '*************************************************
      'B/L내역등록에서 dll로값을 넘기기위한 dummy string
      '*************************************************
      strVal = strVal & "" & Parent.gRowSep
      
      lGrpCnt = lGrpCnt + 1
      
     Case ggoSpread.DeleteFlag
      strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep
      'B/L순번
      .vspdData.Col = C_BLSeq       
      strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep

      lGrpCnt = lGrpCnt + 1
    End Select
   Next

   .txtMaxRows.value = lGrpCnt-1
   .txtSpread.value = strDel & strVal

   Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)     

  End With

  DbSave = True             
 End Function
'========================================================================================================
 Function DbDelete()
 End Function
'========================================================================================================
 Function DbQueryOk()        
	lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False   

	Call ggoOper.LockField(Document, "Q")  
	Call SetToolbar("11101011000111")   

	If frm1.txtRefFlg.value = "M" Then 
		frm1.btnPosting.disabled = True 
	Else  
		If CInt(frm1.txtStatusFlg.value) < 3 Then 
			frm1.btnPosting.disabled = False
		Else 
			frm1.btnPosting.disabled = True 
		End If  
	End If 
    
	If frm1.rdoPostingflg1.checked = True Then
		Call ProtectGrid()
		Call SetToolbar("11100000000111")
	Else 
		Call ReleaseGrid()
	End If 

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  

 End Function
'========================================================================================================
 Function BLHdrQueryOk()       
  Call SetToolbar("11101011000011")   
  If frm1.txtRefFlg.value = "M" Then 
   frm1.btnPosting.disabled = True 
  Else  
   If CInt(frm1.txtStatusFlg.value) < 3 Then 
    frm1.btnPosting.disabled = False
   Else 
    frm1.btnPosting.disabled = True 
   End If  
  End If 
 End Function
'========================================================================================================
 Function DbSaveOk()        
  Call InitVariables
  frm1.txtBLNo.value = frm1.txtHBLNo.value
  Call ggoOper.ClearField(Document, "2")      
  Call FncQuery()
 End Function
'========================================================================================================
 Function DbDeleteOk()           
 End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
 <TABLE <%=LR_SPACE_TYPE_00%>>
  <TR>
   <TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백%></TD>
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
         <TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>{{T_B/L내역}}</font></TD>
         <TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
           </TR>
       </TABLE>
      </TD>
      <TD WIDTH=* align=right><A href="vbscript:OpenSODtRef">{{수주내역참조}}</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCDtRef">{{L/C내역참조}}</A>&nbsp;|&nbsp;<A href="vbscript:OpenCCDtlRef">{{통관내역참조}}</A></TD>
      <TD WIDTH=10>&nbsp;</TD>
     </TR>
    </TABLE>
   </TD>
  </TR>
  <TR HEIGHT=*>
   <TD WIDTH=100% CLASS="Tab11">
    <TABLE <%=LR_SPACE_TYPE_20%>>
     <TR>
      <TD HEIGHT=5 WIDTH=100%></TD>
     </TR>
     <TR>
      <TD HEIGHT=20 WIDTH=100%>
       <FIELDSET CLASS="CLSFLD">
        <TABLE <%=LR_SPACE_TYPE_40%>>
         <TR>
          <TD CLASS=TD5 NOWRAP>{{B/L관리번호}}</TD>
          <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="{{B/L관리번호}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnBLNoOnClick()"></TD>
          <TD CLASS=TD6 NOWRAP></TD>
          <TD CLASS=TD6 NOWRAP></TD>
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
         <TD CLASS=TD5 NOWRAP>{{수입자}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="{{수입자}}">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=25 TAG="24"></TD>
         <TD CLASS=TD5 NOWRAP>{{수주번호}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" TYPE=TEXT SIZE=20 TAG="24XXXU" ALT="{{수주번호}}"></TD>
        </TR>
        <TR>
         <TD CLASS=TD5 NOWRAP>{{영업그룹}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="{{영업그룹}}">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=25 TAG="24"></TD>
         <TD CLASS=TD5 NOWRAP>{{확정여부}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">{{확정}}</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">{{미확정}}</LABEL></TD>
        </TR>
        <TR>
         <TD CLASS=TD5 NOWRAP>{{매출채권형태}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBillType" SIZE=10 MAXLENGTH=20 TAG="24XXXU" ALT="{{매출채권형태}}">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=25 TAG="24"></TD>
         <TD CLASS=TD5 NOWRAP>{{중량단위}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="{{중량단위}}" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
        </TR>
        <TR>
         <TD CLASS=TD5 NOWRAP>{{결제방법}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="{{결제방법}}">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=25 TAG="24"></TD>
         <TD CLASS=TD5 NOWRAP>{{용적단위}}</TD>
         <TD CLASS=TD6 NOWRAP><INPUT NAME="txtVolumnUnit" ALT="{{용적단위}}" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
        </TR>
        <TR>
         <TD CLASS=TD5 NOWRAP>{{B/L금액}}</TD>
         <TD CLASS=TD6 NOWRAP>
          <TABLE CELLSPACING=0 CELLPADDING=0> 
           <TR>
            <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtDocAmt" CLASS=FPDS140 tag="24X2" ALT="{{B/L금액}}" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
            <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="{{화폐}}"></TD>
           </TR>
          </TABLE>
         </TD>
         <TD CLASS=TD5 NOWRAP>{{B/L자국금액}}</TD>
         <TD CLASS=TD6 NOWRAP>
          <TABLE CELLSPACING=0 CELLPADDING=0> 
           <TR>
            <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtLocAmt" CLASS=FPDS140 tag="24X2" ALT="{{B/L자국금액}}" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
            <TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="{{자국화폐}}"></TD>
           </TR>
          </TABLE>
         </TD>
        </TR>                         
        <TR>
         <TD HEIGHT=100% WIDTH=100% COLSPAN=4>
          <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
     <TD><BUTTON NAME="btnPosting" CLASS="CLSMBTN">{{확정}}</BUTTON>&nbsp;
      <BUTTON NAME="btnGLView" CLASS="CLSMBTN">{{전표조회}}</BUTTON>&nbsp;
      <BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">{{선수금현황}}</BUTTON></TD>
     <TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(1, EXBL_HEADER_ENTRY_ID)">{{B/L등록}}</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1, BIZ_BillCollect_JUMP_ID)">{{B/L수금내역등록}}</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(2, EXPORT_CHARGE_ENTRY_ID)">{{판매경비등록}}</A></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TABLE>
   </TD>
  </TR>
  <TR>
   <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
  </TR>
 </TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHBLNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtRefFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtStatusFlg" TAG="24"> 
<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtXchgRate" tag="24">
<INPUT TYPE=HIDDEN NAME="txtXchgRateOp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtVatAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtVatLocAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtIncoterms" tag="24">
<INPUT TYPE=HIDDEN NAME="txtIncotermsNm" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBLIssueDt" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

