<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase 
'*  2. Function Name        : Goods Receipt
'*  3. Program ID           : M9211MA1		
'*  4. Program Name         : 재고이동입고 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : EverForever
'* 10. Modifier (Last)      : KO MYOUNG JIN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>



<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit					

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const BIZ_PGM_ID = "m9211mb1.asp"									
Const BIZ_PGM_JUMP_ID	= "M9111MA1"

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const C_SHEETMAXROWS = 100

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop          
Dim lblnWinEvent
Dim lgOpenFlag
Dim lgRefABCflag
Dim interface_Account

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_GrQty
Dim C_GRUnit
Dim C_TrackingNo
Dim C_DocAmt
Dim C_Cur
Dim C_PlantCd
Dim C_PlantNm
Dim C_SlCd
Dim C_SlNm
Dim C_LotNo
Dim C_LotSeqNo
Dim C_MakerLotNo
Dim C_MakerLotSeqNo
Dim C_GRNo
Dim C_GRSeqNo
Dim C_StoNo
Dim C_StoSeqNo
Dim C_SGiNo
Dim C_SGiSeqNo
Dim C_Base_Unit
DIM C_Mvmt_prc
DIM C_Locamt
DIM C_Mvmt_no
DIM C_Base_Qty
DIM C_PUR_GRP

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################

'--------------------------------------------------------------------
'		Field의 Tag속성을 Protect로 전환,복구 시키는 함수 
'--------------------------------------------------------------------

Function ChangeTag(Byval Changeflg)
	
	Dim index

	If Changeflg = true then	

		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "Q"
		'ggoOper.SetReqAttr	frm1.txtTaxCd, "Q"
		ggoOper.SetReqAttr	frm1.txtGroupCd, "Q"
		frm1.vspdData.ReDraw = false
		For index = 1 to frm1.vspdData.MaxCols
			ggoSpread.SpreadLock index , -1, index , -1
		Next
		frm1.vspdData.ReDraw = true	
	
	Else

		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "N"
		'ggoOper.SetReqAttr	frm1.txtTaxCd, "N"
		ggoOper.SetReqAttr	frm1.txtGroupCd, "N"
		Call ggoOper.LockField(Document, "N")	
		
		ggoOper.SetReqAttr	frm1.txtMvmtNo1, "D"
		'ggoOper.SetReqAttr	frm1.txtTaxCd, "D"
		ggoOper.SetReqAttr	frm1.txtGroupCd, "D"
		Call SetSpreadLock 
		
	End if 
	
End Function 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	
	C_ItemCd		= 1
	C_ItemNm		= 2 
	C_Spec			= 3
	C_GrQty			= 4
	C_GRUnit		= 5
	C_TrackingNo	= 6
	C_DocAmt		= 7
	C_Cur			= 8
	C_PlantCd		= 9
	C_PlantNm		= 10
	C_SlCd			= 11
	C_SlNm			= 12
	C_LotNo			= 13
	C_LotSeqNo		= 14
	C_MakerLotNo	= 15
	C_MakerLotSeqNo	= 16
	C_GRNo			= 17
	C_GRSeqNo		= 18	
	C_StoNo			= 19
	C_StoSeqNo		= 20
	C_SGiNo			= 21
	C_SGiSeqNo		= 22
	C_Base_Unit     = 23
	C_Mvmt_prc		= 24
	C_Locamt		= 25
	C_Mvmt_no		= 26
	C_Base_Qty		= 27
	C_PUR_GRP		= 28
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                
    lgBlnFlgChgValue = False                 
    lgIntGrpCount = 0                        
    
    lgStrPrevKey = ""                        
    lgLngCurRows = 0                         
    frm1.vspdData.MaxRows = 0
    
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
	lgOpenFlag = False    
    lgRefABCflag = ""
	
	'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
	frm1.txtGmDt.Text = UNIConvDateAtoB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.txtGroupCd.Value = Parent.gPurGrp
    Call SetToolBar("1110000100001111")
    frm1.txtMvmtNo.focus 
    Set gActiveElement = document.activeElement    
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true    
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		.MaxCols = C_PUR_GRP+1
		.Col = .MaxCols:	.ColHidden = True
    	.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit			C_ItemCd,		"품목", 10
		ggoSpread.SSSetEdit 		C_ItemNm,		"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,			"품목규격", 20 	
		SetSpreadFloatLocal 		C_GrQty,		"입고수량",15,1, 3
		ggoSpread.SSSetEdit 		C_GRUnit,		"단위", 10
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15 				
		SetSpreadFloatLocal 		C_DocAmt,		"입고금액", 15 ,1, 2	
		ggoSpread.SSSetEdit 		C_Cur,			"화폐", 10
		ggoSpread.SSSetEdit 		C_PlantCd,		"공장", 10
		ggoSpread.SSSetEdit 		C_PlantNm,		"공장명", 20
		ggoSpread.SSSetEdit			C_SlCd,			"창고", 10
		ggoSpread.SSSetEdit 		C_SlNm,			"창고명", 20	    
		ggoSpread.SSSetEdit 		C_LotNo,		"Lot No.", 20,,,,2    
		ggoSpread.SSSetEdit 		C_LotSeqNo,		"LOT NO 순번",10,,,,2
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,,2    
		ggoSpread.SSSetEdit 		C_MakerLotSeqNo,"Maker Lot 순번", 10,,,,2
		ggoSpread.SSSetEdit 		C_GRNo,			"재고처리번호", 20
		ggoSpread.SSSetEdit 		C_GRSeqNo,		"재고처리순번", 10
		ggoSpread.SSSetEdit 		C_StoNo,		"재고이동요청번호", 15
		ggoSpread.SSSetEdit		 	C_StoSeqNo,		"재고이동요청순번", 15,,,,2
		ggoSpread.SSSetEdit 		C_SGiNo,		"출고번호", 15,,,,2
		ggoSpread.SSSetEdit		 	C_SGiSeqNo,		"출고순번", 10,,,,2
		
		ggoSpread.SSSetEdit 		C_Base_Unit,	"단위", 10
		SetSpreadFloatLocal 		C_Mvmt_prc,		"단가", 15 ,1, 4	
		SetSpreadFloatLocal 		C_Locamt,		"자국금액", 15 ,1, 2	
		ggoSpread.SSSetEdit 		C_Mvmt_no,		"번호", 10
		SetSpreadFloatLocal			C_Base_Qty,		"입고수량", 15 ,1, 3	
		ggoSpread.SSSetEdit			C_PUR_GRP,		"구매그룹", 15 ,1, 2	
		ggoSpread.SSSetEdit 		C_SGiSeqNo+1,	"", 10
		
		
		Call ggoSpread.MakePairsColumn(C_ItemCd,C_Spec)
		Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlNm)
		Call ggoSpread.MakePairsColumn(C_LotNo,C_LotSeqNo)
		Call ggoSpread.MakePairsColumn(C_MakerLotNo,C_MakerLotSeqNo)
		Call ggoSpread.MakePairsColumn(C_GRNo,C_GRSeqNo)
		Call ggoSpread.MakePairsColumn(C_StoNo,C_StoSeqNo)
		Call ggoSpread.MakePairsColumn(C_SGiNo,C_SGiSeqNo)
		
		
		Call ggoSpread.SSSetColHidden(C_Base_Unit,C_PUR_GRP+1,True)	

		Call SetSpreadLock()
		.ReDraw = true
    
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    
    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
			 
    With ggoSpread
		.SpreadLock C_ItemCd       , -1 , C_ItemCd         , -1
		.SpreadLock C_ItemNm       , -1 , C_ItemNm         , -1
		.SpreadLock C_Spec         , -1 , C_Spec           , -1
		.SpreadLock C_GrQty        , -1 , C_GrQty          , -1
		.SpreadLock C_GRUnit       , -1 , C_GRUnit         , -1
		.SpreadLock C_TrackingNo   , -1 , C_TrackingNo     , -1
		.SpreadLock C_DocAmt       , -1 , C_DocAmt         , -1
		.SpreadLock C_Cur          , -1 , C_Cur            , -1
		.SpreadLock C_PlantCd      , -1 , C_PlantCd        , -1
		.SpreadLock C_PlantNm      , -1 , C_PlantNm        , -1
		.SpreadLock C_SlCd         , -1 , C_SlCd           , -1
		.SpreadLock C_SlNm         , -1 , C_SlNm           , -1
		.SpreadLock C_LotNo        , -1 , C_LotNo          , -1
		.SpreadLock C_LotSeqNo     , -1 , C_LotSeqNo       , -1
		.SpreadLock C_MakerLotNo   , -1 , C_MakerLotNo     , -1
		.SpreadLock C_MakerLotSeqNo, -1 , C_MakerLotSeqNo  , -1
		.SpreadLock C_GRNo         , -1 , C_GRNo           , -1
		.SpreadLock C_GRSeqNo      , -1 , C_GRSeqNo        , -1
		.SpreadLock C_StoNo        , -1 , C_StoNo          , -1
		.SpreadLock C_StoSeqNo     , -1 , C_StoSeqNo       , -1
		.SpreadLock C_SGiNo        , -1 , C_SGiNo          , -1
		.SpreadLock C_SGiSeqNo     , -1 , C_SGiSeqNo       , -1
    End With
    
    frm1.vspdData.ReDraw = True
    
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		.ReDraw = False
		
    	ggoSpread.SSSetProtected  C_ItemCd        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_ItemNm        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_Spec          , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_GrQty         , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_GRUnit        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_TrackingNo    , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_DocAmt        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_Cur           , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_PlantCd       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_PlantNm       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_SlCd          , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_SlNm          , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_LotNo         , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_LotSeqNo      , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_MakerLotNo    , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_MakerLotSeqNo , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_GRNo          , pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected  C_GRSeqNo       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_StoNo         , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_StoSeqNo      , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_SGiNo         , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_SGiSeqNo      , pvStartRow, pvEndRow		
		
		.Col = 1
		.Row = .ActiveRow
		.Action = 0
		.EditMode = True
		.ReDraw = True
	End With
	
End Sub

'여기네 

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd        = iCurColumnPos(1)
			C_ItemNm        = iCurColumnPos(2)
			C_Spec          = iCurColumnPos(3)
			C_GrQty         = iCurColumnPos(4)
			C_GRUnit        = iCurColumnPos(5)
			C_TrackingNo    = iCurColumnPos(6)
			C_DocAmt        = iCurColumnPos(7)
			C_Cur           = iCurColumnPos(8)
			C_PlantCd       = iCurColumnPos(9)
			C_PlantNm       = iCurColumnPos(10)
			C_SlCd          = iCurColumnPos(11)
			C_SlNm          = iCurColumnPos(12)
			C_LotNo         = iCurColumnPos(13)
			C_LotSeqNo      = iCurColumnPos(14)
			C_MakerLotNo    = iCurColumnPos(15)
			C_MakerLotSeqNo = iCurColumnPos(16)
			C_GRNo          = iCurColumnPos(17)
			C_GRSeqNo       = iCurColumnPos(18)
			C_StoNo         = iCurColumnPos(19)
			C_StoSeqNo      = iCurColumnPos(20)
			C_SGiNo         = iCurColumnPos(21)
			C_SGiSeqNo      = iCurColumnPos(22)
		
	End Select

End Sub	
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************
'========================================== 2.4.1 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'------------------------------------------  OpenPoRef()  -------------------------------------------------
'	Name : OpenPoRef()
'	Description : 
'---------------------------------------------------------------------------------------------------------

Function OpenSGiRef()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","출고참조" )
		Exit Function
	End if 
	
    If Trim(frm1.txtSupplierCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "출고공장", "X")
		frm1.txtSupplierCd.focus()
    	Exit Function
    End IF

	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtGroupCd.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = ""
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
	arrParam(9) = ""
	arrParam(10) = ""
	arrParam(11) = ""
	arrParam(12) = ""
	
	'lgOpenFlag = False
	'Call changeMvmtType   
	'if lgOpenFlag = True THEN EXIT FUNCTION
	'Call changeSpplCd   
	'if lgOpenFlag = True THEN EXIT FUNCTION
	'Call changeTaxCd   
	'if lgOpenFlag = True THEN EXIT FUNCTION
	'Call changeGroupCd
	'if lgOpenFlag = True THEN EXIT FUNCTION
	'msgbox lgOpenFlag
	'strRet = window.showModalDialog("m9211ra1.asp", Array(Window.parent), _	
	'if lgOpenFlag = True THEN EXIT FUNCTION
	
	iCalledAspName = AskPRAspName("m9211ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m9211ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _	
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	
	lgOpenFlag	= False

	If isEmpty(strRet) Then Exit Function	

	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetSGiRef(strRet)
	End If	

End Function

Function SetSGiRef(strRet)

	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim strtemp1,strtemp2
	Dim temp
	Dim SCheck 
	Dim S_txtSupplierCd
	DIM S_txtSupplierNm

	Const C_PO_NO_REF              = 0
	Const C_PO_SEQ_NO_REF          = 1
	Const C_PLANT_CD_REF           = 2
	Const C_PLANT_NM_REF           = 3
	Const C_ITEM_CD_REF            = 4
	Const C_ITEM_NM_REF            = 5
	Const C_spec_REF               = 6
	Const C_GI_QTY_REF             = 7
	Const C_GI_UNIT_REF            = 8
	Const C_PRICE_REF              = 9
	Const C_GI_AMT_REF             = 10
	Const C_bp_cd_REF              = 11
	Const C_bp_nm_REF              = 12
	Const C_DN_NO_REF              = 13
	Const C_DN_SEQ_REF             = 14
	Const C_SL_CD_REF              = 15
	Const C_SL_NM_REF              = 16
	Const C_CUR_REF                = 17
	Const C_TRACKING_NO_REF        = 18
	Const C_trns_lot_no_REF        = 19
	Const C_trns_lot_sub_no_REF    = 20
	Const C_lot_no_REF             = 21
	Const C_lot_sub_no_REF         = 22
	Const C_GI_AMT_LOC_REF         = 23
	Const C_BASE_UNIT_REF          = 24
	Const C_BASE_QTY_REF           = 25
	Const C_PUR_GRP_REF			   = 26	
	
	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true

	with frm1
	SCheck = true
	
	if count1 <> 0 or .vspdData.MaxRows <> 0 then
		for index1 = 0 to Count1
			if .txtSupplierCd.value = "" then
				.txtSupplierCd.value = 	Trim(strRet(index1, 11))
			else				
				IF UCase(Trim(.txtSupplierCd.value)) <> UCase(Trim(strRet(index1, 11))) then
					strMessage = Trim(strRet(index1, 11))
					.txtSupplierCd.value = ""
					Call DisplayMsgBox("174620","X" , "X","X")
					
					exit function
				end if
			end if
		next
	end if
		
	for index1 = 0 to Count1

		.vspdData.Row=Index1+1

		for Index3=1 to .vspdData.MaxRows'count1
		
			.vspdData.Row = index3
			.vspdData.Col=C_SGiNo
			strtemp1 = .vspdData.Text
			.vspdData.Col=C_SGiSeqNo
			strtemp2 = .vspdData.Text
			
			if Trim(strtemp1) = Trim(strRet(index1, 13)) and Trim(strtemp2) = Trim(strRet(index1, 14)) then
				strMessage = Trim(strRet(index1, 13))
				Call DisplayMsgBox("17a005", "X",strMessage,"출고번호")
				exit function
			end if
		Next		
		
		if IntIflg <> False then
			
			Call fncinsertrow(1)

			.vspdData.Redraw = False
			Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow)
			
			.vspdData.Row=.vspdData.ActiveRow 
			
			for index2 = 0 to Count2
				
				Select Case Index2
		
				Case C_PO_NO_REF
					.vspdData.Col=C_StoNo
					.vspdData.Text=strRet(index1,index2)
				Case C_PO_SEQ_NO_REF
					.vspdData.Col=C_StoSeqNo
					.vspdData.Text=strRet(index1,index2)
				Case C_PLANT_CD_REF
					.vspdData.Col=C_PlantCd
					.vspdData.Text=strRet(index1,index2)
				Case C_PLANT_NM_REF
					.vspdData.Col=C_PlantNm
					.vspdData.Text=strRet(index1,index2)
				Case C_ITEM_CD_REF
					.vspdData.Col=C_ItemCd
					.vspdData.Text=strRet(index1,index2)
				Case C_ITEM_NM_REF
					.vspdData.Col=C_ItemNm
					.vspdData.Text=strRet(index1,index2)			
				Case C_spec_REF
					.vspdData.Col=C_Spec
					.vspdData.Text=strRet(index1,index2)				
				Case C_bp_cd_REF

					IF SCheck = TRUE THEN
						S_txtSupplierCd = strRet(index1,index2)
					END IF
				Case C_bp_nm_REF	
					IF SCheck = TRUE THEN
						S_txtSupplierNm = strRet(index1,index2)
					END IF
					SCheck = False						
				Case C_GI_QTY_REF
					.vspdData.Col=C_GrQty
					.vspdData.Text=strRet(index1,index2)
				Case C_GI_UNIT_REF
					.vspdData.Col=C_GRUnit
					.vspdData.Text=strRet(index1,index2)
				Case C_PRICE_REF
					.vspdData.Col=C_Mvmt_prc
					.vspdData.Text=strRet(index1,index2)
				Case C_GI_AMT_REF
					.vspdData.Col=C_DocAmt
					.vspdData.Text=strRet(index1,index2)	
				Case C_DN_NO_REF
					.vspdData.Col=C_SGiNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_DN_SEQ_REF
					.vspdData.Col=C_SGiSeqNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_SL_CD_REF
					.vspdData.Col=C_SlCd
					.vspdData.Text=strRet(index1,index2)		
				Case C_SL_NM_REF
					.vspdData.Col=C_SlNm
					.vspdData.Text=strRet(index1,index2)		
				Case C_CUR_REF
					.vspdData.Col=C_Cur
					.vspdData.Text=strRet(index1,index2)		
				Case C_TRACKING_NO_REF
					.vspdData.Col=C_TrackingNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_trns_lot_no_REF
					.vspdData.Col=C_MakerLotNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_trns_lot_sub_no_REF
					.vspdData.Col=C_MakerLotSeqNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_lot_no_REF
					.vspdData.Col=C_LotNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_lot_sub_no_REF
					.vspdData.Col=C_LotSeqNo
					.vspdData.Text=strRet(index1,index2)		
				Case C_GI_AMT_LOC_REF
					.vspdData.Col=C_Locamt
					.vspdData.Text=strRet(index1,index2)				
				Case C_BASE_UNIT_REF
					.vspdData.Col=C_Base_Unit
					.vspdData.Text=strRet(index1,index2)			
				Case C_BASE_QTY_REF
					.vspdData.Col=C_Base_Qty
					.vspdData.Text=strRet(index1,index2)			
				Case C_PUR_GRP_REF
					.vspdData.Col=C_PUR_GRP
					.vspdData.Text=strRet(index1,index2)	
				End Select
				
			next
			
		Else
			IntIFlg=True
		End if 
	next
	
	.vspdData.ReDraw = True
	
	if Trim(.txtGroupCd.value) = "" then
		Call GroupTxt
	END IF	
	
	'if .txtSupplierCd.value = "" THEN
		.txtSupplierCd.value = S_txtSupplierCd 
		.txtSupplierNm.value = S_txtSupplierNm 
	'END IF
	
	End with
	
	if frm1.vspdData.Maxrows > 0 then
		ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
	end if
	
	Call SetToolBar("11101011000111")
End Function

Function GroupTxt

	frm1.vspdData.row = 1
	frm1.vspdData.col = C_PUR_GRP
	frm1.txtGroupCd.value = frm1.vspdData.text

End Function


'==========================================  SetAflag, SetBflag, ,SetCflag, ResetABCflag =================
'	Name : SetAflag, SetBflag
'	Description : Set when Mouse be over 
'=========================================================================================================
Function SetAflag()

	lgRefABCflag = "A"
	
End Function

Function ResetABCflag()
	lgRefABCflag = ""	
End Function

Function OpenTax()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True or UCase(frm1.txtTaxCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function 

	IsOpenPop = True

'	arrParam(0) = "Tax"						
'	arrParam(1) = "b_tax_code  a, b_tax_code_history b, b_tax_proc_line c, b_tax_item d"			
	'arrParam(4) = "a.io_flag = 'I' AND a.usage = 'Y' and a.tax_code = b.tax_code "
	'arrParam(4) = arrParam(4) & " and b.effective_from = (select max(effective_from) from b_tax_code_history b where effective_from <= '"
	'arrParam(4) = arrParam(4) & UniConvDate(frm1.txtGmDt.text) &  "')"
	'arrParam(4) = arrParam(4) & " and b.tax_procedure = c.tax_procedure and c.tax_item = d.tax_item "
	'arrParam(4) = arrParam(4) & " and b.tax_code >=  '"
	'arrParam(4) = arrParam(4) & Trim(frm1.txtTaxCd.value) &  "'"		
	
'	arrParam(4) = " a.io_flag = 'i' and   a.usage = 'Y' "
 '   arrParam(4) = arrParam(4) & " and   a.tax_code = b.tax_code"
  '  arrParam(4) = arrParam(4) & " and   b.effect_from_dt <= '" & UNIConvDate(frm1.txtGmDt.text)  & "'"
   ' arrParam(4) = arrParam(4) & " and   b.tax_procedure = c.tax_procedure"
    'arrParam(4) = arrParam(4) & " and   c.tax_item *= d.tax_item"
    'arrParam(4) = arrParam(4) & " and   a.tax_code >= '" & trim(frm1.txtTaxCd.value) & "'"    
    
	'arrParam(5) = "Tax"					
    'arrField(0) = "a.tax_code"						
    'arrField(1) = "a.remark"						
    'arrField(2) = "d.tax_rate"
    'arrHeader(0) = "Tax"					
    'arrHeader(1) = "Tax명"					
    'arrHeader(2) = "Tax율"					
    
	arrHeader(0) = "Tax"								 ' Header명(0)
	arrHeader(1) = "Tax명"		               			 ' Header명(1)
	arrHeader(2) = "Tax율"
	arrParam(0) = "Tax"								 ' 팝업 명칭 
	arrParam(1) = "b_Tax_Code"				                     ' TABLE 명칭 
	arrParam(2) = UCase(Trim(frm1.txtTaxCd.value))				 ' Code Condition
	arrParam(4) = "Usage = " & FilterVar("Y", "''", "S") & "  And Io_flag = " & FilterVar("I", "''", "S") & " "	 
	arrParam(5) = "Tax"								 ' TextBox 명칭		
	arrField(0) = "Tax_code"									 ' Field명(0)
	arrField(1) = "Remark"									     ' Field명(1)		
	'arrField(2) = "Tax_rate"
	'arrField(2) = UNIConvDate(frm1.txtGmDt.text)
	
	'arrField(2) = "F2" & parent.gColSep & "Dbo.ufn_a_getTaxRate('" & ucase(trim(frm1.txtTaxCd.value)) & "','I', '" & UNIConvDate(frm1.txtGmDt.text) & "')"    '===>인자는 io_flag(I)와 계산서일자(입고일자)

    arrField(2) = "F2" & parent.gColSep & "Dbo.ufn_a_getTaxRate(Tax_Code," & FilterVar("I", "''", "S") & " ,  " & FilterVar(UNIConvDate(frm1.txtGmDt.text), "''", "S") & ")"   '===>인자는 io_flag(I)와 계산서일자(입고일자)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtTaxCd.value =arrRet(0)
		frm1.txtTaxNM.value=arrRet(1)		
		'frm1.txtTaxRate.text=arrRet(2)
		frm1.txtTaxRate.text=UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)		
	End If	
	
END Function

'------------------------------------------  OpenMvmtType()  ----------------------------------------------
'	Name : OpenMvmtType()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenMvmtType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtMvmtType.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "입고형태"	
	'arrParam(1) = " ( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b "
	'arrParam(1) = arrParam(1) & " where a.rcpt_type = b.io_type_cd    and a.sto_flg = 'N' AND a.USAGE_FLG='Y' "
	'arrParam(1) = arrParam(1) & " and ((b.RCPT_FLG='Y' AND b.RET_FLG='N') or (b.RET_FLG='N' And b.SUBCONTRA_FLG='N')) ) c "   
	
	arrParam(1) = "(select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b"
	arrParam(1) = arrParam(1) & " where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("Y", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & " ) c "
	
	arrParam(2) = Trim(frm1.txtMvmtType.Value)

	'arrParam(4) = "a.rcpt_type = b.io_type_cd    and a.sto_flg = 'Y' AND a.USAGE_FLG='Y' "
	arrParam(5) = "입고형태"			
	
    arrField(0) = " c.IO_Type_Cd"
    arrField(1) = " c.IO_Type_NM"
    
    arrHeader(0) = "입고형태"		
    arrHeader(1) = "입고형태명"
    


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else 
		Call SetMovetype(arrRet)
	End If	

End Function

'------------------------------------------  OpenMvmtNo()  -------------------------------------------------
'	Name : OpenMvmtNo()
'	Description : OpenPoNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
		
		If lblnWinEvent = True Or UCase(frm1.txtMvmtNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		

		arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""'This is for Inspection check, must be nothing.

		iCalledAspName = AskPRAspName("m9211pa1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m9211pa1", "X")
			IsOpenPop = False
			Exit Function
		End If

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _	
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lgOpenFlag	= False

		If isEmpty(strRet) Then Exit Function	
		
		If strRet(0) = "" Then
			Exit Function
		Else
			Call SetMvmtNo(strRet)
		End If	
		
End Function
'------------------------------------------  OpenGroup()  ------------------------------------------------
'	Name : OpenGroup()
'	Description : OpenGroup1 PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If	
	
End Function
'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenSppl()
'	Description :  OpenSppl PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "출고공장"				
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""							
	
	'arrParam(4) = "Bp_Type in ('S','CS') AND usage_flag='Y'"	
	arrParam(4) = "Bp_Type <> " & FilterVar("C", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("I", "''", "S") & " "	
	arrParam(5) = "출고공장"				
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					

	arrHeader(0) = "출고공장"				
	arrHeader(1) = "출고공장명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
	End If	
	
End Function

'==========================================  2.4.2 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'---------------------------------------  SetMvmtNo()  --------------------------------------------------
'	Name : SetMvmtNo()
'	Description : Group Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetMvmtNo(strRet)
	frm1.txtMvmtNo.value = strRet(0)
End Function

'------------------------------------------  SetGroup()  ------------------------------------------------
'	Name : SetGroup()
'	Description : Group Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetGroup(byval arrRet)
	frm1.txtGroupCd.Value= arrRet(0)		
	frm1.txtGroupNm.Value= arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetMovetype()  --------------------------------------------------
'	Name : SetMovetype()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function SetMovetype(byval arrRet)
	frm1.txtMvmtType.Value	= arrRet(0)		
	frm1.txtMvmtTypeNm.Value= arrRet(1)
	Call changeMvmtType()
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'Lot 순번 Maker Lot 순번 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6"				  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub

'==========================================================================================
'   Event Name : setReference()
'   Event Desc : 
'==========================================================================================
Function setReference()

	ggoOper.SetReqAttr	frm1.txtMvmtType, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"

End Function

'==========================================================================================
'   Event Name : CookiePage
'   Event Desc : 
'==========================================================================================
Function CookiePage(Byval Kubun)

	Dim strTemp
	
	'if lgIntFlgMode = Parent.OPMD_CMODE then
	'	Call DisplayMsgBox("900002", "X", "X", "X")
	'	Exit Function
	'End if
	If Kubun = 1 Then
	    
	    if frm1.vspdData.Maxrows <> 0 then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = C_StoNo
		
			WriteCookie "PoNo" , Trim(frm1.vspdData.text)				
		else
			WriteCookie "PoNo" , ""
		end if

		Call PgmJump(BIZ_PGM_JUMP_ID)
		
	'Else
	'	strTemp = Parent.ReadCookie("MvmtNo")
	
	'	If strTemp = "" then Exit Function
	
	'	frm1.txtMvmtNo.value = Parent.ReadCookie("MvmtNo")
	
	'	Call Parent.WriteCookie("MvmtNo" , "")
	
	'	MainQuery()
	End if
	
End Function


 '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables
    
    Call CookiePage(0)
	    
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
  
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
   gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else	
		Call SetPopupMenuItemInf("0001111111")
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
'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   
	ggoSpread.Source = frm1.vspdData
    
	if Col = C_SlCdPop then
    	Call OpenSlCd()
	End if	
    
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    Call ChangeTag(True)
End Sub

'==========================================================================================
'   Event Name : txtGmDt
'   Event Desc :
'==========================================================================================
Sub txtGmDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGmDt.Action = 7
	End if
End Sub
'==========================================================================================
'   Event Name : txtGmDt
'   Event Desc :
'==========================================================================================
Sub txtGmDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
    
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	
		If lgStrPrevKey <> "" Then	
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If

		End If
    End if
    
End Sub

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    On Error Resume Next                                                 
    Err.Clear                                               
    
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
    FncQuery = True											

    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ChangeTag(False)
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")           
    Call SetDefaultVal
    Call InitVariables
        
    FncNew = True                     
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    
	Dim IntRetCD
	
	On Error Resume Next 
    Err.Clear                                               
    
    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    								
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then             
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True                                
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim intIndex
    
    FncSave = False                                 
    
    On Error Resume Next                           
    Err.Clear                                       

	ggoSpread.Source = frm1.vspdData				
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")					
        Exit Function
    End If

    
    If Not chkField(Document, "2") Then									
       Exit Function
    End If

	with frm1
  
		if CompareDateByFormat(.txtGmDt.text,UNIConvDateAtoB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat),.txtGmDt.Alt,.txtGmDt.Alt, _
                   "970025",.txtGmDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtGmDt.text) <> ""  then	
			Call DisplayMsgBox("17a003","X","입고일","X")			
			Exit Function
		End if   
                   
	End with
	

    ggoSpread.Source = frm1.vspdData									
    If Not ggoSpread.SSDefaultCheck Then								
       Exit Function
    End If
    
    If frm1.vspdData.Maxrows < 1 then
    	Exit Function
    End if
    '-----------------------
    'Check content area
    '-----------------------
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0	
	Next
	    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    FncSave = True                                                      
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
	
	if frm1.vspdData.Maxrows < 1 then exit function

	frm1.vspdData.Col=C_LotFlg
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	if frm1.vspdData.Text <> "Y" then exit function
    
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow

	frm1.vspdData.ReDraw = False
	
	if frm1.vspdData.Text <> "Y" then
		ggoSpread.spreadUnlock C_LotNo, frm1.vspdData.Row, C_LotNoPop, frm1.vspdData.Row
		ggoSpread.SSSetRequired		C_LotNo, frm1.vspdData.Row, frm1.vspdData.Row
	else
		ggoSpread.spreadlock C_LotNo, frm1.vspdData.Row, C_LotNoSeq, frm1.vspdData.Row
	end if
	
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    	
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo  
    Set gActiveElement = document.ActiveElement                                                   
    
    if frm1.vspdData.Maxrows = 0 then
		Call SetToolBar("11100001000111")		
		ggoOper.SetReqAttr	frm1.txtSupplierCd, "N"
	end if
End Function
'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow

	On Error Resume Next
	Err.Clear
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End IF
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow , imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
    End With
	
	If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
	
End Function
'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
    
		lDelRows = ggoSpread.DeleteRow
	End With
    
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                
End Function
'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                
End Function
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLEMULTI)		
    Set gActiveElement = document.ActiveElement   						
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)  
    Set gActiveElement = document.ActiveElement                                 
End Function
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Dim strVal
        
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbDelete = False													
    
    frm1.txtMode.value = Parent.UID_M0003
    
    If LayerShowHide(1) = False Then
         Exit Function
    End If

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)									
   
    DbDelete = True                                                     
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()													
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'*********************************************************************************************************
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey
    Dim strVal
    
    DbQuery = False  
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  
    If LayerShowHide(1) = False Then Exit Function
    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtRcptNo=" & .hdnRcptNo.value
		    strVal = strVal & "&txtMvmtNo=" & .hdnMvmtNo.value
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)
		End if

		Call RunMyBizASP(MyBizASP, strVal)									
    End With
    
    DbQuery = True
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()													
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE											
    
    Call ggoOper.LockField(Document, "Q")								
	lgBlnFlgChgValue = False	
	Call SetToolBar("11101011000111")
	Call ChangeTag(True)
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if

End Function
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal,strDel
	Dim intIndex
		
    DbSave = False                                                      

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""			
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text			
				
				Case ggoSpread.InsertFlag 
		   			strVal = strVal & "C" & Parent.gColSep	
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_ItemNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Spec
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GrQty
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep	
					.vspdData.Col = C_GRUnit
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_TrackingNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_DocAmt
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep		
					.vspdData.Col = C_Cur
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_PlantCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_PlantNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SlCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SlNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_LotNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_LotSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_MakerLotNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_MakerLotSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GRNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GRSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_StoNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_StoSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SGiNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SGiSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Base_Unit
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Mvmt_prc
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep
					.vspdData.Col = C_Locamt
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Mvmt_no
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Base_Qty
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep	
					.vspdData.Col = C_PUR_GRP
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep	
				    
				    strVal = strVal & lRow & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		            
			Case ggoSpread.DeleteFlag
					
					strVal = strVal & "D" & Parent.gColSep
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_ItemNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Spec
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GrQty
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep	
					.vspdData.Col = C_GRUnit
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_TrackingNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_DocAmt
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep		
					.vspdData.Col = C_Cur
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_PlantCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_PlantNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SlCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SlNm
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_LotNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_LotSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_MakerLotNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_MakerLotSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GRNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_GRSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_StoNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_StoSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SGiNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_SGiSeqNo
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Base_Unit
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Mvmt_prc
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep
					.vspdData.Col = C_Locamt
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Mvmt_no
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep			
					.vspdData.Col = C_Base_Qty
					strVal = strVal & UNIConvNum(.vspdData.Text,0) & Parent.gColSep	
					.vspdData.Col = C_PUR_GRP
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep	
				    
				    strVal = strVal & lRow & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		            
		   	End Select 
		Next
    	
    	.txtMaxRows.value = lGrpCnt-1
    	
    	if Trim(.txtGroupCd.value) = "" then
    		.vspdData.Row = 1
    		.vspdData.Col = C_PUR_GRP
    		.txtGroupCd.value = .vspdData.text
    	end if
    	
		.txtSpread.value = strVal
		
		If LayerShowHide(1) = False Then
		     Exit Function
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)								
	End With
	
    DbSave = True                                                       
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
   
	Call InitVariables
	Call ChangeTag(true)
	Call fncQuery()
	
End Function

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_ItemNm Or NewCol <= C_ItemNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'==========================================================================================
'   Event Name : changeMvmtType()
'   Event Desc :
'==========================================================================================
Function changeMvmtType()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
	

	IF Trim(frm1.txtMvmtType.value) = "" then
		exit function
	end if

	If gLookUpEnable = False Then
		Exit Function
	End If
	
  	If CheckRunningBizProcess = True Then
		Exit Function
	End If    
	
    changeMvmtType = False                 
  
    If LayerShowHide(1) = False Then
         Exit Function
    End If

    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeMvmtType"
    strVal = strVal & "&txtMvmtType=" & Filtervar(Trim(frm1.txtMvmtType.Value),"","SNM")
        
    Call RunMyBizASP(MyBizASP, strVal)

	lgBlnFlgChgValue = true

    changeMvmtType = True                  

End Function
'==========================================================================================
'   Event Name : changeMvmtIoType()
'   Event Desc :
'==========================================================================================
Function changeSpplCd()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	IF Trim(frm1.txtSupplierCd.value) = "" then
		exit function
	end if

	If gLookUpEnable = False Then
		Exit Function
	End If
	
  	If CheckRunningBizProcess = True Then
		Exit Function
	End If                           
    
    changeSpplCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeSpplCd"
    strVal = strVal & "&txtSupplierCd=" & Filtervar(Trim(frm1.txtSupplierCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true

    changeSpplCd = True            

End Function

'==========================================================================================
'   Event Name : changeMvmtIoType()
'   Event Desc :
'==========================================================================================
Function changeGroupCd()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	IF Trim(frm1.txtGroupCd.value) = "" then
		exit function
	end if
	
	If gLookUpEnable = False Then
		Exit Function
	End If
  	
  	If CheckRunningBizProcess = True Then
		Exit Function
	End If                                          
    
    changeGroupCd = False           
    
	If LayerShowHide(1) = False Then
	    Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeGroupCd"
    strVal = strVal & "&txtGroupCd=" & Filtervar(Trim(frm1.txtGroupCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true

    changeGroupCd = True            

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고이동입고</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSGiRef()" onMouseOver="vbscript:SetAflag" onMouseOut="vbscript:ResetABCflag" onFocus="vbscript:SetAflag" onBlur="vbscript:ResetABCflag">출고참조</A>&nbsp;											
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
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmtNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD CLASS="TD5" NOWRAP>입고형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="23NXXU" OnChange="VBScript:changeMvmtType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMvmtType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>입고일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m9211ma1_fpDateTime1_txtGmDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>출고공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="출고공장" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" OnChange="VBScript:changeSpplCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="출고공장명" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="21XXXU" OnChange="VBScript:changeGroupCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo1" MAXLENGTH=18 SIZE=35 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m9211ma1_OBJECT1_vspdData.js'></script>
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
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td>&nbsp;</td>
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">재고이동요청등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>	
   
