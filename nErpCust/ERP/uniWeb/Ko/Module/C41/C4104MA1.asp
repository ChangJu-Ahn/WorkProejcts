<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : COSTING
*  2. Function Name        : 실제원가관리 
*  3. Program ID           : C4104MA1
*  4. Program Name         : 입고차이금액반영 
*  5. Program Desc         : 입고차이정보 조회, 입고차이금액 반영/취소 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/27
*  8. Modified date(Last)  :
*  9. Modifier (First)     : Cho Ig Sung
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID		= "C4104MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'''Const CookieSplit = 1233
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim C_PlantCd
Dim C_TrnsTypeCd
Dim C_TrnsTypeNm		 
Dim C_MovTypeCd
Dim C_MovTypeNm
Dim C_CostCd
Dim C_ItemAcctNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemDocNo
Dim C_Seq
Dim C_RcptQty		 
Dim C_StdRcptAmt	 
Dim C_ActlRcptAmt
Dim C_RcptDiffAmt	 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Dim lgStrPlantPrevKey
Dim lgStrTrnsPrevKey
Dim lgStrMovPrevKey
Dim lgStrCostPrevKey
Dim lgStrItemPrevKey

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop          


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_PlantCd		= 1
	C_TrnsTypeCd	= 2
	C_TrnsTypeNm	= 3	 
	C_MovTypeCd		= 4
	C_MovTypeNm		= 5
	C_CostCd		= 6
	C_ItemAcctNm	= 7
	C_ItemCd		= 8
	C_ItemNm		= 9
	C_ItemDocNo		=10
	C_Seq			=11
	C_RcptQty		=12 
	C_StdRcptAmt	=13 
	C_ActlRcptAmt	=14
	C_RcptDiffAmt	=15		
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
	lgStrPlantPrevKey	= ""
	lgStrTrnsPrevKey	= ""
	lgStrMovPrevKey		= ""
	lgStrCostPrevKey	= ""
	lgStrItemPrevKey	= ""
	
	lgSortKey         = 1                                       '⊙: initializes sort direction		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgStrPlantPrevKey	= ""
	lgStrItemPrevKey	= ""
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
		
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

	'========================================================================================================
	' Name : CookiePage()
	' Desc : Write or Read cookie value 
	'========================================================================================================
'	Sub CookiePage(Kubun)
'	   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
'	   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
'	End Sub
'
'	'========================================================================================================
'	' Name : MakeKeyStream
'	' Desc : This method set focus to pos of err
'	'========================================================================================================
'	Sub MakeKeyStream(pRow)
'	   
'	   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
'	    lgKeyStream       = Frm1.txtMajorCd.Value & Parent.gColSep                                           'You Must append one character(Parent.gColSep)
'	   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
'	End Sub        

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	    
	With frm1.vspdData
				
		.MaxCols = C_RcptDiffAmt + 1                                                     ' ☜:☜: Add 1 to Maxcols
		.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:

		ggoSpread.Source = Frm1.vspdData
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread 
		   
		ggoSpread.ClearSpreadData

		.ReDraw = false
				
		Call AppendNumberPlace("6","3","0")

		Call GetSpreadColumnPos("A")		


		ggoSpread.SSSetEdit C_PlantCd, "공장코드", 10
		ggoSpread.SSSetEdit C_TrnsTypeCd, "수불구분", 8
		ggoSpread.SSSetEdit C_TrnsTypeNm, "수불구분명", 15
		ggoSpread.SSSetEdit C_MovTypeCd, "수불유형", 10
		ggoSpread.SSSetEdit C_MovTypeNm, "수불유형명", 15        		
		ggoSpread.SSSetEdit C_CostCd, "C/C", 15        		
		ggoSpread.SSSetEdit C_ItemAcctNm, "품목계정", 10
		ggoSpread.SSSetEdit C_ItemCd, "품목코드", 15
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 25
		ggoSpread.SSSetEdit C_ItemDocNo, "수불번호", 15
		ggoSpread.SSSetFloat C_Seq, "수불SEQ", 8, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_RcptQty, "입고수량", 15, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_StdRcptAmt, "표준입고금액", 15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_ActlRcptAmt, "실제입고금액", 15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_RcptDiffAmt, "입고차이금액", 15, Parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

		.ReDraw = true
				
		Call SetSpreadLock 
			    
	End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	With frm1
		    
	.vspdData.ReDraw = False
	ggoSpread.SpreadLock C_PlantCd, -1, C_PlantCd
	ggoSpread.SpreadLock C_TrnsTypeCd, -1, C_TrnsTypeCd
	ggoSpread.SpreadLock C_TrnsTypeNm, -1, C_TrnsTypeNm
	ggoSpread.SpreadLock C_MovTypeCd, -1, C_MovTypeCd
	ggoSpread.SpreadLock C_MovTypeNm, -1, C_MovTypeNm
	ggoSpread.SpreadLock C_CostCd, -1, C_CostCd
	ggoSpread.SpreadLock C_ItemAcctNm, -1, C_ItemAcctNm
	ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
	ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
	ggoSpread.SpreadLock C_ItemDocNo, -1, C_ItemDocNo
	ggoSpread.SpreadLock C_Seq, -1, C_Seq
	ggoSpread.SpreadLock C_RcptQty, -1, C_RcptQty
	ggoSpread.SpreadLock C_StdRcptAmt, -1, C_StdRcptAmt
	ggoSpread.SpreadLock C_ActlRcptAmt, -1, C_ActlRcptAmt
	ggoSpread.SpreadLock C_RcptDiffAmt, -1, C_RcptDiffAmt
	.vspdData.ReDraw = True

	End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
	select case iWhere
		case 1
			arrParam(0) = "공장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_PLANT"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""		' Where Condition
			arrParam(5) = "공장"			
	
			arrField(0) = "PLANT_CD"					' Field명(0)
			arrField(1) = "PLANT_NM"					' Field명(1)
    
			arrHeader(0) = "공장"				' Header명(0)
			arrHeader(1) = "공장명"				' Header명(1)
		case 2
			arrParam(0) = "코스트센터 팝업"				' 팝업 명칭 
			arrParam(1) = "B_COST_CENTER"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtCostCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""		' Where Condition
			arrParam(5) = "코스트센타"			
	
			arrField(0) = "COST_CD"					' Field명(0)
			arrField(1) = "COST_NM"					' Field명(1)
    
			arrHeader(0) = "코스트센터"				' Header명(0)
			arrHeader(1) = "코스트센터명"				' Header명(1)
		case 3
			arrParam(0) = "수불구분 팝업"				' 팝업 명칭 
			arrParam(1) = "B_MINOR a"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtTrnsTypeCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "a.major_cd = " & FilterVar("I0002", "''", "S") & "  and minor_cd in (" & FilterVar("MR", "''", "S") & " ," & FilterVar("PR", "''", "S") & " )"		' Where Condition
			arrParam(5) = "수불구분"			
	
			arrField(0) = "a.MINOR_CD"					' Field명(0)
			arrField(1) = "a.MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "수불구분"				' Header명(0)
			arrHeader(1) = "수불구분명"				' Header명(1)

		case 4
			arrParam(0) = "수불유형 팝업"				' 팝업 명칭 
			arrParam(1) = "B_MINOR a,i_movetype_configuration b"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtMovTypeCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			IF Trim(frm1.txtTrnsTypeCd.value) <> "" Then
				arrParam(4) = "a.major_cd = " & FilterVar("I0001", "''", "S") & "  and b.trns_type in (" & FilterVar("MR", "''", "S") & " ," & FilterVar("PR", "''", "S") & " ) and a.minor_cd = b.mov_type and b.trns_type = " & FilterVar(frm1.txtTrnsTypeCd.value, "''", "S")		' Where Condition
			Else
				arrParam(4) = "a.major_cd = " & FilterVar("I0001", "''", "S") & "  and a.minor_cd = b.mov_type and b.trns_type  in (" & FilterVar("MR", "''", "S") & " ," & FilterVar("PR", "''", "S") & " ) "		' Where Condition
			END IF
			arrParam(5) = "수불유형"			
	
			arrField(0) = "a.MINOR_CD"					' Field명(0)
			arrField(1) = "a.MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "수불유형"				' Header명(0)
			arrHeader(1) = "수불유형명"				' Header명(1)
		case 5
			arrParam(0) = "품목계정 팝업"				' 팝업 명칭 
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('1FINAL','2SEMI') "


			arrParam(5) = "품목계정"			
	
			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "품목계정"				' Header명(0)
			arrHeader(1) = "품목계정명"				' Header명(1)
		case 6
			arrParam(0) = "품목 팝업"				' 팝업 명칭 
			arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			arrParam(4) = "a.item_cd = b.item_cd and b.procur_type <> " & FilterVar("P", "''", "S") & " "		' Where Condition

			arrParam(5) = "품목"			
	
			arrField(0) = "a.ITEM_CD"					' Field명(0)
			arrField(1) = "a.ITEM_NM"					' Field명(1)
    
			arrHeader(0) = "품목"				' Header명(0)
			arrHeader(1) = "품목명"				' Header명(1)
	end select
		
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
	  Select case iWhere
		case 1
			frm1.txtPlantCD.focus		
		case 2
			frm1.txtCostCd.focus		
		case 3
			frm1.txtTrnsTypeCd.focus		
		case 4
			frm1.txtMovTypeCd.focus		
		case 5
			frm1.txtItemAcctCd.focus		
		case 6
			frm1.txtItemCd.focus
	  End Select		
		Exit Function
	Else
		Call SetReturnVal(iWhere,arrRet)
	End If	


End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 

Function SetReturnVal(byVal iWhere,byval arrRet)
	With frm1
		select case iWhere	
			case 1
				.txtPlantCD.focus	
				.txtplantCd.Value	= arrRet(0)
				.txtPlantNm.Value	= arrRet(1)
			case 2
				.txtCostCd.focus	
				.txtCostCd.Value	= arrRet(0)
				.txtCostNm.Value	= arrRet(1)
			case 3
				.txtTrnsTypeCd.focus
				.txtTrnsTypeCd.Value	= arrRet(0)
				.txtTrnsTypeNm.Value	= arrRet(1)
			case 4
				.txtMovTypeCd.focus
				.txtMovTypeCd.Value	= arrRet(0)
				.txtMovTypeNm.Value	= arrRet(1)
			case 5
				.txtItemAcctCd.focus
				.txtItemAcctCd.Value	= arrRet(0)
				.txtItemAcctNm.Value	= arrRet(1)
			case 6
				.txtItemCd.focus
				.txtItemCd.Value	= arrRet(0)
				.txtItemNm.Value	= arrRet(1)
		end select 		
	End With


End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
			C_PlantCd		= iCurColumnPos(1)
			C_TrnsTypeCd	= iCurColumnPos(2)
			C_TrnsTypeNm	= iCurColumnPos(3)	 
			C_MovTypeCd		= iCurColumnPos(4)
			C_MovTypeNm		= iCurColumnPos(5)
			C_CostCd		= iCurColumnPos(6)
			C_ItemAcctNm	= iCurColumnPos(7)
			C_ItemCd		= iCurColumnPos(8)
			C_ItemNm		= iCurColumnPos(9)
			C_ItemDocNo		= iCurColumnPos(10)
			C_Seq			= iCurColumnPos(11)
			C_RcptQty		= iCurColumnPos(12) 
			C_StdRcptAmt	= iCurColumnPos(13) 
			C_ActlRcptAmt	= iCurColumnPos(14)
			C_RcptDiffAmt	= iCurColumnPos(15)		
 
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
'	Call CookiePage (0)                                                              '☜: Check Cookie

	frm1.txtYyyymm.focus
	frm1.BtnExe.disabled = True
	frm1.BtnCnc.disabled = True

			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
	Dim IntRetCD 

	FncQuery = False															 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = Frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	    															
	If Not chkField(Document, "1") Then									         '☜: This function check required field
	   Exit Function
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Call InitVariables                                                           '⊙: Initializes local global variables
'	Call SetDefaultVal
'''	Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery = False Then                                                      '☜: Query db data
	   Exit Function
	End If
	    
	Set gActiveElement = document.ActiveElement   
	FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()
	Dim strVal

	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	Err.Clear                                                                    '☜: Clear err status
	DbQuery = False                                                              '☜: Processing is NG

	Dim intRetCd	
	Dim arrTemp
	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		IF frm1.txtPlantCd.value <> "" Then
			intRetCd = CommonQueryRs("plant_nm","b_plant","plant_cd = " & FilterVar(Trim(frm1.txtPlantCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtPlantNm.value = arrTemp(0)
			ELSE
				frm1.txtPlantNm.value = ""
			ENd IF
		ELSE
			frm1.txtPlantNm.value = ""
		ENd IF

		IF frm1.txtCostCd.value <> "" Then
			intRetCd = CommonQueryRs("cost_nm","b_cost_center","cost_cd = " & FilterVar(Trim(frm1.txtCostCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtCostNm.value = arrTemp(0)
			ELSE
				frm1.txtCostNm.value = ""
			ENd IF
		ELSE
			frm1.txtCostNm.value = ""
		ENd IF

		IF frm1.txtTrnsTypeCd.value <> "" Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd = " & FilterVar("I0002", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtTrnsTypeCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtTrnsTypeNm.value = arrTemp(0)
			ELSE
				frm1.txtTrnsTypeNm.value = ""
			ENd IF
		ELSE
			frm1.txtTrnsTypeNm.value = ""
		ENd IF

		IF frm1.txtMovTypeCd.value <> ""  Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd =" & FilterVar("I0001", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtMovTypeCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtMovTypeNm.value = arrTemp(0)
			ELSE
				frm1.txtMovTypeNm.value = ""
			ENd IF
		ELSE
			frm1.txtMovTypeNm.value = ""
		ENd IF

		IF frm1.txtItemAcctCd.value <> "" Then
			intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd = " & FilterVar("P1001", "''", "S") & "  and minor_cd = " & FilterVar(Trim(frm1.txtItemAcctCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtItemAcctNm.value = arrTemp(0)
			ELSE
				frm1.txtItemAcctNm.value = ""
			ENd IF
		ELSE
			frm1.txtItemAcctNm.value = ""
		ENd IF

		IF frm1.txtItemCd.value <> "" Then
			intRetCd = CommonQueryRs("item_nm","b_item","item_cd = " & FilterVar(Trim(frm1.txtItemCd.value),"''","S" ),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			IF intRetcd = True Then
				arrTemp = Split(lgF0,Chr(11))
				frm1.txtItemNm.value = arrTemp(0)
			ELSE
				frm1.txtItemNm.value = ""
			ENd IF
		ELSE
			frm1.txtItemNm.value = ""
		ENd IF
		
	ENd If


	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	    
	frm1.BtnExe.disabled = True
	frm1.BtnCnc.disabled = True


	With Frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtYyyymm=" & .hYyyymm.value				
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.hCostCd.value)
			strVal = strVal & "&txtTrnsTypeCd=" & Trim(.hTrnsTypeCd.value)
			strVal = strVal & "&txtMovTypeCd=" & Trim(.hMovTypeCd.value)
			strVal = strVal & "&txtItemAcctCd=" & Trim(.hItemAcctCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
			strVal = strVal & "&lgStrPlantPrevKey=" & lgStrPlantPrevKey
			strVal = strVal & "&lgStrTrnsPrevKey=" & lgStrTrnsPrevKey
			strVal = strVal & "&lgStrMovPrevKey=" & lgStrMovPrevKey
			strVal = strVal & "&lgStrCostPrevKey=" & lgStrCostPrevKey
			strVal = strVal & "&lgStrItemPrevKey=" & lgStrItemPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS_D
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtYyyymm=" & strYear & strMonth
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.txtCostCd.value)
			strVal = strVal & "&txtTrnsTypeCd=" & Trim(.txtTrnsTypeCd.value)
			strVal = strVal & "&txtMovTypeCd=" & Trim(.txtMovTypeCd.value)
			strVal = strVal & "&txtItemAcctCd=" & Trim(.txtItemAcctCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&lgStrPlantPrevKey=" & lgStrPlantPrevKey
			strVal = strVal & "&lgStrTrnsPrevKey=" & lgStrTrnsPrevKey
			strVal = strVal & "&lgStrMovPrevKey=" & lgStrMovPrevKey
			strVal = strVal & "&lgStrCostPrevKey=" & lgStrCostPrevKey
			strVal = strVal & "&lgStrItemPrevKey=" & lgStrItemPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
'			strVal = strVal & "&lgMaxCount=" & C_SHEETMAXROWS_D
		End If
	End With

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	DbQuery = True                                                               '☜: Processing is OK
	Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("1100000000011111")                                              '☆: Developer must customize
	frm1.BtnExe.disabled = False
	frm1.BtnCnc.disabled = False

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'=======================================================================================================
'	Name : ExeReflect()
'	Description : 평가금액 반영작업 
'=======================================================================================================
Function ExeReflect()
	Dim IntRetCD 
	    
	Dim strVal
	Dim strYear
	Dim strMonth
	Dim strDay

	If Not chkField(Document, "1") Then									         '☜: This function check required field
	   Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If


	ExeReflect = False															 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	strVal = BIZ_PGM_ID & "?txtMode=" & "ExeReflect"
	strVal = strVal & "&txtYyyymm=" & strYear & strMonth

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	    
	ExeReflect = True                                                              '☜: Processing is OK

	Set gActiveElement = document.ActiveElement   
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc :
'=======================================================================================================
Function ExeReflectOk()
Dim IntRetCD 

	window.status = "반영 작업 완료"

	IntRetCD =DisplayMsgBox("990000","X","X","X")

	Call FncQuery
			
End Function

'=======================================================================================================
'	Name : ExeCancel()
'	Description : 평가금액 반영작업 
'=======================================================================================================
Function ExeCancel()
	Dim IntRetCD 
	    
	Dim strVal
	Dim strYear
	Dim strMonth
	Dim strDay

	If Not chkField(Document, "1") Then									         '☜: This function check required field
	   Exit Function
	End If

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	ExeCancel = False															 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	strVal = BIZ_PGM_ID & "?txtMode=" & "ExeCancel"
	strVal = strVal & "&txtYyyymm=" & strYear & strMonth

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	    
	ExeCancel = True                                                              '☜: Processing is OK

	Set gActiveElement = document.ActiveElement   

End Function

'======================================================================================================
' Function Name : ExeCancelOk
' Function Desc :
'=======================================================================================================
Function ExeCancelOk()
Dim IntRetCD

	window.status = "취소 작업 완료"

	IntRetCD =DisplayMsgBox("990000","X","X","X")

	Call MainQuery

End Function

Function OpenPopupGL()
	Dim iCalledAspName
	Dim IntRetCD

 
	Dim arrRet
	Dim arrParam(1)	
    
	If lgIsOpenPop = True Then Exit Function

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemDocNo
	
	if Trim(frm1.vspdData.Value) = "" then
       Call DisplayMsgBox("169804","X", "X", "X")    '수불번호가 필요합니다 
       Exit Function
    End If
	
		   	
	arrParam(0) = ""			'회계전표번호 
	arrParam(1) = Trim(frm1.vspdData.Value) & "-" & Trim(frm1.txtYyyymm.Year)   '수불번호'
    
	lgIsOpenPop = True

	
	iCalledAspName = AskPRAspName("A5120RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5120RA1","x")
		IsOpenPop = False
		Exit Function
	End If
   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

 	
End Function


Function OpenMoveDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim strYear,strMonth,strDay
	Dim arrRet
	Dim Param1  '사업장코드 
	Dim Param2  '사업장명 
	Dim Param3  '수불번호 
	Dim Param4	'수불발생일 
	Dim Param5	'회계전표발생일 
	Dim Param6  '수불구분 
	Dim Param7  '수불구분 
	
	If lgIsOpenPop = True Then Exit Function


	IF lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("800167","X", "X", "X")    '조회를 먼저 하세요 
		Exit Function	    
	ENd IF	
	
	With frm1.vspdData	
	


		.Row = .ActiveRow 
		.Col = C_MovTypeNm
		
	    Param7 = .Value


		
		
		Param4 = UniDateAdd("D",-1,UniDateAdd("m", 1,UniConvYYYYMMDDToDate(parent.gDateFormat ,frm1.txtYyyymm.Year, frm1.txtYyyymm.Month, "01"),parent.gDateFormat),parent.gDateFormat) 
		Param5 = UniDateAdd("D",-1,UniDateAdd("m", 1,UniConvYYYYMMDDToDate(parent.gDateFormat ,frm1.txtYyyymm.Year, frm1.txtYyyymm.Month, "01"),parent.gDateFormat),parent.gDateFormat) 

		.Row = .ActiveRow 
		.Col = C_TrnsTypeNm
		Param6 = Trim(.Value)
		
  
	    
		If .MaxRows = 0 Then
			Call DisplayMsgBox("169804","X", "X", "X")    '수불번호가 필요합니다 
			Exit Function
		else
		   .Col = C_ItemDocNo			: .Row = .ActiveRow : Param3 = Trim(.Text )
			IF Param3 = "" Then
				Call DisplayMsgBox("169804","X", "X", "X")    '수불번호가 필요합니다 
				Exit Function
			END IF
		End If	

		ggoSpread.Source = frm1.vspdData    
		.Row = .ActiveRow 
		.Col = C_ItemDocNo
		
		IntRetCD = CommonQueryRs("a.biz_area_cd,b.biz_area_nm","i_goods_movement_header a,b_biz_area b","a.biz_area_cd = b.biz_area_cd and a.item_document_no = " & FilterVar(.Value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If IntRetCD = False Then
			Call DisplayMsgBox("169803","X", "X", "X")    '사업장자료가 필요합니다 
			Exit Function	    
		Else
		    Param1 = Trim(Replace(lgF0,Chr(11),""))
		    Param2 = Trim(Replace(lgF1,Chr(11),""))
		End If
		
    End With
    	
    if Param3 = "" then
       Call DisplayMsgBox("169804","X", "X", "X")    '수불번호가 필요합니다 
    	Exit Function
    End If
	
	lgIsOpenPop = True

	iCalledAspName = AskPRAspName("I1711RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1711RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
    		'Call SetPartRef(arrRet)
	End If	
	
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row )
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData


    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    Else
    	frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_ItemCd
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub



Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
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
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
  

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPlantPrevKey <> "" Then                         
      	   DbQuery
    	End If
    End if
End Sub

Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입고차이반영</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A  href="vbscript:OpenMoveDtlRef()">수불상세 |</A>&nbsp;<A  href="vbscript:OpenPopupGL()">회계전표정보 </A></TD>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="작업년월" id=txtYyyymm> </OBJECT>');</SCRIPT>
									</TD>								
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(1)">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">코스트센타</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(2)">
										<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=30 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">수불구분</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtTrnsTypeCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="수불구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrnsTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(3)">
										<INPUT TYPE=TEXT NAME="txtTrnsTypeNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">수불유형</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtMovTypeCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovTypeCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(4)">
										<INPUT TYPE=TEXT NAME="txtMovTypeNm" SIZE=30 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtItemAcctCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(5)">
										<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(6)">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14X">
									</TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TDT">
									<TD CLASS="TD6">
									<TD CLASS="TD5" NOWRAP>총 계</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSum style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 160px" title="FPDOUBLESINGLE" ALT="총계" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
	                                </TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
					    <TD WIDTH=10>&nbsp;</TD>
						<TD>
						    <BUTTON NAME="btnExe" CLASS="CLSSBTN" ONCLICK="vbscript:ExeReflect()" >반영</BUTTON>&nbsp;<BUTTON NAME="btnCnc" CLASS="CLSSBTN" ONCLICK="vbscript:ExeCancel()" >취소</BUTTON>
						</TD>
						<TD WIDTH=*>&nbsp;</TD>

					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"    tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hTrnsTypeCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hMovTypeCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

