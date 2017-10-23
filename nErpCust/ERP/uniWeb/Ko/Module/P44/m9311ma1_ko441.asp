<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2230ma1_KO441
'*  4. Program Name			: 일별생산계획수립(S)
'*  5. Program Desc			:
'*  6. Business ASP List	: +p2230ma1_KO441.asp		'☆: List MPS
							  +p2230mb1_KO441.asp		'☆: MPS(query, save)
'*  7. Modified date(First)	:
'*  8. Modified date(Last)	:
'*  9. Modifier (First)		: Lee Ho Jun
'* 10. Modifier (Last)		: 
'* 11. Comment				: 일별생산계획수립(S)
'* 12. History              : 일별로 생산계획을 수립한다.
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID        = "m9311mb1_KO441.asp"
Const BIZ_PGM_SAVE_ID1  = "m9311mb2_KO441.asp"
Const BIZ_PGM_SAVE_ID2  = "m9311mb3_KO441.asp"


Dim LocSvrDate
Dim strDate
Dim EndDate
Dim lblnWinEvent
Dim releaseFlag

EndDate = UNIConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)

LocSvrDate = "<%=GetSvrDate%>"

strDate = UniConvDateAToB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat) 	'☆: 초기화면에 뿌려지는 마지막 날짜 

Dim IsOpenPop
Dim queryboolean
queryboolean = False

'Dim C_Item_Cd
'Dim C_Item_Nm
Dim C_Tracking_No
Dim C_Type_cd
Dim C_Type
Dim C_Qty_Day_0
Dim C_Qty_Day_1
Dim C_Qty_Day_2
Dim C_Qty_Day_3
Dim C_Qty_Day_4
Dim C_Qty_Day_5
Dim C_Qty_Day_6
Dim C_Qty_Day_7
Dim C_Qty_Day_8
Dim C_Qty_Day_9
Dim C_Qty_Day_10
Dim C_Qty_Day_11
Dim C_Qty_Day_12
Dim C_Qty_Day_13
Dim C_Qty_Day_14
Dim C_Qty_Day_15
Dim C_Qty_Day_16
Dim C_Qty_Day_17
Dim C_Qty_Day_18
Dim C_Qty_Day_19
Dim C_Qty_Day_20
Dim C_Qty_Day_21
Dim C_Qty_Day_22
Dim C_Qty_Day_23
Dim C_Qty_Day_24
Dim C_Qty_Day_25
Dim C_Qty_Day_26
Dim C_Qty_Day_27
Dim C_Qty_Day_28
Dim C_Qty_Day_29
Dim C_Qty_Day_30
Dim C_Qty_Month_1
Dim C_Qty_Month_2
Dim C_Qty_Month_3
Dim C_Plant_Cd
Dim C_Qty_Day_0_Hidden
Dim C_Qty_Day_1_Hidden
Dim C_Qty_Day_2_Hidden
Dim C_Qty_Day_3_Hidden
Dim C_Qty_Day_4_Hidden
Dim C_Qty_Day_5_Hidden
Dim C_Qty_Day_6_Hidden
Dim C_Qty_Day_7_Hidden
Dim C_Qty_Day_8_Hidden
Dim C_Qty_Day_9_Hidden
Dim C_Qty_Day_10_Hidden
Dim C_Qty_Day_11_Hidden
Dim C_Qty_Day_12_Hidden
Dim C_Qty_Day_13_Hidden
Dim C_Qty_Day_14_Hidden
Dim C_Qty_Day_15_Hidden
Dim C_Qty_Day_16_Hidden
Dim C_Qty_Day_17_Hidden
Dim C_Qty_Day_18_Hidden
Dim C_Qty_Day_19_Hidden
Dim C_Qty_Day_20_Hidden
Dim C_Qty_Day_21_Hidden
Dim C_Qty_Day_22_Hidden
Dim C_Qty_Day_23_Hidden
Dim C_Qty_Day_24_Hidden
Dim C_Qty_Day_25_Hidden
Dim C_Qty_Day_26_Hidden
Dim C_Qty_Day_27_Hidden
Dim C_Qty_Day_28_Hidden
Dim C_Qty_Day_29_Hidden
Dim C_Qty_Day_30_Hidden

Dim C_ITEM_SEQ
Dim C_CONFIG_FLAG
Dim C_LOC
Dim C_ITEM_CD
Dim C_POPUP2
Dim C_ITEM_NM
Dim C_BASIC_UNIT
Dim C_REQ_QTY
Dim C_ISSUE_QTY
Dim C_PRODT_ORDER_NO
Dim C_REMARK1

Dim C_ISSUE_REQ_NO
Dim C_TRNS_TYPE
Dim C_REQ_DT
Dim C_ISSUE_TYPE
Dim C_MOV_TYPE
Dim C_DEPT_CD
Dim C_EMP_NO
Dim C_REMARK


Dim dayAr, dayAr2
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
    Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat, "2")
    Call ggoOper.LockField(Document, "N")                                           '⊙: Lock Field

    Call FormatDATEField(frm1.txtPoDt)
    Call LockObjectField(frm1.txtPoDt, "R")
    Call SetDefaultVal
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox
    Call InitSpreadComboBox
    Call CookiePage

'    Call SetToolbar("1100100100001111")                                             '버튼 툴바 제어
    call SetToolBar("1110110100001111")
    
    ggoOper.SetReqAttr	frm1.txtPlantCd, "Q"        '2008-03-26 5:56오후 :: hanc

    Set gActiveElement = document.activeElement
    
End Sub

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call LoadInfTB19029A("I", "P", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' Function Name : SetDefaultVal
' Function Desc : Set Default Values
'==========================================================================================================
Sub SetDefaultVal()
    Dim strYear
    Dim strMonth
    Dim strDay
  
    frm1.txtYYYYMM.Text = UniConvDateAToB(strDate, Parent.gDateFormat, Parent.gDateFormatYYYYMM)

    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
    frm1.txtPoDt.Text = EndDate

    queryboolean = False
End Sub

'==========================================================================================
'   Event Name : txtPoDt
'==========================================================================================
Sub txtPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoDt.Action = 7
		SetFocusToDocument("M")	
		frm1.txtPoDt.focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtPoDt
'==========================================================================================
Sub txtPoDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i, startDate
    Dim startDate1, startDate2
    i = 0
    ReDim dayAr(30)

' 캡션 날짜로 표시
    startDate = Replace(frm1.txtYYYYMM.Text, "-", "")

    if  CommonQueryRs(" Month(DATEADD(m, +1, '"&startDate&"'))","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
    Else
        startDate1 = Replace(lgF0, Chr(11), "") + "월"
    End If
    
    if  CommonQueryRs(" Month(DATEADD(m, +2, '"&startDate&"'))","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
    Else
        startDate2 = Replace(lgF0, Chr(11), "") + "월"
    End If

    For i = 0 To 30
        dayAr(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
    Next

    Call InitSpreadPosVariables()

    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

        .ReDraw = false

        .MaxCols = C_LOC + 1                                      <%'☜: 최대 Columns의 항상 1개 증가시킴 %>

        .Col = .MaxCols                                                         <%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0

        ggoSpread.Source = Frm1.vspdData

        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos()

        ggoSpread.SSSetEdit 	C_ITEM_CD         , "품목",         18,,,18,2
    	ggoSpread.SSSetButton 	C_POPUP2                
        ggoSpread.SSSetEdit 	C_ITEM_NM         , "품목명",       18,,,18,2
        ggoSpread.SSSetEdit 	C_BASIC_UNIT      , "단위",         18,,,18,2
        ggoSpread.SSSetFloat    C_REQ_QTY         , "불출의뢰수량", 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_ISSUE_QTY       , "불출수량" ,    10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        
        ggoSpread.SSSetEdit 	C_PRODT_ORDER_NO  , "제조오더번호", 18,,,18,2
        ggoSpread.SSSetEdit 	C_REMARK1         , "비고",         18,,,100,2
        
        ggoSpread.SSSetEdit 	C_ISSUE_REQ_NO    , "불출의뢰번호", 18,,,18,2
        ggoSpread.SSSetEdit 	C_TRNS_TYPE       , "출고구분",     18,,,18,2
        ggoSpread.SSSetEdit 	C_REQ_DT          , "불출의뢰일",   18,,,18,2
        ggoSpread.SSSetEdit 	C_ISSUE_TYPE      , "불출의뢰유형", 18,,,18,2
        ggoSpread.SSSetEdit 	C_DEPT_CD         , "의뢰부서",     18,,,18,2
        ggoSpread.SSSetEdit 	C_EMP_NO          , "의뢰담당자",   18,,,18,2
        ggoSpread.SSSetEdit 	C_REMARK          , "특이사항",     18,,,18,2
        ggoSpread.SSSetEdit 	C_ITEM_SEQ        , "순번",         18,1
        ggoSpread.SSSetEdit 	C_CONFIG_FLAG     , "확정여부",     18,,,18,2
        ggoSpread.SSSetEdit 	C_LOC             , "LOC",          18,,,18,2
        
        

		call ggoSpread.MakePairsColumn(C_ITEM_CD, C_POPUP2)

        Call ggoSpread.SSSetColHidden(C_ISSUE_REQ_NO      , C_ISSUE_REQ_NO      , True)  
        Call ggoSpread.SSSetColHidden(C_TRNS_TYPE         , C_TRNS_TYPE         , True)  
        Call ggoSpread.SSSetColHidden(C_REQ_DT            , C_REQ_DT            , True)  
        Call ggoSpread.SSSetColHidden(C_ISSUE_TYPE        , C_ISSUE_TYPE        , True)  
        Call ggoSpread.SSSetColHidden(C_DEPT_CD           , C_DEPT_CD           , True)  
        Call ggoSpread.SSSetColHidden(C_EMP_NO            , C_EMP_NO            , True)  
        Call ggoSpread.SSSetColHidden(C_REMARK            , C_REMARK            , True)  
        Call ggoSpread.SSSetColHidden(C_ITEM_SEQ            , C_ITEM_SEQ            , True)  
        Call ggoSpread.SSSetColHidden(C_CONFIG_FLAG       , C_CONFIG_FLAG          , True)  
        Call ggoSpread.SSSetColHidden(C_LOC       , C_LOC          , True)  

'        ggoSpread.SSSetEdit     C_Item_Cd       , "품목"   , 10
'        ggoSpread.SSSetEdit     C_Item_Nm       , "품목명" , 20
'        ggoSpread.SSSetEdit     C_Tracking_No   , "Tracking No"   , 10
'        ggoSpread.SSSetEdit     C_Type_cd	    , "구분"   , 10
'        ggoSpread.SSSetEdit     C_Type		    , "구분"   , 10
'        
'        for i = 1 To 31
'			'ggoSpread.SSSetFloat    i + 5     , dayAr(i-1) , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'			ggoSpread.SSSetFloat    i + 5     , Cstr(i) + "일" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'        Next
'        
'       	ggoSpread.SSSetFloat    C_Qty_Month_1     , "월합계" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'  		ggoSpread.SSSetFloat    C_Qty_Month_2     , startDate1 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'		ggoSpread.SSSetFloat    C_Qty_Month_3     , startDate2 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'		ggoSpread.SSSetEdit     C_Plant_Cd		  , "구분"   , 10
'		
'		for i = 1 To 31
'			ggoSpread.SSSetFloat    i + 40     , dayAr(i-1), 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'        Next
'
'        Call ggoSpread.SSSetColHidden(C_Plant_Cd         , C_Plant_Cd         , True)
'        Call ggoSpread.SSSetColHidden(C_Type_Cd          , C_Type_Cd          , True)
'        Call ggoSpread.SSSetColHidden(C_Qty_Day_0_Hidden , C_Qty_Day_30_Hidden , True)
'
'
'        .Col = C_Item_Cd        : .ColMerge = 2
'        .Col = C_Item_Nm        : .ColMerge = 2
'        .Col = C_Tracking_no    : .ColMerge = 2
'       '.Col = C_Item_Stock_Qty : .ColMerge = 2
'
'        ggoSpread.SSSetSplit2(5)

        Call SetSpreadLock()

        frm1.vspdData.ReDraw = True
        
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
        frm1.vspdData.ReDraw = False
        
        ggoSpread.SpreadLock     C_BASIC_UNIT       , -1, C_BASIC_UNIT
        ggoSpread.SpreadLock     C_Item_Nm       , -1, C_Item_Nm
        ggoSpread.SpreadLock     C_ISSUE_QTY       , -1, C_ISSUE_QTY

        ggoSpread.SpreadLock     C_PRODT_ORDER_NO   , -1, C_PRODT_ORDER_NO
        ggoSpread.SpreadLock     C_Type_cd		 , -1, C_Type_cd
        ggoSpread.SpreadLock     C_Type          , -1, C_Type
        ggoSpread.SpreadLock     C_Qty_Month_1   , -1, C_Qty_Month_1
        ggoSpread.SpreadLock     C_Qty_Month_2   , -1, C_Qty_Month_2
        ggoSpread.SpreadLock     C_Qty_Month_3   , -1, C_Qty_Month_3
        ggoSpread.SpreadLock     C_Plant_Cd      , -1, C_Plant_Cd         

		ggoSpread.SpreadUNLock   C_ITEM_CD     , -1, C_ITEM_CD
		ggoSpread.SpreadUNLock   C_REQ_QTY     , -1, C_REQ_QTY
'		ggoSpread.SpreadUNLock   C_ISSUE_QTY     , -1, C_ISSUE_QTY

        ggoSpread.SSSetRequired  C_ITEM_CD, -1
        ggoSpread.SSSetRequired  C_REQ_QTY, -1

        frm1.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()

    C_ITEM_CD          = 1 
    C_POPUP2		   = 2     
    C_ITEM_NM          = 3      
    C_BASIC_UNIT       = 4      
    C_REQ_QTY          = 5      
    C_ISSUE_QTY        = 6      
    C_PRODT_ORDER_NO   = 7      
    C_REMARK1          = 8      
    C_ISSUE_REQ_NO     = 9 
    C_TRNS_TYPE        = 10
    C_REQ_DT           = 11
    C_ISSUE_TYPE       = 12
    C_DEPT_CD          = 13
    C_EMP_NO           = 14
    C_REMARK           = 15
    C_ITEM_SEQ         = 16
    C_CONFIG_FLAG      = 17
    C_LOC               = 18


'    C_Item_Cd           = 1 
'    C_Item_Nm           = 2
'    C_Tracking_No       = 3
'    C_Type_cd           = 4
'    C_Type		        = 5
'    C_Qty_Day_0         = 6
'    C_Qty_Day_1         = 7
'    C_Qty_Day_2         = 8
'    C_Qty_Day_3         = 9
'    C_Qty_Day_4         = 10
'    C_Qty_Day_5         = 11
'    C_Qty_Day_6         = 12
'    C_Qty_Day_7         = 13
'    C_Qty_Day_8         = 14
'    C_Qty_Day_9         = 15
'    C_Qty_Day_10        = 16
'    C_Qty_Day_11        = 17
'    C_Qty_Day_12        = 18
'    C_Qty_Day_13        = 19
'    C_Qty_Day_14        = 20
'    C_Qty_Day_15        = 21
'    C_Qty_Day_16        = 22
'    C_Qty_Day_17        = 23
'    C_Qty_Day_18        = 24
'    C_Qty_Day_19        = 25
'    C_Qty_Day_20        = 26
'    C_Qty_Day_21        = 27
'    C_Qty_Day_22        = 28
'    C_Qty_Day_23        = 29
'    C_Qty_Day_24        = 30
'    C_Qty_Day_25        = 31
'    C_Qty_Day_26        = 32
'    C_Qty_Day_27        = 33
'    C_Qty_Day_28        = 34
'    C_Qty_Day_29        = 35
'    C_Qty_Day_30        = 36
'    C_Qty_Month_1       = 37
'    C_Qty_Month_2       = 38
'    C_Qty_Month_3       = 39
'    C_Plant_Cd          = 40
'    C_Qty_Day_0_Hidden  = 41
'    C_Qty_Day_1_Hidden  = 42
'    C_Qty_Day_2_Hidden  = 43
'    C_Qty_Day_3_Hidden  = 44
'    C_Qty_Day_4_Hidden  = 45
'    C_Qty_Day_5_Hidden  = 46
'    C_Qty_Day_6_Hidden  = 47
'    C_Qty_Day_7_Hidden  = 48
'    C_Qty_Day_8_Hidden  = 49
'    C_Qty_Day_9_Hidden  = 50
'    C_Qty_Day_10_Hidden = 51
'    C_Qty_Day_11_Hidden = 52
'    C_Qty_Day_12_Hidden = 53
'    C_Qty_Day_13_Hidden = 54
'    C_Qty_Day_14_Hidden = 55
'    C_Qty_Day_15_Hidden = 56
'    C_Qty_Day_16_Hidden = 57
'    C_Qty_Day_17_Hidden = 58
'    C_Qty_Day_18_Hidden = 59
'    C_Qty_Day_19_Hidden = 60
'    C_Qty_Day_20_Hidden = 61
'    C_Qty_Day_21_Hidden = 62
'    C_Qty_Day_22_Hidden = 63
'    C_Qty_Day_23_Hidden = 64
'    C_Qty_Day_24_Hidden = 65
'    C_Qty_Day_25_Hidden = 66
'    C_Qty_Day_26_Hidden = 67
'    C_Qty_Day_27_Hidden = 68
'    C_Qty_Day_28_Hidden = 69
'    C_Qty_Day_29_Hidden = 70
'    C_Qty_Day_30_Hidden = 71
    
End SuB

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos

    ggoSpread.Source = frm1.vspdData

    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    C_ITEM_CD          = iCurColumnPos(1)
    C_POPUP2		   = iCurColumnPos(2)
    C_ITEM_NM          = iCurColumnPos(3)
    C_BASIC_UNIT       = iCurColumnPos(4)
    C_REQ_QTY          = iCurColumnPos(5)
    C_ISSUE_QTY        = iCurColumnPos(6)
    C_PRODT_ORDER_NO   = iCurColumnPos(7)
    C_REMARK1          = iCurColumnPos(8)
    C_ISSUE_REQ_NO     = iCurColumnPos(9) 
    C_TRNS_TYPE        = iCurColumnPos(10)
    C_REQ_DT           = iCurColumnPos(11)
    C_ISSUE_TYPE       = iCurColumnPos(12)
    C_DEPT_CD          = iCurColumnPos(13)
    C_EMP_NO           = iCurColumnPos(14)
    C_REMARK           = iCurColumnPos(15)
    C_ITEM_SEQ         = iCurColumnPos(16)
    C_CONFIG_FLAG      = iCurColumnPos(17)
    C_LOC      = iCurColumnPos(18)


'    C_Item_Cd           = iCurColumnPos(1)  
'    C_Item_Nm           = iCurColumnPos(2)
'    C_Tracking_no       = iCurColumnPos(3)
'    C_Type_cd		    = iCurColumnPos(4)
'    C_Type              = iCurColumnPos(5)
'    C_Qty_Day_0         = iCurColumnPos(6)
'    C_Qty_Day_1         = iCurColumnPos(7)
'    C_Qty_Day_2         = iCurColumnPos(8)
'    C_Qty_Day_3         = iCurColumnPos(9)
'    C_Qty_Day_4         = iCurColumnPos(10)
'    C_Qty_Day_5         = iCurColumnPos(11)
'    C_Qty_Day_6         = iCurColumnPos(12)
'    C_Qty_Day_7         = iCurColumnPos(13)
'    C_Qty_Day_8         = iCurColumnPos(14)
'    C_Qty_Day_9         = iCurColumnPos(15)
'    C_Qty_Day_10        = iCurColumnPos(16)
'    C_Qty_Day_11        = iCurColumnPos(17)
'    C_Qty_Day_12        = iCurColumnPos(18)
'    C_Qty_Day_13        = iCurColumnPos(19)
'    C_Qty_Day_14        = iCurColumnPos(20)
'    C_Qty_Day_15        = iCurColumnPos(21)
'    C_Qty_Day_16        = iCurColumnPos(22)
'    C_Qty_Day_17        = iCurColumnPos(23)
'    C_Qty_Day_18        = iCurColumnPos(24)
'    C_Qty_Day_19        = iCurColumnPos(25)
'    C_Qty_Day_20        = iCurColumnPos(26)
'    C_Qty_Day_21        = iCurColumnPos(27)
'    C_Qty_Day_22        = iCurColumnPos(28)
'    C_Qty_Day_23        = iCurColumnPos(29)
'    C_Qty_Day_24        = iCurColumnPos(30)
'    C_Qty_Day_25        = iCurColumnPos(31)
'    C_Qty_Day_26        = iCurColumnPos(32)
'    C_Qty_Day_27        = iCurColumnPos(33)
'    C_Qty_Day_28        = iCurColumnPos(34)
'    C_Qty_Day_29        = iCurColumnPos(35)
'    C_Qty_Day_30        = iCurColumnPos(36)
'    C_Qty_Month_1       = iCurColumnPos(37)
'    C_Qty_Month_2       = iCurColumnPos(38)
'    C_Qty_Month_3       = iCurColumnPos(39)
'    C_Plant_Cd          = iCurColumnPos(40)
'    C_Qty_Day_0_Hidden  = iCurColumnPos(41)
'    C_Qty_Day_1_Hidden  = iCurColumnPos(42)
'    C_Qty_Day_2_Hidden  = iCurColumnPos(43)
'    C_Qty_Day_3_Hidden  = iCurColumnPos(44)
'    C_Qty_Day_4_Hidden  = iCurColumnPos(45)
'    C_Qty_Day_5_Hidden  = iCurColumnPos(46)
'    C_Qty_Day_6_Hidden  = iCurColumnPos(47)
'    C_Qty_Day_7_Hidden  = iCurColumnPos(48)
'    C_Qty_Day_8_Hidden  = iCurColumnPos(49)
'    C_Qty_Day_9_Hidden  = iCurColumnPos(50)
'    C_Qty_Day_10_Hidden = iCurColumnPos(51)
'    C_Qty_Day_11_Hidden = iCurColumnPos(52)
'    C_Qty_Day_12_Hidden = iCurColumnPos(53)
'    C_Qty_Day_13_Hidden = iCurColumnPos(54)
'    C_Qty_Day_14_Hidden = iCurColumnPos(55)
'    C_Qty_Day_15_Hidden = iCurColumnPos(56)
'    C_Qty_Day_16_Hidden = iCurColumnPos(57)
'    C_Qty_Day_17_Hidden = iCurColumnPos(58)
'    C_Qty_Day_18_Hidden = iCurColumnPos(59)
'    C_Qty_Day_19_Hidden = iCurColumnPos(60)
'    C_Qty_Day_20_Hidden = iCurColumnPos(61)
'    C_Qty_Day_21_Hidden = iCurColumnPos(62)
'    C_Qty_Day_22_Hidden = iCurColumnPos(63)
'    C_Qty_Day_23_Hidden = iCurColumnPos(64)
'    C_Qty_Day_24_Hidden = iCurColumnPos(65)
'    C_Qty_Day_25_Hidden = iCurColumnPos(66)
'    C_Qty_Day_26_Hidden = iCurColumnPos(67)
'    C_Qty_Day_27_Hidden = iCurColumnPos(68)
'    C_Qty_Day_28_Hidden = iCurColumnPos(69)
'    C_Qty_Day_29_Hidden = iCurColumnPos(70)
'    C_Qty_Day_30_Hidden = iCurColumnPos(71)
    
End Sub

'==========================================================================================================
' Function Name : InitVariables
' Function Desc : Initialize value
'==========================================================================================================
Sub InitVariables()
    lgIntFlgMode     = parent.OPMD_CMODE                       '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                   '⊙: Indicates that no value changed
    lgIntGrpCount    = 0                                       '⊙: Initializes Group View Size
    lgStrPrevKey     = ""                                      '⊙: initializes Previous Key
    lgSortKey        = 1                                       '⊙: initializes sort direction
End Sub

'=========================================================================================================
' Function Name : InitComboBox
' Function Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
End Sub

'=========================================================================================================
' Function Name : InitSpreadComboBox
' Function Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'=========================================================================================================
' Function Name : SetSpreadColor
' Function Description : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    Dim i, ptfForMps, startDate
    
    i = 0

    With frm1.vspdData

        .Redraw = False


        ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetProtected C_ITEM_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ISSUE_QTY		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PRODT_ORDER_NO		, pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected C_BASIC_UNIT	, pvStartRow, pvEndRow	
		
		
        For i = 1 To frm1.vspdData.MaxRows
			
			.row = i
			.Col = C_Type
			
			If .text = "생판요청" Then
			   ggoSpread.SpreadLock C_Qty_Day_0 ,  i, C_Qty_Day_30_Hidden , i 
			End If 
			      
        Next
    End With
'            If (i Mod 7) = 5 Or (i Mod 7) = 6 Then
'                If startDate <= dayAr(0)  And  dayAr(0)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_0 ,  i, C_Qty_Day_0 , i End If 
'                If startDate <= dayAr(1)  And  dayAr(1)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_1 ,  i, C_Qty_Day_1 , i End If 
'                If startDate <= dayAr(2)  And  dayAr(2)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_2 ,  i, C_Qty_Day_2 , i End If 
'                If startDate <= dayAr(3)  And  dayAr(3)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_3 ,  i, C_Qty_Day_3 , i End If 
'                If startDate <= dayAr(4)  And  dayAr(4)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_4 ,  i, C_Qty_Day_4 , i End If 
'                If startDate <= dayAr(5)  And  dayAr(5)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_5 ,  i, C_Qty_Day_5 , i End If 
'                If startDate <= dayAr(6)  And  dayAr(6)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_6 ,  i, C_Qty_Day_6 , i End If 
'                If startDate <= dayAr(7)  And  dayAr(7)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_7 ,  i, C_Qty_Day_7 , i End If 
'                If startDate <= dayAr(8)  And  dayAr(8)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_8 ,  i, C_Qty_Day_8 , i End If 
'                If startDate <= dayAr(9)  And  dayAr(9)  < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_9 ,  i, C_Qty_Day_9 , i End If 
'                If startDate <= dayAr(10) And  dayAr(10) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_10,  i, C_Qty_Day_10, i End If 
'                If startDate <= dayAr(11) And  dayAr(11) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_11,  i, C_Qty_Day_11, i End If 
'                If startDate <= dayAr(12) And  dayAr(12) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_12,  i, C_Qty_Day_12, i End If 
'                If startDate <= dayAr(13) And  dayAr(13) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_13,  i, C_Qty_Day_13, i End If 
'                If startDate <= dayAr(14) And  dayAr(14) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_14,  i, C_Qty_Day_14, i End If 
'                If startDate <= dayAr(15) And  dayAr(15) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_15,  i, C_Qty_Day_15, i End If 
'                If startDate <= dayAr(16) And  dayAr(16) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_16,  i, C_Qty_Day_16, i End If 
'                If startDate <= dayAr(17) And  dayAr(17) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_17,  i, C_Qty_Day_17, i End If 
'                If startDate <= dayAr(18) And  dayAr(18) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_18,  i, C_Qty_Day_18, i End If 
'                If startDate <= dayAr(19) And  dayAr(19) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_19,  i, C_Qty_Day_19, i End If 
'                If startDate <= dayAr(20) And  dayAr(20) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_20,  i, C_Qty_Day_20, i End If 
'                If startDate <= dayAr(21) And  dayAr(21) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_21,  i, C_Qty_Day_21, i End If 
'                If startDate <= dayAr(22) And  dayAr(22) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_22,  i, C_Qty_Day_22, i End If 
'                If startDate <= dayAr(23) And  dayAr(23) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_23,  i, C_Qty_Day_23, i End If 
'                If startDate <= dayAr(24) And  dayAr(24) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_24,  i, C_Qty_Day_24, i End If 
'                If startDate <= dayAr(25) And  dayAr(25) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_25,  i, C_Qty_Day_25, i End If 
'                If startDate <= dayAr(26) And  dayAr(26) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_26,  i, C_Qty_Day_26, i End If 
'                If startDate <= dayAr(27) And  dayAr(27) < endDate Then ggoSpread.SpreadUnLock C_Qty_Day_27,  i, C_Qty_Day_27, i End If 
'            End If
'        Next
'
'        .Col = 1
'        .Row = .ActiveRow
'        .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
'        .EditMode = True
'
'        .Redraw = True
'
'    End With
End Sub

'------------------------------------------ OpenPlant()  --------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant Popup
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "공장팝업"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "PLANT_CD"
	arrField(1) = "PLANT_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If

End Function

'------------------------------------------ OpenItem()  --------------------------------------------------
'	Name : OpenItem()
'	Description : Item Popup
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)

	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantCd.focus
		Exit function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtItemCd.Value)
	arrParam(2) = ""
	arrParam(3) = ""

	arrField(0) = 1
	arrField(1) = 2
	arrField(2) = 9
	arrField(3) = 6

	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItem(arrRet)
	End If
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : OpenPlant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
End Function
'------------------------------------------  SetItem()  --------------------------------------------------
'	Name : SetItem()
'	Description : OpenItem Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItem(ByRef arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus
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

    call SetToolBar("11101111001111")        '20080305::hanc
    
    FncNew = True                     
	Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData

    '20080218::hanc
    If Not chkFieldByCell(frm1.txtPlantCd, "A",1)	then
       Exit Function
    End If
    If Not chkFieldByCell(frm1.txtPoNo, "A",1)	then
       Exit Function
    End If

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")                     '☜: Data is changed.  Do you want to display it?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.ClearSpreadData
'    If Not chkField(Document, "1") Then                                          '☜: This function check required field
'       Exit Function
'    End If

    frm1.hPlantCd.value = Trim(frm1.txtPlantCd.value)
    'frm1.hYYYYMM.value = Replace(Trim(frm1.txtYYYYMM.Text),"-", "")
    frm1.hYYYYMM.value = Trim(frm1.txtYYYYMM.Text)
    frm1.hItemCd.value = Trim(frm1.txtItemCd.value)
    frm1.hTrackingNo.value = Trim(frm1.txtTrackingNo.value)

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream

    Call DisableToolBar(parent.TBC_QUERY)

    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()
    'lgKeyStream =frm1.hPlantCd.value & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream =frm1.hYYYYMM.value & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & frm1.hItemCd.value & parent.gColSep
    lgKeyStream = lgKeyStream & frm1.hTrackingNo.value & parent.gColSep
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    DbQuery = False
    Err.Clear                                                                        '☜: Clear err status

    If LayerShowHide(1) = False Then
         Exit Function
    End If

    Dim strVal
    
    call SetToolBar("11100000000111")
    
    With frm1
        strVal = BIZ_PGM_ID & "?txtMode="      & parent.UID_M0001
        strVal = strVal     & "&txtKeyStream=" & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtPlantCd=" & Trim(.txtPlantCd.value)                       '☜: Query Key
        strVal = strVal     & "&txtPoNo=" & Trim(.txtPoNo.value)                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="   & frm1.vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
        'strVal = strVal     & "&queryFlag="    & queryFlag                 '☜: Next key tag
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True
End Function

'20080305::hanc
Function DbSaveDelOk()
    call FncNew()
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    'queryFlag = "P"
'    queryFlag = "Q"
    Dim strVal
    Dim rowcnt

    lgIntFlgMode = parent.OPMD_UMODE
'    Call ggoOper.LockField(Document, "Q")                                       '⊙: Lock field
    Call ggoOper.LockField(Document, "N")           '20080313::hanc

	ggoOper.SetReqAttr	frm1.txtPoNo1, "Q"


    Call InitData()
	Call vspdData_Click1(1, 1)
    queryboolean = true

	frm1.vspdData.Col = C_CONFIG_FLAG
	frm1.vspdData.Row = 1 

'MSGBOX "Trim(frm1.vspdData.Text) : " & Trim(frm1.vspdData.Text)	
	if Trim(frm1.vspdData.Text) = "Y" then
		call SetToolBar("11100000000111")
'        call ggoOper.LockField(Document, "Q")
	    frm1.btnCfmSel.disabled = False
        ggoSpread.SpreadLock     C_ITEM_CD       , -1, C_CONFIG_FLAG
		frm1.btnCfm.value = "확정취소"

        ''20080313::hanc:: 확정이면 protected
	    ggoOper.SetReqAttr	frm1.txtPoDt, "Q"
	    ggoOper.SetReqAttr	frm1.txtPoTypeCd, "Q"
	    ggoOper.SetReqAttr	frm1.txtEmp_no, "Q"
	    ggoOper.SetReqAttr	frm1.txtDept_cd, "Q"
	    ggoOper.SetReqAttr	frm1.txtRemark, "Q"
	    ggoOper.SetReqAttr	frm1.txtLoc, "Q"

	else
		call SetToolBar("11101111001111")'200309 헤더삭제는 디테일삭제가 완료되며 자동으로 삭제됨.
'		call ggoOper.LockField(Document, "N")
		frm1.btnCfm.value = "확정"
		frm1.btnCfmSel.disabled = False	
	    Call SetSpreadLock()
	end if

    rowcnt = frm1.vspdData.MaxRows      '20080305::hanc

    IF rowcnt = 0 THEN
        call DbSaveDel
    END IF


    'Call SetQuerySpreadColor()
    Call SetSpreadColor(0, 0)

    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DBPOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
'Function DBPOk()
'    queryFlag = "Q"
'    Call FncQuery
'End Function

'========================================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function SetQuerySpreadColor() 
    Dim iDx

    With frm1.vspdData
        For iDx = (frm1.vspdData.MaxRows-1) To frm1.vspdData.MaxRows
            .Col = -1
            .Row =  iDx
            .BackColor = RGB(255,255,204) '하늘색 
        Next
    End With
End Function

'========================================================================================================
' Function Name : ChangeCaption
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function ChangeCaption()
    
    Dim i, startDate
    Dim startDate1, startDate2
    Dim dayAr1
    i = 0
    ReDim dayAr1(30)
    ReDim dayAr2(30)

' 캡션 날짜로 표시
    startDate = Replace(frm1.txtYYYYMM.Text, "-", "")

    if  CommonQueryRs(" Month(DATEADD(m, +1, '"&startDate&"'+'01'))","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
    Else
        startDate1 = Replace(lgF0, Chr(11), "") + "월"
    End If

    if  CommonQueryRs(" Month(DATEADD(m, +2, '"&startDate&"'+'01'))","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
    Else
        startDate2 = Replace(lgF0, Chr(11), "") + "월"
    End If

'    dayAr(0) = startDate

    For i = 0 To 30
        dayAr1(i) =  Cstr(i+1) + "일"	'UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
        dayAr2(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
    Next

    With frm1.vspdData
        .Col = C_Qty_Day_0  : .Row = 0 : .Text = dayAr1(0)
        .Col = C_Qty_Day_1  : .Row = 0 : .Text = dayAr1(1)
        .Col = C_Qty_Day_2  : .Row = 0 : .Text = dayAr1(2)
        .Col = C_Qty_Day_3  : .Row = 0 : .Text = dayAr1(3)
        .Col = C_Qty_Day_4  : .Row = 0 : .Text = dayAr1(4)
        .Col = C_Qty_Day_5  : .Row = 0 : .Text = dayAr1(5)
        .Col = C_Qty_Day_6  : .Row = 0 : .Text = dayAr1(6)
        .Col = C_Qty_Day_7  : .Row = 0 : .Text = dayAr1(7)
        .Col = C_Qty_Day_8  : .Row = 0 : .Text = dayAr1(8)
        .Col = C_Qty_Day_9  : .Row = 0 : .Text = dayAr1(9)
        .Col = C_Qty_Day_10 : .Row = 0 : .Text = dayAr1(10)
        .Col = C_Qty_Day_11 : .Row = 0 : .Text = dayAr1(11)
        .Col = C_Qty_Day_12 : .Row = 0 : .Text = dayAr1(12)
        .Col = C_Qty_Day_13 : .Row = 0 : .Text = dayAr1(13)
        .Col = C_Qty_Day_14 : .Row = 0 : .Text = dayAr1(14)
        .Col = C_Qty_Day_15 : .Row = 0 : .Text = dayAr1(15)
        .Col = C_Qty_Day_16 : .Row = 0 : .Text = dayAr1(16)
        .Col = C_Qty_Day_17 : .Row = 0 : .Text = dayAr1(17)
        .Col = C_Qty_Day_18 : .Row = 0 : .Text = dayAr1(18)
        .Col = C_Qty_Day_19 : .Row = 0 : .Text = dayAr1(19)
        .Col = C_Qty_Day_20 : .Row = 0 : .Text = dayAr1(20)
        .Col = C_Qty_Day_21 : .Row = 0 : .Text = dayAr1(21)
        .Col = C_Qty_Day_22 : .Row = 0 : .Text = dayAr1(22)
        .Col = C_Qty_Day_23 : .Row = 0 : .Text = dayAr1(23)
        .Col = C_Qty_Day_24 : .Row = 0 : .Text = dayAr1(24)
        .Col = C_Qty_Day_25 : .Row = 0 : .Text = dayAr1(25)
        .Col = C_Qty_Day_26 : .Row = 0 : .Text = dayAr1(26)
        .Col = C_Qty_Day_27 : .Row = 0 : .Text = dayAr1(27)
        .Col = C_Qty_Day_28 : .Row = 0 : .Text = dayAr1(28)
        .Col = C_Qty_Day_29 : .Row = 0 : .Text = dayAr1(29)
        .Col = C_Qty_Day_30 : .Row = 0 : .Text = dayAr1(30)
        
        .Col = C_Qty_Day_0_Hidden  : .Row = 0 : .Text = dayAr2(0)
        .Col = C_Qty_Day_1_Hidden  : .Row = 0 : .Text = dayAr2(1)
        .Col = C_Qty_Day_2_Hidden  : .Row = 0 : .Text = dayAr2(2)
        .Col = C_Qty_Day_3_Hidden  : .Row = 0 : .Text = dayAr2(3)
        .Col = C_Qty_Day_4_Hidden  : .Row = 0 : .Text = dayAr2(4)
        .Col = C_Qty_Day_5_Hidden  : .Row = 0 : .Text = dayAr2(5)
        .Col = C_Qty_Day_6_Hidden  : .Row = 0 : .Text = dayAr2(6)
        .Col = C_Qty_Day_7_Hidden  : .Row = 0 : .Text = dayAr2(7)
        .Col = C_Qty_Day_8_Hidden  : .Row = 0 : .Text = dayAr2(8)
        .Col = C_Qty_Day_9_Hidden  : .Row = 0 : .Text = dayAr2(9)
        .Col = C_Qty_Day_10_Hidden : .Row = 0 : .Text = dayAr2(10)
        .Col = C_Qty_Day_11_Hidden : .Row = 0 : .Text = dayAr2(11)
        .Col = C_Qty_Day_12_Hidden : .Row = 0 : .Text = dayAr2(12)
        .Col = C_Qty_Day_13_Hidden : .Row = 0 : .Text = dayAr2(13)
        .Col = C_Qty_Day_14_Hidden : .Row = 0 : .Text = dayAr2(14)
        .Col = C_Qty_Day_15_Hidden : .Row = 0 : .Text = dayAr2(15)
        .Col = C_Qty_Day_16_Hidden : .Row = 0 : .Text = dayAr2(16)
        .Col = C_Qty_Day_17_Hidden : .Row = 0 : .Text = dayAr2(17)
        .Col = C_Qty_Day_18_Hidden : .Row = 0 : .Text = dayAr2(18)
        .Col = C_Qty_Day_19_Hidden : .Row = 0 : .Text = dayAr2(19)
        .Col = C_Qty_Day_20_Hidden : .Row = 0 : .Text = dayAr2(20)
        .Col = C_Qty_Day_21_Hidden : .Row = 0 : .Text = dayAr2(21)
        .Col = C_Qty_Day_22_Hidden : .Row = 0 : .Text = dayAr2(22)
        .Col = C_Qty_Day_23_Hidden : .Row = 0 : .Text = dayAr2(23)
        .Col = C_Qty_Day_24_Hidden : .Row = 0 : .Text = dayAr2(24)
        .Col = C_Qty_Day_25_Hidden : .Row = 0 : .Text = dayAr2(25)
        .Col = C_Qty_Day_26_Hidden : .Row = 0 : .Text = dayAr2(26)
        .Col = C_Qty_Day_27_Hidden : .Row = 0 : .Text = dayAr2(27)
        .Col = C_Qty_Day_28_Hidden : .Row = 0 : .Text = dayAr2(28)
        .Col = C_Qty_Day_29_Hidden : .Row = 0 : .Text = dayAr2(29)
        .Col = C_Qty_Day_30_Hidden : .Row = 0 : .Text = dayAr2(30)
        .Col = C_Qty_Month_1: .Row = 0 : .Text = "월합계"
        .Col = C_Qty_Month_2: .Row = 0 : .Text = startDate1
        .Col = C_Qty_Month_3: .Row = 0 : .Text = startDate2
    End With
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim rowcnt
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    

    rowcnt = frm1.vspdData.MaxRows
	If rowcnt = 0 Then
		Msgbox "품목을 등록하십시오.",vbInformation, parent.gLogoName
		Exit Function
	End If
    
    
    '20080218::hanc
    If Not chkFieldByCell(frm1.txtPoTypeCd, "A",1)	then
       Exit Function
    End If
    If Not chkFieldByCell(frm1.txtDept_cd, "A",1)	then
       Exit Function
    End If

	If Trim(frm1.txtEmp_no.value) = "" Then
		Msgbox "의뢰담당자를 확인하십시오.",vbInformation, parent.gLogoName
		Exit Function
	End If

'    If Not chkFieldByCell(frm1.txtEmp_no, "A",1)	then
'       Exit Function
'    End If

'    ggoSpread.Source = frm1.vspdData
'    If ggoSpread.SSCheckChange = False Then
'        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
'        Exit Function
'    End If

	'----------- 불출수량 > 0 검사. ----------
	Dim r
	With Frm1
		For r = 1 To .vspdData.MaxRows	
           .vspdData.Row = r
           .vspdData.Col = 0
			Select Case .vspdData.Text
				Case ggoSpread.DeleteFlag  
					.vspdData.Col = C_ISSUE_QTY    
					If Trim(.vspdData.Text) > 0 Then
						Call DisplayMsgBox("ZZ0008", "X", "X", "X")
						Exit Function
					End If
			End Select  
		Next
	End With


	

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

'    If Not chkField(Document, "1") Then                                          '☜: This function check required field
'       Exit Function
'    End If

    Call DisableToolBar(parent.TBC_SAVE)

    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    FncSave = True                                                              '☜: Processing is OK

End Function

'2008305::hanc
Function DbSaveDel()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim strcurr_dt
	

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSaveDel = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                                 '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                         '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	strcurr_dt = replace(LocSvrDate, "-","")	
'msgbox " 여긴 300 "
	With frm1
    
       For lRow = 1 To 1
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
'msgbox " 여긴 insert "

                                                     strVal = strVal & "D"                       & Parent.gColSep
                                                     strVal = strVal & lRow                      & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
       Next

'msgbox " 여긴 200 "
			.txtMaxRows.value     = lGrpCnt-1	
			.txtSpread.value      = strDel & strVal
			.txtcurr_dt.value     = strcurr_dt
	
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID1)

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSaveDel = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim strcurr_dt
	

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                                 '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                         '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	strcurr_dt = replace(LocSvrDate, "-","")	
'msgbox " 여긴 300 "
	With frm1
    
       For lRow = 1 To 1
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
'msgbox " 여긴 insert "

                                                     strVal = strVal & "C"                       & Parent.gColSep
                                                     strVal = strVal & lRow                      & Parent.gColSep
                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '2
                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep     '3
                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '4
                    'msgbox Trim(.vspdData.Text)
                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7

                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
       Next

'msgbox " 여긴 200 "
			.txtMaxRows.value     = lGrpCnt-1	
			.txtSpread.value      = strDel & strVal
			.txtcurr_dt.value     = strcurr_dt
	
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID1)
'msgbox " 여긴 400 "
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
'	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
'	
'    ggoSpread.Source = frm1.vspdData
'	Set gActiveElement = document.ActiveElement       
'	
'
'    strVal = ""
'    strDel = ""
'    lGrpCnt = 1
'
'	With Frm1
'    
'       For lRow = 1 To .vspdData.MaxRows
'    
'           .vspdData.Row = lRow
'           .vspdData.Col = 0
'        
'           Select Case .vspdData.Text
' 
'               Case ggoSpread.InsertFlag                                      '☜: Update
''msgbox " 여긴aa insert "
'
'                                                     strVal = strVal & "C"                       & Parent.gColSep
'                                                     strVal = strVal & lRow                      & Parent.gColSep
'                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '2
'                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep     '3
'                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '4
'                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
'                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
'                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7
'
'                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
'                    lGrpCnt = lGrpCnt + 1
'               Case ggoSpread.UpdateFlag                                      '☜: Update
''msgbox " 여긴aa update "
'                                                     strVal = strVal & "U"                       & Parent.gColSep
'                                                     strVal = strVal & lRow                      & Parent.gColSep
'                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
'                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
'                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
'                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
'                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7
'
'                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
'                    lGrpCnt = lGrpCnt + 1
'               Case ggoSpread.DeleteFlag                                      '☜: Delete
''msgbox " 여긴aa delete "
'                                                     'strDel = strDel & "D"                       & Parent.gColSep	'20080220
'                                                     'strDel = strDel & lRow                      & Parent.gColSep	'20080220
'                                                     strVal = strVal & "D"                       & Parent.gColSep
'                                                     strVal = strVal & lRow                      & Parent.gColSep                                                     
'                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
'                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
'                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
'                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
'                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7
'
'                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
'                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
'                    lGrpCnt = lGrpCnt + 1
'           End Select
'       Next
'
''msgbox " 여긴 200 "
'	   .txtMaxRows.value     = lGrpCnt-1	
'	   '.txtSpread.value      = strDel & strVal	'20080220
'		.txtSpread.value      = strVal
'	End With
'	
''msgbox " 여긴 300 "
'	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID2)
''msgbox " 여긴 400 "
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

Function DbSave_GRD2()
'MSGBOX "DBSave_DGD2"
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim strcurr_dt
	

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave_GRD2 = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                                 '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                         '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData
	Set gActiveElement = document.ActiveElement       
	

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
'msgbox " 여긴aa insert "

                                                     strVal = strVal & "C"                       & Parent.gColSep
                                                     strVal = strVal & lRow                      & Parent.gColSep
                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '2
                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep     '3
                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '4
                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7

                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
'msgbox " 여긴aa update "
                                                     strVal = strVal & "U"                       & Parent.gColSep
                                                     strVal = strVal & lRow                      & Parent.gColSep
                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7

                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
'msgbox " 여긴aa delete "
                                                     'strDel = strDel & "D"                       & Parent.gColSep	'20080220
                                                     'strDel = strDel & lRow                      & Parent.gColSep	'20080220
                                                     strVal = strVal & "D"                       & Parent.gColSep
                                                     strVal = strVal & lRow                      & Parent.gColSep                                                     
                    .vspdData.Col = C_ITEM_SEQ     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
                    .vspdData.Col = C_PRODT_ORDER_NO : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep
                    .vspdData.Col = C_ITEM_CD      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_REQ_QTY      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '5
                    .vspdData.Col = C_ISSUE_QTY    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '6
                    .vspdData.Col = C_REMARK1      : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep       '7

                    .vspdData.Col = C_EnrollDT     : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_GraduatedDT  : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoney       : strVal = strVal & Trim(.vspdData.Text)      & Parent.gColSep   
                    .vspdData.Col = C_SMoneyCnt    : strVal = strVal & Trim(.vspdData.Text)      & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

'msgbox " 여긴 200 "
	   .txtMaxRows.value     = lGrpCnt-1	
	   '.txtSpread.value      = strDel & strVal	'20080220
		.txtSpread.value      = strVal
	End With
	
'msgbox " 여긴 300 "
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID2)
'msgbox " 여긴 400 "
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave_GRD2 = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

Function DbSaveOk111()
    CALL DbSave_GRD2()
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables                                                          '⊙: Initializes local global variables
    Call MainQuery()
End Function

Function DbSaveOk1()
'    ggoSpread.Source = frm1.vspdData
'    ggoSpread.ClearSpreadData
'    Call InitVariables                                                          '⊙: Initializes local global variables
'    Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,parent.gColSep)

    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))

       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           frm1.vspdData.Col = iDx
           frm1.vspdData.Row = iRow
           If frm1.vspdData.ColHidden <> True And frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              frm1.vspdData.Col = iDx
              frm1.vspdData.Row = iRow
              frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
       Next
    End If
End Sub

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD,imRow

    On Error Resume Next
    FncInsertRow = False

    if IsNumeric(Trim(pvRowCnt)) Then
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
        ggoSpread.InsertRow,imRow
        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
	    ggoSpread.spreadUnlock C_ITEM_CD,.vspdData.ActiveRow,C_ITEM_CD,.vspdData.ActiveRow + imRow - 1
	    ggoSpread.sssetrequired C_ITEM_CD,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1        
	    ggoSpread.sssetrequired C_REQ_QTY,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1        '20080305::HANC
		ggoSpread.spreadUnlock C_POPUP2,.vspdData.ActiveRow, C_POPUP2,.vspdData.ActiveRow + imRow - 1        
		
       .vspdData.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    If Err.number =0 Then
        FncInsertRow = True
    End if
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    If frm1.vspdData.MaxRows < 1 then
       Exit function
    End if
    With frm1.vspdData
        .focus
        ggoSpread.Source = frm1.vspdData
        lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

    Call InitData()
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    With frm1.VspdData
         .ReDraw = False
         If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow

            .ReDraw = True
            .Focus
         End If
    End With

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        'If Row < 1 Then Exit Sub
        
 		If Row > 0 Then
	        .Col = Col
	        .Row = Row
	        
			Select Case Col 
				Case C_POPUP2
					Call popUpItem()
			End Select
	        
	    End If


    End With
End Sub

Sub txtItemCd_Onchange()

	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
    End If


		

    frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click1(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData


    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
    End If


		
		frm1.vspdData.Row = Row
		  		
		With frm1
		    '불출의뢰번호
			.vspdData.Col = C_ISSUE_REQ_NO  
			.vspdData.Row = .vspdData.ActiveRow 	
			.txtPoNo1.value = .vspdData.Text
		    '출고구분
			.vspdData.Col = C_TRNS_TYPE
			if .vspdData.Text = "OI" then
    		    .rdoReleaseflg2.checked = True
			else
    		    .rdoReleaseflg1.checked = True
			end if
		    
		    '불출의뢰일

            '불출의뢰유형
			.vspdData.Col = C_ISSUE_TYPE
			.txtPoTypeCd.value = .vspdData.Text
			 
			 '의뢰부서
			.vspdData.Col = C_DEPT_CD
			.txtDept_cd.value = .vspdData.Text
			 
			 '의뢰담당자
			.vspdData.Col = C_EMP_NO
			.txtEmp_no.value = .vspdData.Text
	
	        '특이사항
			.vspdData.Col = C_REMARK
			.txtRemark.value = .vspdData.Text

	        '불출의뢰일
			.vspdData.Col  = C_REQ_DT
			.txtPoDt.Text = UniDateClientFormat(.vspdData.Text) 
            
	        'LOC
			.vspdData.Col = C_LOC
			.txtLoc.value = .vspdData.Text
			
			
		

		End With   



    frm1.vspdData.Row = Row
End Sub

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave()
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore()
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
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
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()

    Call SetQuerySpreadColor()
    Call SetSpreadColor(0, 0)
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

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

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos
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
' Function Name : vspdData_MouseDown
' Function Desc : popup grid
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
    Dim IntRetCD

    FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")            '⊙: Data is changed.  Do you want to exit?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function

'=======================================================================================================
'   Event Name : txtYYYYMM_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
        SetFocusToDocument("M")
        frm1.txtYYYYMM.Focus

'        if  CommonQueryRs(" DATEADD(ww, DATEDIFF(ww, 0, '"&frm1.txtYYYYMM.Text&"'), 0) ","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
'            frm1.txtYYYYMM.Text = Replace(lgF0, Chr(11), "")
'        End If
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYYYYMM_onchange(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Function txtYYYYMM_OnBlur()
'    if  CommonQueryRs(" DATEADD(ww, DATEDIFF(ww, 0, '"&frm1.txtYYYYMM.Text&"'), 0) ","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
'        frm1.txtYYYYMM.Text = Replace(lgF0, Chr(11), "")
'    End If
End Function

'==========================================================================================
'   Event Name : txtYYYYMM_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtYYYYMM_KeyDown(KeyCode, Shift)
    If KeyCode = 13 Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : CookiePage()
'   Event Desc :
'==========================================================================================
Function CookiePage()
End Function

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim iCalledAspName, IntRetCD

'	If frm1.txtPlantCd.value= "" Then
'		Call DisplayMsgBox("971012","X", "공장","X")
'		frm1.txtPlantCd.focus 
'		Set gActiveElement = document.activeElement
'		IsOpenPop = False 
'		Exit Function
'	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	If UCase(Parent.gPlant) <> "" Then
		arrParam(0) = UCase(Parent.gPlant)
		'arrParam(0) = Trim(frm1.txtPlantCd.value)
	Else
		arrParam(0) = Trim(frm1.txtPlantCd.value)
	End If
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
		
End Function


'------------------------------------------  popUpItem()  -------------------------------------------------
Function popUpItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	if  Trim(Trim(frm1.txtPlantCd.Value)) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	IsOpenPop = True

	' -- 그리드에 있는 값을 참조하기에 추가하였음	
	frm1.vspdData.Col = C_ITEM_CD
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.vspdData.text)		' Item Code
	arrParam(2) = "36!PP"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec	
	arrField(3) = 4	' -- 단위
    
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ITEM_CD, frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		frm1.vspdData.Col	= C_ITEM_CD
		frm1.vspdData.Text	= arrRet(0)
		frm1.vspdData.Col	= C_ITEM_NM
		frm1.vspdData.Text 	= arrRet(1)
		frm1.vspdData.Col	= C_BASIC_UNIT
		frm1.vspdData.Text 	= arrRet(3)
		'Call ChangeItemPlant(frm1.vspdData.ActiveRow)
		Call SetActiveCell(frm1.vspdData, C_REQ_QTY, frm1.vspdData.ActiveRow,"M","X","X")
	End If	
End Function


'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)

		.txtDept_nm.value = arrRet(2)

        Call CommonQueryRs(" DEPT_CD "," HAA010T "," EMP_NO =  " & FilterVar(arrRet(0), "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        .txtDept_cd.value = Replace(lgF0, Chr(11), "")

		ggoSpread.Source = Frm1.vspdData    
		'ggoSpread.ClearSpreadData     
		
		.txtEmp_no.focus

		lgBlnFlgChgValue = False
	End With
End Sub
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			<%' 조건부에서 누른 경우 Code Condition%>
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			<%' Grid에서 누른 경우 Code Condition%>
	End If
	arrParam(1) = ""								<%' Name Cindition%>
    arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtDept_cd.focus
		Else 'spread
			frm1.vspdData.Col = C_Dept
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtDept_cd.value = arrRet(0)
			.txtDept_Nm.value = arrRet(1)
			.txtDept_cd.focus
		Else 'spread
			.vspdData.Col = C_DeptNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_Dept
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function


'20080211::hanc
Function OpenIssueReq()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	
	IsOpenPop = True


    arrParam(0) = "불출의뢰번호 팝업"                   ' 팝업 명칭
    arrParam(1) = "M_ISSUE_REQ_HDR_KO441 A, HAA010T B"               ' TABLE 명칭
    arrParam(2) = Trim(frm1.txtPoNo.Value)
    arrParam(3) = ""                                    ' Name Cindition
    arrParam(4) = "A.EMP_NO = B.EMP_NO"
    arrParam(5) = "불출의뢰번호"

    arrField(0) = "A.ISSUE_REQ_NO"                    ' Field명(0)
    arrField(1) = "DD" & parent.gColSep & "A.REQ_DT"   ' "convert(char(10), REQ_DT, 120)"                     	    ' Field명(1)
'    arrField(1) = "REQ_DT"                     	    ' Field명(1)
    arrField(2) = "ED13" & Parent.gColSep &"DBO.UFN_GETDEPTNAME(DBO.UFN_H_GET_DEPT_CD(A.EMP_NO,GETDATE()),GETDATE())"                     	' Field명(2)
    arrField(3) = "ED12" & Parent.gColSep &"B.NAME"                    ' Field명(0)
    arrField(4) = "ED10" & Parent.gColSep &"(CASE WHEN A.TRNS_TYPE = 'ST' THEN '재고이동' ELSE '출고' END) TRNS_TYPE"                    
    arrField(5) = "ED17" & Parent.gColSep &"DBO.ufn_GetTransTypeNM(A.TRNS_TYPE, A.ISSUE_TYPE, A.MOV_TYPE) MOV_TYPE"                    
    arrField(6) = "ED17" & Parent.gColSep &"ISNULL(A.LOC, '') LOC"                    



    arrHeader(0) = "불출의뢰번호"                     	' Header명(0)
    arrHeader(1) = "블츨의뢰일"                  	 	' Header명(1)
    arrHeader(2) = "불출의뢰부서명"                  	 	' Header명(2)
    arrHeader(3) = "불출의뢰담당자"                  	 	' Header명(2)
    arrHeader(4) = "출고구분"                  	 	' Header명(2)
    arrHeader(5) = "불출의뢰유형"                  	 	' Header명(2)
    arrHeader(6) = "LOCATION"                  	 	' Header명(2)
    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=820px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetIssueReq(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPoNo.focus
	
End Function

Function SetIssueReq(byval arrRet)

	frm1.txtPoNo.value          = arrRet(0)
	frm1.txtDept_cd.Value       = arrRet(2)
	
End Function

'20080218::hanc
Function OpenPartRef()
'msgbox "OpenPartRef"
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

'	If lgIntFlgMode = parent.OPMD_CMODE Then
''		If lgBlnFlgChgValue = False Then
'			Call DisplayMsgBox("900002", "x", "x", "x")
'			Exit Function
''		End If
'	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4311RA1_KO441")     '20080218::hanc
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	arrParam(1) = ""  'Trim(frm1.txtProdOrderNo1.value)	'☜: 조회 조건 데이타 

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	If arrRet(0,0) = "" Then
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetPoRef2(arrRet)
	End If	

'	Call SetPoRef2(arrRet)

End Function

'20080213::hanc
Function SetPoRef2(strRet)
'msgbox "SetPoRef2"
	Dim Index1, Count1, Row1
	Dim temp
			
	Const C_ItemCd_Ref		= 0
	Const C_ItemNm_Ref		= 1
	Const C_Req_Qty_Ref		= 3
	Const C_Unit_Ref		= 4
	Const C_Issue_Qty_Ref   = 7
	Const C_ITEM_SEQ_Ref    = 14
	Const C_PRODT_ORDER_NO_Ref = 2
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			
            
			Call .SetText(C_ITEM_CD,		Row1, strRet(index1,C_ItemCd_Ref))
			Call .SetText(C_ITEM_NM,		Row1, strRet(index1,C_ItemNm_Ref))
			Call .SetText(C_BASIC_UNIT,		Row1, strRet(index1,C_Unit_Ref))
			Call .SetText(C_REQ_QTY,		Row1, UNICDbl(strRet(index1,C_Req_Qty_Ref)))
'			Call .SetText(C_ISSUE_QTY,		Row1, UNICDbl(strRet(index1,C_Issue_Qty_Ref)))
			Call .SetText(C_ITEM_SEQ,		Row1, strRet(index1,C_ITEM_SEQ_Ref))
			Call .SetText(C_PRODT_ORDER_NO,		Row1, strRet(index1,C_PRODT_ORDER_NO_Ref))

		Next
	
		
		.ReDraw = True
		
	End with

End Function


Function OpenReqRef()

'msgbox "OpenReqRef"
	Dim strRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD	
	
	if frm1.txtRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = ""
	arrParam(1) = ""
'	arrParam(2) = Trim(frm1.txtGroupCd.value)
'	arrParam(3) = Trim(frm1.txtGroupNm.value)
	arrParam(4) = "P"
	arrParam(5) = "Y"
	arrParam(6) = Trim(frm1.txtSupplierCd.Value)

	iCalledAspName = AskPRAspName("M2111RA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If
		
End Function

Function SetReqRef(strRet)

Dim Index1,index2,Index3,Count1,Count2
Dim IntIflg
Dim strMessage
Dim intstartRow,intEndRow

Const C_ReqNo_Ref		= 0
Const C_PlantCd_Ref		= 1
Const C_PlantNm_Ref		= 2
Const C_ItemCd_Ref		= 3
Const C_ItemNm_Ref		= 4
Const C_SpplSpec_Ref    = 5                         '품목 규격 추가 
Const C_Qty_Ref			= 6
Const C_Unit_Ref		= 7
Const C_DlvyDt_Ref		= 8
Const C_Pr_Type_Ref		= 9 
Const C_Pr_Type_Nm_Ref	= 10
Const C_SoNo_Ref		= 11
Const C_SoSeqNo_Ref		= 12
Const C_TrackingNo_Ref	= 13
Const C_SLCd_Ref		= 14
Const C_SLNm_Ref		= 15 
Const C_HSCd_Ref		= 16
Const C_HSNm_Ref		= 17


	Count1 = Ubound(strRet,1)
	Count2 = UBound(strRet,2)
	strMessage = ""
	IntIflg=true
	
	with frm1
	
	intStartRow = .vspdData.MaxRows + 1
	
	for index1 = 0 to Count1
	
		.vspdData.Row=Index1+1
	
		for Index3=0 to .vspdData.MaxRows
			.vspdData.Row = index3+1
			.vspdData.Col=C_PrNo
			if .vspdData.Text = strRet(index1,C_ReqNo_Ref) then
				strMessage = strMessage & strRet(Index1,C_ReqNo_Ref) & ";"
				intIflg=False
				Exit for
			End if
		Next
		
		if IntIflg <> False then
		    ggoSpread.Source = .vspdData
	         .vspdData.ReDraw = False
	         ggoSpread.InsertRow
		    Call SetSpreadColor(.vspdData.ActiveRow,.vspdData.ActiveRow)
			.vspdData.Row=.vspdData.ActiveRow 

			'Call SetState("C",.vspdData.ActiveRow)
			
			for index2 = 0 to Count2 - 1 
		
				Select Case Index2
				Case C_ItemCd_Ref
					.vspdData.Col=C_itemCd
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_ItemCd,.vspdData.ActiveRow,C_ItemCd,.vspdData.ActiveRow
					ggoSpread.spreadlock C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow
				Case C_ItemNm_Ref
					.vspdData.Col=C_itemNm
					.vspdData.Text=strRet(index1,index2)
					
				Case C_SpplSpec_Ref                              '품목규격 추가 
				    .vspdData.Col=C_SpplSpec
					.vspdData.Text=strRet(index1,index2)			
					
				Case C_Qty_Ref
					.vspdData.Col=C_OrderQty
					.vspdData.Text=strRet(index1,index2)
				Case C_Unit_Ref
					.vspdData.Col=C_OrderUnit
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_OrderUnit,.vspdData.ActiveRow,C_Popup3,.vspdData.ActiveRow
					'ggoSpread.spreadlock C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow
				Case C_DlvyDt_Ref
					.vspdData.Col=C_DlvyDT
					.vspdData.Text=strRet(index1,index2)
					ggoSpread.spreadlock C_DlvyDT, .vspdData.ActiveRow, C_DlvyDT ,.vspdData.ActiveRow
				Case C_SLCd_Ref
					.vspdData.Col=C_SLCd
					.vspdData.Text=strRet(index1,index2)
				Case C_SLNm_Ref
					.vspdData.Col=C_SLNm
					.vspdData.Text=strRet(index1,index2)
				Case C_TrackingNo_Ref
					.vspdData.Col=C_TrackingNo
					.vspdData.Text=strRet(index1,index2)
				     ggoSpread.spreadlock C_TrackingNo, .vspdData.ActiveRow, C_TrackingNoPop ,.vspdData.ActiveRow
				Case C_ReqNo_Ref
					.vspdData.Col=C_PrNo
					.vspdData.Text=strRet(index1,index2)	
                Case C_SoNo_Ref
					.vspdData.Col=C_So_No
					.vspdData.Text=strRet(index1,index2)					
			    Case C_SoSeqNo_Ref
					.vspdData.Col=C_So_Seq_No
					.vspdData.Text=strRet(index1,index2)		
				End Select
				
			next
				
		Else
			IntIFlg=True
		End if 
	next
	
	intEndRow = .vspdData.ActiveRow
	
	if strMessage<>"" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if
	
	'.vspdData.Col 	= C_Stateflg
	'.vspdData.Text = "C"
	
	.vspdData.ReDraw = True
	
	End with

			
End Function


'------------------------------------------  SetTrackingNo()  --------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetTrackingNo(Byval arrRet)

    With frm1
			.txtTrackingNo.Value = arrRet(0)
			.txtTrackingNo.focus
			Set gActiveElement = document.activeElement
	End With
	
End Function

'20080219::hanc
Function OpenPoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPoTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	
	IsOpenPop = True


    arrParam(0) = "불출의뢰유형 팝업"                 ' 팝업 명칭
    arrParam(1) = "I_MOVETYPE_CONFIGURATION C, B_MINOR D"                ' TABLE 명칭
    arrParam(2) = Trim(frm1.txtPoTypeCd.Value)
    arrParam(3) = ""                                ' Name Cindition
	If frm1.rdoReleaseflg(0).checked Then
        arrParam(4) = "C.MOV_TYPE = D.MINOR_CD AND D.MAJOR_CD = 'I0001' AND D.MINOR_TYPE = 'U' AND C.TRNS_TYPE = 'ST' "
	Else
        arrParam(4) = "C.MOV_TYPE = D.MINOR_CD AND D.MAJOR_CD = 'I0001' AND D.MINOR_TYPE = 'U' AND C.TRNS_TYPE = 'OI' "
	End If
    arrParam(5) = "불출의뢰유형"
    
    arrField(0) = "C.MOV_TYPE"                         ' Field명(0)
    arrField(1) = "D.MINOR_NM"                     	' Field명(1)

    arrHeader(0) = "불출의뢰유형"                     	' Header명(0)
    arrHeader(1) = "불출의뢰유형명"                  	 	' Header명(1)


    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPoType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPoTypeCd.focus
	
End Function

'20080219::hanc
Function SetPoType(byval arrRet)
	frm1.txtPoTypeCd.Value      = arrRet(0)
	frm1.txtPoTypeCdNm.Value    = arrRet(1)
End Function
'==========================================   Release()  ======================================
'	Name : Release()
'===================================================================================================
Sub Release()

    Err.Clear
    
    If CheckRunningBizProcess = True Then	
		Exit Sub
	End If                
    
    Dim strVal
    
    strVal = BIZ_PGM_SAVE_ID1 & "?txtMode=" & Trim(frm1.hdnMode.Value)	
    strVal = strVal & "&txtReleaseFlag=" & Trim(releaseFlag)
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo1.Value)
    strVal = strVal & "&txtUpdtUserId=" & Parent.gUsrID   
    
    If LayerShowHide(1) = False Then Exit Sub
	Call RunMyBizASP(MyBizASP, strVal)								
	
End Sub
'==========================================   btnCfm()  ======================================
Sub Cfm()
    Dim IntRetCD , i
    
    Err.Clear                                                       
    
    if ggoSpread.SSCheckChange = True then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if

	frm1.vspdData.Col = C_CONFIG_FLAG
	frm1.vspdData.Row = 1 

	if Trim(frm1.vspdData.Text) = "Y" then
        With frm1.vspdData
            For i = 1 To frm1.vspdData.MaxRows
    			.row = i
    			.Col = C_ISSUE_QTY
    			
    			If .text > 0 Then
            		Call DisplayMsgBox("ZZ0007", "X", "X", "X")
        			Exit Sub
    			End If 
    			      
            Next
            
        End With
	End if
        


	frm1.vspdData.Col = C_CONFIG_FLAG
	frm1.vspdData.Row = 1 
	
	if Trim(frm1.vspdData.Text) = "N" then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		frm1.hdnMode.Value = "Release"
		releaseFlag        = "Y"
					                                                
	elseif Trim(frm1.vspdData.Text) = "Y" then
			
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		
		frm1.hdnMode.Value = "UnRelease"
		releaseFlag        = "N"
		
	End if
	Call Release()
	
End Sub

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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
                            </TR>
                        </TABLE>
                    </TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef">제조오더예약부품참조</A></TD>
					<TD WIDTH=10>&nbsp;</TD>					
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR  HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
                </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE  <%=LR_SPACE_TYPE_40%>>
								<TR>
                                    <TD CLASS="TD5" NOWRAP>공장</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="12XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=29 MAXLENGTH=40 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>불출의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=29 MAXLENGTH=18 ALT="불출의뢰번호" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIssueReq()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
                <TR STYLE="display:none">
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                        <TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
                                <TD CLASS="TD5" NOWRAP>생산계획월</TD>
                                <TD CLASS="TD6" NOWRAP>
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtYYYYMM CLASSID=<%=gCLSIDFPDT%> ALT="기준일자" tag="11X1" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>
                                </TD>
                                <TD CLASS="TD5" NOWRAP>품목</TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>
                            </TR>
                            <TR>
                                <TD CLASS="TD5" NOWRAP></TD>
                                <TD CLASS="TD6" NOWRAP></TD>
                                <TD CLASS="TD5" NOWRAP></TD>
                                <TD CLASS="TD6" NOWRAP></TD>
                                <!--<TD CLASS="TD5" NOWRAP>Tracking No.</TD>      
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>-->
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
                        <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
							    <TD CLASS="TD5" NOWRAP>불출의뢰번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="불출의뢰번호" NAME="txtPoNo1" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
							    <TD CLASS=TD5 NOWRAP>출고구분</TD>
							    <TD CLASS=TD6 NOWRAP>
								<input type=radio CLASS="RADIO" name="rdoReleaseflg" id="rdoReleaseflg1" tag = "11" Value="ST"   checked><label for="rdoConfirmFlg_Yes">재고이동</label>&nbsp;&nbsp;
								<input type=radio CLASS = "RADIO" name="rdoReleaseflg" id="rdoReleaseflg2"   tag = "11" Value="OI" ><label for="rdoConfirmFlg_No">출고</label>
							</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>불출의뢰일</TD>
								<TD CLASS="TD656" NOWRAP>
										<script language =javascript src='./js/m9311ma1_txtPoDt.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>불출의뢰유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="불출의뢰유형" NAME="txtPoTypeCd" SIZE=10 MAXLENGTH=5 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="불출의뢰유형" NAME="txtPoTypeCdNm" SIZE=20 tag="24X"></TD>
							</TR>
							<TR>
			    	    		<TD CLASS="TD5" NOWRAP>의뢰담당자</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" MAXLENGTH=13 SiZE=10 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">  &nbsp;<INPUT NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14XXXU></TD>
								<TD CLASS=TD5 NOWRAP>의뢰부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=10 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(0)">&nbsp;<INPUT NAME="txtDept_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
								</TD>
							</TR>
							<TR>
				                <TD CLASS="TD5" NOWRAP>특이사항</TD>
				                <TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT ALT="비고" NAME="txtRemark" MAXLENGTH=200 SIZE=54 tag="21XXXU"></TD>
							    <TD CLASS="TD5" NOWRAP>Location</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="Location" NAME="txtLoc" MAXLENGTH=30 SIZE=13 tag="21XXXU"></TD>
							</TR>
							<TR STYLE="display:none">
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="23NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
							</TR>
							<TR STYLE="display:none">
				                <TD CLASS="TD5" NOWRAP>공급처담당</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처담당" NAME="txtSuppPrsn" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>긴급연락처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="긴급연락처" NAME="txtTel" MAXLENGTH=18 SIZE=35 tag="21XXXU"></TD>
							</TR>
                            <TR>
                                <TD HEIGHT="100%" colspan=4>
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
                                </TD>
                            </TR>
                            <TR>
                                <TD HEIGHT=5 WIDTH=100% colspan=4></TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
    </TR>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">확정</button></a></td>					
					<td WIDTH="*" align=right><!-- |<a href="VBSCRIPT:CookiePage(1)">재고이동입고</a> <a href="VBSCRIPT:CookiePage(2)">경비등록</a>--></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd"  tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUsrId" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd"   tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo"   tag="24">
<INPUT TYPE=HIDDEN NAME="txtTrackingNo"   tag="24">
<INPUT TYPE=HIDDEN NAME="txtRelease" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMode" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="txtcurr_dt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>