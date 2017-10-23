<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: s3314ma1_KO441
'*  4. Program Name			: 고객사별일별수주등록
'*  5. Program Desc			:
'*  6. Business ASP List	: +s3314ma1_KO441.asp		'☆: List MPS
							  +s3314mb1_KO441.asp		'☆: MPS(query, save)
'*  7. Modified date(First)	:
'*  8. Modified date(Last)	:
'*  9. Modifier (First)		: HAN cheol
'* 10. Modifier (Last)		: 
'* 11. Comment				: 고객사별일별수주등록
'* 12. History              : 
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID    = "p2250mb1_KO441.asp"
Const BIZ_PGM_ID2   = "p2250mb2_KO441.asp"						  '20080129::hanc         '☆: Biz Logic ASP Name
Const BIZ_PGM_ID3   = "p2250mb3_KO441.asp"						  '20080303::hanc         '☆: Biz Logic ASP Name

Dim arrColVal_header        '20080303::hanc
Dim LocSvrDate
Dim strDate
LocSvrDate = "<%=GetSvrDate%>"

strDate = UniConvDateAToB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat) 	'☆: 초기화면에 뿌려지는 마지막 날짜 

Dim IsOpenPop
Dim queryboolean
queryboolean = False

Dim C_bp_Cd
Dim C_bp_Nm

Dim C_Item_Cd
Dim C_Item_Nm
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
Dim C_Qty_Day_31        '20080303::hanc
Dim C_Qty_Day_32        '20080303::hanc
Dim C_Qty_Day_33        '20080303::hanc
Dim C_Qty_Day_34        '20080303::hanc
Dim C_Qty_Day_35        '20080303::hanc

Dim C_Qty_Day_36        '20080303::hanc
Dim C_Qty_Day_37        '20080303::hanc
Dim C_Qty_Day_38        '20080303::hanc
Dim C_Qty_Day_39        '20080303::hanc

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

Dim C_Qty_Day_31_Hidden
Dim C_Qty_Day_32_Hidden
Dim C_Qty_Day_33_Hidden
Dim C_Qty_Day_34_Hidden
Dim C_Qty_Day_35_Hidden
Dim C_Qty_Day_36_Hidden
Dim C_Qty_Day_37_Hidden
Dim C_Qty_Day_38_Hidden
Dim C_Qty_Day_39_Hidden

Dim dayAr, dayAr2
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status

    Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat, "1")
    Call ggoOper.LockField(Document, "N")                                           '⊙: Lock Field

    Call DbQueryPeriod      '20080303::hanc
    Call SetDefaultVal
'    Call InitSpreadSheet                                                            'Setup the Spread sheet 
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox
    Call InitSpreadComboBox
    Call CookiePage
'    Call SetToolbar("1100100100001111")                                             '버튼 툴바 제어
    Call SetToolbar("1100101100001111")                                             '버튼 툴바 제어
    
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

    queryboolean = False
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i, startDate
    Dim startDate1, startDate2
    Dim iPeriod
    i = 0
    ReDim dayAr(40)

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

    For i = 0 To 39
        dayAr(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
    Next

    Call InitSpreadPosVariables()

    iPeriod = Cint(frm1.txtPeriod.value)        '20080303::hanc
       
    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

        .ReDraw = false

        .MaxCols = C_Qty_Day_39_Hidden + 1                                      <%'☜: 최대 Columns의 항상 1개 증가시킴 %>

        .Col = .MaxCols                                                         <%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0

        ggoSpread.Source = Frm1.vspdData

        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos()

        ggoSpread.SSSetEdit     C_bp_Cd       , "CUSTOM"   , 9
        ggoSpread.SSSetEdit     C_bp_Nm       , "CUSTOM명" , 17

        ggoSpread.SSSetEdit     C_Item_Cd       , "품목"   , 9
        ggoSpread.SSSetEdit     C_Item_Nm       , "품목명" , 17
        ggoSpread.SSSetEdit     C_Tracking_No   , "Tracking No"   , 10
        ggoSpread.SSSetEdit     C_Type_cd	    , "구분"   , 10
        ggoSpread.SSSetEdit     C_Type		    , "구분"   , 13

        for i = 1 To iPeriod
'			'ggoSpread.SSSetFloat    i + 5     , dayAr(i-1) , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat    i + 5     , Cstr(i) + "일" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
'            ggoSpread.SSSetEdit     C_Type+Cstr(i)		    , Cstr(i) + "일"   , 10
        Next
        
       	ggoSpread.SSSetFloat    C_Qty_Month_1     , "월합계" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
  		ggoSpread.SSSetFloat    C_Qty_Month_2     , startDate1 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_Qty_Month_3     , startDate2 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit     C_Plant_Cd		  , "구분"   , 10
		
		for i = 1 To iPeriod
			ggoSpread.SSSetFloat    i + 49     , dayAr(i-1), 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        Next

        Call ggoSpread.SSSetColHidden(C_Plant_Cd         , C_Plant_Cd         , True)
        Call ggoSpread.SSSetColHidden(C_Type_Cd          , C_Type_Cd          , True)
        Call ggoSpread.SSSetColHidden(C_Type             , C_Type          , True)
        Call ggoSpread.SSSetColHidden(C_Tracking_No         , C_Tracking_No          , True)   '20080303::hanc
        Call ggoSpread.SSSetColHidden(C_Qty_Day_34          , C_Qty_Day_39          , True)   '20080624::hanc
        
        Call ggoSpread.SSSetColHidden(C_Qty_Month_1          , C_Qty_Month_3          , True)   '20080303::hanc
        Call ggoSpread.SSSetColHidden(C_Qty_Day_0_Hidden , C_Qty_Day_39_Hidden , True)


        .Col = C_Item_Cd        : .ColMerge = 2
        .Col = C_Item_Nm        : .ColMerge = 2
        .Col = C_Tracking_no    : .ColMerge = 2
       '.Col = C_Item_Stock_Qty : .ColMerge = 2

        ggoSpread.SSSetSplit2(5)

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
        
        ggoSpread.SpreadLock     C_Bp_Cd       , -1, C_bp_Cd
        ggoSpread.SpreadLock     C_bp_Nm       , -1, C_bp_Nm

        ggoSpread.SpreadLock     C_Item_Cd       , -1, C_Item_Cd
        ggoSpread.SpreadLock     C_Item_Nm       , -1, C_Item_Nm
        ggoSpread.SpreadLock     C_Tracking_No   , -1, C_Tracking_No
        ggoSpread.SpreadLock     C_Type_cd		 , -1, C_Type_cd
        ggoSpread.SpreadLock     C_Type          , -1, C_Type
		'ggoSpread.SpreadLock     C_Qty_Day_0     , -1, C_Qty_Day_30
        ggoSpread.SpreadLock     C_Qty_Month_1   , -1, C_Qty_Month_1
        ggoSpread.SpreadLock     C_Qty_Month_2   , -1, C_Qty_Month_2
        ggoSpread.SpreadLock     C_Qty_Month_3   , -1, C_Qty_Month_3
        ggoSpread.SpreadLock     C_Plant_Cd      , -1, C_Plant_Cd         
	    'ggoSpread.SpreadLock     C_Qty_Day_0_Hidden , -1, C_Qty_Day_30_Hidden 

        frm1.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()

    C_bp_Cd           = 1 
    C_bp_Nm           = 2

    C_Item_Cd           = 3 
    C_Item_Nm           = 4 
    C_Tracking_No       = 5 
    C_Type_cd           = 6 
    C_Type		        = 7 
    C_Qty_Day_0         = 8 
    C_Qty_Day_1         = 9 
    C_Qty_Day_2         = 10
    C_Qty_Day_3         = 11
    C_Qty_Day_4         = 12
    C_Qty_Day_5         = 13
    C_Qty_Day_6         = 14
    C_Qty_Day_7         = 15
    C_Qty_Day_8         = 16
    C_Qty_Day_9         = 17
    C_Qty_Day_10        = 18
    C_Qty_Day_11        = 19
    C_Qty_Day_12        = 20
    C_Qty_Day_13        = 21
    C_Qty_Day_14        = 22
    C_Qty_Day_15        = 23
    C_Qty_Day_16        = 24
    C_Qty_Day_17        = 25
    C_Qty_Day_18        = 26
    C_Qty_Day_19        = 27
    C_Qty_Day_20        = 28
    C_Qty_Day_21        = 29
    C_Qty_Day_22        = 30
    C_Qty_Day_23        = 31
    C_Qty_Day_24        = 32
    C_Qty_Day_25        = 33
    C_Qty_Day_26        = 34
    C_Qty_Day_27        = 35
    C_Qty_Day_28        = 36
    C_Qty_Day_29        = 37
    C_Qty_Day_30        = 38

    C_Qty_Day_31        = 39        '20080303::hanc
    C_Qty_Day_32        = 40
    C_Qty_Day_33        = 41
    C_Qty_Day_34        = 42
    C_Qty_Day_35        = 43

    C_Qty_Day_36        = 44
    C_Qty_Day_37        = 45
    C_Qty_Day_38        = 46
    C_Qty_Day_39        = 47

    C_Qty_Month_1       = 48
    C_Qty_Month_2       = 49
    C_Qty_Month_3       = 50
    C_Plant_Cd          = 51
    C_Qty_Day_0_Hidden  = 52
    C_Qty_Day_1_Hidden  = 53
    C_Qty_Day_2_Hidden  = 54
    C_Qty_Day_3_Hidden  = 55
    C_Qty_Day_4_Hidden  = 56
    C_Qty_Day_5_Hidden  = 57
    C_Qty_Day_6_Hidden  = 58
    C_Qty_Day_7_Hidden  = 59
    C_Qty_Day_8_Hidden  = 60
    C_Qty_Day_9_Hidden  = 61
    C_Qty_Day_10_Hidden = 62
    C_Qty_Day_11_Hidden = 63
    C_Qty_Day_12_Hidden = 64
    C_Qty_Day_13_Hidden = 65
    C_Qty_Day_14_Hidden = 66
    C_Qty_Day_15_Hidden = 67
    C_Qty_Day_16_Hidden = 68
    C_Qty_Day_17_Hidden = 69
    C_Qty_Day_18_Hidden = 70
    C_Qty_Day_19_Hidden = 71
    C_Qty_Day_20_Hidden = 72
    C_Qty_Day_21_Hidden = 73
    C_Qty_Day_22_Hidden = 74
    C_Qty_Day_23_Hidden = 75
    C_Qty_Day_24_Hidden = 76
    C_Qty_Day_25_Hidden = 77
    C_Qty_Day_26_Hidden = 78
    C_Qty_Day_27_Hidden = 79
    C_Qty_Day_28_Hidden = 80
    C_Qty_Day_29_Hidden = 81
    C_Qty_Day_30_Hidden = 82


    C_Qty_Day_31_Hidden = 83
    C_Qty_Day_32_Hidden = 84
    C_Qty_Day_33_Hidden = 85
    C_Qty_Day_34_Hidden = 86
    C_Qty_Day_35_Hidden = 87
    C_Qty_Day_36_Hidden = 88
    C_Qty_Day_37_Hidden = 89
    C_Qty_Day_38_Hidden = 90
    C_Qty_Day_39_Hidden = 91    
End SuB

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos

    ggoSpread.Source = frm1.vspdData

    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    C_bp_Cd           = iCurColumnPos(1)  
    C_bp_Nm           = iCurColumnPos(2)

    C_Item_Cd           = iCurColumnPos(3)  
    C_Item_Nm           = iCurColumnPos(4) 
    C_Tracking_no       = iCurColumnPos(5) 
    C_Type_cd		    = iCurColumnPos(6) 
    C_Type              = iCurColumnPos(7) 
    C_Qty_Day_0         = iCurColumnPos(8) 
    C_Qty_Day_1         = iCurColumnPos(9) 
    C_Qty_Day_2         = iCurColumnPos(10)
    C_Qty_Day_3         = iCurColumnPos(11)
    C_Qty_Day_4         = iCurColumnPos(12)
    C_Qty_Day_5         = iCurColumnPos(13)
    C_Qty_Day_6         = iCurColumnPos(14)
    C_Qty_Day_7         = iCurColumnPos(15)
    C_Qty_Day_8         = iCurColumnPos(16)
    C_Qty_Day_9         = iCurColumnPos(17)
    C_Qty_Day_10        = iCurColumnPos(18)
    C_Qty_Day_11        = iCurColumnPos(19)
    C_Qty_Day_12        = iCurColumnPos(20)
    C_Qty_Day_13        = iCurColumnPos(21)
    C_Qty_Day_14        = iCurColumnPos(22)
    C_Qty_Day_15        = iCurColumnPos(23)
    C_Qty_Day_16        = iCurColumnPos(24)
    C_Qty_Day_17        = iCurColumnPos(25)
    C_Qty_Day_18        = iCurColumnPos(26)
    C_Qty_Day_19        = iCurColumnPos(27)
    C_Qty_Day_20        = iCurColumnPos(28)
    C_Qty_Day_21        = iCurColumnPos(29)
    C_Qty_Day_22        = iCurColumnPos(30)
    C_Qty_Day_23        = iCurColumnPos(31)
    C_Qty_Day_24        = iCurColumnPos(32)
    C_Qty_Day_25        = iCurColumnPos(33)
    C_Qty_Day_26        = iCurColumnPos(34)
    C_Qty_Day_27        = iCurColumnPos(35)
    C_Qty_Day_28        = iCurColumnPos(36)
    C_Qty_Day_29        = iCurColumnPos(37)
    C_Qty_Day_30        = iCurColumnPos(38)
    C_Qty_Day_31        = iCurColumnPos(39)     '20080303::hanc
    C_Qty_Day_32        = iCurColumnPos(40)
    C_Qty_Day_33        = iCurColumnPos(41)
    C_Qty_Day_34        = iCurColumnPos(42)
    C_Qty_Day_35        = iCurColumnPos(43)
    C_Qty_Day_36        = iCurColumnPos(44)
    C_Qty_Day_37        = iCurColumnPos(45)
    C_Qty_Day_38        = iCurColumnPos(46)
    C_Qty_Day_39        = iCurColumnPos(47)
    C_Qty_Month_1       = iCurColumnPos(48)
    C_Qty_Month_2       = iCurColumnPos(49)
    C_Qty_Month_3       = iCurColumnPos(50)
    C_Plant_Cd          = iCurColumnPos(51)
    C_Qty_Day_0_Hidden  = iCurColumnPos(52)
    C_Qty_Day_1_Hidden  = iCurColumnPos(53)
    C_Qty_Day_2_Hidden  = iCurColumnPos(54)
    C_Qty_Day_3_Hidden  = iCurColumnPos(55)
    C_Qty_Day_4_Hidden  = iCurColumnPos(56)
    C_Qty_Day_5_Hidden  = iCurColumnPos(57)
    C_Qty_Day_6_Hidden  = iCurColumnPos(58)
    C_Qty_Day_7_Hidden  = iCurColumnPos(59)
    C_Qty_Day_8_Hidden  = iCurColumnPos(60)
    C_Qty_Day_9_Hidden  = iCurColumnPos(61)
    C_Qty_Day_10_Hidden = iCurColumnPos(62)
    C_Qty_Day_11_Hidden = iCurColumnPos(63)
    C_Qty_Day_12_Hidden = iCurColumnPos(64)
    C_Qty_Day_13_Hidden = iCurColumnPos(65)
    C_Qty_Day_14_Hidden = iCurColumnPos(66)
    C_Qty_Day_15_Hidden = iCurColumnPos(67)
    C_Qty_Day_16_Hidden = iCurColumnPos(68)
    C_Qty_Day_17_Hidden = iCurColumnPos(69)
    C_Qty_Day_18_Hidden = iCurColumnPos(70)
    C_Qty_Day_19_Hidden = iCurColumnPos(71)
    C_Qty_Day_20_Hidden = iCurColumnPos(72)
    C_Qty_Day_21_Hidden = iCurColumnPos(73)
    C_Qty_Day_22_Hidden = iCurColumnPos(74)
    C_Qty_Day_23_Hidden = iCurColumnPos(75)
    C_Qty_Day_24_Hidden = iCurColumnPos(76)
    C_Qty_Day_25_Hidden = iCurColumnPos(77)
    C_Qty_Day_26_Hidden = iCurColumnPos(78)
    C_Qty_Day_27_Hidden = iCurColumnPos(79)
    C_Qty_Day_28_Hidden = iCurColumnPos(80)
    C_Qty_Day_29_Hidden = iCurColumnPos(81)
    C_Qty_Day_30_Hidden = iCurColumnPos(82)
    C_Qty_Day_31_Hidden = iCurColumnPos(83)
    C_Qty_Day_32_Hidden = iCurColumnPos(84)
    C_Qty_Day_33_Hidden = iCurColumnPos(85)
    C_Qty_Day_34_Hidden = iCurColumnPos(86)
    C_Qty_Day_35_Hidden = iCurColumnPos(87)
    C_Qty_Day_36_Hidden = iCurColumnPos(88)
    C_Qty_Day_37_Hidden = iCurColumnPos(89)
    C_Qty_Day_38_Hidden = iCurColumnPos(90)
    C_Qty_Day_39_Hidden = iCurColumnPos(91)
    
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
    
    Dim i, ptfForMps, startDate, endDate
    
    i = 0
'
'    if  CommonQueryRs(" PTF_FOR_MPS ", " B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
'        ptfForMps = Replace(lgF0, Chr(11), "")
'    End If
'
'    if  CommonQueryRs(" DATEADD(ww, DATEDIFF(ww, 0, '"&LocSvrDate&"'), 0) ","","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
'        startDate = Replace(lgF0, Chr(11), "")
'    End If
'
'    endDate = UNIDateAdd("d", ptfForMps , startDate, parent.gDateFormat)
'
    With frm1.vspdData

        .Redraw = False

        ggoSpread.Source = frm1.vspdData

        For i = 1 To frm1.vspdData.MaxRows
			
			.row = i
			.Col = C_Type
			
			If .text = "생판요청" Then
			   ggoSpread.SpreadLock C_Qty_Day_0 ,  i, C_Qty_Day_39_Hidden , i 
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

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")                     '☜: Data is changed.  Do you want to display it?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    ggoSpread.ClearSpreadData
    If Not chkField(Document, "1") Then                                          '☜: This function check required field
       Exit Function
    End If

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
	lgKeyStream = lgKeyStream & Trim(frm1.hFileName.Value) & parent.gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.txtPeriod.Value) & parent.gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.txtPlantCd.Value) & parent.gColSep
	lgKeyStream = lgKeyStream & Trim(frm1.txtVersion.Value) & parent.gColSep

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

    With frm1
        strVal = BIZ_PGM_ID & "?txtMode="      & parent.UID_M0001
        strVal = strVal     & "&txtKeyStream=" & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="   & frm1.vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
        'strVal = strVal     & "&queryFlag="    & queryFlag                 '☜: Next key tag
    End With

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQuery = True
End Function

'20080303::hanc----------------------------------------------------------------------
'생산계획기간 가져오기---------------------------------------------------------------
Function DbQueryPeriod()
    DbQueryPeriod = False
    Err.Clear                                                                        '☜: Clear err status

    Dim strVal

    With frm1
        strVal = BIZ_PGM_ID3 & "?txtMode="      & parent.UID_M0004
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQueryPeriod = True
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    'queryFlag = "P"
'    queryFlag = "Q"
    Dim strVal

    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")                                       '⊙: Lock field
    Call InitData()
'    Call SetToolbar("1100100100001111")
    Call SetToolbar("1100101100001111")

    queryboolean = true

    'Call SetQuerySpreadColor()
'20080626::hanc    Call ChangeCaption()
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
    ReDim dayAr1(40)
    ReDim dayAr2(40)

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

    For i = 0 To 40
        dayAr1(i) =  Cstr(i+1) + "일"	'UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
        dayAr2(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text , parent.gDateFormat)
    Next

    With frm1.vspdData
        .Col = C_Qty_Day_0  : .Row = 0 : .Text = dayAr2(0)
        .Col = C_Qty_Day_1  : .Row = 0 : .Text = dayAr2(1)
        .Col = C_Qty_Day_2  : .Row = 0 : .Text = dayAr2(2)
        .Col = C_Qty_Day_3  : .Row = 0 : .Text = dayAr2(3)
        .Col = C_Qty_Day_4  : .Row = 0 : .Text = dayAr2(4)
        .Col = C_Qty_Day_5  : .Row = 0 : .Text = dayAr2(5)
        .Col = C_Qty_Day_6  : .Row = 0 : .Text = dayAr2(6)
        .Col = C_Qty_Day_7  : .Row = 0 : .Text = dayAr2(7)
        .Col = C_Qty_Day_8  : .Row = 0 : .Text = dayAr2(8)
        .Col = C_Qty_Day_9  : .Row = 0 : .Text = dayAr2(9)
        .Col = C_Qty_Day_10 : .Row = 0 : .Text = dayAr2(10)
        .Col = C_Qty_Day_11 : .Row = 0 : .Text = dayAr2(11)
        .Col = C_Qty_Day_12 : .Row = 0 : .Text = dayAr2(12)
        .Col = C_Qty_Day_13 : .Row = 0 : .Text = dayAr2(13)
        .Col = C_Qty_Day_14 : .Row = 0 : .Text = dayAr2(14)
        .Col = C_Qty_Day_15 : .Row = 0 : .Text = dayAr2(15)
        .Col = C_Qty_Day_16 : .Row = 0 : .Text = dayAr2(16)
        .Col = C_Qty_Day_17 : .Row = 0 : .Text = dayAr2(17)
        .Col = C_Qty_Day_18 : .Row = 0 : .Text = dayAr2(18)
        .Col = C_Qty_Day_19 : .Row = 0 : .Text = dayAr2(19)
        .Col = C_Qty_Day_20 : .Row = 0 : .Text = dayAr2(20)
        .Col = C_Qty_Day_21 : .Row = 0 : .Text = dayAr2(21)
        .Col = C_Qty_Day_22 : .Row = 0 : .Text = dayAr2(22)
        .Col = C_Qty_Day_23 : .Row = 0 : .Text = dayAr2(23)
        .Col = C_Qty_Day_24 : .Row = 0 : .Text = dayAr2(24)
        .Col = C_Qty_Day_25 : .Row = 0 : .Text = dayAr2(25)
        .Col = C_Qty_Day_26 : .Row = 0 : .Text = dayAr2(26)
        .Col = C_Qty_Day_27 : .Row = 0 : .Text = dayAr2(27)
        .Col = C_Qty_Day_28 : .Row = 0 : .Text = dayAr2(28)
        .Col = C_Qty_Day_29 : .Row = 0 : .Text = dayAr2(29)
        .Col = C_Qty_Day_30 : .Row = 0 : .Text = dayAr2(30)
        .Col = C_Qty_Day_31 : .Row = 0 : .Text = dayAr2(31)
        .Col = C_Qty_Day_32 : .Row = 0 : .Text = dayAr2(32)
        .Col = C_Qty_Day_33 : .Row = 0 : .Text = dayAr2(33)
        .Col = C_Qty_Day_34 : .Row = 0 : .Text = dayAr2(34)
        .Col = C_Qty_Day_35 : .Row = 0 : .Text = dayAr2(35)
        .Col = C_Qty_Day_36 : .Row = 0 : .Text = dayAr2(36)
        .Col = C_Qty_Day_37 : .Row = 0 : .Text = dayAr2(37)
        .Col = C_Qty_Day_38 : .Row = 0 : .Text = dayAr2(38)
        .Col = C_Qty_Day_39 : .Row = 0 : .Text = dayAr2(39)
        
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

        .Col = C_Qty_Day_31_Hidden : .Row = 0 : .Text = dayAr2(31)  'hanc
        .Col = C_Qty_Day_32_Hidden : .Row = 0 : .Text = dayAr2(32)
        .Col = C_Qty_Day_33_Hidden : .Row = 0 : .Text = dayAr2(33)
        .Col = C_Qty_Day_34_Hidden : .Row = 0 : .Text = dayAr2(34)
        .Col = C_Qty_Day_35_Hidden : .Row = 0 : .Text = dayAr2(35)
        .Col = C_Qty_Day_36_Hidden : .Row = 0 : .Text = dayAr2(36)
        .Col = C_Qty_Day_37_Hidden : .Row = 0 : .Text = dayAr2(37)
        .Col = C_Qty_Day_38_Hidden : .Row = 0 : .Text = dayAr2(38)
        .Col = C_Qty_Day_39_Hidden : .Row = 0 : .Text = dayAr2(39)
        
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
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
	If Trim(frm1.txtPlantCd.value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd.Focus
		Exit Function
	End If


'    If Not chkField(Document, "1") Then                                          '☜: This function check required field
'       Exit Function
'    End If

'	IntRetCD = DisplayMsgBox("ZZ0009", parent.VB_YES_NO,"x","x")   '900016 ZZ0009  20080304::HANC::IntRetCD = MsgBox("기존데이터가 존재할 경우 지워집니다. 작업하시겠습니까?", vbYesNo)
'	If IntRetCD = vbNo Then
'		Exit Function
'	End If

    Call DisableToolBar(parent.TBC_SAVE)

    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    FncSave = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()

    Dim lRow
    Dim lGrpCnt
    Dim strVal, strDel
    Dim iPeriod, iPeriodCnt
    Dim iInsCnt                 '20080312::hanc::입력 개수를 파악해서 입력일경우만 기존data 지워진다는 경고메시지 뿌림
    Dim IntRetCD


    DbSave = False

    If LayerShowHide(1) = False Then
         Exit Function
    End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    iInsCnt = 0
    iPeriod = Cint(frm1.txtPeriod.value)        '20080303::hanc
    
    With frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.UpdateFlag                                      '☜: Update

                                                   strVal = strVal & "U" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_Plant_Cd			:	strVal = strVal & Trim(frm1.txtPlantCd.value) & parent.gColSep
                                                            strVal = strVal & Trim(frm1.txtVersion.value) & parent.gColSep
                    .vspdData.Col = C_Item_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_bp_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Tracking_No		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(1)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_0         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_0_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(2)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_1         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_1_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(3)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_2         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_2_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(4)    & parent.gColSep
					
                    .vspdData.Col = C_Qty_Day_3         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_3_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(5)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_4         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_4_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(6)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_5         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_5_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(7)    & parent.gColSep
					
                    .vspdData.Col = C_Qty_Day_6         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_6_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(8)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_7         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_7_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(9)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_8         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_8_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(10)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_9         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_9_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(11)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_10        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_10_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(12)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_11        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_11_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(13)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_12        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_12_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(14)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_13        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_13_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(15)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_14        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_14_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(16)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_15        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_15_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(17)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_16        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_16_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(18)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_17        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_17_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(19)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_18        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_18_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(20)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_19        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_19_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(21)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_20        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_20_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(22)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_21        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_21_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(23)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_22        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_22_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(24)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_23        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_23_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(25)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_24        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_24_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(26)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_25        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_25_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(27)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_26        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_26_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(28)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_27        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_27_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(29)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_28        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_28_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(30)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_29        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_29_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(31)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_30        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_30_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

									                        strVal = strVal & arrColVal_header(32)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_31        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_31_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(33)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_32        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_32_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									                        strVal = strVal & arrColVal_header(34)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_33        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_33_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
'									                        strVal = strVal & arrColVal_header(2)    & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_34        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_34_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                                                            strVal = strVal & dayAr2(35)            & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_35        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_35_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                                                            strVal = strVal & dayAr2(36)            & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_36        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_36_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                                                            strVal = strVal & dayAr2(37)            & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_37        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_37_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                                                            strVal = strVal & dayAr2(38)            & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_38        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_38_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                                                            strVal = strVal & dayAr2(39)            & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_39        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Qty_Day_39_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    'msgbox strVal

                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                            strVal = strVal & "C" & parent.gColSep
                                                            strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_Plant_Cd			:	strVal = strVal & Trim(frm1.txtPlantCd.value) & parent.gColSep
                    .vspdData.Col = C_bp_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Item_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Tracking_No		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & frm1.txtPeriod.value & parent.gColSep


									                        strVal = strVal & arrColVal_header(iPeriodCnt+1)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_0         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+2)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_1         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+3)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_2         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+4)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_3         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+5)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_4         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+6)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_5         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+7)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_6         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+8)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_7         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+9)    & parent.gColSep
                    .vspdData.Col = C_Qty_Day_8         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+10)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_9         :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep

									                        strVal = strVal & arrColVal_header(iPeriodCnt+11)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_10        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+12)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_11        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+13)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_12        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+14)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_13        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+15)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_14        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+16)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_15        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+17)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_16        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+18)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_17        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+19)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_18        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+20)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_19        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep

									                        strVal = strVal & arrColVal_header(iPeriodCnt+21)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_20        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+22)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_21        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+23)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_22        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+24)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_23        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+25)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_24        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+26)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_25        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+27)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_26        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+28)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_27        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+29)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_28        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+30)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_29        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep

									                        strVal = strVal & arrColVal_header(iPeriodCnt+31)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_30        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+32)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_31        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+33)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_32        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+34)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_33        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+35)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_34        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+36)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_35        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+37)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_36        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+38)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_37        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+39)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_38        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep
									                        strVal = strVal & arrColVal_header(iPeriodCnt+40)   & parent.gColSep
                    .vspdData.Col = C_Qty_Day_39        :   strVal = strVal & Trim(.vspdData.Text)              & parent.gColSep

                                                            strVal = strVal &  parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
                    iInsCnt = iInsCnt + 1       '20080312::hanc
                    
               Case ggoSpread.DeleteFlag                                      '☜: delete
                                                            strVal = strVal & "D" & parent.gColSep
                                                            strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_Plant_Cd			:	strVal = strVal & Trim(frm1.txtPlantCd.value) & parent.gColSep
                    .vspdData.Col = C_Item_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Tracking_No		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & frm1.txtPeriod.value & parent.gColSep
                                                            strVal = strVal & frm1.txtVersion.value & parent.gColSep
                    .vspdData.Col = C_bp_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep


                                                            strVal = strVal &  parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
       .txtMaxRows.value     = lGrpCnt - 1
       .txtSpread.value      = strDel & strVal

    End With

    '20080312::hanc::입력일 경우만 기존데이터 지워진다는 경고 창 띄운다.	
'    if iInsCnt > 0 then
'    	IntRetCD = DisplayMsgBox("ZZ0009", parent.VB_YES_NO,"x","x")   '900016 ZZ0009  20080304::HANC::IntRetCD = MsgBox("기존데이터가 존재할 경우 지워집니다. 작업하시겠습니까?", vbYesNo)
'    	If IntRetCD = vbNo Then
'    	    call LayerShowHide(0)
'    		Exit Function
'    	End If
'    end if

    Call ExecMyBizASP(frm1, BIZ_PGM_ID3)

    DbSave = True
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
        If Row < 1 Then Exit Sub

    End With
End Sub


'20080302::hanc
Sub SetHeader(ByVal lgstrData_header)
    arrColVal_header = Split(lgstrData_header, parent.gColSep)

           frm1.vspdData.Row = 0
           frm1.vspdData.Col = C_Qty_Day_0  :            frm1.vspdData.Text = arrColVal_header(1)
           frm1.vspdData.Col = C_Qty_Day_1  :            frm1.vspdData.Text = arrColVal_header(2)
           frm1.vspdData.Col = C_Qty_Day_2  :            frm1.vspdData.Text = arrColVal_header(3)
           frm1.vspdData.Col = C_Qty_Day_3  :            frm1.vspdData.Text = arrColVal_header(4)
           frm1.vspdData.Col = C_Qty_Day_4  :            frm1.vspdData.Text = arrColVal_header(5)
           frm1.vspdData.Col = C_Qty_Day_5  :            frm1.vspdData.Text = arrColVal_header(6)
           frm1.vspdData.Col = C_Qty_Day_6  :            frm1.vspdData.Text = arrColVal_header(7)
           frm1.vspdData.Col = C_Qty_Day_7  :            frm1.vspdData.Text = arrColVal_header(8)
           frm1.vspdData.Col = C_Qty_Day_8  :            frm1.vspdData.Text = arrColVal_header(9)
           frm1.vspdData.Col = C_Qty_Day_9  :            frm1.vspdData.Text = arrColVal_header(10)
           frm1.vspdData.Col = C_Qty_Day_10 :            frm1.vspdData.Text = arrColVal_header(11)
           frm1.vspdData.Col = C_Qty_Day_11 :            frm1.vspdData.Text = arrColVal_header(12)
           frm1.vspdData.Col = C_Qty_Day_12 :            frm1.vspdData.Text = arrColVal_header(13)
           frm1.vspdData.Col = C_Qty_Day_13 :            frm1.vspdData.Text = arrColVal_header(14)
           frm1.vspdData.Col = C_Qty_Day_14 :            frm1.vspdData.Text = arrColVal_header(15)
           frm1.vspdData.Col = C_Qty_Day_15 :            frm1.vspdData.Text = arrColVal_header(16)
           frm1.vspdData.Col = C_Qty_Day_16 :            frm1.vspdData.Text = arrColVal_header(17)
           frm1.vspdData.Col = C_Qty_Day_17 :            frm1.vspdData.Text = arrColVal_header(18)
           frm1.vspdData.Col = C_Qty_Day_18 :            frm1.vspdData.Text = arrColVal_header(19)
           frm1.vspdData.Col = C_Qty_Day_19 :            frm1.vspdData.Text = arrColVal_header(20)
           frm1.vspdData.Col = C_Qty_Day_20 :            frm1.vspdData.Text = arrColVal_header(21)
           frm1.vspdData.Col = C_Qty_Day_21 :            frm1.vspdData.Text = arrColVal_header(22)
           frm1.vspdData.Col = C_Qty_Day_22 :            frm1.vspdData.Text = arrColVal_header(23)
           frm1.vspdData.Col = C_Qty_Day_23 :            frm1.vspdData.Text = arrColVal_header(24)
           frm1.vspdData.Col = C_Qty_Day_24 :            frm1.vspdData.Text = arrColVal_header(25)
           frm1.vspdData.Col = C_Qty_Day_25 :            frm1.vspdData.Text = arrColVal_header(26)
           frm1.vspdData.Col = C_Qty_Day_26 :            frm1.vspdData.Text = arrColVal_header(27)
           frm1.vspdData.Col = C_Qty_Day_27 :            frm1.vspdData.Text = arrColVal_header(28)
           frm1.vspdData.Col = C_Qty_Day_28 :            frm1.vspdData.Text = arrColVal_header(29)
           frm1.vspdData.Col = C_Qty_Day_29 :            frm1.vspdData.Text = arrColVal_header(30)
           frm1.vspdData.Col = C_Qty_Day_30 :            frm1.vspdData.Text = arrColVal_header(31)
           frm1.vspdData.Col = C_Qty_Day_31 :            frm1.vspdData.Text = arrColVal_header(32)
           frm1.vspdData.Col = C_Qty_Day_32 :            frm1.vspdData.Text = arrColVal_header(33)
        if arrColVal_header(34) = "1900-01-01" then
           frm1.vspdData.Col = C_Qty_Day_33 :            frm1.vspdData.Text = ""
        else
           frm1.vspdData.Col = C_Qty_Day_33 :            frm1.vspdData.Text = arrColVal_header(34)
        end if

'           frm1.vspdData.Col = C_Qty_Day_34 :            frm1.vspdData.Text = arrColVal_header(35)
'           frm1.vspdData.Col = C_Qty_Day_35 :            frm1.vspdData.Text = arrColVal_header(36)
'           frm1.vspdData.Col = C_Qty_Day_36 :            frm1.vspdData.Text = arrColVal_header(37)
'           frm1.vspdData.Col = C_Qty_Day_37 :            frm1.vspdData.Text = arrColVal_header(38)
'           frm1.vspdData.Col = C_Qty_Day_38 :            frm1.vspdData.Text = arrColVal_header(39)
'           frm1.vspdData.Col = C_Qty_Day_39 :            frm1.vspdData.Text = arrColVal_header(40)
           

End Sub
'======================================================================================================
'	Name : DBAutoQueryOk()
'	Description : s3314mb2_ko441.asp 이후 Query OK해 줌
'=======================================================================================================
Sub DBAutoQueryOk()
    Dim lRow
	Dim intIndex
	Dim daytimeVal 
	Dim strSub_type 
    
    With Frm1
        .vspdData.ReDraw = false
         ggoSpread.Source = .vspdData
   
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            .vspdData.Text =  ggoSpread.InsertFlag
       Next
            .vspdData.ReDraw = TRUE
        
    End With 
    ggoSpread.ClearSpreadData "T"
     Set gActiveElement = document.ActiveElement   
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

'===============================================================================================
'   by Shin hyoung jae 
'	Name : GetOpenFilePath()
'	Description : GetTextFilePath	
'================================================================================================= 
Function GetOpenFilePath()
	Dim dlg
    Dim sPath
 
	On Error Resume Next
	Set dlg = CreateObject("uni2kCM.SaveFile")
	
	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If
	
    sPath = dlg.GetOpenFilePath()

	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If

	lgFilePath2 = sPath
	frm1.txtFileName2.Value = ExtractFileName(sPath)

    Set dlg = Nothing
	frm1.hFileName.value = sPath
End Function

Function ExtractFileName(byVal strPath)
	strPath = StrReverse(strPath)
	strPath = Left(strPath, InStr(strPath, "\") - 1)
	ExtractFileName = StrReverse(strPath)
End Function

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect() 
    Dim strVal
    Dim IntRetCD

	
	ggoSpread.ClearSpreadData

	If Trim(frm1.txtPlantCd.Value) = "" then        '20080303::hanc
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If


	If trim(frm1.txtFileName2.value) = "" Then
		call DisplayMsgBox("970029","X" , frm1.txtFileName2.Alt, "X")
		frm1.txtFileName2.focus 	
		Exit Function
	Else
		
		if (ggoSaveFile.fileExists(frm1.hFileName.value) = 0)  = false  then
			IntRetCD = DisplayMsgBox("115191","x","x","x")                           '☜:There is no picture
			Exit Function
		end if
			
	End If		    
    
    ExeReflect = False                                                              '☜: Processing is NG
	
    Call MakeKeyStream()    '20080129::hanc:: "X" 제외
  
	Call RemovedivTextArea 	

    If LayerShowHide(1) = false Then
        Exit Function
    End If
    
    strVal =""
    
	strVal = BIZ_PGM_ID2 & "?txtMode="			& Parent.UID_M0001						'☜: Query
	strVal = strVal      & "&txtKeyStream="     & lgKeyStream							'☜: Query Key
	strVal = strVal      & "&lgStrPrevKey="		& lgStrPrevKey							'☜: Next key tag
	strVal = strVal      & "&txtMaxRows="       & Frm1.vspdData.MaxRows					'☜: Max fetched data	
	strVal = strVal      & "&htxtFileGubun="    & "A"
	
    Call RunMyBizASP(MyBizASP, strVal)													'☜:  Run biz logic

    ExeReflect = True																	'☜: Processing is NG		

End Function

'========================================================================================
 ' Function Name : RemovedivTextArea
 ' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
 Function RemovedivTextArea()
 
 	Dim ii
 		
 	For ii = 1 To divTextArea.children.length
 	    divTextArea.removeChild(divTextArea.children(0))
 	Next
 
 End Function
 
 '2008-06-26 9:10오전 :: hanc
 Function OpenVersion()
	Dim iCalledAspName
	Dim IntRetCD

	If GetSetupMod(Parent.gSetupMod, "p") <> "Y" Then
    	Call DisplayMsgBox("169936","X", "X", "X")
		Exit Function
	End if
				
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6 

	If IsOpenPop = True Then Exit Function
	

	Param1 = Trim(frm1.txtVersion.value)
	Param2 = ""
'20080312::hanc	If Trim(frm1.txtSLCd1.value) = "" then
'20080312::hanc		Call ClickTab1()   
'20080312::hanc		Call DisplayMsgBox("169902","X","X","X")
'20080312::hanc		frm1.txtSLCd1.Focus
'20080312::hanc	    Set gActiveElement = document.activeElement
'20080312::hanc		Exit Function
'20080312::hanc	End If
	
	Param3 = ""
	Param4 = ""
	
	
'20080312::hanc	If Trim(frm1.txtSLCd2.value) = "" then
'20080312::hanc		
'20080312::hanc		If ClickTab2_1() Then                       '20080306::hanc::ClickTab2_1 만든이유 : 자재불출의뢰정보 참조팝업 시 수불유형 선택하지 않고 팝업창 띄우기 위함.
'20080312::hanc			Call DisplayMsgBox("169937","X","X","X")
'20080312::hanc			'frm1.txtSLCd2.Focus
'20080312::hanc			'Set gActiveElement = document.activeElement
'20080312::hanc		End if
'20080312::hanc	    Exit Function
'20080312::hanc	End If
	
	Param5 = ""
	Param6 = ""
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("p2250XA1_KO441")     '20080226::HANC
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1311RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6), _
		 "dialogWidth=250px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    

	IsOpenPop = False
	
	If arrRet(0,0) = "" Then
		frm1.txtSLCd2.focus
		Exit Function
	Else
		Call SetMoveInvRef1(arrRet)
	End If
	
End Function
'20080226::hanc***********************************************************
Function SetMoveInvRef1(arrRet)
	Dim TempRow
	Dim intLoopCnt
	Dim intCnt
	Dim iRow

		frm1.txtVersion.value = arrRet(0, 0)	
		

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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=*>&nbsp;</TD>
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
                                <TD CLASS="TD5" NOWRAP>공장</TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="12XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=29 MAXLENGTH=40 tag="14"></TD>
                                <TD CLASS="TD5" NOWRAP>Ver.</TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=14 NAME="txtVersion" MAXLENGTH="14" tag="12XXXU" ALT = "Ver"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVersion"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenVersion()"></TD>
                                <TD CLASS="TD5" NOWRAP style="display:none">생산계획기간</TD>
                                <TD CLASS="TD6" NOWRAP style="display:none">
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px;  HEIGHT: 20px" name=txtYYYYMM CLASSID=<%=gCLSIDFPDT%> ALT="기준일자" tag="22X1" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>&nbsp;
  											</TD>
											<TD>
                                                <INPUT TYPE=TEXT  NAME="txtPeriod" SIZE=2 MAXLENGTH=2 tag="14" ALT = "생산계획기간">일
											</TD>
							            </TR>
					               </TABLE>
                                </TD>
                            </TR>
                            <TR>
                                <TD CLASS="TD5" NOWRAP STYLE="DISPLAY:NONE">품목</TD>
                                <TD CLASS="TD6" NOWRAP STYLE="DISPLAY:NONE"><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>
                                <TD CLASS="TD5" NOWRAP STYLE="DISPLAY:NONE">Tracking No.</TD>      
								<TD CLASS="TD6" NOWRAP STYLE="DISPLAY:NONE"><INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>
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
                                <TD HEIGHT="100%">
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
    <TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
					<TD WIDTH=10>&nbsp;</TD>
	                <TD WIDTH=10 style="display:none"><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: FncSave" >실행</BUTTON>&nbsp;</TD>
					<TD WIDTH=10 style="display:none"><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: FncDelete" >취소</BUTTON>&nbsp;</TD>
					<TD CLASS=TD5 NOWRAP>파일경로</TD>
					<TD WIDTH=210><INPUT TYPE=TEXT ID="txtFileName2" NAME="txtFileName2" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="화일명" tag="14X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
					<TD WIDTH=10 align=left><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>Import</BUTTON></TD>
	                <TD Width=*>&nbsp;</TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>   
    </TR>

    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd"   tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo"   tag="24">

<INPUT TYPE=HIDDEN NAME="hFileName" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>