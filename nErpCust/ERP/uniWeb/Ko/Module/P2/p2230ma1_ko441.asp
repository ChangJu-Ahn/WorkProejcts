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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "P2230mb1_KO441.asp"

Dim LocSvrDate
Dim strDate
LocSvrDate = "<%=GetSvrDate%>"

strDate = UniConvDateAToB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat) 	'☆: 초기화면에 뿌려지는 마지막 날짜 

Dim IsOpenPop
Dim queryboolean
queryboolean = False

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
    Call ggoOper.FormatDate(frm1.txtYYYYMM, parent.gDateFormat, "1")    '20080304::HANC:: 2->1
    Call ggoOper.LockField(Document, "N")                                           '⊙: Lock Field

    Call DbQueryPeriod      '20080303::hanc

    Call SetDefaultVal
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox
    Call InitSpreadComboBox
    Call CookiePage

    Call SetToolbar("1100100100001111")                                             '버튼 툴바 제어
    
    Set gActiveElement = document.activeElement
    
End Sub

'20080303::hanc----------------------------------------------------------------------
'생산계획기간 가져오기---------------------------------------------------------------
Function DbQueryPeriod()
    DbQueryPeriod = False
    Err.Clear                                                                        '☜: Clear err status

    Dim strVal

    With frm1
        strVal = BIZ_PGM_ID & "?txtMode="      & parent.UID_M0004
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbQueryPeriod = True
End Function


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

    For i = 0 To 39         'hanc :: 30
        dayAr(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
    Next

    Call InitSpreadPosVariables()

    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

        .ReDraw = false

        .MaxCols = C_Qty_Day_39_Hidden + 1      'hanc                                <%'☜: 최대 Columns의 항상 1개 증가시킴 %>

        .Col = .MaxCols                                                         <%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0

        ggoSpread.Source = Frm1.vspdData

        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos()

        ggoSpread.SSSetEdit     C_Item_Cd       , "품목"   , 10
        ggoSpread.SSSetEdit     C_Item_Nm       , "품목명" , 20
        ggoSpread.SSSetEdit     C_Tracking_No   , "Tracking No"   , 10
        ggoSpread.SSSetEdit     C_Type_cd	    , "구분"   , 10
        ggoSpread.SSSetEdit     C_Type		    , "구분"   , 10
        
        for i = 1 To 40   '20080304::hanc:: 31
			'ggoSpread.SSSetFloat    i + 5     , dayAr(i-1) , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			ggoSpread.SSSetFloat    i + 5     , Cstr(i) + "일" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        Next

       	ggoSpread.SSSetFloat    C_Qty_Month_1     , "월합계" , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
  		ggoSpread.SSSetFloat    C_Qty_Month_2     , startDate1 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat    C_Qty_Month_3     , startDate2 , 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit     C_Plant_Cd		  , "구분"   , 10
		
		for i = 1 To 40   '20080304::hanc:: 31
			ggoSpread.SSSetFloat    i + 49     , dayAr(i-1), 10, Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        Next

        Call ggoSpread.SSSetColHidden(C_Plant_Cd         , C_Plant_Cd         , True)
        Call ggoSpread.SSSetColHidden(C_Type_Cd          , C_Type_Cd          , True)
        Call ggoSpread.SSSetColHidden(C_Qty_Day_0_Hidden , C_Qty_Day_39_Hidden , True)  'hanc
        Call ggoSpread.SSSetColHidden(C_Tracking_No      , C_Tracking_No          , True)
        Call ggoSpread.SSSetColHidden(C_Qty_Month_1      , C_Qty_Month_3          , True)


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

    C_Item_Cd           = 1 
    C_Item_Nm           = 2
    C_Tracking_No       = 3
    C_Type_cd           = 4
    C_Type		        = 5
    C_Qty_Day_0         = 6
    C_Qty_Day_1         = 7
    C_Qty_Day_2         = 8
    C_Qty_Day_3         = 9
    C_Qty_Day_4         = 10
    C_Qty_Day_5         = 11
    C_Qty_Day_6         = 12
    C_Qty_Day_7         = 13
    C_Qty_Day_8         = 14
    C_Qty_Day_9         = 15
    C_Qty_Day_10        = 16
    C_Qty_Day_11        = 17
    C_Qty_Day_12        = 18
    C_Qty_Day_13        = 19
    C_Qty_Day_14        = 20
    C_Qty_Day_15        = 21
    C_Qty_Day_16        = 22
    C_Qty_Day_17        = 23
    C_Qty_Day_18        = 24
    C_Qty_Day_19        = 25
    C_Qty_Day_20        = 26
    C_Qty_Day_21        = 27
    C_Qty_Day_22        = 28
    C_Qty_Day_23        = 29
    C_Qty_Day_24        = 30
    C_Qty_Day_25        = 31
    C_Qty_Day_26        = 32
    C_Qty_Day_27        = 33
    C_Qty_Day_28        = 34
    C_Qty_Day_29        = 35
    C_Qty_Day_30        = 36


    C_Qty_Day_31        = 37        '20080303::hanc
    C_Qty_Day_32        = 38
    C_Qty_Day_33        = 39
    C_Qty_Day_34        = 40
    C_Qty_Day_35        = 41

    C_Qty_Day_36        = 42
    C_Qty_Day_37        = 43
    C_Qty_Day_38        = 44
    C_Qty_Day_39        = 45
    
    C_Qty_Month_1       = 46
    C_Qty_Month_2       = 47
    C_Qty_Month_3       = 48
    C_Plant_Cd          = 49
    C_Qty_Day_0_Hidden  = 50
    C_Qty_Day_1_Hidden  = 51
    C_Qty_Day_2_Hidden  = 52
    C_Qty_Day_3_Hidden  = 53
    C_Qty_Day_4_Hidden  = 54
    C_Qty_Day_5_Hidden  = 55
    C_Qty_Day_6_Hidden  = 56
    C_Qty_Day_7_Hidden  = 57
    C_Qty_Day_8_Hidden  = 58
    C_Qty_Day_9_Hidden  = 59
    C_Qty_Day_10_Hidden = 60
    C_Qty_Day_11_Hidden = 61
    C_Qty_Day_12_Hidden = 62
    C_Qty_Day_13_Hidden = 63
    C_Qty_Day_14_Hidden = 64
    C_Qty_Day_15_Hidden = 65
    C_Qty_Day_16_Hidden = 66
    C_Qty_Day_17_Hidden = 67
    C_Qty_Day_18_Hidden = 68
    C_Qty_Day_19_Hidden = 69
    C_Qty_Day_20_Hidden = 70
    C_Qty_Day_21_Hidden = 71
    C_Qty_Day_22_Hidden = 72
    C_Qty_Day_23_Hidden = 73
    C_Qty_Day_24_Hidden = 74
    C_Qty_Day_25_Hidden = 75
    C_Qty_Day_26_Hidden = 76
    C_Qty_Day_27_Hidden = 77
    C_Qty_Day_28_Hidden = 78
    C_Qty_Day_29_Hidden = 79
    C_Qty_Day_30_Hidden = 80

    C_Qty_Day_31_Hidden = 81
    C_Qty_Day_32_Hidden = 82
    C_Qty_Day_33_Hidden = 83
    C_Qty_Day_34_Hidden = 84
    C_Qty_Day_35_Hidden = 85
    C_Qty_Day_36_Hidden = 86
    C_Qty_Day_37_Hidden = 87
    C_Qty_Day_38_Hidden = 88
    C_Qty_Day_39_Hidden = 89

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

    C_Item_Cd           = iCurColumnPos(1)  
    C_Item_Nm           = iCurColumnPos(2)
    C_Tracking_no       = iCurColumnPos(3)
    C_Type_cd		    = iCurColumnPos(4)
    C_Type              = iCurColumnPos(5)
    C_Qty_Day_0         = iCurColumnPos(6)
    C_Qty_Day_1         = iCurColumnPos(7)
    C_Qty_Day_2         = iCurColumnPos(8)
    C_Qty_Day_3         = iCurColumnPos(9)
    C_Qty_Day_4         = iCurColumnPos(10)
    C_Qty_Day_5         = iCurColumnPos(11)
    C_Qty_Day_6         = iCurColumnPos(12)
    C_Qty_Day_7         = iCurColumnPos(13)
    C_Qty_Day_8         = iCurColumnPos(14)
    C_Qty_Day_9         = iCurColumnPos(15)
    C_Qty_Day_10        = iCurColumnPos(16)
    C_Qty_Day_11        = iCurColumnPos(17)
    C_Qty_Day_12        = iCurColumnPos(18)
    C_Qty_Day_13        = iCurColumnPos(19)
    C_Qty_Day_14        = iCurColumnPos(20)
    C_Qty_Day_15        = iCurColumnPos(21)
    C_Qty_Day_16        = iCurColumnPos(22)
    C_Qty_Day_17        = iCurColumnPos(23)
    C_Qty_Day_18        = iCurColumnPos(24)
    C_Qty_Day_19        = iCurColumnPos(25)
    C_Qty_Day_20        = iCurColumnPos(26)
    C_Qty_Day_21        = iCurColumnPos(27)
    C_Qty_Day_22        = iCurColumnPos(28)
    C_Qty_Day_23        = iCurColumnPos(29)
    C_Qty_Day_24        = iCurColumnPos(30)
    C_Qty_Day_25        = iCurColumnPos(31)
    C_Qty_Day_26        = iCurColumnPos(32)
    C_Qty_Day_27        = iCurColumnPos(33)
    C_Qty_Day_28        = iCurColumnPos(34)
    C_Qty_Day_29        = iCurColumnPos(35)
    C_Qty_Day_30        = iCurColumnPos(36)

    C_Qty_Day_31        = iCurColumnPos(37)     '20080303::hanc
    C_Qty_Day_32        = iCurColumnPos(38)
    C_Qty_Day_33        = iCurColumnPos(39)
    C_Qty_Day_34        = iCurColumnPos(40)
    C_Qty_Day_35        = iCurColumnPos(41)

    C_Qty_Day_36        = iCurColumnPos(42)
    C_Qty_Day_37        = iCurColumnPos(43)
    C_Qty_Day_38        = iCurColumnPos(44)
    C_Qty_Day_39        = iCurColumnPos(45)

    C_Qty_Month_1       = iCurColumnPos(46)
    C_Qty_Month_2       = iCurColumnPos(47)
    C_Qty_Month_3       = iCurColumnPos(48)
    C_Plant_Cd          = iCurColumnPos(49)
    C_Qty_Day_0_Hidden  = iCurColumnPos(50)
    C_Qty_Day_1_Hidden  = iCurColumnPos(51)
    C_Qty_Day_2_Hidden  = iCurColumnPos(52)
    C_Qty_Day_3_Hidden  = iCurColumnPos(53)
    C_Qty_Day_4_Hidden  = iCurColumnPos(54)
    C_Qty_Day_5_Hidden  = iCurColumnPos(55)
    C_Qty_Day_6_Hidden  = iCurColumnPos(56)
    C_Qty_Day_7_Hidden  = iCurColumnPos(57)
    C_Qty_Day_8_Hidden  = iCurColumnPos(58)
    C_Qty_Day_9_Hidden  = iCurColumnPos(59)
    C_Qty_Day_10_Hidden = iCurColumnPos(60)
    C_Qty_Day_11_Hidden = iCurColumnPos(61)
    C_Qty_Day_12_Hidden = iCurColumnPos(62)
    C_Qty_Day_13_Hidden = iCurColumnPos(63)
    C_Qty_Day_14_Hidden = iCurColumnPos(64)
    C_Qty_Day_15_Hidden = iCurColumnPos(65)
    C_Qty_Day_16_Hidden = iCurColumnPos(66)
    C_Qty_Day_17_Hidden = iCurColumnPos(67)
    C_Qty_Day_18_Hidden = iCurColumnPos(68)
    C_Qty_Day_19_Hidden = iCurColumnPos(69)
    C_Qty_Day_20_Hidden = iCurColumnPos(70)
    C_Qty_Day_21_Hidden = iCurColumnPos(71)
    C_Qty_Day_22_Hidden = iCurColumnPos(72)
    C_Qty_Day_23_Hidden = iCurColumnPos(73)
    C_Qty_Day_24_Hidden = iCurColumnPos(74)
    C_Qty_Day_25_Hidden = iCurColumnPos(75)
    C_Qty_Day_26_Hidden = iCurColumnPos(76)
    C_Qty_Day_27_Hidden = iCurColumnPos(77)
    C_Qty_Day_28_Hidden = iCurColumnPos(78)
    C_Qty_Day_29_Hidden = iCurColumnPos(79)
    C_Qty_Day_30_Hidden = iCurColumnPos(80)

    C_Qty_Day_31_Hidden = iCurColumnPos(81)
    C_Qty_Day_32_Hidden = iCurColumnPos(82)
    C_Qty_Day_33_Hidden = iCurColumnPos(83)
    C_Qty_Day_34_Hidden = iCurColumnPos(84)
    C_Qty_Day_35_Hidden = iCurColumnPos(85)
    C_Qty_Day_36_Hidden = iCurColumnPos(86)
    C_Qty_Day_37_Hidden = iCurColumnPos(87)
    C_Qty_Day_38_Hidden = iCurColumnPos(88)
    C_Qty_Day_39_Hidden = iCurColumnPos(89)

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
			   ggoSpread.SpreadLock C_Qty_Day_0 ,  i, C_Qty_Day_39_Hidden , i       'hanc
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
    lgKeyStream = lgKeyStream & frm1.txtPlantCd.value & parent.gColSep      '20080304::HANC
    lgKeyStream = lgKeyStream & frm1.txtYYYYMM.value & parent.gColSep      '20080304::HANC
    
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
    Call SetToolbar("1100100100001111")

    queryboolean = true

    'Call SetQuerySpreadColor()
    Call ChangeCaption()
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

    For i = 0 To 39     '20080304::HANC::  30
        dayAr1(i) =  Cstr(i+1) + "일"	'UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
'20080304::HANC        dayAr2(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text + "-01", parent.gDateFormat)
'MSGBOX "frm1.txtyyyymm.text : " & frm1.txtyyyymm.text
        dayAr2(i) =  UNIDateAdd("d", i , frm1.txtyyyymm.text, parent.gDateFormat)
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
        
'        .Col = C_Qty_Day_31_Hidden : .Row = 0 : .Text = dayAr2(31)
'        .Col = C_Qty_Day_32_Hidden : .Row = 0 : .Text = dayAr2(32)
'        .Col = C_Qty_Day_33_Hidden : .Row = 0 : .Text = dayAr2(33)
'        .Col = C_Qty_Day_34_Hidden : .Row = 0 : .Text = dayAr2(34)
'        .Col = C_Qty_Day_35_Hidden : .Row = 0 : .Text = dayAr2(35)
'        .Col = C_Qty_Day_36_Hidden : .Row = 0 : .Text = dayAr2(36)
'        .Col = C_Qty_Day_37_Hidden : .Row = 0 : .Text = dayAr2(37)
'        .Col = C_Qty_Day_38_Hidden : .Row = 0 : .Text = dayAr2(38)
'        .Col = C_Qty_Day_39_Hidden : .Row = 0 : .Text = dayAr2(39)
        
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

    If Not chkField(Document, "1") Then                                          '☜: This function check required field
       Exit Function
    End If

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

    DbSave = False

    If LayerShowHide(1) = False Then
         Exit Function
    End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    With frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                   strVal = strVal & "U" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_Plant_Cd			:	strVal = strVal & Trim(.txtPlantCd.Value) & parent.gColSep
                    .vspdData.Col = C_Item_Cd			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Tracking_No		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
															strVal = strVal & dayAr2(0)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_0         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_0_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(1)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_1         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_1_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(2)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_2         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_2_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(3)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_3         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_3_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(4)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_4         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_4_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(5)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_5         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_5_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(6)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_6         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_6_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(7)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_7         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_7_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(8)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_8         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_8_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(9)             & parent.gColSep
                    .vspdData.Col = C_Qty_Day_9         :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_9_Hidden  :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(10)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_10        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_10_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(11)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_11        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_11_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(12)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_12        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_12_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(13)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_13        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_13_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(14)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_14        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_14_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(15)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_15        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_15_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(16)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_16        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_16_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(17)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_17        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_17_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(18)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_18        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_18_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(19)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_19        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_19_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(20)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_20        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_20_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(21)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_21        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_21_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(22)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_22        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_22_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(23)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_23        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_23_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(24)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_24        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_24_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(25)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_25        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_25_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(26)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_26        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_26_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(27)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_27        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_27_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(28)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_28        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_28_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(29)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_29        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_29_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(30)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_30        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_30_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

                                                            strVal = strVal & dayAr2(31)           & parent.gColSep
                    .vspdData.Col = C_Qty_Day_31        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_31_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(32)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_32        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_32_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(33)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_33        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_33_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(34)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_34        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_34_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(35)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_35        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_35_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(36)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_36        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_36_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(37)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_37        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_37_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(38)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_38        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_38_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                            strVal = strVal & dayAr2(39)            & parent.gColSep
                    .vspdData.Col = C_Qty_Day_39        :   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Qty_Day_39_Hidden :   strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
       .txtMaxRows.value     = lGrpCnt - 1
       .txtSpread.value      = strDel & strVal

    End With

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)

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

'2008-06-05::hanc::수주계획일괄복사
Function RegProd() 
    Dim lRow, ptfForMps, startDate, endDate
    Dim s_C_Qty_Day_0,  s_C_Qty_Day_1,  s_C_Qty_Day_2,  s_C_Qty_Day_3,  s_C_Qty_Day_4
    Dim s_C_Qty_Day_5,  s_C_Qty_Day_6,  s_C_Qty_Day_7,  s_C_Qty_Day_8,  s_C_Qty_Day_9
    Dim s_C_Qty_Day_10,  s_C_Qty_Day_11,  s_C_Qty_Day_12,  s_C_Qty_Day_13,  s_C_Qty_Day_14
    Dim s_C_Qty_Day_15,  s_C_Qty_Day_16,  s_C_Qty_Day_17,  s_C_Qty_Day_18,  s_C_Qty_Day_19
    Dim s_C_Qty_Day_20,  s_C_Qty_Day_21,  s_C_Qty_Day_22,  s_C_Qty_Day_23,  s_C_Qty_Day_24
    Dim s_C_Qty_Day_25,  s_C_Qty_Day_26,  s_C_Qty_Day_27,  s_C_Qty_Day_28,  s_C_Qty_Day_29
    Dim s_C_Qty_Day_30,  s_C_Qty_Day_31,  s_C_Qty_Day_32,  s_C_Qty_Day_33,  s_C_Qty_Day_34
    Dim s_C_Qty_Day_35,  s_C_Qty_Day_36,  s_C_Qty_Day_37,  s_C_Qty_Day_38,  s_C_Qty_Day_39
    

    With frm1

        .vspdData.Redraw = False

        ggoSpread.Source = frm1.vspdData

        For lRow = 1 To frm1.vspdData.MaxRows step 2

           .vspdData.Row = lRow
           .vspdData.Col = 0

                    .vspdData.Col = C_Qty_Day_0         :   s_C_Qty_Day_0           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_1         :   s_C_Qty_Day_1           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_2         :   s_C_Qty_Day_2           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_3         :   s_C_Qty_Day_3           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_4         :   s_C_Qty_Day_4           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_5         :   s_C_Qty_Day_5           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_6         :   s_C_Qty_Day_6           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_7         :   s_C_Qty_Day_7           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_8         :   s_C_Qty_Day_8           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_9         :   s_C_Qty_Day_9           =  Trim(.vspdData.Text) 
                    
                    .vspdData.Col = C_Qty_Day_10         :   s_C_Qty_Day_10           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_11         :   s_C_Qty_Day_11           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_12         :   s_C_Qty_Day_12           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_13         :   s_C_Qty_Day_13           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_14         :   s_C_Qty_Day_14           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_15         :   s_C_Qty_Day_15           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_16         :   s_C_Qty_Day_16           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_17         :   s_C_Qty_Day_17           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_18         :   s_C_Qty_Day_18           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_19         :   s_C_Qty_Day_19           =  Trim(.vspdData.Text) 

                    .vspdData.Col = C_Qty_Day_20         :   s_C_Qty_Day_20           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_21         :   s_C_Qty_Day_21           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_22         :   s_C_Qty_Day_22           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_23         :   s_C_Qty_Day_23           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_24         :   s_C_Qty_Day_24           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_25         :   s_C_Qty_Day_25           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_26         :   s_C_Qty_Day_26           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_27         :   s_C_Qty_Day_27           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_28         :   s_C_Qty_Day_28           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_29         :   s_C_Qty_Day_29           =  Trim(.vspdData.Text) 

                    .vspdData.Col = C_Qty_Day_30         :   s_C_Qty_Day_30           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_31         :   s_C_Qty_Day_31           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_32         :   s_C_Qty_Day_32           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_33         :   s_C_Qty_Day_33           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_34         :   s_C_Qty_Day_34           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_35         :   s_C_Qty_Day_35           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_36         :   s_C_Qty_Day_36           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_37         :   s_C_Qty_Day_37           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_38         :   s_C_Qty_Day_38           =  Trim(.vspdData.Text) 
                    .vspdData.Col = C_Qty_Day_39         :   s_C_Qty_Day_39           =  Trim(.vspdData.Text) 

           .vspdData.Row = lRow + 1
           .vspdData.Col = 0

                    .vspdData.Col = C_Qty_Day_0         :   .vspdData.Text  =   s_C_Qty_Day_0
                    .vspdData.Col = C_Qty_Day_1         :   .vspdData.Text  =   s_C_Qty_Day_1
                    .vspdData.Col = C_Qty_Day_2         :   .vspdData.Text  =   s_C_Qty_Day_2
                    .vspdData.Col = C_Qty_Day_3         :   .vspdData.Text  =   s_C_Qty_Day_3
                    .vspdData.Col = C_Qty_Day_4         :   .vspdData.Text  =   s_C_Qty_Day_4
                    .vspdData.Col = C_Qty_Day_5         :   .vspdData.Text  =   s_C_Qty_Day_5
                    .vspdData.Col = C_Qty_Day_6         :   .vspdData.Text  =   s_C_Qty_Day_6
                    .vspdData.Col = C_Qty_Day_7         :   .vspdData.Text  =   s_C_Qty_Day_7
                    .vspdData.Col = C_Qty_Day_8         :   .vspdData.Text  =   s_C_Qty_Day_8
                    .vspdData.Col = C_Qty_Day_9         :   .vspdData.Text  =   s_C_Qty_Day_9
                    
                    .vspdData.Col = C_Qty_Day_10         :   .vspdData.Text  =   s_C_Qty_Day_10
                    .vspdData.Col = C_Qty_Day_11         :   .vspdData.Text  =   s_C_Qty_Day_11
                    .vspdData.Col = C_Qty_Day_12         :   .vspdData.Text  =   s_C_Qty_Day_12
                    .vspdData.Col = C_Qty_Day_13         :   .vspdData.Text  =   s_C_Qty_Day_13
                    .vspdData.Col = C_Qty_Day_14         :   .vspdData.Text  =   s_C_Qty_Day_14
                    .vspdData.Col = C_Qty_Day_15         :   .vspdData.Text  =   s_C_Qty_Day_15
                    .vspdData.Col = C_Qty_Day_16         :   .vspdData.Text  =   s_C_Qty_Day_16
                    .vspdData.Col = C_Qty_Day_17         :   .vspdData.Text  =   s_C_Qty_Day_17
                    .vspdData.Col = C_Qty_Day_18         :   .vspdData.Text  =   s_C_Qty_Day_18
                    .vspdData.Col = C_Qty_Day_19         :   .vspdData.Text  =   s_C_Qty_Day_19
                    
                    .vspdData.Col = C_Qty_Day_20         :   .vspdData.Text  =   s_C_Qty_Day_20
                    .vspdData.Col = C_Qty_Day_21         :   .vspdData.Text  =   s_C_Qty_Day_21
                    .vspdData.Col = C_Qty_Day_22         :   .vspdData.Text  =   s_C_Qty_Day_22
                    .vspdData.Col = C_Qty_Day_23         :   .vspdData.Text  =   s_C_Qty_Day_23
                    .vspdData.Col = C_Qty_Day_24         :   .vspdData.Text  =   s_C_Qty_Day_24
                    .vspdData.Col = C_Qty_Day_25         :   .vspdData.Text  =   s_C_Qty_Day_25
                    .vspdData.Col = C_Qty_Day_26         :   .vspdData.Text  =   s_C_Qty_Day_26
                    .vspdData.Col = C_Qty_Day_27         :   .vspdData.Text  =   s_C_Qty_Day_27
                    .vspdData.Col = C_Qty_Day_28         :   .vspdData.Text  =   s_C_Qty_Day_28
                    .vspdData.Col = C_Qty_Day_29         :   .vspdData.Text  =   s_C_Qty_Day_29
                    
                    .vspdData.Col = C_Qty_Day_30         :   .vspdData.Text  =   s_C_Qty_Day_30
                    .vspdData.Col = C_Qty_Day_31         :   .vspdData.Text  =   s_C_Qty_Day_31
                    .vspdData.Col = C_Qty_Day_32         :   .vspdData.Text  =   s_C_Qty_Day_32
                    .vspdData.Col = C_Qty_Day_33         :   .vspdData.Text  =   s_C_Qty_Day_33
                    .vspdData.Col = C_Qty_Day_34         :   .vspdData.Text  =   s_C_Qty_Day_34
                    .vspdData.Col = C_Qty_Day_35         :   .vspdData.Text  =   s_C_Qty_Day_35
                    .vspdData.Col = C_Qty_Day_36         :   .vspdData.Text  =   s_C_Qty_Day_36
                    .vspdData.Col = C_Qty_Day_37         :   .vspdData.Text  =   s_C_Qty_Day_37
                    .vspdData.Col = C_Qty_Day_38         :   .vspdData.Text  =   s_C_Qty_Day_38
                    .vspdData.Col = C_Qty_Day_39         :   .vspdData.Text  =   s_C_Qty_Day_39
                    
                    ggoSpread.UpdateRow lRow + 1
                    
			      
        Next

        .vspdData.Redraw = True
        .vspdData.focus

    End With
    
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
                                <TD CLASS="TD5" NOWRAP>생산계획시작일</TD>
                                <TD CLASS="TD6" NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtYYYYMM CLASSID=<%=gCLSIDFPDT%> ALT="기준일자" tag="12X1" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>&nbsp;
  											</TD>
											<TD>
                                                <INPUT TYPE=TEXT  NAME="txtPeriod" SIZE=2 MAXLENGTH=2 tag="14" ALT = "생산계획기간">일
											</TD>
							            </TR>
					               </TABLE>
                                </TD>
<!--                                <TD CLASS="TD6" NOWRAP>
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtYYYYMM CLASSID=<%=gCLSIDFPDT%> ALT="기준일자" tag="12X1" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>
                                </TD> -->
                                <TD CLASS="TD5" NOWRAP>품목</TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>
                            </TR>
                            <TR>
                                <TD CLASS="TD5" NOWRAP>공장</TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="12XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=29 MAXLENGTH=40 tag="14"></TD>
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
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD  HEIGHT=3></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR >
					<TD>&nbsp;</TD>
					<TD><!--<BUTTON NAME="btnRec" CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()" >MES정보수신</BUTTON> -->
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd()" >수주계획일괄복사</BUTTON>
					<TD>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd"   tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo"   tag="24">
<INPUT TYPE=HIDDEN NAME="txtTrackingNo"   tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>