<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Mold Resources
*  2. Function Name        : 금형별실적조회 
*  3. Program ID           : P6310MA
*  4. Program Name         :
*  5. Program Desc         : 금형별실적조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2005/02/21
*  8. Modified date(Last)  : 2005/02/21
*  9. Modifier (First)     : Lee Sang-Ho
* 10. Modifier (Last)      : Lee Sang-Ho
* 11. Comment              : Who Let the dog out?
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "P6310MB1.asp"                                      'Biz Logic ASP
Const C_SHEETMAXROWS    = 100	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
<%'========================================================================================================%>
Dim lsConcd
Dim IsOpenPop

Dim gSelframeFlg			   ' 현재 TAB의 위치를 나타내는 Flag
Dim gCounts
Dim isFirst   '첫화면이 열리는지 여부 
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgPageNo_C
Dim lgOldRow_A
Dim lgOldRow_B
Dim lgOldRow_C


Dim C_CAST_CD
Dim C_CAST_NM
Dim C_WC_CD
Dim C_WC_NM
Dim C_CAR_KIND_CD
Dim C_CAR_KIND_NM
Dim C_ITEM_CD_1
Dim C_ITEM_NM_1
Dim C_CUR_ACCNT

Dim C_PROD_DT
Dim C_PROD_ORDER_NO2
Dim C_OPR_NO
Dim C_RESULT_SEQ
Dim C_ITEM_CD_2
Dim C_ITEM_NM_2
Dim C_CAST_QTY
Dim C_REPORT_TYPE
Dim C_RESULT_QTY
Dim C_UNIT

'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<%'========================================================================================================%>

Dim iDBSYSDate
Dim EndDate, StartDate

	'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
	EndDate = "<%=GetSvrDate%>"
	'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
	StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
	EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize the position
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)

    If pvSpdNo = "A" Then
		C_CAST_CD			=  1
		C_CAST_NM			=  2
		C_WC_CD				=  3
		C_WC_NM				=  4
		C_CAR_KIND_CD		=  5
		C_CAR_KIND_NM		=  6
		C_ITEM_CD_1			=  7
		C_ITEM_NM_1			=  8
		C_CUR_ACCNT			=  9
		
    ElseIf pvSpdNo = "B" Then
		C_PROD_DT				=  1
		C_PROD_ORDER_NO2		=  2
		C_OPR_NO				=  3
		C_RESULT_SEQ			=  4
		C_ITEM_CD_2				=  5
		C_ITEM_NM_2				=  6
		C_CAST_QTY				=  7
		C_REPORT_TYPE			=  8
		C_RESULT_QTY			=  9
		C_UNIT					=  10
    End If

End Sub

'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

	lgIntFlgMode      = parent.OPMD_CMODE						'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKey1	  = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow_A = 0
	lgOldRow_B = 0
	lgPageNo_A = 0
	lgPageNo_B = 0

End Sub

'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtProd_Dt_Fr.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtProd_Dt_Fr.Year = strYear 		 '년월일 default value setting
	frm1.txtProd_Dt_Fr.Month = strMonth 
	frm1.txtProd_Dt_Fr.Day = "01"
	
	frm1.txtProd_Dt_To.Year = strYear 		 '년월일 default value setting
	frm1.txtProd_Dt_To.Month = strMonth 
	frm1.txtProd_Dt_To.Day = strDay

End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check

End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)

End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	If pvSpdNo = "" OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")

		With frm1.vspdData

			    ggoSpread.Source = frm1.vspdData
			    ggoSpread.Spreadinit "V20051217",,parent.gAllowDragDropSpread

			    .ReDraw = false

			    .MaxCols = C_CUR_ACCNT + 1                                                <%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			    .ColHidden = True

			    .MaxRows = 0
			    ggoSpread.ClearSpreadData

				Call AppendNumberPlace("6","2","0")
				Call GetSpreadColumnPos("A")

				ggoSpread.SSSetEdit		C_CAST_CD,				"금형코드",	    18,,,20,2
				ggoSpread.SSSetEdit		C_CAST_NM,				"금형명", 		20,,,30,2
				ggoSpread.SSSetEdit		C_WC_CD,				"작업장", 	    10,,,7,2
				ggoSpread.SSSetEdit		C_WC_NM,				"작업장명", 	15,,,15,2
				ggoSpread.SSSetEdit		C_CAR_KIND_CD,			"적용모델", 	18,,,20,2
				ggoSpread.SSSetEdit		C_CAR_KIND_NM,			"적용모델명", 	30,,,30,2
				ggoSpread.SSSetEdit		C_ITEM_CD_1,			"품목", 	    18,,,20,2
				ggoSpread.SSSetEdit		C_ITEM_NM_1,			"품목명", 	    30,,,30,2
				ggoSpread.SSSetEdit		C_CUR_ACCNT,		    "현재타수", 	10,,,10,2
				
				Call ggoSpread.SSSetColHidden(C_ITEM_CD_1,  C_ITEM_NM_1	, True)
				Call ggoSpread.SSSetColHidden(.MaxCols	,  .MaxCols	, True)

				.ReDraw = true

				Call SetSpreadLock
		End With

	End if

    If pvSpdNo = "" OR pvSpdNo = "B" Then
		
		Call AppendNumberPlace("6","11","0")
		
		Call initSpreadPosVariables("B")
		
		With frm1.vspdData1
			
		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.Spreadinit "V20041010",,parent.gAllowDragDropSpread

		    .ReDraw = false
		    .MaxCols = C_UNIT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True

		    .MaxRows = 0

			Call AppendNumberPlace("6","2","0")
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetDate		C_PROD_DT,				 "실적일",			12,2,gDateFormat
			ggoSpread.SSSetEdit		C_PROD_ORDER_NO2,		 "제조오더번호",	    20
			ggoSpread.SSSetEdit		C_OPR_NO,				 "공정",				5
			ggoSpread.SSSetEdit		C_RESULT_SEQ,			 "실적순번",			5
			ggoSpread.SSSetEdit		C_ITEM_CD_2,			"품목", 	    15
			ggoSpread.SSSetEdit		C_ITEM_NM_2,			"품목명", 	    18
			ggoSpread.SSSetEdit		C_REPORT_TYPE,			 "양/불",				8		
			ggoSpread.SSSetFloat	C_CAST_QTY,				 "반영수량",			10, "6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RESULT_QTY,			 "실적수량",			10,parent.ggQtyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_UNIT,					 "단위",				8
			.ReDraw = true

		Call SetSpreadLock1

		End With
    End if
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()


        ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()


End Sub

'======================================================================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock1()

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If

       Next

    End If
End Sub

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
				C_CAST_CD			= iCurColumnPos(1)
				C_CAST_NM			= iCurColumnPos(2)
				C_WC_CD				= iCurColumnPos(3)
				C_WC_NM				= iCurColumnPos(4)
				C_CAR_KIND_CD      	= iCurColumnPos(5)
				C_CAR_KIND_NM      	= iCurColumnPos(6)
				C_ITEM_CD_1        	= iCurColumnPos(7)
				C_ITEM_NM_1       	= iCurColumnPos(8)
				C_CUR_ACCNT       	= iCurColumnPos(9)
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_PROD_DT           = iCurColumnPos(1)
				C_PROD_ORDER_NO2		= iCurColumnPos(2)
				C_OPR_NO			=   iCurColumnPos(3)
				C_RESULT_SEQ			= iCurColumnPos(4)
				C_ITEM_CD_2        	= iCurColumnPos(5)
				C_ITEM_NM_2       	= iCurColumnPos(6)
				C_CAST_QTY			= iCurColumnPos(7)
				C_REPORT_TYPE			= iCurColumnPos(8)
				C_RESULT_QTY			= iCurColumnPos(9)
				C_UNIT					= iCurColumnPos(10)
	End Select
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Err.Clear                                                                       '☜: Clear err status

	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

	Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

	Call InitSpreadSheet("")                                                            'Setup the Spread sheet
	Call InitVariables                                                              'Initializes local global variables
	Call SetDefaultVal
	Call SetToolbar("1100000000011111")										        '버튼 툴바 제어 
	gCounts = 0
	isFirst = true
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtCastCd.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()

    Dim IntRetCD
    Dim ChgOK

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ChgOK = false

	Call InitVariables                                                           '⊙: Initializes local global variables

    gCounts = 0
    isFirst = true
    lgCurrentSpd = "M"  ' Master

	Call  DisableToolBar( parent.TBC_QUERY)
	
	ggoSpread.Source       = Frm1.vspdData
	ggoSpread.ClearSpreadData	
	
	ggoSpread.Source       = Frm1.vspdData1
	ggoSpread.ClearSpreadData

	If ValidDateCheck(frm1.txtProd_Dt_Fr, frm1.txtProd_Dt_To) = False Then Exit Function
		
	Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	
	IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
		Call DisplayMsgBox("971012", "X", "공장코드", "X")
		frm1.txtPlantCd.focus
		frm1.txtPlantNm.value = ""
		Exit Function
	ELSE
	   frm1.txtPlantNm.value = left(lgF0, len(lgF0) -1)
	END IF

	IF frm1.txtProd_Item_Cd.value <> "" THEN
		Call  CommonQueryRs(" ITEM_NM "," B_ITEM "," ITEM_CD = '" & frm1.txtProd_Item_Cd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "생산품목", "X")
			frm1.txtProd_Item_Cd.focus
			Set gActiveElement = document.activeElement
			frm1.txtProd_Item_Nm.value = ""
			Exit Function
		ELSE
		   frm1.txtProd_Item_Nm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtProd_Item_Nm.value = ""		
	END IF

	IF frm1.txtCastCd.value <> "" THEN
		Call  CommonQueryRs(" cast_nm "," y_cast "," SET_PLANT = '" & frm1.txtPlantCd.value & "' AND cast_cd = '" & frm1.txtCastCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "금형코드", "X")
			frm1.txtCastCd.focus
			Set gActiveElement = document.ActiveElement
			frm1.txtCastNm.value = ""
			Exit Function
		ELSE
			frm1.txtCastNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtCastNm.value = ""
	END IF	

  If Not chkField(Document, "1") Then									         '☜: This function check required field
     Exit Function
  End If
    
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()

End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()

End Function

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
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================================
Function FncFind()
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")
		Case "vaSpread1"
			Call InitSpreadSheet("B")
		Case "vaSpread2"
	End Select
	Call ggoSpread.ReOrderingSpreadData()
	 
End Sub
'========================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
    Dim IntRetCD

	FncExit = False

     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

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

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="			& parent.UID_M0001
        strVal = strVal     & "&lgCurrentSpd="		& lgCurrentSpd                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="		& lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtPlantCd="		& Frm1.txtPlantCd.Value              '☜: Query Key
	    strVal = strVal     & "&txtCast_Cd="		& Frm1.txtCastCd.Value              '☜: Query Key
	    strVal = strVal     & "&txtProd_Dt_Fr="		& Frm1.txtProd_Dt_Fr.Text			'☜: Query Key
	    strVal = strVal     & "&txtProd_Dt_To="		& Frm1.txtProd_Dt_To.Text			'☜: Query Key
	    strVal = strVal     & "&txtProd_Item_Cd="	& Frm1.txtProd_Item_Cd.Value		'☜: Query Key
        strVal = strVal     & "&txtMaxRows="		& .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey						'☜: Next key tag
		strVal = strVal     & "&lgPageNo_A="		& lgPageNo_A                        '☜: Next key tag
		strVal = strVal     & "&txtType="			& "A"								'☜: Next key tag
    End With

	Call RunMyBizASP(MyBizASP, strVal)													'☜: Run Biz Logic
    DbQuery = True
End Function

'========================================================================================================
' Name : DbDtlQuery1
' Desc : This function is called by FncQuery
'========================================================================================================

Function DbDtlQuery1()

    DbDtlQuery1 = False

    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="			& parent.UID_M0001
        strVal = strVal     & "&lgCurrentSpd="		& "S1"								'☜: Next key tag
        strVal = strVal     & "&txtKeyStream="		& lgKeyStream                       '☜: Query Key
	    strVal = strVal     & "&txtCastCd1="		& Frm1.hCastCd.value              '☜: Query Key
	    strVal = strVal     & "&txtProd_Dt_Fr="		& Frm1.txtProd_Dt_Fr.Text              '☜: Query Key
	    strVal = strVal     & "&txtProd_Dt_To="		& Frm1.txtProd_Dt_To.Text              '☜: Query Key
        strVal = strVal     & "&txtMaxRows="		& .vspdData1.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" 		& lgStrPrevKey1						'☜: Next key tag
		strVal = strVal     & "&lgPageNo_B="		& lgPageNo_B                        '☜: Next key tag
		strVal = strVal     & "&txtType="			& "B"								'☜: Next key tag
    End With

	Call RunMyBizASP(MyBizASP, strVal)													'☜: Run Biz Logic

    DbDtlQuery1 = True
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
                                                     '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgOldRow_A = 0
	lgOldRow_B = 0
    lgIntFlgMode =  parent.OPMD_UMODE
    Call  ggoOper.LockField(Document, "Q")													'⊙: Lock field
     

    Call SetToolbar("1100000000011111")

	isFirst = false		' 첫화면이 열리고나서 오른쪽 그리드 세팅하기 위해 
	Call DisableToolBar(parent.TBC_QUERY)

	Call vspdData_click(1,frm1.vspdData.activerow)

	frm1.vspdData.focus

End Function

'========================================================================================================
' Function Name : DbDtlQueryOk1
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbDtlQueryOk1()
    lgIntFlgMode =  parent.OPMD_UMODE

	 
    Call SetToolbar("1100000000011111")

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
'	frm1.vspdData1.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

 '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True   Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_Plant"


	arrParam(2) = Trim(frm1.txtPlantCd.Value)

	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_CD"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		
		frm1.txtPlantCd.focus
		
		Exit Function
	Else
		frm1.txtPlantCd.Value  = arrRet(0)
		frm1.txtPlantNm.Value  = arrRet(1)
		frm1.txtPlantCd.focus
	End If
End Function


'------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenItem()
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	IsOpenPop = True

	arrParam(0) = "품목팝업"
	arrParam(1) = "B_Item_By_Plant,	B_Item"
	arrParam(2) = Trim(frm1.txtProd_Item_Cd.Value)
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = 'N' "

	arrParam(5) = "품목"

	arrField(0) = "B_Item_By_Plant.Item_Cd"
	arrField(1) = "B_Item.Item_NM"


	iCalledAspName = AskPRAspName("m1111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItem(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProd_Item_Cd.focus
End Function

'------------------------------------------ SetItem()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetItem(byval arrRet)
	frm1.txtProd_Item_Cd.Value    = arrRet(0)		
	frm1.txtProd_Item_Nm.Value    = arrRet(1)	
End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim flagTxt
	Call SetPopupMenuItemInf("1100000000")

	IF lgBlnFlgChgValue = False and frm1.vspdData.Maxrows = 0 then
		Call SetToolbar("1100000000011111")
	End if

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	ggoSpread.Source = frm1.vspdData

	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
	    Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData

		If lgSortKey = 1 Then
		    ggoSpread.SSSort Col               'Sort in ascending
		    lgSortKey = 2
		Else
		    ggoSpread.SSSort Col, lgSortKey    'Sort in descending
		    lgSortKey = 1
		End If

		Exit Sub
	End If

	lgCurrentSpd = "M"
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""

	If lgOldRow_A <> Row Then
		Call  DisableToolBar( parent.TBC_QUERY)
		ggoSpread.Source = frm1.vspdData
		frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_CAST_CD
		frm1.hCastCd.Value = frm1.vspdData.Value
	    ggoSpread.Source       = Frm1.vspdData1
	    ggoSpread.ClearSpreadData
		lgPageNo_B = 0
		
		Call DbDtlQuery1
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           Call DbQuery
	    End If
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           Call DbDtlQuery1
	    End If
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

    If Row <= 0 Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

    If Row <= 0 Then
        Exit Sub
    End If

    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)
       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
     End If
End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData
		 ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			    Case C_Req_Dept_POP
			    	.Col = Col - 1
			    	.Row = Row
                	Call OpenCode(.text, C_Req_Dept_POP, Row)
			    Case C_Insp_Dept_POP
			    	.Col = Col - 1
			    	.Row = Row
                	Call OpenCode(.text, C_Insp_Dept_POP, Row)
			    Case C_CASTPop
			    	.Col = Col - 1
			    	.Row = Row
                	Call OpenFacility_Popup("2")
			End Select
		End If

	End With
End Sub


'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData1
		 ggoSpread.Source = frm1.vspdData1
		If Row > 0 Then
			Select Case Col
			    Case C_Sury_Assy_Pop
					.Col = C_Sury_Assy
					.Row = Row
					Call OpenSItem(.text)
			    End Select
		End If

	End With
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

    iColumnLimit  = 5

    If  gMouseClickStatus = "SPCR" Then
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0  :	iRet =  DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
          Exit Function
       End If

       Frm1.vspdData.ScrollBars =  parent.SS_SCROLLBAR_NONE

        ggoSpread.Source = Frm1.vspdData

        ggoSpread.SSSetSplit(ACol)

       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow

       Frm1.vspdData.Action = 0

       Frm1.vspdData.ScrollBars =  parent.SS_SCROLLBAR_BOTH
    End If

    If  gMouseClickStatus = "SP1CR" Then
       ACol = Frm1.vspdData1.ActiveCol
       ARow = Frm1.vspdData1.ActiveRow

       If ACol > iColumnLimit Then
          Frm1.vspdData1.Col = iColumnLimit : Frm1.vspdData1.Row = 0  :	iRet =  DisplayMsgBox("900030", "X", Trim(frm1.vspdData1.Text), "X")
          Exit Function
       End If

       Frm1.vspdData1.ScrollBars =  parent.SS_SCROLLBAR_NONE

        ggoSpread.Source = Frm1.vspdData1

        ggoSpread.SSSetSplit(ACol)

       Frm1.vspdData1.Col = ACol
       Frm1.vspdData1.Row = ARow


       Frm1.vspdData1.Action = 0

       Frm1.vspdData1.ScrollBars =  parent.SS_SCROLLBAR_BOTH
    End If
    
    If  gMouseClickStatus = "SP2CR" Then
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit Then
          Frm1.vspdData2.Col = iColumnLimit : Frm1.vspdData2.Row = 0  :	iRet =  DisplayMsgBox("900030", "X", Trim(frm1.vspdData2.Text), "X")
          Exit Function
       End If

       Frm1.vspdData2.ScrollBars =  parent.SS_SCROLLBAR_NONE

        ggoSpread.Source = Frm1.vspdData2

        ggoSpread.SSSetSplit(ACol)

       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow

       Frm1.vspdData2.Action = 0

       Frm1.vspdData2.ScrollBars =  parent.SS_SCROLLBAR_BOTH
    End If
 End Function

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
	lgActiveSpd      = "M"
	lgCurrentSpd	="M"
End Sub
'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S1"
	lgCurrentSpd	="S1"
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           Call DbQuery
	    End If
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           Call DbDtlQuery1
	    End If
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub



'==========================================================================================
'   Event Name : txtProd_dt_Fr
'   Event Desc :
'==========================================================================================

 Sub txtProd_Dt_Fr_DblClick(Button)
	if Button = 1 then
		frm1.txtProd_dt_Fr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtProd_dt_Fr.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtProd_dt_To
'   Event Desc :
'==========================================================================================

 Sub txtProd_dt_To_DblClick(Button)
	if Button = 1 then
		frm1.txtProd_dt_To.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtProd_dt_To.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================

Sub txtProd_dt_Fr_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================

Sub txtProd_dt_To_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub




'------------------------------------------  OpenCast()  ------------------------------------------------
'	Name : OpenCast()
'	Description : Cast PopUp
'------------------------------------------------------------------------------------------------------------
Function OpenCast()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	IF frm1.txtPlantCd.value <> "" THEN
		Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			frm1.txtPlantNm.value = ""
			IsOpenPop = False
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.ActiveElement
			Exit Function
		ELSE
			frm1.txtPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtPlantNm.value = ""
		IsOpenPop = False
		Call DisplayMsgBox("971012", "X", "공장코드", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.ActiveElement
		Exit Function
	END IF 

		arrParam(0) = "금형코드"								' 팝업 명칭 
		arrParam(1) = "Y_CAST"											' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtCastCd.Value)		' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = "SET_PLANT = " & FilterVar(frm1.txtPlantCd.value, "''", "S")								' Where Condition
		arrParam(5) = "금형코드"								' TextBox 명칭 

    arrField(0) = "ED15" & parent.gcolsep & "CAST_CD"							' Field명(0)
    arrField(1) = "ED15" & parent.gcolsep & "CAST_NM"							' Field명(1)
    arrField(2) = "ED20" & parent.gcolsep & "(SELECT ITEM_GROUP_NM FROM B_ITEM_GROUP WHERE ITEM_GROUP_CD = CAR_KIND )"						' Field명(2)
    arrField(3) = "ED20" & parent.gcolsep & "(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = ITEM_CD_1 )"						' Field명(3)
    arrField(4) = "F3"   & parent.gcolsep & "EXT1_QTY"						' Field명(4)

    arrHeader(0) = "금형코드"					' Header명(0)
    arrHeader(1) = "금형코드명"					' Header명(1)
    arrHeader(2) = "모델명"						' Header명(2)
    arrHeader(3) = "품목명"						' Header명(3)
    arrHeader(4) = "차수"						' Header명(4)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCast(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtCastCd.focus
End Function

'------------------------------------------ SetCast()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast(byval arrRet)
	frm1.txtCastCd.Value    = arrRet(0)		
	frm1.txtCastNm.Value    = arrRet(1)	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>금형별실적조회</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
										<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X">
									</TD>
									<TD CLASS="TD5" NOWRAP>금형코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ID=txtCastCd NAME="txtCastCd" ALT="금형코드" TYPE="Text" SiZE="18" MAXLENGTH="18" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCast()">
															<INPUT ID=txtCastNm NAME="txtCastNm" ALT="금형코드명" TYPE="Text" SiZE="25" MAXLENGTH="40" tag="14XXXU"></TD>
								<TR>
									<TD CLASS="TD5" NOWRAP>생산품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProd_Item_Cd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="생산품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem_cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtProd_Item_Nm" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP>실적일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/p6310ma1_OBJECT1_txtProd_Dt_Fr.js'></script>
												</td>
												<td>&nbsp;~&nbsp;</td>
												<td>
													<script language =javascript src='./js/p6310ma1_OBJECT2_txtProd_Dt_To.js'></script>
												</td>
											<tr>
										</table>

									</TD>
								<TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD WIDTH="30%">
									<script language =javascript src='./js/p6310ma1_vaSpread_vspdData.js'></script>
								</TD>
								<TD WIDTH="60%">
									<script language =javascript src='./js/p6310ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
						</Table>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrder_No" tag="24">
<INPUT TYPE=HIDDEN NAME="hWc_Cd" tag="24">
<INPUT TYPE=HIDDEN NAME="hCastCd" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
