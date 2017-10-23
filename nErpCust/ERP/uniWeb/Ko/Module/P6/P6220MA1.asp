<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 금형관리 
*  2. Function Name        : 금형점검내용등록 
*  3. Program ID           : P6220MA1
*  4. Program Name         :
*  5. Program Desc         : 금형점검내용등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2005/01/17
*  8. Modified date(Last)  : 2005/01/17
*  9. Modifier (First)     : Lee Sang-Ho
* 10. Modifier (Last)      : Lee Sang-Je
* 11. Comment              : 
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
Const BIZ_PGM_ID = "P6220mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_QUERY_ID = "P6220mb2.asp"
Const BIZ_PGM_QUERY2_ID = "P6220mb3.asp"
Const BIZ_PGM_SAVE_ID = "P6220mb4.asp"
Const C_SHEETMAXROWS    = 100	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
'========================================================================================================
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
Dim C_SET_PLANT
Dim C_SET_PLANT_NM
Dim C_CAR_KIND_Nm
Dim C_Plan_Dt
Dim C_Insp_Text
Dim C_Insp_Hour
Dim C_Insp_Min
Dim C_Req_Dept
Dim C_Req_Dept_POP
Dim C_Req_Dept_Nm
Dim C_Insp_Dept
Dim C_Insp_Dept_POP
Dim C_Insp_Dept_Nm
Dim C_Insp_Emp_Qty
Dim C_Payroll
Dim C_Matl_Cost
Dim C_Insp_Flag
Dim C_INSP_PRID

Dim C_Seq
Dim C_Zinsp_PartCd
Dim C_Zinsp_PartNm
Dim C_Insp_PartCd
Dim C_Insp_PartNm
Dim C_Insp_MethCd
Dim C_Insp_MethNm
Dim C_Insp_DeciCd
Dim C_Insp_DeciNm
Dim C_St_GoGubunCd
Dim C_St_GoGubunNm
Dim C_Sury_Assy
Dim C_Sury_Assy_Pop
Dim C_Sury_Assy_Nm
Dim C_S_Qty
Dim C_Price
Dim C_Sury_Amt
Dim C_Sury_Type
Dim C_Sury_Type_Nm

Dim C_Insp_Emp_Gb
Dim C_Insp_Emp_GbNm
Dim C_Insp_Emp_Cd
Dim C_Insp_Emp_Pop
Dim C_Insp_Emp_Nm
Dim C_Cust_Cd
Dim C_Cust_Pop
Dim C_Cust_Nm
Dim C_Insp_Hour2
Dim C_Insp_Min2
Dim C_Payroll2



'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================


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
		C_CAST_CD		 	= 1
		C_CAST_NM			= 2
		C_SET_PLANT        	= 3
		C_SET_PLANT_NM  	= 4
		C_CAR_KIND_NM		= 5
		C_Plan_Dt          	= 6
		C_Insp_Text        	= 7
		C_Insp_Hour        	= 8
		C_Insp_Min         	= 9
		C_Req_Dept         	= 10
		C_Req_Dept_POP     	= 11
		C_Req_Dept_Nm      	= 12
		C_Insp_Dept        	= 13
		C_Insp_Dept_POP    	= 14
		C_Insp_Dept_Nm     	= 15
		C_Insp_Emp_Qty     	= 16
		C_Payroll          	= 17
        C_Matl_Cost        	= 18
        C_Insp_Flag        	= 19
        C_INSP_PRID      	= 20
        
    ElseIf pvSpdNo = "B" Then
		C_Seq            	= 1
		C_Zinsp_PartCd		= 2 
		C_Zinsp_PartNm		= 3 
		C_Insp_PartCd		= 4 
		C_Insp_PartNm		= 5 
		C_Insp_MethCd		= 6 
		C_Insp_MethNm		= 7 
		C_Insp_DeciCd	= 8 
		C_Insp_DeciNm	= 9 
		C_St_GoGubunCd		= 10
		C_St_GoGubunNm		= 11
		C_Sury_Assy      	= 12
		C_Sury_Assy_Pop  	= 13
		C_Sury_Assy_Nm   	= 14
		C_S_Qty          	= 15
		C_Price          	= 16
		C_Sury_Amt       	= 17
		C_Sury_Type      	= 18
		C_Sury_Type_Nm      = 19
		
    ElseIf pvSpdNo = "C" Then
		C_Seq            	= 1 
		C_Insp_Emp_Gb    	= 2 
		C_Insp_Emp_GbNm  	= 3 
		C_Insp_Emp_Cd    	= 4 
		C_Insp_Emp_Pop   	= 5 
		C_Insp_Emp_Nm    	= 6 
		C_Cust_Cd        	= 7 
		C_Cust_Pop       	= 8 
		C_Cust_Nm        	= 9 
		C_Insp_Hour2     	= 10
		C_Insp_Min2      	= 11
		C_Payroll2       	= 12
    End If

End Sub

'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKey1	  = ""                                      '⊙: initializes Previous Key Index
    lgStrPrevKey2	  = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow_A = 0
	lgOldRow_B = 0
	lgPageNo_A = 0
	lgPageNo_B = 0
	lgPageNo_C = 0

End Sub

'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtWork_Dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtWork_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtWork_dt.Month = strMonth 
	frm1.txtWork_dt.Day = "01"

End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
<%'========================================================================================================%>
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    If lgCurrentSpd = "M" Then
       lgKeyStream = UNIConvDate(Trim(Frm1.txtWork_Dt.text)) & parent.gColSep
       lgKeyStream = lgKeyStream & Frm1.txtPlantCd.value & parent.gColSep
       lgKeyStream = lgKeyStream & Frm1.txtCarKind.value & parent.gColSep
       lgKeyStream = lgKeyStream & Frm1.txtCastCd.value & parent.gColSep
    Else

    	frm1.vspdData.Row = pRow
		frm1.vspdData.Col = C_CAST_CD
        lgKeyStream = frm1.vspdData.Text & parent.gColSep     'You Must append one character( parent.gColSep)
		frm1.vspdData.Col = C_Plan_Dt
        lgKeyStream = lgKeyStream & UNIConvDate(Trim(frm1.vspdData.Text)) & parent.gColSep     'You Must append one character( parent.gColSep)
    End If
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
  Dim iCodeArr
  Dim iNameArr

  ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_INSP_FLAG

End Sub

Sub InitComboBox1()
  Dim iCodeArr
  Dim iNameArr

	ggoSpread.Source = frm1.vspdData1

	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z425' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Zinsp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Zinsp_PartNm
	
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z411' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Insp_PartCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Insp_PartNm
	
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z412' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Insp_MethCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Insp_MethNm
	
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z418' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Insp_DeciCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Insp_DeciNm
	
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Z419' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_St_GoGubunCd
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_St_GoGubunNm
	
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'Y6003' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Sury_Type
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Sury_Type_Nm

End Sub

Sub InitComboBox2()
  Dim iCodeArr
  Dim iNameArr
  
	ggoSpread.Source = frm1.vspdData2
	Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = 'P1003' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	iCodeArr = lgF0
	iNameArr = lgF1
	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_Insp_Emp_Gb
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_Insp_Emp_GbNm

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex

	     ggoSpread.Source = frm1.vspdData1
	    With frm1.vspdData1
	    	For intRow = 1 To .MaxRows
	    		.Row = intRow

' 	    		.Col = C_ALLOW_CD         ' 수당코드 
' 	    		intIndex = .value
' 	    		.col = C_ALLOW_CD_NM
' 	    		.value = intindex

	    	Next
	    End With
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
			    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

			    .ReDraw = false

			    .MaxCols = C_INSP_PRID + 1                                                <%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			    .ColHidden = True

			    .MaxRows = 0
			    ggoSpread.ClearSpreadData
				
				Call GetSpreadColumnPos("A")

				ggoSpread.SSSetEdit		C_CAST_CD,				"금형코드",	    15,,,18,2
				ggoSpread.SSSetEdit		C_CAST_NM,				"금형코드명", 		20,,,40,2
				ggoSpread.SSSetEdit		C_SET_PLANT,  			"설치공장",   10,,,10,2
				ggoSpread.SSSetEdit		C_SET_PLANT_NM,			"공장명", 15,,,20,2
				ggoSpread.SSSetEdit 	C_CAR_KIND_NM,			"적용모델", 		15,,,20,2
				ggoSpread.SSSetDate   	C_Plan_Dt, 				"작업일자", 12,2,gDateFormat
				ggoSpread.SSSetEdit		C_Insp_Text,			"점검/수리내용",	    15,,,40,1
				ggoSpread.SSSetFloat	C_Insp_Hour,			"소요시간", 11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","9999"
				ggoSpread.SSSetFloat	C_Insp_Min,				"소요분"  , 11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
				ggoSpread.SSSetEdit		C_Req_Dept,				"의뢰부서",	    15,,,10,2
				ggoSpread.SSSetButton   C_Req_Dept_POP
				ggoSpread.SSSetEdit		C_Req_Dept_Nm,			"의뢰부서명",	    15,,,40,2
				ggoSpread.SSSetEdit		C_Insp_Dept,			"수리부서",	    15,,,10,2
				ggoSpread.SSSetButton   C_Insp_Dept_POP
				ggoSpread.SSSetEdit		C_Insp_Dept_Nm,			"수리부서명",	    15,,,40,2
				ggoSpread.SSSetFloat	C_Insp_Emp_Qty,			"수리인원",     20,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				ggoSpread.SSSetFloat	C_Payroll,				"인건비",     20,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				ggoSpread.SSSetFloat	C_Matl_Cost,			"소모자재비",     20,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				ggoSpread.SSSetCombo 	C_Insp_Flag,			"점검여부",  10, 0, False
				ggoSpread.SSSetFloat 	C_INSP_PRID,			"점검타수",     20,  parent.ggQtyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
				
				Call ggoSpread.MakePairsColumn(C_CAST_CD,  C_CAST_NM)
				Call ggoSpread.MakePairsColumn(C_SET_PLANT	 ,  C_SET_PLANT_NM)
				
				.ReDraw = true
				
				Call SetSpreadLock

		End With

	End if

    If pvSpdNo = "" OR pvSpdNo = "B" Then
		Call initSpreadPosVariables("B")
		With frm1.vspdData1

		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

		    .ReDraw = false
		    .MaxCols = C_Sury_Type_Nm + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True

		    .MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetFloat	C_Seq,				"순서",    8, "7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

			ggoSpread.SSSetCombo 	C_Zinsp_PartCd,		"부위",  10, 0, True
			ggoSpread.SSSetCombo	C_Zinsp_PartNm,		"부위",  10, 0, False
			ggoSpread.SSSetCombo	C_Insp_PartCd,		"점검항목",  10, 0, True
			ggoSpread.SSSetCombo	C_Insp_PartNm,		"점검항목",  10, 0, False
			ggoSpread.SSSetCombo	C_Insp_MethCd,		"점검방법",  10, 0, True
			ggoSpread.SSSetCombo	C_Insp_MethNm,		"점검방법",  10, 0, False
			ggoSpread.SSSetCombo	C_Insp_DeciCd,		"판정기준",  10, 0, True
			ggoSpread.SSSetCombo	C_Insp_DeciNm,		"판정기준",  10, 0, False
			ggoSpread.SSSetCombo	C_St_GoGubunCd,		"운/휴구분",  14, 0, True
			ggoSpread.SSSetCombo	C_St_GoGubunNm,		"운/휴구분",  14, 0, False
			ggoSpread.SSSetEdit		C_Sury_Assy,		"부품코드",	    15,,,18,2
			ggoSpread.SSSetButton   C_Sury_Assy_Pop
			ggoSpread.SSSetEdit		C_Sury_Assy_Nm,		"부품명",	    20,,,20,2
			ggoSpread.SSSetFloat	C_S_Qty,			"수량",     20,  parent.ggQtyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Price,			"단가",			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
			ggoSpread.SSSetFloat	C_Sury_Amt,			"금액",     20,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo 	C_Sury_Type,		"조치유형",   14, 0, True
			ggoSpread.SSSetCombo 	C_Sury_Type_Nm,		"조치유형",   14, 0, False
			
			Call ggoSpread.SSSetColHidden(C_Zinsp_PartCd, C_Zinsp_PartCd, True)
			Call ggoSpread.SSSetColHidden(C_Insp_PartCd, C_Insp_PartCd, True)
			Call ggoSpread.SSSetColHidden(C_Insp_MethCd, C_Insp_MethCd, True)
			Call ggoSpread.SSSetColHidden(C_Insp_DeciCd, C_Insp_DeciCd, True)
			Call ggoSpread.SSSetColHidden(C_St_GoGubunCd, C_St_GoGubunCd, True)
			Call ggoSpread.SSSetColHidden(C_Sury_Type, C_Sury_Type, True)
			.ReDraw = true

		Call SetSpreadLock1

		End With
    End if
    If pvSpdNo = "" OR pvSpdNo = "C" Then

		Call initSpreadPosVariables("C")
		With frm1.vspdData2

		    ggoSpread.Source = frm1.vspdData2
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread

		    .ReDraw = false
		    .MaxCols = C_Payroll2 + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True

		    .MaxRows = 0


			Call GetSpreadColumnPos("C")

			ggoSpread.SSSetFloat	C_Seq,				"순서",    8, "7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetCombo 	C_Insp_Emp_Gb,		"점검/수리구분",  10, 0, True
			ggoSpread.SSSetCombo 	C_Insp_Emp_GbNm,	"점검/수리구분",  15, 0, False
			ggoSpread.SSSetEdit		C_Insp_Emp_Cd,		"수리자",	    10,,,13,2
			ggoSpread.SSSetButton   C_Insp_Emp_Pop
			ggoSpread.SSSetEdit		C_Insp_Emp_Nm,		"이름",	    15,,,20,2
			ggoSpread.SSSetEdit		C_Cust_Cd,			"수리업체",	    15,,,20,2
			ggoSpread.SSSetButton   C_Cust_Pop
			ggoSpread.SSSetEdit		C_Cust_Nm,			"업체명",	    15,,,20,2
			ggoSpread.SSSetFloat	C_Insp_Hour2,		"소요시간", 11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
			ggoSpread.SSSetFloat	C_Insp_Min2,		"소요분", 11,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
			ggoSpread.SSSetFloat	C_Payroll2,			"인건비",     20,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

			Call ggoSpread.SSSetColHidden(C_Insp_Emp_Gb,  C_Insp_Emp_Gb, True)

		.ReDraw = true

		Call SetSpreadLock2

		End With
    End if
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData

		.ReDraw = False
		ggoSpread.SpreadLock			C_CAST_CD				, -1, C_Plan_Dt
		ggoSpread.SpreadLock			C_Insp_Hour			, -1, C_Insp_Hour
		ggoSpread.SpreadLock			C_Insp_Min			, -1, C_Insp_Min
		ggoSpread.SpreadLock			C_Req_Dept_Nm		, -1, C_Req_Dept_Nm
		ggoSpread.SpreadLock			C_Insp_Dept_Nm	, -1, C_Insp_Dept_Nm
		ggoSpread.SpreadLock			C_Payroll				, -1, C_Payroll
		ggoSpread.SpreadLock			C_Matl_Cost			, -1, C_Matl_Cost
		ggoSpread.SpreadLock			C_Insp_Emp_Qty			, -1, C_Insp_Emp_Qty
		ggoSpread.SpreadLock			C_INSP_PRID			, -1, C_INSP_PRID


		ggoSpread.SSSetProtected	.MaxCols				, -1, -1
		.ReDraw = True

	End With

End Sub

Sub SetSpreadLock1()

	With frm1.vspdData1

		ggoSpread.Source = frm1.vspdData1

		.ReDraw = False
		ggoSpread.SpreadLock			C_SEQ						, -1, C_SEQ
		ggoSpread.SpreadLock			C_Sury_Assy_Nm	, -1, C_Sury_Assy_Nm
		ggoSpread.SSSetProtected	.MaxCols				, -1, -1
		ggoSpread.SSSetRequired		C_Zinsp_PartNm	, -1
		ggoSpread.SSSetRequired		C_Insp_PartNm		, -1
		ggoSpread.SSSetRequired		C_Insp_MethNm		, -1
		ggoSpread.SSSetRequired		C_Insp_DeciNm		, -1
		ggoSpread.SSSetRequired		C_St_GoGubunNm	, -1
		.ReDraw = True

	End With

End Sub


Sub SetSpreadLock2()

	With frm1.vspdData2

		ggoSpread.Source = frm1.vspdData2

		.ReDraw = False
		ggoSpread.SpreadLock			C_SEQ						, -1, C_SEQ
		ggoSpread.SpreadLock			C_Cust_Nm				, -1, C_Cust_Nm
		ggoSpread.SpreadLock			C_Insp_Emp_Nm		, -1, C_Insp_Emp_Nm
		ggoSpread.SSSetProtected	.MaxCols				, -1, -1
		ggoSpread.SSSetRequired		C_Insp_Emp_GbNm	, -1
		ggoSpread.SSSetRequired		C_Insp_Emp_Cd		, -1
		ggoSpread.SSSetRequired		C_Cust_Cd				, -1
		.ReDraw = True

	End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData

		.ReDraw = False

		ggoSpread.SSSetProtected	C_CAST_CD	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CAST_NM		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANT				, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_SET_PLANT_NM			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_CAR_KIND_NM			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_Req_Dept_Nm			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_Insp_Dept_Nm			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_Plan_Dt			, pvStartRow, pvEndRow

		.ReDraw = True

	End With

End Sub

Sub SetSpreadColor1(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData1

        ggoSpread.Source = frm1.vspdData1

        .ReDraw = False

			ggoSpread.SSSetProtected	C_Sury_Assy_Nm			, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Seq		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Zinsp_PartNm		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Insp_PartNm		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Insp_MethNm		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Insp_DeciNm		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_St_GoGubunNm		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected						.MaxCols, pvStartRow, pvEndRow
        .ReDraw = True

    End With

End Sub

Sub SetSpreadColor2(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData2

        ggoSpread.Source = frm1.vspdData2

        .ReDraw = False

			ggoSpread.SSSetProtected	C_Insp_Emp_Nm			, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Seq		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Insp_Emp_GbNm		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Insp_Emp_Cd		, pvStartRow, pvEndRow
			ggoSpread.SSSetRequired		C_Cust_Cd		, pvStartRow, pvEndRow
			ggoSpread.SSSetProtected	    C_Cust_Nm       , pvStartRow, pvEndRow

             ggoSpread.SSSetProtected						.MaxCols, pvStartRow, pvEndRow
        .ReDraw = True

    End With

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

				C_CAST_CD      		= iCurColumnPos(1 )
				C_CAST_NM      		= iCurColumnPos(2 )
				C_Set_Plant        	= iCurColumnPos(3 )
				C_Set_Plant_Nm     	= iCurColumnPos(4 )
				C_CAR_KIND_Nm		= iCurColumnPos(5 )
				C_Plan_Dt          	= iCurColumnPos(6 )
				C_Insp_Text        	= iCurColumnPos(7 )
				C_Insp_Hour        	= iCurColumnPos(8 )
				C_Insp_Min         	= iCurColumnPos(9 )
				C_Req_Dept         	= iCurColumnPos(10)
				C_Req_Dept_POP     	= iCurColumnPos(11)
				C_Req_Dept_Nm      	= iCurColumnPos(12)
				C_Insp_Dept        	= iCurColumnPos(13)
				C_Insp_Dept_POP    	= iCurColumnPos(14)
				C_Insp_Dept_Nm     	= iCurColumnPos(15)
				C_Insp_Emp_Qty     	= iCurColumnPos(16)
				C_Payroll          	= iCurColumnPos(17)
				C_Matl_Cost        	= iCurColumnPos(18)
				C_Insp_Flag        	= iCurColumnPos(19)
				C_INSP_PRID      	= iCurColumnPos(20)

       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_Seq            	= iCurColumnPos(1 )
				C_Zinsp_PartCd		= iCurColumnPos(2 )
				C_Zinsp_PartNm		= iCurColumnPos(3 )
				C_Insp_PartCd		= iCurColumnPos(4 )
				C_Insp_PartNm		= iCurColumnPos(5 )
				C_Insp_MethCd		= iCurColumnPos(6 )
				C_Insp_MethNm		= iCurColumnPos(7 )
				C_Insp_DeciCd	= iCurColumnPos(8 )
				C_Insp_DeciNm	= iCurColumnPos(9 )
				C_St_GoGubunCd		= iCurColumnPos(10)
				C_St_GoGubunNm		= iCurColumnPos(11)
				C_Sury_Assy      	= iCurColumnPos(12)
				C_Sury_Assy_Pop  	= iCurColumnPos(13)
				C_Sury_Assy_Nm   	= iCurColumnPos(14)
				C_S_Qty          	= iCurColumnPos(15)
				C_Price          	= iCurColumnPos(16)
				C_Sury_Amt       	= iCurColumnPos(17)
				C_Sury_Type      	= iCurColumnPos(18)
				C_Sury_Type_Nm      	= iCurColumnPos(19)				

       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_Seq            	= iCurColumnPos(1 )
				C_Insp_Emp_Gb    	= iCurColumnPos(2 )
				C_Insp_Emp_GbNm  	= iCurColumnPos(3 )
				C_Insp_Emp_Cd    	= iCurColumnPos(4 )
				C_Insp_Emp_Pop   	= iCurColumnPos(5 )
				C_Insp_Emp_Nm    	= iCurColumnPos(6 )
				C_Cust_Cd        	= iCurColumnPos(7 )
				C_Cust_Pop       	= iCurColumnPos(8 )
				C_Cust_Nm        	= iCurColumnPos(9 )
				C_Insp_Hour2     	= iCurColumnPos(10)
				C_Insp_Min2      	= iCurColumnPos(11)
				C_Payroll2       	= iCurColumnPos(12)

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
	
	Call AppendNumberPlace("6","2","0")
	Call AppendNumberPlace("7","5","0")

    Call InitSpreadSheet("")                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables
	Call SetDefaultVal
    Call InitComboBox
    Call InitComboBox1
    Call InitComboBox2

    Call SetToolbar("1100110100011111")										        '버튼 툴바 제어 
    gCounts = 0
    isFirst = true
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtWork_Dt.focus
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
	Dim iNameArr
	FncQuery = False															 '☜: Processing is NG
	Err.Clear                                                                    '☜: Clear err status

	ChgOK = false

	ggoSpread.Source = Frm1.vspdData2
	If  ggoSpread.SSCheckChange = True Then
		ChgOK = True
	End If

	ggoSpread.Source = Frm1.vspdData1
	If  ggoSpread.SSCheckChange = True Then
		ChgOK = True
	End If

	ggoSpread.Source = Frm1.vspdData
	If  ggoSpread.SSCheckChange = True Then
		ChgOK = True
	End If
		
	If  ChgOK Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")		'☜: Data is changed.  Do you want to display it?
			
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
		
	Call  ggoOper.ClearField(Document, "2")

		
	Call InitVariables                                                           '⊙: Initializes local global variables
	lgCurrentSpd = "M"
			
	Call MakeKeyStream("M")
			
	gCounts = 0
	isFirst = true
		
	lgCurrentSpd = "M"  ' Master
			
	Call  DisableToolBar( parent.TBC_QUERY)
			
	Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
		
	IF frm1.txtPlantCd.value <> "" THEN
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement
			frm1.txtPlantNm.value = ""
			Exit Function
		ELSE
			frm1.txtPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtPlantNm.value = ""
	END IF
				
	IF frm1.txtCarKind.value <> "" THEN
		Call  CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD = '" & frm1.txtCarKind.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "적용모델", "X")
			frm1.txtcarKind.focus
			Set gActiveElement = document.activeElement
			frm1.txtCarKindNm.value = ""
			Exit Function
		ELSE
			frm1.txtCarKindNm.value = left(lgF0, len(lgF0) -1)
		END IF	
	ELSE
		frm1.txtCarKindNm.value = ""
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
    Dim intRetCD

    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD
	dim lRow
	DIM strCD, strNm
	Dim tmpCost
	Dim tmpPayroll
	Dim tmpEmpQty

	FncSave = False                                                              '☜: Processing is NG
	Err.Clear

	frm1.ChgSave1.value = "F"
	frm1.ChgSave2.value = "F"
	frm1.ChgSave3.value = "F"
	
	
  ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
  If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
		ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
		If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
			ggoSpread.Source = frm1.vspdData2							'⊙: Preset spreadsheet pointer 
			If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
			    IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
			    Exit Function
			End If
		End If
  End If
		
	ggoSpread.Source = frm1.vspdData1
	FOR lRow = 1 TO frm1.vspdData1.MaxRows	
		frm1.vspdData1.Row = lRow
		frm1.vspdData1.Col = C_SURY_AMT
		tmpCost = tmpCost + cdbl(frm1.vspdData1.value)
	next

	ggoSpread.Source = frm1.vspdData2
	FOR lRow = 1 TO frm1.vspdData2.MaxRows
		frm1.vspdData2.Row = lRow
		frm1.vspdData2.Col = C_PAYROLL2
		tmpPayroll = tmpPayroll + cdbl(frm1.vspdData2.value)
	next

	tmpEmpQty = frm1.vspdData2.MaxRows
		
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.activerow
	frm1.vspdData.Col = C_INSP_EMP_QTY
	frm1.vspdData.value = UNIConvNum (Trim(tmpEmpQty),0)	
	frm1.vspdData.Col = C_MATL_COST
	frm1.vspdData.value = UNIConvNum (Trim(tmpCost),0)
	frm1.vspdData.Col = C_PAYROLL
	frm1.vspdData.value = UNIConvNum (Trim(tmpPayroll),0)
	ggoSpread.updaterow frm1.vspdData.activerow

	ggoSpread.Source = frm1.vspdData
	If  ggoSpread.SSCheckChange = True Then
		frm1.ChgSave1.value = "T"
	End If

	ggoSpread.Source = Frm1.vspdData1
	If  ggoSpread.SSCheckChange = True Then
		frm1.ChgSave2.value = "T"
	End If

	ggoSpread.Source = Frm1.vspdData2
	If  ggoSpread.SSCheckChange = True Then
		frm1.ChgSave3.value = "T"
	End If

	If frm1.ChgSave1.value = "F" and frm1.ChgSave2.value="F" and frm1.ChgSave3.value="F" Then
		IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
	Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData1
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData2
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
	End If

	lgCurrentSpd = "M"
	Call  DisableToolBar( parent.TBC_SAVE)
	  
	If DbSave = False Then
		Call  RestoreToolBar()
	Exit Function
	End If

	FncSave = True                                                              '☜: Processing is OK

End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "S1"

            If Frm1.vspdData1.MaxRows < 1 Then
                Exit Function
            End If

	        With Frm1.vspdData1

	        	If .ActiveRow > 0 Then
	        		.ReDraw = False

	        		ggoSpread.Source = frm1.vspdData1
	        		ggoSpread.CopyRow
                    SetSpreadColor1 .ActiveRow, .ActiveRow

                    .Col  = C_SEQ
 					.Row  = .ActiveRow
 					.Text = cint(.Text) + 1
					ggoSpread.SpreadLock C_SEQ				, -1, C_SEQ
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With
        Case  "S2"

            If Frm1.vspdData2.MaxRows < 1 Then
                Exit Function
            End If

	        With Frm1.vspdData2

	        	If .ActiveRow > 0 Then
	        		.ReDraw = False

	        		ggoSpread.Source = frm1.vspdData2
	        		ggoSpread.CopyRow
                    SetSpreadColor2 .ActiveRow, .ActiveRow

                    .Col  = C_SEQ
 					.Row  = .ActiveRow
 					.Text = Cint(.Text) + 1
					ggoSpread.SpreadLock C_SEQ				, -1, C_SEQ					
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With

    End Select

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
    if lgCurrentSpd = "M" then
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.EditUndo
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		Call DbDtlQuery1()
    elseif lgCurrentSpd = "S1" then
		ggoSpread.Source = Frm1.vspdData1
		ggoSpread.EditUndo
    else
		ggoSpread.Source = Frm1.vspdData2
		ggoSpread.EditUndo
    end if
    Call initdata()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

    Dim imRow
    Dim iRow
	Dim IntRetCD
	Dim iTemp
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                         '☜: Processing is NG

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
 		  Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
      frm1.txtPlantCd.focus
      Set gActiveElement = document.activeElement
      Exit Function
    End If


    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
		Case	"M"
			With Frm1
				.vspdData1.ReDraw	=	False
				.vspdData1.Focus
				ggoSpread.Source = .vspdData1
				ggoSpread.InsertRow	.vspdData1.ActiveRow,	imRow
				SetSpreadColor1	.vspdData1.ActiveRow,	.vspdData1.ActiveRow + imRow - 1
				iTemp	=	0
				For	iRow =	.vspdData1.ActiveRow to	.vspdData1.ActiveRow + imRow - 1
					iTemp	=	iTemp	+	1
					.vspdData1.Row = iRow
					.vspdData1.Col=	C_Seq
					.vspdData1.Text	=	.vspdData1.Maxrows + iTemp - imRow
				Next
				.vspdData1.ReDraw	=	True
			End	With
        Case  "S1"
              With Frm1
					.vspdData1.ReDraw = False
					.vspdData1.Focus
					ggoSpread.Source = .vspdData1
					ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
					SetSpreadColor1 .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
					iTemp = 0
					For iRow =  .vspdData1.ActiveRow to .vspdData1.ActiveRow + imRow - 1
						iTemp = iTemp + 1
						.vspdData1.Row = iRow
						.vspdData1.Col= C_Seq
						.vspdData1.Text = .vspdData1.Maxrows + iTemp - imRow
					Next
					ggoSpread.SpreadLock C_SEQ				, -1, C_SEQ
					.vspdData1.ReDraw = True
              End With
        Case "S2"
              With Frm1
					.vspdData2.ReDraw = False
					.vspdData2.Focus
					ggoSpread.Source = .vspdData2
					ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
					SetSpreadColor2 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
					iTemp = 0
					For iRow =  .vspdData2.ActiveRow to .vspdData2.ActiveRow + imRow - 1
						iTemp = iTemp + 1
						.vspdData2.Row = iRow
						.vspdData2.Col= C_Seq
						.vspdData2.Text = .vspdData2.Maxrows + iTemp - imRow
					Next
					ggoSpread.SpreadLock C_SEQ				, -1, C_SEQ
					.vspdData2.ReDraw = True
              End With
    End Select

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    if  lgCurrentSpd = "M" then
        If Frm1.vspdData.MaxRows < 1 then
           Exit function
	    End if
        With Frm1.vspdData
        	.focus
        	 ggoSpread.Source = frm1.vspdData
        	lDelRows =  ggoSpread.DeleteRow
        End With
    ELSEif lgCurrentSpd = "S1" then
        If Frm1.vspdData1.MaxRows < 1 then
           Exit function
	    End if
        With Frm1.vspdData1
        	.focus
        	 ggoSpread.Source = frm1.vspdData1
        	lDelRows =  ggoSpread.DeleteRow
        End With
    ELSE
        If Frm1.vspdData2.MaxRows < 1 then
           Exit function
	    End if
        With Frm1.vspdData2
        	.focus
        	 ggoSpread.Source = frm1.vspdData2
        	lDelRows =  ggoSpread.DeleteRow
        End With
    END IF
    Set gActiveElement = document.ActiveElement
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
			Call InitComboBox1
		Case "vaSpread2"
			Call InitSpreadSheet("C")
			Call InitComboBox2
	End Select
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
			strVal = strVal		& "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal     & "&lgCurrentSpd="		& lgCurrentSpd                      '☜: Next key tag
			strVal = strVal     & "&txtKeyStream="		& lgKeyStream                       '☜: Query Key
			strVal = strVal     & "&txtWork_Dt="		& Frm1.txtWork_Dt.text                     '☜: Query Key
			strVal = strVal     & "&txtCastCd="			& Frm1.txtCastCd.value      '☜: Query Key
			strVal = strVal     & "&txtPlantCd="		& Frm1.txtPlantCd.value      '☜: Query Key
			strVal = strVal     & "&txtCarkind="		& Frm1.txtCarkind.value      '☜: Query Key

			strVal = strVal     & "&txtMaxRows="		& .vspdData.MaxRows
			strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey                 '☜: Next key tag
			strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey                 '☜: Next key tag
			strVal = strVal     & "&lgPageNo_A="		& lgPageNo_A                          '☜: Next key tag
			strVal = strVal     & "&txtType="			& "A"                          '☜: Next key tag
    End With

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    DbQuery = True
    Call DbQueryOk()
End Function

'========================================================================================================
' Name : DbDtlQuery1
' Desc : This function is called by FncQuery
'========================================================================================================

Function DbDtlQuery1()
	
    DbDtlQuery1 = False

    Err.Clear    
                                                                        '☜: Clear err status
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

    With Frm1
		strVal = BIZ_PGM_QUERY_ID & "?txtMode="			& parent.UID_M0001
        strVal = strVal     & "&lgCurrentSpd="		& "S1"                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="		& lgKeyStream                       '☜: Query Key
	    strVal = strVal     & "&txtWorkDt="			& .txthWorkDt.Value                     '☜: Query Key
	    strVal = strVal     & "&txtCastCd="			& .txthCastCd.Value     '☜: Query Key
	    strVal = strVal     & "&txtPlantCd="		& ""
	    strVal = strVal     & "&CboFacility_Accnt="	& ""
        strVal = strVal     & "&txtMaxRows="		& .vspdData1.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" 		& lgStrPrevKey1                 '☜: Next key tag
		strVal = strVal     & "&lgPageNo_B="		& lgPageNo_B                          '☜: Next key tag
		strVal = strVal     & "&txtType="			& "B"                          '☜: Next key tag
    End With
'msgbox strval
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbDtlQuery1 = True
End Function

'========================================================================================================
' Name : DbDtlQuery2
' Desc : This function is called by FncQuery
'========================================================================================================

Function DbDtlQuery2()

    DbDtlQuery2 = False

    Err.Clear                                                                      '☜: Clear err status

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

    With Frm1
		strVal = BIZ_PGM_QUERY2_ID & "?txtMode="			& parent.UID_M0001
        strVal = strVal     & "&lgCurrentSpd="		& "S2"                      '☜: Next key tag
        strVal = strVal     & "&txtKeyStream="		& lgKeyStream                       '☜: Query Key
	    strVal = strVal     & "&txtWorkDt="			& .txthWorkDt.value                    '☜: Query Key
	    strVal = strVal     & "&txtCastCd="			& .txthCastCd.value      '☜: Query Key
        strVal = strVal     & "&txtMaxRows="		& .vspdData2.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" 		& lgStrPrevKey2                 '☜: Next key tag
		strVal = strVal     & "&lgPageNo_C="		& lgPageNo_C                          '☜: Next key tag
		strVal = strVal     & "&txtType="			& "C"                          '☜: Next key tag    End With
    End With

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

    DbDtlQuery2 = True
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal, strDel
	
    DbSave = False

    If LayerShowHide(1) = False Then
			Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    if frm1.ChgSave1.value = "T"  then
  		ggoSpread.Source = frm1.vspdData
	    With Frm1
           For lRow = 1 To .vspdData.MaxRows
               .vspdData.Row = lRow
               .vspdData.Col = 0
               Select Case .vspdData.Text
				   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & "M" & parent.gColSep
                        .vspdData.Col = C_CAST_CD   : strVal = strVal & 			Trim(.vspdData.Text)	& parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strVal = strVal & UNIConvDate(Trim(.vspdData.Text))   & parent.gColSep
                        .vspdData.Col = C_Insp_Text		: strVal = strVal &             Trim(.vspdData.Text) 	& parent.gColSep
                        .vspdData.Col = C_Insp_Hour		: strVal = strVal & UNIConvNum (Trim(.vspdData.Text),0)	& parent.gColSep
                        .vspdData.Col = C_Insp_Min		: strVal = strVal & UNIConvNum (Trim(.vspdData.Text),0)	& parent.gColSep
                        .vspdData.Col = C_Req_Dept		: strVal = strVal & 			Trim(.vspdData.Text)	& parent.gColSep
                        .vspdData.Col = C_Insp_Dept		: strVal = strVal & 			Trim(.vspdData.Text)	& parent.gColSep
                        .vspdData.Col = C_Insp_Emp_Qty	: strVal = strVal & UNIConvNum (Trim(.vspdData.Text),0)	& parent.gColSep
                        .vspdData.Col = C_Payroll		: strVal = strVal & UNIConvNum (Trim(.vspdData.Text),0)	& parent.gColSep
                        .vspdData.Col = C_Matl_Cost		: strVal = strVal & UNIConvNum (Trim(.vspdData.Text),0)	& parent.gColSep
                        .vspdData.Col = C_Insp_Flag		: strVal = strVal &			    Trim(.vspdData.Text)    & parent.gColSep
                        .vspdData.Col = C_INSP_PRID  	: strVal = strVal &	UNIConvNum (Trim(.vspdData.Text),0)	& parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & "M" & parent.gColSep
                        .vspdData.Col = C_CAST_CD	    : strDel = strDel &				Trim(.vspdData.Text)	& parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strDel = strDel & UNIConvDate(Trim(.vspdData.Text))	& parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
           .txtMode.value        =  parent.UID_M0002
	       .txtMaxRows.value     = lGrpCnt-1
	       .txtSpread.value      = strDel & strVal
	    End With
	end if

    if  frm1.ChgSave2.value = "T" then
 		ggoSpread.Source = frm1.vspdData1
	    With Frm1
           For lRow = 1 To .vspdData1.MaxRows
               .vspdData1.Row = lRow
               .vspdData1.Col = 0
               Select Case .vspdData1.Text
                    Case  ggoSpread.InsertFlag                                      '☜: Create
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & "S1" & parent.gColSep
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
                        .vspdData.Col = C_CAST_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData1.Col = C_Seq 			: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Zinsp_PartCd	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_PartCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_MethCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_DeciCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_St_GoGubunCd : strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Sury_Assy	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_S_Qty		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Price		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Amt		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Type	: strVal = strVal &				Trim(.vspdData1.Text)	& parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & "S1" & parent.gColSep
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_CAST_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData1.Col = C_Seq 			: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Zinsp_PartCd	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_PartCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_MethCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Insp_DeciCd 	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_St_GoGubunCd : strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_Sury_Assy	: strVal = strVal & 			Trim(.vspdData1.Text)	& parent.gColSep
						.vspdData1.Col = C_S_Qty		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Price		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Amt		: strVal = strVal & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gColSep
						.vspdData1.Col = C_Sury_Type	: strVal = strVal &				Trim(.vspdData1.Text)	& parent.gRowSep
                        lGrpCnt = lGrpCnt + 1

                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & "S1" & parent.gColSep
                        frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_CAST_CD       : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData1.Col = C_Seq 			: strDel = strDel & UNIConvNum( Trim(.vspdData1.Text),0) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next

           .txtMode.value        =  parent.UID_M0002
	       .txtMaxRows.value     = lGrpCnt-1
	       .txtSpread.value      = strDel & strVal

	    End With
    end if

    if  frm1.ChgSave3.value = "T" then
		ggoSpread.Source = frm1.vspdData2
	    With Frm1
           For lRow = 1 To .vspdData2.MaxRows
               .vspdData2.Row = lRow
               .vspdData2.Col = 0
               Select Case .vspdData2.Text

                   Case  ggoSpread.InsertFlag                                      '☜: Create
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & "S2" & parent.gColSep
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_CAST_CD	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData2.Col = C_Seq 			: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
                        .vspdData2.Col = C_Insp_Emp_Gb	: strVal = strVal &             Trim(.vspdData2.Text)   & parent.gColSep
						.vspdData2.Col = C_Insp_Emp_Cd	: strVal = strVal & 			Trim(.vspdData2.Text)	& parent.gColSep
						.vspdData2.Col = C_Cust_Cd	 	: strVal = strVal & 			Trim(.vspdData2.Text)	& parent.gColSep
						.vspdData2.Col = C_Insp_Hour2	: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Insp_Min2	: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Payroll2		: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & "S2" & parent.gColSep
						frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_CAST_CD   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData2.Col = C_Seq 			: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
                        .vspdData2.Col = C_Insp_Emp_Gb	: strVal = strVal &             Trim(.vspdData2.Text)   & parent.gColSep
						.vspdData2.Col = C_Insp_Emp_Cd	: strVal = strVal & 			Trim(.vspdData2.Text)	& parent.gColSep
						.vspdData2.Col = C_Cust_Cd	 	: strVal = strVal & 			Trim(.vspdData2.Text)	& parent.gColSep
						.vspdData2.Col = C_Insp_Hour2	: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Insp_Min2	: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
						.vspdData2.Col = C_Payroll2		: strVal = strVal & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1

                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & "S2" & parent.gColSep
                        frm1.vspdData.Row = frm1.vspdData.ActiveRow
                        .vspdData.Col = C_CAST_CD   : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_Plan_Dt	    : strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep
                        .vspdData2.Col = C_Seq 			: strDel = strDel & UNIConvNum( Trim(.vspdData2.Text),0) & parent.gColSep
                        .vspdData2.Col = C_Insp_Emp_Gb	: strDel = strDel &             Trim(.vspdData2.Text)   & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next

           .txtMode.value        =  parent.UID_M0002
	       .txtMaxRows.value     = lGrpCnt-1
	       .txtSpread.value      = strDel & strVal
	    End With
    end if
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
    DbSave = True

End Function


'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd

    FncDelete = False                                                      '⊙: Processing is NG

    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------
		Exit Function
	End If


    Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	lgOldRow_A = 0
	lgOldRow_B = 0
	lgOldRow_C = 0
  lgIntFlgMode =  parent.OPMD_UMODE
  Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
  Call InitData()

  Call SetToolbar("1100111100011111")

	if lgStrPrevKey1 <> "" and isFirst = false then
		exit function
	end if
	
	if isFirst = TRUE	Then	' 첫화면이 열리고나서 오른쪽 그리드 세팅하기 위해 
		Call DisableToolBar(parent.TBC_QUERY)
		Call vspdData_click(1,frm1.vspdData.activerow)
	end if
	
'  SetSpreadLock1()
  
	frm1.vspdData.focus
End Function

Function DbDtlQueryOk1()
    lgIntFlgMode =  parent.OPMD_UMODE

	Call InitData()
    Call SetToolbar("1100111100011111")

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
	Call DbDtlQuery2
'	frm1.vspdData1.focus
End Function

Function DbDtlQueryOk2()
    lgIntFlgMode =  parent.OPMD_UMODE

	Call InitData()
    Call SetToolbar("1100111100011111")

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
'	frm1.vspdData2.focus
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")

    Call InitVariables															'⊙: Initializes local global variables
    lgCurrentSpd = "M"
    Call MakeKeyStream("X")

	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function



'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 1 Then 'TextBox(Condition)
	    frm1.vspdData2.Col = C_Insp_Emp_Cd
		arrParam(0) = frm1.vspdData2.Text			' Code Condition
        frm1.vspdData2.Col = C_Insp_Emp_Nm
	    arrParam(1) = ""'frm1.vspdData2.Text		' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If
		
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")	
	
	IsOpenPop = False
		
	If arrRet(0) = "" Then	
		If iWhere = 0 Then
			frm1.C_Insp_Emp_Cd.focus
		Else
			frm1.vspdData2.Col = C_Insp_Emp_Cd
			frm1.vspdData2.action =0 
		End If
		Exit Function
	Else	
		Call SubSetCondEmp(arrRet, iWhere)	
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
    
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)

		Else 'spread
			.vspdData2.row = .vspdData2.ActiveRow
			.vspdData2.Col = C_Insp_Emp_Cd
			.vspdData2.Text = arrRet(0)
			.vspdData2.Col = C_Insp_Emp_Nm
			.vspdData2.Text = arrRet(1)
			.vspdData2.action =0 
			Call SetActiveCell(frm1.vspdData2,C_CUST_CD,frm1.vspdData2.ActiveRow,"M","X","X")
			Set gActiveElement = document.activeElement
		End If
	End With
End Sub



 '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant(byval strComp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  And strComp <> "Plant1"  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_Plant"


	arrParam(2) = Trim(frm1.txtPlantCd.Value)

	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_Cd"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If strComp="Plant1" Then
			frm1.txtPlantCd.focus
		Else
		End If
		Exit Function
	Else
		If strComp="Plant1" Then
			frm1.txtPlantCd.Value  = arrRet(0)
			frm1.txtPlantNm.Value  = arrRet(1)
			frm1.txtPlantCd.focus
		Else
		End If
	End If
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then

		Call EScCode(iwhere)
		Exit Function
	Else

		Call SetBp(arrRet, iWhere)
		if iWhere <> 0 then
       	  ggoSpread.Source = frm1.vspdData2
           ggoSpread.UpdateRow frm1.vspdData2.ActiveRow
        End if
	End If	
End Function

'========================================================================================================
'	Name : SetBp()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetBp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		    Case 3
				.vspdData2.row    = .vspdData2.ActiveRow
		        .vspdData2.Col    = C_Cust_Cd
		    	.vspdData2.text   = arrRet(0)
		    	.vspdData2.Col    = C_Cust_Nm
		    	.vspdData2.text   = arrRet(1)
		    	Call SetActiveCell(.vspdData2,C_INSP_HOUR2,.vspdData2.ActiveRow ,"M","X","X")
		    	Set gActiveElement = document.activeElement
        End Select

	End With

End Function


'------------------------------------------  OpenSItem()  -------------------------------------------------
' Name : OpenSItem()
' Description : SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSItem(byval strCon)

	Dim arrRet
	Dim arrParam(5), arrField(15)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_SET_PLANT
	
	arrParam(0) = Trim(frm1.vspdData.text)	' Plant Code
	arrParam(1) =strCon						' Item Code
	arrParam(2) = ""						' Combo Set Data:"1029!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 	'ITEM_CD		' Field명(0)
	arrField(1) = 2 	'ITEM_NM		' Field명(1)

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData1 
			.Row = .ActiveRow 
			.Col = C_Sury_Assy_Nm
			.text = arrRet(1)
			.Col = C_Sury_Assy
			.text = arrRet(0) 
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
		   	Call SetActiveCell(frm1.vspdData1,C_S_QTY, frm1.vspdData1.ActiveRow,"M","X","X")
			Set gActiveElement = document.activeElement
			
		End With 
	End If 
End Function

'========================================================================================================
'	Name : EScCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function EScCode(Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 3
		    	Call SetActiveCell(.vspdData2,C_INSP_HOUR,.vspdData2.ActiveRow ,"M","X","X")
        End Select

	End With

End Function
'======================================================================================================
'	Name : OpenCode()
'	Description :
'=======================================================================================================

Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_Insp_Dept_POP, C_Req_Dept_POP
	        arrParam(0) = "부서코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "H_CURRENT_DEPT"		  			    ' TABLE 명칭 
	    	arrParam(2) = strCode                   	        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = ""	                		    	' Where Condition
	    	arrParam(5) = "부서코드" 			            ' TextBox 명칭 

	    	arrField(0) = "dept_cd"						    	' Field명(0)
	    	arrField(1) = "dept_nm"    					    	' Field명(1)

	    	arrHeader(0) = "부서코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "부서코드명"	    		        ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	   	frm1.vspdData.action=0
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If

End Function

'======================================================================================================
'	Name : SetCode()
'	Description :
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_Insp_Dept_POP
		    	.vspdData.Col = C_Insp_Dept_NM
		    	.vspdData.text = arrRet(1)
		        .vspdData.Col = C_Insp_Dept
		    	.vspdData.text = arrRet(0)
		    	.vspdData.action=0
		    Case C_Req_Dept_POP
		    	.vspdData.Col = C_Req_Dept_NM
		    	.vspdData.text = arrRet(1)
		        .vspdData.Col = C_Req_Dept
		    	.vspdData.text = arrRet(0)
		    	.vspdData.action=0
        End Select
	End With
End Function



'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

    lgCurrentSpd = "M"

End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
    Dim intIndex
    Dim tmpSqty
    Dim tmpPrice
    Dim tmpAmt
    
   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

	Select Case Col
		Case  C_Zinsp_PartNm
			Frm1.vspdData1.col = C_Zinsp_PartNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Zinsp_PartCd
			Frm1.vspdData1.value = intindex
		Case  C_Insp_PartNm
			Frm1.vspdData1.col = C_Insp_PartNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_PartCd
			Frm1.vspdData1.value = intindex
		Case  C_Insp_MethNm
			Frm1.vspdData1.col = C_Insp_MethNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_MethCd
			Frm1.vspdData1.value = intindex
		Case  C_Insp_DeciNm
			Frm1.vspdData1.col = C_Insp_DeciNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_DeciCd
			Frm1.vspdData1.value = intindex
		Case  C_St_GoGubunNm
			Frm1.vspdData1.col = C_St_GoGubunNm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_St_GoGubunCd
			Frm1.vspdData1.value = intindex
		Case  C_Zinsp_PartCd
			Frm1.vspdData1.col = C_Zinsp_PartCd
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Zinsp_PartNm
			Frm1.vspdData1.value = intindex
		Case  C_Insp_PartCd
			Frm1.vspdData1.col = C_Insp_PartCd
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_PartNm
			Frm1.vspdData1.value = intindex
		Case  C_Insp_MethCd
			Frm1.vspdData1.col = C_Insp_MethCd
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_MethNm
			Frm1.vspdData1.value = intindex
		Case  C_Insp_DeciCd
			Frm1.vspdData1.col = C_Insp_DeciCd
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_Insp_DeciNm
			Frm1.vspdData1.value = intindex
		Case  C_St_GoGubunCd
			Frm1.vspdData1.col = C_St_GoGubunCd
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_St_GoGubunNm
			Frm1.vspdData1.value = intindex		
		Case  C_SURY_TYPE
			Frm1.vspdData1.col = C_SURY_TYPE
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_SURY_TYPE_NM
			Frm1.vspdData1.value = intindex				
		Case  C_SURY_TYPE_Nm
			Frm1.vspdData1.col = C_SURY_TYPE_Nm
			intIndex = Frm1.vspdData1.value
			Frm1.vspdData1.Col = C_SURY_TYPE
			Frm1.vspdData1.value = intindex				
		Case C_S_QTY
			frm1.vspdData1.col = C_S_QTY
			tmpSqty = frm1.vspdData1.value
			frm1.vspdData1.col = C_PRICE
			tmpPrice = frm1.vspdData1.value
			tmpAmt = tmpSqty * tmpPrice 
			frm1.vspdData1.col = C_SURY_AMT
			frm1.VspdData1.value = tmpAmt
		Case C_PRICE
			frm1.vspdData1.col = C_S_QTY
			tmpSqty = frm1.vspdData1.value
			frm1.vspdData1.col = C_PRICE
			tmpPrice = frm1.vspdData1.value
			tmpAmt = tmpSqty * tmpPrice 
			frm1.vspdData1.col = C_SURY_AMT
			frm1.VspdData1.value = tmpAmt			
	End Select
	

   	If Frm1.vspdData1.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData1.text) <  UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData1
     ggoSpread.UpdateRow Row

    lgCurrentSpd = "S1"

End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
    Dim intIndex

   	Frm1.vspdData2.Row = Row
   	Frm1.vspdData2.Col = Col

	Select Case Col
		Case  C_Insp_Emp_GbNm
			Frm1.vspdData2.col = C_Insp_Emp_GbNm
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Zinsp_PartCd
			Frm1.vspdData2.value = intindex
		Case  C_Insp_Emp_Gb
			Frm1.vspdData2.col = C_Insp_Emp_Gb
			intIndex = Frm1.vspdData2.value
			Frm1.vspdData2.Col = C_Insp_Emp_GbNm
			Frm1.vspdData2.value = intindex
    End Select

   	If Frm1.vspdData2.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData2.text) <  UNICDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
	End If

	 ggoSpread.Source = frm1.vspdData2
     ggoSpread.UpdateRow Row

    lgCurrentSpd = "S2"

End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim flagTxt
	Dim txtCastCd
	Dim txtWorkDt
    Call SetPopupMenuItemInf("1101111111")

	IF lgBlnFlgChgValue = False and frm1.vspdData.Maxrows = 0 then
		Call SetToolbar("1100110100011111")
	End if

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	With Frm1
		.vspdData.Row = Row
		.vspdData.Col = 0
		flagTxt = .vspdData.Text
		If flagTxt =  ggoSpread.InsertFlag or flagTxt =  ggoSpread.UpdateFlag or flagTxt =  ggoSpread.DeleteFlag Then
			Exit Sub
		End If
	End With

    gMouseClickStatus = "SPC"

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


	'lgCurrentSpd = "S1"
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""


	If lgOldRow_A <> Row Then
		frm1.vspdData.Col = C_Plan_Dt
		frm1.txthWorkDt.value = frm1.vspdData.text
		frm1.vspdData.Col = C_CAST_CD
		frm1.txthCastCd.value = frm1.vspdData.text
		lgOldRow_A = Row

		Call  DisableToolBar( parent.TBC_QUERY)
	    ggoSpread.Source       = Frm1.vspdData1
	    ggoSpread.ClearSpreadData
	    ggoSpread.Source       = Frm1.vspdData2
	    ggoSpread.ClearSpreadData

		lgPageNo_B = 0
		lgPageNo_C = 0

		Call DbDtlQuery1

	End if
End Sub



'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101011111")

     gMouseClickStatus = "SP1C"

    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1

       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending
           lgSortKey = 1
       End If

       Exit Sub
    End If
    lgCurrentSpd = "S1"
    Set gActiveSpdSheet = frm1.vspdData1
End Sub
'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101011111")

     gMouseClickStatus = "SP2C"

    Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2

       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending
           lgSortKey = 1
       End If

       Exit Sub
    End If
    lgCurrentSpd = "S2"
    Set gActiveSpdSheet = frm1.vspdData2
End Sub

'========================================================================================================
'		Event	Name : vspdData_ScriptLeaveCell
'		Event	Desc : This	function is	called when	cursor leave cell
'========================================================================================================
Sub	vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	Dim	iRet
	If NewRow <= 0 Or	Row	=	NewRow Then	Exit Sub
	
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row	=	NewRow
	frm1.vspdData.Col	=	C_Plan_Dt
	frm1.txthWorkDt.value	=	frm1.vspdData.text
	frm1.vspdData.Col	=	C_CAST_CD
	frm1.txthCastCd.value	=	frm1.vspdData.text

	lgOldRow_A = NewRow

	Call	DisableToolBar(	parent.TBC_QUERY)
	ggoSpread.Source			 = Frm1.vspdData1
	ggoSpread.ClearSpreadData
	ggoSpread.Source			 = Frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	lgPageNo_B = 0
	lgPageNo_C = 0
	If DbDtlQuery1() = False Then	Exit Sub

	
End	Sub


'========================================================================================================
'		Event	Name : vspdData_TopLeftChange
'		Event	Desc : This	function is	data query with	spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If CheckRunningBizProcess = True Then
		Exit Sub
	End If

	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgPageNo_A <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			Call Dbquery
		End If
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
'    If OldLeft <> NewLeft Then
'        Exit Sub
'    End If
'    
'	If CheckRunningBizProcess = True Then
'	   Exit Sub
'	End If
'    
'    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	'☜: 재쿼리 체크'
'
'		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'           Call DisableToolBar(Parent.TBC_QUERY)
'           Call DbDtlQuery1
'	    End If
'	End if
End Sub
'========================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
'    If OldLeft <> NewLeft Then
'        Exit Sub
'    End If
'    
'	If CheckRunningBizProcess = True Then
'	   Exit Sub
'	End If
'    
'    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	'☜: 재쿼리 체크'
'		If lgPageNo_C <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'           Call DisableToolBar(Parent.TBC_QUERY)
'           Call DbDtlQuery2
'	    End If
'	End if
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
'   Event Name : vspdData2_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

    If Row <= 0 Then
        Exit Sub
    End If

    If frm1.vspdData2.MaxRows = 0 Then
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
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
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("C")
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


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData2
		 ggoSpread.Source = frm1.vspdData2
		If Row > 0 Then
			Select Case Col
			    Case C_Insp_Emp_Pop
		                Call OpenEmptName("1")
			    Case C_Cust_Pop
			         frm1.vspdData2.Col = C_Cust_CD   
		            Call OpenBp(frm1.vspdData2.Text, 3)
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
'   Event Name : vspdData2_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_OnFocus()
    lgActiveSpd      = "S2"
	lgCurrentSpd	="S2"
End Sub

'========================================================================================================
'   Event Name : txtOcpt_type_Onchange
'   Event Desc :
'========================================================================================================
Function txtOcpt_type_Onchange()
    gCounts = 0
End Function
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
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub


'==========================================================================================
'   Event Name : txtAppFrDt
'   Event Desc :
'==========================================================================================

 Sub txtWork_Dt_DblClick(Button)
	if Button = 1 then
		frm1.txtWork_Dt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtWork_Dt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================

Sub txtWork_Dt_KeyDown(KeyCode, Shift)
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

'------------------------------------------  OpenCarKind()  -------------------------------------------------
'	Name : OpenCarKind()
'	Description : Condition CarKind PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCarKind()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "적용모델"						' 팝업 명칭 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCarKind.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "적용모델"						' TextBox 명칭 
	
    arrField(0) = "ITEM_GROUP_CD"						' Field명(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field명(1)
    
    arrHeader(0) = "적용모델"						' Header명(0)
    arrHeader(1) = "적용모델명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCarKind(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCarKind.focus
End Function

'------------------------------------------  OpenCarKind()  -------------------------------------------------
'	Name : OpenCarKind1()
'	Description : Condition CarKind PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCarKind1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "적용모델"						' 팝업 명칭 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCarKind1.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "적용모델"						' TextBox 명칭 
	
    arrField(0) = "ITEM_GROUP_CD"						' Field명(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field명(1)
    
    arrHeader(0) = "적용모델"						' Header명(0)
    arrHeader(1) = "적용모델명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCarKind1(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCarKind.focus
	
End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast(byval arrRet)
	frm1.txtCastCd.Value    = arrRet(0)		
	frm1.txtCastNm.Value    = arrRet(1)	

End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast1()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast1(byval arrRet)
	frm1.txtCastCd1.Value    = arrRet(0)		
	frm1.txtCastNm1.Value    = arrRet(1)	
	lgBlnFlgChgValue = True		
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind()
'	Description : Condition CarKind Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind(byval arrRet)
	frm1.txtCarKind.Value    = arrRet(0)		
	frm1.txtCarKindNm.Value  = arrRet(1)

End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind1()
'	Description : Condition CarKind Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind1(byval arrRet)
	frm1.txtCarKind1.Value    = arrRet(0)		
	frm1.txtCarKindNm1.Value  = arrRet(1)
	lgBlnFlgChgValue = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>금형점검내역등록</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant('Plant1')" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
															<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X">
									<TD CLASS="TD5" NOWRAP>작업일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p6220ma1_txtWork_Dt_txtWork_Dt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>적용모델</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCarKind" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="적용모델"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCarKind" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCarKind()">&nbsp;<INPUT TYPE=TEXT NAME="txtCarKindNm" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP>금형코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ID=txtCastCd NAME="txtCastCd" ALT="금형코드" TYPE="Text" SiZE="18" MAXLENGTH="18" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCast()">
															<INPUT ID=txtCastNm NAME="txtCastNm" ALT="금형코드명" TYPE="Text" SiZE="25" MAXLENGTH="40" tag="14XXXU"></TD>
								</TR>
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
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p6220ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="20%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p6220ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="20%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p6220ma1_vaSpread2_vspdData2.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode"          tag="24"> <INPUT TYPE=HIDDEN NAME="txtInsrtUserId"   tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"       tag="24"> <INPUT TYPE=HIDDEN NAME="ChgSave1"         tag="24">
<INPUT TYPE=HIDDEN NAME="ChgSave2"         tag="24"> <INPUT TYPE=HIDDEN NAME="ChgSave3"         tag="24">
<INPUT TYPE=HIDDEN NAME="txthCastCd"       tag="24"> <INPUT TYPE=HIDDEN NAME="txthWorkDt"       tag="24">
<INPUT TYPE=HIDDEN NAME="txthSuryAssy"     tag="24"> <INPUT TYPE=HIDDEN NAME="txtInspEmpCd"     tag="24">
<INPUT TYPE=HIDDEN NAME="txtCustCd"        tag="24"> <INPUT TYPE=HIDDEN NAME="txtRequestDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRepairDeptCd"  tag="24"> 
<INPUT TYPE=HIDDEN NAME="hWork_Dt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFacility_Cd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

