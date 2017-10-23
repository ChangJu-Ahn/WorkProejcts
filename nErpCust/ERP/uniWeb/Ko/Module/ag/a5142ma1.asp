<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5123MA1
'*  4. Program Name         : 회계전표일괄생성 
'*  5. Program Desc         : 각 모쥴에서 생성한 자료를 토대로 일괄적으로 전표처리.
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/09/26 : ..........
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit  

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID1 = "a5142mb1.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "a5142mb2.asp"												'☆: 비지니스 로직 ASP명 
Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_Confirm  															'☆: Spread Sheet의 Column별 상수 
Dim C_BatchDt  														'☆: Spread Sheet의 Column별 상수  
Dim C_Refno  
Dim C_BizCD  
Dim C_BizNm  
Dim C_GLInputType  
Dim C_GLInputTypeNm  
Dim C_BatchNo  
Dim C_ItemAmt  
Dim C_ItemLocAmt  
Dim C_GlDesc   
'☜:--------Spreadsheet #2-----------------------------------------------------------------------------
Dim C_Confirm1  															'☆: Spread Sheet의 Column별 상수 
Dim C_BatchDt1  														'☆: Spread Sheet의 Column별 상수  
Dim C_Chainno1  
Dim C_Refno1  
Dim C_BizCD1  
Dim C_BizNm1  
Dim C_GLInputType1  
Dim C_GLInputTypeNm1  
Dim C_BatchNo1  
Dim C_GlDt1  
Dim C_GlNo1  
Dim C_ItemAmt1  
Dim C_ItemLocAmt1  
Dim C_GlDesc1   
Dim C_TEMP_Gl_FG1  
'☜:--------Spreadsheet #3-----------------------------------------------------------------------------      
Dim C_Confirm2  															'☆: Spread Sheet의 Column별 상수 
Dim C_BatchDt2  														'☆: Spread Sheet의 Column별 상수  
Dim C_Chainno2  
Dim C_Refno2  
Dim C_BizCD2  
Dim C_BizNm2  
Dim C_GLInputType2  
Dim C_GLInputTypeNm2  
Dim C_BatchNo2  
Dim C_GlDt2  
Dim C_GlNo2  
Dim C_TEMP_Gl_FG2  


Const C_SHEETMAXROWS = 70		' : 한 화면에 보여지는 최대갯수*1.5

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgStrPrevKeyTempGlDt
Dim lgStrPrevKeyBatchNo

Dim lgQueryFlag					' 신규조회 및 추가조회 구분 Flag
Dim  gSelframeFlg

Dim  IsOpenPop          

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
     lgStrPrevKeyTempGlDt = ""              
    lgStrPrevKeyBatchNo = ""                       'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

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
	Dim StartDate, EndDate
	Dim strYear, strMonth, strDay

	Call	ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
	
	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")      '☆: 초기화면에 뿌려지는 시작 날짜 
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)   '☆: 초기화면에 뿌려지는 마지막 날짜 

	frm1.txtFromReqDt.text =  StartDate
	frm1.txtToReqDt.text   =  EndDate
	frm1.GIDate.text   =  EndDate
	frm1.txtFromReqDt1.text =  StartDate
	frm1.txtToReqDt1.text   =  EndDate
	frm1.txtTransType.focus	
End Sub
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_Confirm = 1															'☆: Spread Sheet의 Column별 상수 
	C_BatchDt = 2														'☆: Spread Sheet의 Column별 상수  
	C_Refno = 3
	C_BizCD = 4
	C_BizNm = 5
	C_GLInputType = 6
	C_GLInputTypeNm = 7
	C_BatchNo = 8
	C_ItemAmt = 9
	C_ItemLocAmt = 10
	C_GlDesc  = 11
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables1()  
	
'☜:--------Spreadsheet #2-----------------------------------------------------------------------------
	C_Confirm1 = 1															'☆: Spread Sheet의 Column별 상수 
	C_BatchDt1 = 2														'☆: Spread Sheet의 Column별 상수  
	C_Chainno1 = 3
	C_Refno1 = 4
	C_BizCD1 = 5
	C_BizNm1 = 6
	C_GLInputType1 = 7
	C_GLInputTypeNm1 = 8
	C_BatchNo1 = 9
	C_GlDt1 = 10
	C_GlNo1 = 11
	C_ItemAmt1 = 12
	C_ItemLocAmt1 = 13
	C_GlDesc1  = 14
	C_TEMP_Gl_FG1 = 15
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables2()  	
'☜:--------Spreadsheet #3-----------------------------------------------------------------------------      
	C_Confirm2 = 1															'☆: Spread Sheet의 Column별 상수 
	C_BatchDt2 = 2														'☆: Spread Sheet의 Column별 상수  
	C_Chainno2 = 3
	C_Refno2 = 4
	C_BizCD2 = 5
	C_BizNm2 = 6
	C_GLInputType2 = 7
	C_GLInputTypeNm2 = 8
	C_BatchNo2 = 9
	C_GlDt2 = 10
	C_GlNo2 = 11
	C_TEMP_Gl_FG2 = 12	
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet(ByVal pvSpdNo)

  Select Case UCase(pvSpdNo)
    Case "A"
	
	Call initSpreadPosVariables() 
	With frm1.vspdData
	
		.MaxCols = C_GlDesc+1									'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols										'☆: 사용자 별 Hidden Column
		.ColHidden = True
		          
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030308",,parent.gAllowDragDropSpread
		
		Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
		
		.ReDraw = false
		
		Call GetSpreadColumnPos("A")
	

		'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)    
		ggoSpread.SSSetCheck C_Confirm,     "",     8,  -10, "", True, -1 
		ggoSpread.SSSetDate C_BatchDt,     "발생일", 10,,Parent.gDateFormat
		ggoSpread.SSSetEdit C_Refno, "참조번호", 25,,,30                                
		ggoSpread.SSSetEdit C_BizCD,    "사업장", 10,,,10
		ggoSpread.SSSetEdit C_BizNm,   "사업장명", 18,,,20
		ggoSpread.SSSetEdit C_GLInputType,   "거래유형",       10,,,20
		ggoSpread.SSSetEdit C_GLInputTypeNm,"거래유형명", 18,,,50
		ggoSpread.SSSetEdit C_BatchNo,     "배치번호", 20,,,20
		ggoSpread.SSSetFloat  C_ItemAmt,    "금액",       15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)", 15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_GlDesc,   "적  요", 30, , , 128
		

		.ReDraw = true
    End With
    Call SetSpreadLock

    Case "B" 
    
    Call initSpreadPosVariables1()  
        
    With frm1.vspdData1
		.MaxCols = C_TEMP_Gl_FG1+1									'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols										'☆: 사용자 별 Hidden Column
		.ColHidden = True

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20030308",,parent.gAllowDragDropSpread
		
		Call ggoSpread.ClearSpreadData()

		.ReDraw = false
		Call GetSpreadColumnPos("B")


		'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)    
		ggoSpread.SSSetCheck C_Confirm1,     "",     8,  -10, "", True, -1 
		ggoSpread.SSSetDate C_BatchDt1,     "발생일", 10,,Parent.gDateFormat
		ggoSpread.SSSetEdit C_Chainno1, "참조번호", 25,,,30    
		ggoSpread.SSSetEdit C_Refno1, "참조번호", 25,,,30                                
		ggoSpread.SSSetEdit C_BizCD1,    "사업장", 10,,,10
		ggoSpread.SSSetEdit C_BizNm1,   "사업장명", 15,,,20
		ggoSpread.SSSetEdit C_GLInputType1,   "거래유형",       10,,,20
		ggoSpread.SSSetEdit C_GLInputTypeNm1,"거래유형명", 18,,,50
		ggoSpread.SSSetEdit C_BatchNo1,     "배치번호", 15,,,20
		ggoSpread.SSSetDate C_GlDt1, "전표일",   10,,Parent.gDateFormat
		ggoSpread.SSSetEdit C_GlNo1, "전표번호", 15,,,20
		ggoSpread.SSSetFloat  C_ItemAmt1,    "금액",       15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt1, "금액(자국)", 15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_GlDesc1,   "적  요", 30, , , 128
		ggoSpread.SSSetEdit C_TEMP_Gl_FG1, "", 2,,,2
		
		Call ggoSpread.SSSetColHidden(C_Chainno1 ,C_Chainno1	,True)
		Call ggoSpread.SSSetColHidden(C_TEMP_Gl_FG1 ,C_TEMP_Gl_FG1	,True)
				
		.ReDraw = true
	End With
    Call SetSpreadLock1

    Case "C" 
    Call initSpreadPosVariables2() 	
	
	With frm1.vspdData2
		.MaxCols = C_TEMP_Gl_FG2+1									'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols										'☆: 사용자 별 Hidden Column
		.ColHidden = True

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20030308",,parent.gAllowDragDropSpread
		
		Call ggoSpread.ClearSpreadData()

		.ReDraw = false
	
		Call GetSpreadColumnPos("C")

		'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)    
		ggoSpread.SSSetCheck C_Confirm2,     "",     8,  -10, "", True, -1 
		ggoSpread.SSSetDate C_BatchDt2,     "발생일", 10,,Parent.gDateFormat
		ggoSpread.SSSetEdit C_Chainno2, "참조번호", 25,,,30    
		ggoSpread.SSSetEdit C_Refno2, "참조번호", 25,,,30                                
		ggoSpread.SSSetEdit C_BizCD2,    "사업장", 10,,,10
		ggoSpread.SSSetEdit C_BizNm2,   "사업장명", 15,,,20
		ggoSpread.SSSetEdit C_GLInputType2,   "거래유형",       10,,,20
		ggoSpread.SSSetEdit C_GLInputTypeNm2,"거래유형명", 18,,,50
		ggoSpread.SSSetEdit C_BatchNo2,     "배치번호", 15,,,20
		ggoSpread.SSSetDate C_GlDt2, "전표일",   10,,Parent.gDateFormat
		ggoSpread.SSSetEdit C_GlNo2, "전표번호", 15,,,20
		ggoSpread.SSSetEdit C_TEMP_Gl_FG2, "", 2,,,2
		
		Call ggoSpread.SSSetColHidden(C_Chainno2 ,C_Chainno2	,True)
		Call ggoSpread.SSSetColHidden(C_TEMP_Gl_FG2 ,C_TEMP_Gl_FG2	,True)
				
		.ReDraw = true
	End With
	
    Call SetSpreadLock2 
  End Select  
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1
    ggoSpread.SpreadLock C_BatchDt,			-1, C_BatchDt
    ggoSpread.spreadlock C_BizCD,			-1, C_BizCD
    ggoSpread.spreadlock C_BizNm,			-1, C_BizNm
    ggoSpread.spreadlock C_Refno,			-1, C_Refno
    ggoSpread.spreadlock C_GLInputType,		-1, C_GLInputType
    ggoSpread.spreadlock C_GLInputTypeNm,	-1, C_GLInputTypeNm
    ggoSpread.spreadlock C_BatchNo,			-1, C_BatchNo
    ggoSpread.spreadlock C_ItemAmt,			-1,	C_ItemAmt
    ggoSpread.spreadlock C_ItemLocAmt,		-1,	C_ItemLocAmt
    ggoSpread.spreadlock C_GlDesc,			-1,	C_GlDesc   
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    End With
End Sub
Sub SetSpreadLock1()
    With frm1
    ggoSpread.SpreadLock C_BatchDt1,		-1,	C_BatchDt1
    ggoSpread.spreadlock C_BizCD1,			-1,	C_BizCD1
    ggoSpread.spreadlock C_BizNm1,			-1,	C_BizNm1
    ggoSpread.spreadlock C_Refno1,			-1,	C_Refno1
    ggoSpread.spreadlock C_GLInputType1,	-1,	C_GLInputType1
    ggoSpread.spreadlock C_GLInputTypeNm1,	-1,	C_GLInputTypeNm1
    ggoSpread.spreadlock C_BatchNo1,		-1, C_BatchNo1    
    ggoSpread.spreadlock C_GlDt1,			-1,	C_GlDt1
    ggoSpread.spreadlock C_GlNo1,			-1,	C_GlNo1 
    ggoSpread.spreadlock C_ItemAmt1,		-1,	C_ItemAmt1
    ggoSpread.spreadlock C_ItemLocAmt1,		-1,	C_ItemLocAmt1
    ggoSpread.spreadlock C_GlDesc1,			-1,	C_GlDesc1
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    End With
End Sub
Sub SetSpreadLock2()
    With frm1
	ggoSpread.SpreadLock C_Confirm2,		-1,	C_Confirm2
    ggoSpread.SpreadLock C_BatchDt2,		-1,	C_BatchDt2
    ggoSpread.spreadlock C_BizCD2,			-1,	C_BizCD2
    ggoSpread.spreadlock C_BizNm2,			-1,	C_BizNm2
    ggoSpread.spreadlock C_Refno2,			-1,	C_Refno2
    ggoSpread.spreadlock C_GLInputType2,	-1,	C_GLInputType2
    ggoSpread.spreadlock C_GLInputTypeNm2,	-1,	C_GLInputTypeNm2
    ggoSpread.spreadlock C_BatchNo2,		-1, C_BatchNo2    
    ggoSpread.spreadlock C_GlDt2,			-1,	C_GlDt2
    ggoSpread.spreadlock C_GlNo2,			-1,	C_GlNo2 
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1       
    End With
End Sub



'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal lRow)
    With frm1
    
    .vspdData.ReDraw = False    
    ggoSpread.SSSetProtected	C_BatchDt, lRow, lRow
    'ggoSpread.SSSetProtected	C_BatchNo, lRow, lRow
    ggoSpread.SSSetProtected	C_BizCD, lRow, lRow
    ggoSpread.SSSetProtected	C_BizNm, lRow, lRow
    ggoSpread.SSSetProtected	C_Refno, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputType, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputTypeNm, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemAmt, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemLocAmt, lRow, lRow
    ggoSpread.SSSetProtected	C_GlDesc, lRow, lRow 

    .vspdData.ReDraw = True
    
    .vspdData1.ReDraw = False    
    ggoSpread.SSSetProtected	C_BatchDt1, lRow, lRow
    'ggoSpread.SSSetProtected	C_BatchNo, lRow, lRow
    ggoSpread.SSSetProtected	C_BizCD1, lRow, lRow
    ggoSpread.SSSetProtected	C_BizNm1, lRow, lRow
    ggoSpread.SSSetProtected	C_Refno1, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputType1, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputTypeNm1, lRow, lRow
    ggoSpread.SSSetProtected	C_GlDt1, lRow, lRow
    ggoSpread.SSSetProtected	C_GlNo1, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemAmt1, lRow, lRow
    ggoSpread.SSSetProtected	C_ItemLocAmt1, lRow, lRow
    ggoSpread.SSSetProtected	C_GlDesc1, lRow, lRow 
    .vspdData1.ReDraw = True
    
    .vspdData2.ReDraw = False    
    ggoSpread.SSSetProtected	C_BatchDt2, lRow, lRow
    'ggoSpread.SSSetProtected	C_BatchNo, lRow, lRow
    ggoSpread.SSSetProtected	C_BizCD2, lRow, lRow
    ggoSpread.SSSetProtected	C_BizNm2, lRow, lRow
    ggoSpread.SSSetProtected	C_Refno2, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputType2, lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputTypeNm2, lRow, lRow
    ggoSpread.SSSetProtected	C_GlDt2, lRow, lRow
    ggoSpread.SSSetProtected	C_GlNo2, lRow, lRow
    .vspdData2.ReDraw = True
        
    
    End With
End Sub

 '========================================================================================
'                       InitComboBox_cond()
' ========================================================================================  
Sub InitComboBox_cond()
	Dim intRetCd,intLoopCnt
	Dim ArrayTemp1
	Dim ArrayTemp2
	IntRetCd = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
	
	If IntRetCD=False  Then
	    Call DisplayMsgBox("122300","X","X","X")                         '☜ : Minor코드정보가 없습니다.
	Else
		ArrayTemp1 = Split(lgF0,Chr(11))
		ArrayTemp2 = Split(lgF1,Chr(11))


	End If
End Sub


'=======================================================================================================
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromReqDt.focus
    End If
End Sub
'=======================================================================================================

Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToReqDt.focus

    End If
End Sub
'=======================================================================================================

Sub txtFromReqDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt1.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromReqDt1.focus

    End If
End Sub
'=======================================================================================================

Sub txtToReqDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt1.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToReqDt1.focus

    End If
End Sub

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0, 2			
			arrParam(0) = "거래유형"					' 팝업 명칭 
			arrParam(1) = "A_ACCT_TRANS_TYPE" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = " MO_CD <> " & FilterVar("A", "''", "S") & "  and MO_CD <> " & FilterVar("F", "''", "S") & " "							' Where Condition
			arrParam(5) = "거래유형"						' 조건필드의 라벨 명칭 

			arrField(0) = "TRANS_TYPE"						' Field명(0)
			arrField(1) = "TRANS_NM"						' Field명(1)
    
			arrHeader(0) = "거래유형"	
			arrHeader(1) = "거래유형명"
		Case 1, 3
			arrParam(0) = "사업장팝업"  				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA"	 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "사업장"	    				' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"						' Field명(0)
			arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
			arrHeader(0) = "사업장"	     				' Header명(0)
			arrHeader(1) = "사업장명"					' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then     
		Call SetPopup(arrRet, iWhere)	
	End if

	Call FocusAfterPopup(iWhere)
		
End Function
'=======================================================================================================

Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)
	Dim IntRetCD	
	Dim iCalledAspName

	frm1.vspdData1.Col =  C_TEMP_Gl_FG1
    IF Trim(frm1.vspdData1.Text) = "T" THEN	
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	Else	
'   if Trim(frm1.vspdData1.Text) = "G" THEN
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
			IsOpenPop = False
			Exit Function
		End If	
	End If
	
	If IsOpenPop = True Then Exit Function

	With frm1.vspdData1
		.Row = .ActiveRow
		.Col =  C_GlNo1
		arrParam(0) = Trim(.Text)	'결의전표번호 
		arrParam(1) = ""			'Reference번호 

		if arrParam(0) = "" THEN Exit Function
			
	End With

	IsOpenPop = True

		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
End Function


'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				frm1.txtTransType.value = arrRet(0)
				frm1.txtTransTypeNm.value = arrRet(1)								
			Case 1
				frm1.txtBizCd.value  = arrRet(0)
				frm1.txtBizNm.value  = arrRet(1)			    
			Case 2
				frm1.txtTransType1.value = arrRet(0)
				frm1.txtTransTypeNm1.value = arrRet(1)								
			Case 3
				frm1.txtBizCd1.value  = arrRet(0)
				frm1.txtBizNm1.value  = arrRet(1)	
		End Select

	End With
	
End Function
'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtTransType.focus
			Case 1 
				.txtBizCd.focus
			Case 2
				.txtTransType1.focus
			Case 3
				.txtBizCd1.focus
		End Select    
	End With

End Function
'=======================================================================================================

Sub txtBizCd_onBlur()
	
	if frm1.txtBizCd.value = "" then
		frm1.txtBizNm.value = ""
	end if
End Sub	

Sub txtTransType_onBlur()
	
	if frm1.txtTransType.value = "" then
		frm1.txtTransTypeNm.value = ""
	end if
End Sub	

Sub txtVatType_onBlur()
	if frm1.txtVatType.value = "" then
		frm1.txtVatTypeNm.value = ""
	end if
End Sub	

'=======================================================================================================
Function OpenVatType(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
      
	
	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True
	
  Select Case iWhere
	Case 1			
	arrParam(0) = "부가세유형팝업"	                ' 팝업 명칭 
	arrParam(1) = "B_MINOR"			                	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType.Value)
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "			
	arrParam(5) = "부가세코드"			        '조건필드의 라벨 명칭 
	
    arrField(0) = "MINOR_CD"	                           ' Field명(0)
    arrField(1) = "MINOR_NM"	                           ' Field명(1)
    
    arrHeader(0) = "부가세유형"		               ' Header명(0)
    arrHeader(1) = "부가세유형명"		               ' Header명(1)
    
    Case 2
    arrParam(0) = "부가세유형팝업"	                ' 팝업 명칭 
	arrParam(1) = "B_MINOR"			                	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType1.Value)
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "			
	arrParam(5) = "부가세코드"			        '조건필드의 라벨 명칭 
	
    arrField(0) = "MINOR_CD"	                           ' Field명(0)
    arrField(1) = "MINOR_NM"	                           ' Field명(1)
    
    arrHeader(0) = "부가세유형"		               ' Header명(0)
    arrHeader(1) = "부가세유형명"		               ' Header명(1)
    
  End Select  
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetVatType(arrRet, iWhere)
	End If	
	Call FocusAfterVATPopup (iWhere)
	
End Function
'=======================================================================================================

Function FocusAfterVATPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1  
				.txtVatType.focus
			Case 2 
				.txtVatType1.focus
		End Select    
	End With

End Function
'=======================================================================================================
Function SetVatType(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
			frm1.txtVatType.Value    = arrRet(0)		
			frm1.txtVatTypeNm.Value    = arrRet(1)		
			lgBlnFlgChgValue = True
			Case 2
			frm1.txtVatType1.Value    = arrRet(0)		
			frm1.txtVatTypeNm1.Value    = arrRet(1)		
			lgBlnFlgChgValue = True
		End Select
	End With
End Function
'=======================================================================================================

Function OpenDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)
	
	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.readOnly = true then
		IsOpenPop = False
		Exit Function
	End If
	
	
'	iCalledAspName = AskPRAspName("DeptPopupDtA2")

'	If Trim(iCalledAspName) = "" Then
'		IntRetCD = DisplayMsgBox("900040", parent.Parent.VB_INFORMATION, "DeptPopupDtA2", "X")
'		IsOpenPop = False
'		Exit Function
'	End If

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.GIDate.Text
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  
	'arrParam(3) = "T"									' 결의일자 상태 Condition  

	arrParam(3) = "F"									' 결의일자 상태 Condition  

'	arrRet = window.showModalDialog(../../comasp/DeptPopupDtA2.asp, Array(window.parent, arrParam), _
'	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If	
	frm1.txtDeptCd.focus	
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDeptCd.Value = arrRet(0)
               .txtDeptNm.Value = arrRet(1)
               .txtInternalCd.Value = arrRet(2)
               .GIDate.text = arrRet(3)
				call txtDeptCd_OnChange()  
				
        End Select
	End With
End Function    
'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.GIDate.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.GIDate.Text, Parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.Value = ""
			frm1.txtDeptNm.Value = ""
			frm1.hOrgChangeId.Value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.Value = Trim(arrVal2(2))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub
 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function fnBttnConf()	
	Dim IntRetCd
	
	IntRetCD = DisplayMsgBox("112190", Parent.VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End if	
      
	fnBttnConf = False                                                          '⊙: Processing is NG
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value		  = Parent.UID_M0002
		.htxtWorkFg.value	  = "CONF"		
		.txtUpdtUserId.value  = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID    				
    END With
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID2)									'☜: 비지니스 ASP 를 가동 
    
    fnBttnConf = True             

End Function


Function fnBttnUnConf()
	Dim IntRetCd
	
	IntRetCD = DisplayMsgBox("112191", Parent.VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then
		Exit Function
	End if	
      
	fnBttnUnConf = False                                                          '⊙: Processing is NG
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value		  = Parent.UID_M0002
		.htxtWorkFg.value	  = "UNCONF"
		.txtUpdtUserId.value  = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID    				
    END With
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID2)									'☜: 비지니스 ASP 를 가동 
    
    fnBttnUnConf = True             

End Function


'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	
'   Call SetToolbar("1110000000001111")										'⊙: 버튼 툴바 제어 
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 	
	gSelframeFlg = TAB1

' click시 tab2의 조회조건을 옮겨온다.
	frm1.txtFromReqDt.text = frm1.txtFromReqDt1.text
	frm1.txtToReqDt.text = frm1.txtToReqDt1.text
	frm1.txtBizCd.value = frm1.txtBizCd1.value	
	frm1.txtBizNm.value = frm1.txtBizNm1.value	
	frm1.txtTransType.value = frm1.txtTransType1.value					 
	frm1.txtTransTypeNm.value = frm1.txtTransTypeNm1.value					 
End Function

Function ClickTab2()

'	If lgIntFlgMode <> Parent.OPMD_UMODE Then
'		Call SetToolBar("1110000000001111")
'	ELSE                 
'		Call SetToolBar("1111000000001111")
'	END IF	

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB2
	
	' click시 tab1의 조회조건을 옮겨온다.
	frm1.txtFromReqDt1.text = frm1.txtFromReqDt.text
	frm1.txtToReqDt1.text = frm1.txtToReqDt.text
	frm1.txtBizCd1.value = frm1.txtBizCd.value	
	frm1.txtBizNm1.value = frm1.txtBizNm.value	
	frm1.txtTransType1.value = frm1.txtTransType.value					 
	frm1.txtTransTypeNm1.value = frm1.txtTransTypeNm.value					 

	
End Function

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
			C_Confirm 			= iCurColumnPos(1)
			C_BatchDt 			= iCurColumnPos(2)
			C_Refno 			= iCurColumnPos(3)    
			C_BizCD				= iCurColumnPos(4)
			C_BizNm 			= iCurColumnPos(5)
			C_GLInputType 		= iCurColumnPos(6)
			C_GLInputTypeNm 	= iCurColumnPos(7)
			C_BatchNo  			= iCurColumnPos(8)
			C_ItemAmt  			= iCurColumnPos(9)
			C_ItemLocAmt  		= iCurColumnPos(10)
			C_GlDesc   			= iCurColumnPos(11)
		Case "B"
			ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Confirm1   			= iCurColumnPos(1)
			C_BatchDt1   			= iCurColumnPos(2)
			C_Chainno1   			= iCurColumnPos(3)    
			C_Refno1    			= iCurColumnPos(4)    
    		C_BizCD1    			= iCurColumnPos(5)
			C_BizNm1    			= iCurColumnPos(6)
			C_GLInputType1     		= iCurColumnPos(7)
			C_GLInputTypeNm1     	= iCurColumnPos(8)
			C_BatchNo1     			= iCurColumnPos(9)
			C_GlDt1     			= iCurColumnPos(10)
			C_GlNo1     			= iCurColumnPos(11)
			C_ItemAmt1     			= iCurColumnPos(12)
			C_ItemLocAmt1     		= iCurColumnPos(13)
			C_GlDesc1      			= iCurColumnPos(14)
			C_TEMP_Gl_FG1     		= iCurColumnPos(15)
		Case "C"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
           	C_Confirm2   			= iCurColumnPos(1)
			C_BatchDt2   			= iCurColumnPos(2)
			C_Chainno2   			= iCurColumnPos(3)    
			C_Refno2    			= iCurColumnPos(4)    
    		C_BizCD2    			= iCurColumnPos(5)
			C_BizNm2    			= iCurColumnPos(6)
			C_GLInputType2     		= iCurColumnPos(7)
			C_GLInputTypeNm2     	= iCurColumnPos(8)
			C_BatchNo2     			= iCurColumnPos(9)
			C_GlDt2     			= iCurColumnPos(10)
			C_GlNo2     			= iCurColumnPos(11)
			C_TEMP_Gl_FG2      		= iCurColumnPos(12)
   End Select    
End Sub


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

	'Set ggoSpread = CreateObject("Uni2KCM.Spread")
'	Call GetGlobalVar
'    Call ClassLoad                                                          '⊙: Load Common DLL    
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet("A")                                                    '⊙: Setup the Spread sheet
    Call InitSpreadSheet("B")
    Call InitSpreadSheet("C")

    Call InitVariables                                                      '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox_Cond
	Call SetDefaultVal
    Call SetToolbar("110000000000111")
    gSelframeFlg = TAB1        

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


 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt.focus
		Call FncQuery
	End If
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	End If
End Sub
Sub txtFromReqDt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt1.focus
		Call FncQuery
	End If
End Sub

Sub txtToReqDt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt1.focus
		Call FncQuery
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
     
    if frm1.vspdData.maxrows = 0 then exit sub
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If 

	Select Case Col
	
		Case C_Confirm 							
			ggoSpread.Source = frm1.vspdData
'			ggoSpread.UpdateRow Row	
			lgBlnFlgChgValue = True						
	End Select 	
	
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData1
    
     If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
 
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SP3C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData2
   
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
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_Cost_Nm Or NewCol <= C_Cost_Nm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_Cost_Nm Or NewCol <= C_Cost_Nm Then
     '   Cancel = True
     '   Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_Cost_Nm Or NewCol <= C_Cost_Nm Then
     '   Cancel = True
     '   Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row	
End Sub

'==========================================================================================
'   Event Name : vspdData_scriptLeaveCell
'   Event Desc : This event is spread sheet data Button Clicked
'==========================================================================================
Sub vspdData_scriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

	    If Row >= NewRow Then
			Exit Sub
		End If
		

	'	If NewRow = .MaxRows Then
	'        DbQuery
	'    End if    

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspdData_KeyPress(index , KeyAscii )
     lgBinFlgChgValue = True                                                 '⊙: Indicates that value changed
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow         
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyBatchNo <> "" Then                         
      	   Call DbQuery
    	End If
    End if
        
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow         
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyBatchNo <> "" Then                         
      	   Call DbQuery
    	End If
    End if
        
    
End Sub

'========================================================================================== 
' Event Name : vspdData1_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData1_scriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row <> NewRow And NewRow > 0 Then
    
       Call DbQuery2(NewRow)		       
    End If
End Sub
'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This event is spread sheet data BUTTON CLICK
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

	If Col = C_Confirm1 Then
		If BUTTONDOWN = 0 or BUTTONDOWN = 1 then
			Call SetVspdData2Checked(Row)
		End If
	End If

End Sub
'=======================================================================================================
'   Event Name : SetVspdData2Checked
'   Event Desc : Called When check box is clicked
'=======================================================================================================
Sub SetVspdData2Checked(Byval Row)
'$$
	Dim i
	Dim StrConf1
	Dim iCol, iRow
	Dim strChgFlag	
	frm1.vspdData1.Row	=	Row
	frm1.vspdData1.col	=	C_Confirm1
	
	IF 	frm1.vspdData1.text = "1" Then
			If frm1.vspdData2.MaxRows > 0 Then 
				For iRow = 1 To frm1.vspdData2.MaxRows
					frm1.vspdData2.Row = iRow
					frm1.vspdData2.col	= C_Confirm2
					frm1.vspdData2.text = "1" 
				Next
			End If	
	Else
		If frm1.vspdData2.MaxRows > 0 Then 
			For iRow = 1 To frm1.vspdData2.MaxRows
				frm1.vspdData2.Row = iRow
				frm1.vspdData2.col	= C_Confirm2	
				frm1.vspdData2.text = "0" 
			Next
		End If	
	End If			
		
	'//체크한목록이 있는지 확인 
	
	strChgFlag = False
	For iRow=1 To frm1.vspdData1.MaxRows 
		frm1.vspdData.row = iRow
		frm1.vspdData.Col = C_Confirm1
		If frm1.vspdData1.text = "1" Then
			strChgFlag = True
			Exit for
		End If	
	Next
	
	If strChgFlag = True Then
		lgBlnFlgChgValue = True
	Else
		lgBlnFlgChgValue = False
	End If
	
	
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
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    
    '-----------------------
    'Erase contents area
    '-----------------------
	If gSelframeFlg = TAB1 Then
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
	Else 
	
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
		
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

	End If
	    
    Call InitVariables 															'⊙: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------

    If gSelframeFlg = TAB1 Then
		If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		   Exit Function
		End If
    
		If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
		                    "970025",frm1.txtFromReqDt.UserDefinedFormat,Parent.gComDateType,True) = False Then	
			frm1.txtFromReqDt.focus
			Exit Function
		End If

	Else
		If Not chkField(Document, "2") Then									'⊙: This function check indispensable field
		   Exit Function
		End If
    
		If CompareDateByFormat(frm1.txtFromReqDt1.text,frm1.txtToReqDt1.text,frm1.txtFromReqDt1.Alt,frm1.txtToReqDt1.Alt, _
		                    "970025",frm1.txtFromReqDt1.UserDefinedFormat,Parent.gComDateType,True) = False Then	
			frm1.txtFromReqDt1.focus
			Exit Function
		End If

	End If

    
	lgQueryFlag = "New"		' 신규조회 및 추가조회 구분 Flag (현재는 신규임)
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X") '☜ 바뀐부분    
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
   
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True                                                           '⊙: Processing is OK

End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
	Dim ii
	Dim count
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------

    
	If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
		    IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
		    Exit Function
		End If

		If Not chkField(Document, "3") Then               '⊙: Check required field(Single area)
		   Exit Function
		End If

		frm1.vspddata.Col = C_Confirm
		count = 0
		For ii = 1 To frm1.vspddata.MaxRows 
			frm1.vspddata.Row = ii
			If frm1.vspddata.text = 1 Then 
				count = count + 1
			End If 
		Next 
		If count = 0 Then
		    IntRetCD = DisplayMsgBox("230118","X","X","X")  '☜ 바뀐부분 
		     Exit Function		
		End If
	Else
		frm1.vspddata1.Col = C_Confirm1
		count = 0
		For ii = 1 To frm1.vspddata1.MaxRows 
			frm1.vspddata1.Row = ii
			If frm1.vspddata1.text = 1 Then 
				count = count + 1
			End If 
		Next 
		If count = 0 Then
		    IntRetCD = DisplayMsgBox("230118","X","X","X")  '☜ 바뀐부분 
		     Exit Function		
		End If

	End If    

	IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    IF DbSave	= False Then			                                                  '☜: Save db data
		 Exit Function
    End If
    
   	
    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

    if frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function
'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    'On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    'On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 '☜: 화면 유형 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      '☜:화면 유형, Tab 유무 
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

	on Error Resume Next
	Err.Clear 

    ggoSpread.Source = gActiveSpdSheet
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")      
'			Call InitComboBox_cond
			Call ggoSpread.ReOrderingSpreadData()
		Case "VSPDDATA1"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")      
'			Call InitComboBox_cond
			Call ggoSpread.ReOrderingSpreadData()
		Case "VSPDDATA2"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("C")      
'			Call InitComboBox_cond
			Call ggoSpread.ReOrderingSpreadData()
	End Select			
End Sub


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    Call LayerShowHide(1)

    Err.Clear                                                               '☜: Protect system from crashing
    
    With frm1

    If gSelframeFlg = TAB1 Then
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID1 & "?txtMode=" & Parent.UID_M0001						'☜:조회표시 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&lgMaxCount="         & CStr(C_SHEETMAXROWS)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd="         & Trim(.hBizCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtTransType="   & Trim(.hGlTransType.value)
			strVal = strVal & "&txtFromReqDt="     & (.txtFromReqDt.Text)
			strVal = strVal & "&txtToReqDt="       & (.txtToReqDt.Text)
			strVal = strVal & "&txtVatType="       & (.txtVatType.value)
			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID1 & "?txtMode="     & Parent.UID_M0001						'☜:조회표시 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&lgMaxCount="         & CStr(C_SHEETMAXROWS)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd="         & Trim(.txtBizCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtTransType="   & Trim(.txtTransType.value)
			strVal = strVal & "&txtFromReqDt="     & (.txtFromReqDt.Text)		
			strVal = strVal & "&txtToReqDt="       & (.txtToReqDt.Text)
			strVal = strVal & "&txtVatType="       & (.txtVatType.value)
			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		End If
	Else

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID1 & "?txtMode=" & Parent.UID_M0004						'☜:조회표시 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&lgMaxCount="         & CStr(C_SHEETMAXROWS)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd1="         & Trim(.hBizCd.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtTransType1="   & Trim(.hGlTransType.value)
			strVal = strVal & "&txtFromReqDt1="     & (.txtFromReqDt1.Text)
			strVal = strVal & "&txtToReqDt1="       & (.txtToReqDt1.Text)
			strVal = strVal & "&txtVatType1="       & (.txtVatType1.value)
			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID1 & "?txtMode="     & Parent.UID_M0004						'☜:조회표시 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&lgMaxCount="         & CStr(C_SHEETMAXROWS)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd1="         & Trim(.txtBizCd1.value)	 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtTransType1="   & Trim(.txtTransType1.value)
			strVal = strVal & "&txtFromReqDt1="     & (.txtFromReqDt1.Text)		
			strVal = strVal & "&txtToReqDt1="       & (.txtToReqDt1.Text)
			strVal = strVal & "&txtVatType1="       & (.txtVatType1.value)
			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		End If
	End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동        
    End With    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Call LayerShowHide(0)

    Call SetToolbar("110010000001111")
   	
    If gSelframeFlg = TAB1 Then

	Else
        If frm1.vspdData1.MaxRows > 0 Then
            Call DbQuery2(frm1.vspdData1.ActiveRow)
        End If
	End If
				
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row)
	Dim UpperChainNo
	Dim UpperBatchNo
	Dim strVal

    Call LayerShowHide(1)	
	frm1.vspdData2.MaxRows = 0
	With frm1.vspdData1
		.Row = Row
		.Col = C_Chainno1
		UpperChainNo = .Text
		.Col = C_BatchNo1
		UpperBatchNo = .Text

		strVal = BIZ_PGM_ID1 & "?txtMode="    & Parent.UID_M0005						'☜:조회표시 			
		strVal = strVal & "&lgMaxCount="      & CStr(C_SHEETMAXROWS)	 			    '☆: 조회 조건 데이타 
		strVal = strVal & "&UpperChainNo="    & Trim(UpperChainNo)	 			    '☆: 조회 조건 데이타 
		strVal = strVal & "&UpperBatchNo="	  & Trim(UpperBatchNo)
		strVal = strVal & "&txtMaxRows="      & .MaxRows

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동        
    End With    
  
End Function 
'========================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk2()														'☆: 조회 성공후 실행로직 
    call SetVspdData2Checked(frm1.vspdData1.ActiveRow)
    Call LayerShowHide(0)
    Call SetToolbar("110010000001111")    
End Function

'========================================================================================
' Function Name : SetGridFocus
' Function Desc : This function is setting a cursor after query 
'========================================================================================
Function SetGridFocus()
	with frm1 
		.vspdData.Col = 1
		.vspdData.Row = 1
		.vspdData.Action = 1
	end with 
End Function 

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal

    DbSave = False                                                          '⊙: Processing is NG
    Call LayerShowHide(1)
    
    'On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1

		If gSelframeFlg = TAB1 Then
			.txtMode.value = Parent.UID_M0001
			.txtUpdtUserId.value = Parent.gUsrID
			.txtInsrtUserId.value = Parent.gUsrID
			
			'-----------------------
			'Data manipulate area
			'-----------------------
			lGrpCnt = 1
			strVal = ""

			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = C_Confirm
				If .vspdData.text = "1" THEN
						strVal = strVal &Parent.gColSep & lRow & Parent.gColSep	
						.vspdData.Col = C_BatchNo		'
						strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
						.vspdData.Col = C_Refno 		'
						strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
						lGrpCnt = lGrpCnt + 1					
				End if
			Next
			.txtMaxRows.value = lGrpCnt-1
			.txtSpread.value = strVal
		Else
			.txtMode.value = Parent.UID_M0002
			.txtUpdtUserId.value = Parent.gUsrID
			.txtInsrtUserId.value = Parent.gUsrID
			
			'-----------------------
			'Data manipulate area
			'-----------------------
			lGrpCnt = 1
			strVal = ""

			For lRow = 1 To .vspdData1.MaxRows
				.vspdData1.Row = lRow
				.vspdData1.Col = C_Confirm1
				If .vspdData1.text = "1" THEN
						strVal = strVal &Parent.gColSep & lRow & Parent.gColSep	
						.vspdData1.Col = C_BatchNo1		'4
						strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep
						lGrpCnt = lGrpCnt + 1					
				End if
			Next

			.txtMaxRows.value = lGrpCnt-1
			.txtSpread.value = strVal
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID2)									'☜: 비지니스 ASP 를 가동 
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    Call LayerShowHide(0)

	If gSelframeFlg = TAB1 Then
		Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
		gSelframeFlg = TAB2
		Call InitVariables	
		Call InitSpreadSheet("A")                                                    '⊙: Setup the Spread sheet		
		Call InitSpreadSheet("B")
		Call InitSpreadSheet("C")                                                    '⊙: Setup the Spread sheet		
		Call DBQuery()				
	Else
		Call InitVariables	
		Call InitSpreadSheet("A")                                                    '⊙: Setup the Spread sheet		
		Call InitSpreadSheet("B")
		Call InitSpreadSheet("C")                                                    '⊙: Setup the Spread sheet		
		Call DBQuery()				
	End If

End Function
'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()

End Function
'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'========================================================================================================= 
Sub fnBttnConf()	
	Dim ii 
	
	If gSelframeFlg = TAB1 Then
		With frm1
			For ii = 1 To .vspddata.MaxRows
				.vspddata.row = ii
				.vspddata.col = C_Confirm
				.vspddata.value = "1"
			Next	
		End With		
	else
		With frm1
			For ii = 1 To .vspddata1.MaxRows
				.vspddata1.row = ii
				.vspddata1.col = C_Confirm1
				.vspddata1.value = "1"
			Next	
		End With		
	end if
		
    lgBlnFlgChgValue = True	
End Sub

'========================================================================================================= 
Function fnBttnUnConf()
	Dim ii 
	If gSelframeFlg = TAB1 Then
		With frm1
			For ii = 1 To .vspddata.MaxRows
				.vspddata.row = ii
				.vspddata.col = C_Confirm
				.vspddata.value = "0"
			Next	
		End With
	else
		With frm1
			For ii = 1 To .vspddata1.MaxRows
				.vspddata1.row = ii
				.vspddata1.col = C_Confirm1
				.vspddata1.value = "0"
			Next	
		End With		
	end if
	
	lgBlnFlgChgValue = True		
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'----------  Coding part  -------------------------------------------------------------

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전표일괄생성</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>전표일괄생성취소</td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>						
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<!-- 첫번째 탭 내용  -->
			<DIV ID="TabDiv" SCROLL="no">			
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5"NOWRAP>발생일자</TD>
										<TD CLASS="TD6"NOWRAP>
											<script language =javascript src='./js/a5142ma1_fpDateTime1_txtFromReqDt.js'></script>
	~ 
											<script language =javascript src='./js/a5142ma1_fpDateTime2_txtToReqDt.js'></script>										
										<TD CLASS="TD5"NOWRAP>사업장</TD>
										<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizCd.Value, 1)">
											 <INPUT TYPE=TEXT ID="txtBizNm" NAME="txtBizNm" SIZE=20 tag="14X" ALT="사업장명">
										</TD>
									</TR>
									<TR>
										<TD CLASS="TD5"NOWRAP>거래유형</TD>
										<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTransType" SIZE=10  MAXLENGTH=10 tag="21XXXU" ALT="거래유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType.Value, 0)">
											 <INPUT TYPE=TEXT ID="txtTransTypeNm" NAME="txtTransTypeNm" SIZE=20 tag="14X" ALT="거래유형명">
										</TD>
										<TD CLASS="TD5" NOWRAP>부가세유형</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="부가세유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType(1)">&nbsp;
											<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="부가세유형"></TD>
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5"NOWRAP>전표일자</TD>
										<TD CLASS="TD6"NOWRAP COLSPAN = 3>
											<script language =javascript src='./js/a5142ma1_GIDate_GIDate.js'></script></TD>
										<TD CLASS=TD5 NOWRAP>부서</TD>								
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="32XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">
															 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="34X"></TD>
															 <INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
											
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=* valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>				
								<TR>
									<TD HEIGHT="100%"><script language =javascript src='./js/a5142ma1_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>							
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</DIV>			
			<!-- 두번째 탭 내용  -->	
			<DIV ID="TabDiv"  SCROLL="no">			
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5"NOWRAP>발생일자</TD>
										<TD CLASS="TD6"NOWRAP>
											<script language =javascript src='./js/a5142ma1_fpDateTime1_txtFromReqDt1.js'></script>
	~ 
											<script language =javascript src='./js/a5142ma1_fpDateTime2_txtToReqDt1.js'></script>										
										<TD CLASS="TD5"NOWRAP>사업장</TD>
										<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBizCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizCd.Value, 3)">
											 <INPUT TYPE=TEXT ID="txtBizNm1" NAME="txtBizNm1" SIZE=20 tag="14X" ALT="사업장명">
										</TD>
									</TR>
									<TR>
										</TD>
										<TD CLASS="TD5"NOWRAP>거래유형</TD>
										<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTransType1" SIZE=10  MAXLENGTH=10 tag="21XXXU" ALT="거래유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType1.Value, 2)">
											 <INPUT TYPE=TEXT ID="txtTransTypeNm1" NAME="txtTransTypeNm1" SIZE=20 tag="14X" ALT="거래유형명">
										<TD CLASS="TD5" NOWRAP>부가세유형</TD>
								        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="부가세유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType(2)">&nbsp;
											<INPUT TYPE=TEXT NAME="txtVatTypeNm1" SIZE=20 tag="24" ALT="부가세유형"></TD>	 
										</TD>									

									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=* valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>				
								<TR>
									<TD HEIGHT="50%"><script language =javascript src='./js/a5142ma1_vaSpread1_vspdData1.js'></script>
									</TD>
								</TR>		
								<TR>
									<TD HEIGHT="50%"><script language =javascript src='./js/a5142ma1_vaSpread1_vspdData2.js'></script>
									</TD>
								</TR>														
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</DIV>			
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE01%>></TD>
	</TR>			
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>				
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnConf" CLASS="CLSSBTN" OnClick="VBScript:Call fnBttnConf()" >일괄선택</BUTTON>&nbsp;<BUTTON NAME="btnUnCon" CLASS="CLSSBTN" OnClick="VBScript:Call fnBttnUnConf()">일괄취소</BUTTON></TD>		        					
					<TD WIDTH=10>&nbsp;</TD>
				</TR>	
			</TABLE>	
		</TD>						
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		<!--<TD WIDTH=100% HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>-->
	</TR>
</TABLE>
<TEXTAREA TABINDEX="-1" CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hGlTransType" tag="24">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24">
<INPUT TYPE=HIDDEN NAME="hBizCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtVatType" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtWorkFg" tag="24">

<script language =javascript src='./js/a5142ma1_fpDateTime1_hFromReqDt.js'></script>
<script language =javascript src='./js/a5142ma1_fpDateTime2_hToReqDt.js'></script>										
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
