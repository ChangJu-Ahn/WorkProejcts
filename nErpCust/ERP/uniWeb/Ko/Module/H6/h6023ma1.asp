<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 급여자동기표처리결과확인 
*  3. Program ID           	: H6023ma1
*  4. Program Name         	: H6023ma1
*  5. Program Desc         	: 급여관리 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2003/05/21
*  8. Modified date(Last)  	: 2003/06/13
*  9. Modifier (First)     	: SBK
* 10. Modifier (Last)     	: Lee SiNa
* 11. Comment              	:
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H6023mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                 
Dim gblnWinEvent                                                 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lsInternal_cd
Dim lgType

Dim C_BIZ_AREA_NM
Dim C_ACCNT_NM
Dim C_ALLOW_CD_NM
Dim C_DEPT_NM
Dim C_EMP_NO
Dim C_NAME
Dim C_ALLOW_AMT
Dim C_DED_AMT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
    If pvSpdNo = "A" Then
	    C_ACCNT_NM	    = 1  
	    C_BIZ_AREA_NM	= 2
	    C_ALLOW_AMT	    = 3
	    C_DED_AMT	    = 4
    ElseIf pvSpdNo = "B" Then
	    C_BIZ_AREA_NM	= 1
	    C_ACCNT_NM	    = 2  
	    C_DEPT_NM		= 3
	    C_ALLOW_AMT	    = 4	
	    C_DED_AMT	    = 5
    ElseIf pvSpdNo = "C" Then
	    C_BIZ_AREA_NM	= 1
	    C_ACCNT_NM	    = 2  
	    C_ALLOW_CD_NM	= 3     
	    C_DEPT_NM		= 4
	    C_EMP_NO		= 5       
	    C_NAME	        = 6
	    C_ALLOW_AMT	    = 7	
	    C_DED_AMT	    = 8
    End If
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
	lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtprov_dt.focus
	frm1.txtprov_dt.Text = UNIConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat , Parent.gDateFormat)
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim rbo_type

    If frm1.rbo_type(0).checked Then
        rbo_type="1"
    ElseIf frm1.rbo_type(1).checked THEN
        rbo_type="2"
    ElseIf frm1.rbo_type(2).checked THEN
        rbo_type="3"
    End If
    
    lgKeyStream  = frm1.txtprov_dt.Year & right("0" & frm1.txtprov_dt.Month,2) & right("0" & frm1.txtprov_dt.Day,2) & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtProv_type.value & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtBizAreaCd.value & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtAccntCd.value & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.cboAllow_kind.value & parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtAllow_cd.value  & parent.gColSep
    lgKeyStream  = lgKeyStream & rbo_type
End Sub 

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0121", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.cboAllow_kind, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	With frm1
        ggoSpread.Source = .vspdData

		For intRow = 1 To .vspdData.MaxRows			
			.vspdData.Row = intRow

			.vspdData.Col = C_DEPT_NM
			If Trim(.vspdData.Value) = "소계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Sub_Total)
			End If

			.vspdData.Col = C_ACCNT_NM
			If Trim(.vspdData.Value) = "소계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Sub_Total)
			End If 

			.vspdData.Col = C_ACCNT_NM
			If Trim(.vspdData.Value) = "합계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Total)
			End If 

			.vspdData.Col = C_BIZ_AREA_NM
			If Trim(.vspdData.Value) = "합계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Total)
			End If 

			.vspdData.Col = C_ACCNT_NM
			If Trim(.vspdData.Value) = "총계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Grand_Total)
			End If

			.vspdData.Col = C_BIZ_AREA_NM
			If Trim(.vspdData.Value) = "총계" Then
                Call SetSpreadBackColor(frm1.vspdData,intRow,-1,intRow,-1,parent.C_RGB_Grand_Total)
			End If
		Next

    End With
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables(lgType)	
	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20030520",,parent.gAllowDragDropSpread    

	    .ReDraw = false
	    .MaxCols = 0
        .MaxCols = C_DED_AMT + 1											'☜: 최대 Columns의 항상 1개 증가시킴 

	    .Col = .MaxCols													'공통콘트롤 사용 Hidden Column
        .ColHidden = True
        
        .MaxRows = 0	
        ggoSpread.ClearSpreadData
       
        Call  GetSpreadColumnPos(lgType)

		 If lgType = "A" Then
		    ggoSpread.SSSetEdit    C_ACCNT_NM,      "계정코드명",  30,,,50
		    ggoSpread.SSSetEdit    C_BIZ_AREA_NM,   "사업장명", 22,,,50
            ggoSpread.SSSetFloat   C_ALLOW_AMT,     "수당액(차변)" ,20, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat   C_DED_AMT,       "공제액(대변)" ,20, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		 ElseIf lgType = "B" Then
		    ggoSpread.SSSetEdit    C_BIZ_AREA_NM,   "사업장명", 20,,,50
		    ggoSpread.SSSetEdit    C_ACCNT_NM,      "계정코드명",  28,,,50
		    ggoSpread.SSSetEdit    C_DEPT_NM,       "부서명", 30,,,50
            ggoSpread.SSSetFloat   C_ALLOW_AMT,     "수당액(차변)" ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat   C_DED_AMT,       "공제액(대변)" ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		 ElseIf lgType = "C" Then
		    ggoSpread.SSSetEdit    C_BIZ_AREA_NM,   "사업장명", 14,,,50
		    ggoSpread.SSSetEdit    C_ACCNT_NM,      "계정코드명",  16,,,50
		    ggoSpread.SSSetEdit    C_ALLOW_CD_NM,   "수당/공제명", 16,,,50
		    ggoSpread.SSSetEdit    C_DEPT_NM,       "부서명", 16,,,50
		    ggoSpread.SSSetEdit    C_EMP_NO,        "사번", 13,,,50
		    ggoSpread.SSSetEdit    C_NAME,          "성명", 13,,,50
            ggoSpread.SSSetFloat   C_ALLOW_AMT,     "수당액(차변)" ,14, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat   C_DED_AMT,       "공제액(대변)" ,14, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		 End If
		 	 		
		.ReDraw = true
		
		Call SetSpreadLock 
	
	End With

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
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1         
        .vspdData.ReDraw = False
         ggoSpread.SSSetProtected	C_BIZ_AREA_NM	, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected	C_ACCNT_NM		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_ALLOW_CD_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_NAME		    , pvStartRow, pvEndRow  
		 ggoSpread.SSSetProtected	C_EMP_NO		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_DEPT_NM	    , pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_ALLOW_AMT	  	, pvStartRow, pvEndRow  
         ggoSpread.SSSetProtected	C_DED_AMT	  	, pvStartRow, pvEndRow  
        .vspdData.ReDraw = True
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                

    Select Case UCase(pvSpdNo)
       Case "A"
			C_ACCNT_NM		= iCurColumnPos(1)  
            C_BIZ_AREA_NM	= iCurColumnPos(2)
			C_ALLOW_AMT		= iCurColumnPos(3)
			C_DED_AMT		= iCurColumnPos(4)
       Case "B"
            C_BIZ_AREA_NM	= iCurColumnPos(1)
			C_ACCNT_NM		= iCurColumnPos(2)  
			C_DEPT_NM		= iCurColumnPos(3)
			C_ALLOW_AMT		= iCurColumnPos(4)
			C_DED_AMT		= iCurColumnPos(5)
       Case "C"
            C_BIZ_AREA_NM	= iCurColumnPos(1)
			C_ACCNT_NM		= iCurColumnPos(2)  
			C_ALLOW_CD_NM	= iCurColumnPos(3)
			C_DEPT_NM		= iCurColumnPos(4)
			C_EMP_NO		= iCurColumnPos(5)       
			C_NAME		    = iCurColumnPos(6)
			C_ALLOW_AMT		= iCurColumnPos(7)
			C_DED_AMT		= iCurColumnPos(8)  
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call AppendNumberPlace("7", "7", "3")
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    
    lgType = "A"
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    
    Call CookiePage (0)                                                             '☜: Check Cookie
    
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
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If   ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    														'⊙: Initializes local global variables
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If txtBizAreaCd_Onchange() Then          'enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtAccntCd_Onchange() Then          'enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtprov_type_Onchange() Then       
        Exit Function
    End if
    
    If txtAllow_cd_Onchange() Then       
        Exit Function
    End if

    Call InitVariables	
    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
          
    FncQuery = True																'☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If   LayerShowHide(1) = False Then
     	Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey="       & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
    
    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim IRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Call  ggoOper.LockField(Document, "Q")
    Call InitData()

    Set gActiveElement = document.ActiveElement   
    lgBlnFlgChgValue = False
    frm1.vspdData.Focus

End Function
	
Sub cboAllow_kind_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub cboProv_type_OnChange()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	    Case "1"
	        arrParam(0) = "수당코드팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtAllow_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value		' Name Cindition
	        arrParam(4) = "" ' Where Condition
	        arrParam(5) = "수당코드"			    ' TextBox 명칭 
	
            arrField(0) = "allow_cd"					' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당코드"				' Header명(0)
            arrHeader(1) = "수당코드명"			    ' Header명(1)
	    Case "2"
	        arrParam(0) = "사업장팝업"			' 팝업 명칭 
	        arrParam(1) = "B_BIZ_AREA"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtBizAreaCd.value		    ' Code Condition
	        arrParam(3) = ""		' Name Cindition
	        arrParam(4) = ""        ' Where Condition
	        arrParam(5) = "사업장코드"			    ' TextBox 명칭 
	
            arrField(0) = "BIZ_AREA_CD"					' Field명(0)
            arrField(1) = "BIZ_AREA_NM"				    ' Field명(1)
    
            arrHeader(0) = "사업장코드"				' Header명(0)
            arrHeader(1) = "사업장명"			    ' Header명(1)
	    Case "3"
	        arrParam(0) = "계정코드팝업"			' 팝업 명칭 
	        arrParam(1) = "A_JNL_ITEM"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtAccntCd.value		    ' Code Condition
	        arrParam(3) = ""
	        arrParam(4) = "JNL_TYPE = " & FilterVar("HR", "''", "S") & ""             ' Where Condition
	        arrParam(5) = "계정코드"			    ' TextBox 명칭 
	
            arrField(0) = "JNL_CD"					' Field명(0)
            arrField(1) = "JNL_NM"				    ' Field명(1)
    
            arrHeader(0) = "계정코드"				' Header명(0)
            arrHeader(1) = "계정코드명"			    ' Header명(1)
	    Case "4"
	        arrParam(0) = "지급구분팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtProv_type.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtAllow_nm.value		' Name Cindition
	        arrParam(4) = "MAJOR_CD = " & FilterVar("H0040", "''", "S") & ""          ' Where Condition
	        arrParam(5) = "지급구분"			    ' TextBox 명칭 
	
            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)
    
            arrHeader(0) = "지급구분"				' Header명(0)
            arrHeader(1) = "지급구분명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "1"
		        frm1.txtAllow_cd.focus
		    Case "2"
		        frm1.txtBizAreaCd.focus
		    Case "3"
		        frm1.txtAccntCd.focus
		    Case "4"
		        frm1.txtProv_type.focus
        End Select	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtAllow_cd.value = arrRet(0)
		        .txtAllow_nm.value = arrRet(1)		
		        .txtAllow_cd.focus
		    Case "2"
		        .txtBizAreaCd.value = arrRet(0)
		        .txtBizAreaNm.value = arrRet(1)		
		        .txtBizAreaCd.focus
		    Case "3"
		        .txtAccntCd.value = arrRet(0)
		        .txtAccntNm.value = arrRet(1)		
		        .txtAccntCd.focus
		    Case "4"
		        .txtProv_type.value = arrRet(0)
		        .txtProv_type_Nm.value = arrRet(1)		
		        .txtProv_type.focus
        End Select
	End With
End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_change
'   Event Desc :
'========================================================================================================
Function txtBizAreaCd_Onchange()
    Dim IntRetCd
    
    If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
    Else
        IntRetCd = CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA "," BIZ_AREA_CD= " & FilterVar(frm1.txtBizAreaCd.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("124200","X","X","X")	
			 frm1.txtBizAreaNm.value = ""
             frm1.txtBizAreaCd.focus
            Set gActiveElement = document.ActiveElement
            txtBizAreaCd_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtAccntCd_change
'   Event Desc :
'========================================================================================================
Function txtAccntCd_Onchange()
    Dim IntRetCd
    
    If frm1.txtAccntCd.value = "" Then
		frm1.txtAccntNm.value = ""
    Else
        IntRetCd = CommonQueryRs(" JNL_NM "," A_JNL_ITEM "," JNL_TYPE=" & FilterVar("HR", "''", "S") & " AND JNL_CD= " & FilterVar(frm1.txtAccntCd.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("120400","X","X","X")	
			 frm1.txtAccntNm.value = ""
             frm1.txtAccntCd.focus
            Set gActiveElement = document.ActiveElement
            txtAccntCd_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtAccntNm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtprov_type_change
'   Event Desc :
'========================================================================================================
Function txtprov_type_Onchange()
    Dim IntRetCd
    
    If frm1.txtprov_type.value = "" Then
		frm1.txtprov_type_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0040", "''", "S") & " AND MINOR_CD= " & FilterVar(frm1.txtprov_type.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800248","X","X","X")	
			 frm1.txtprov_type_nm.value = ""
             frm1.txtprov_type.focus
            Set gActiveElement = document.ActiveElement
            txtprov_type_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtprov_type_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtAllow_cd_change
'   Event Desc :
'========================================================================================================
Function txtAllow_cd_Onchange()
    Dim IntRetCd
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
    Else
        IntRetCd = CommonQueryRs("ALLOW_NM","HDA010T","PAY_CD=" & FilterVar("*", "''", "S") & "  AND ALLOW_CD= " & FilterVar(frm1.txtAllow_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
   		   	Call DisplayMsgBox("800092","X","X","X")	'수당정보에 등록되지 않은 코드입니다.
		    	
		   	frm1.txtAllow_nm.value = ""
            frm1.txtAllow_cd.focus
            Set gActiveElement = document.ActiveElement
            txtAllow_cd_Onchange = true 
                
            Exit Function          
        Else
		  	frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  

    Set gActiveElement = document.ActiveElement
    frm1.txtAllow_cd.focus

End Function

Sub rbo_type1_OnClick()
    lgType = "A"
    frm1.txtDrLocAmt.Value = 0
    frm1.txtCrLocAmt.Value = 0
    
    Call InitSpreadSheet
End Sub

Sub rbo_type2_OnClick()
    lgType = "B"
    frm1.txtDrLocAmt.Value = 0
    frm1.txtCrLocAmt.Value = 0

    Call InitSpreadSheet
End Sub

Sub rbo_type3_OnClick()
    lgType = "C"
    frm1.txtDrLocAmt.Value = 0
    frm1.txtCrLocAmt.Value = 0

    Call InitSpreadSheet
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000010111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	frm1.vspdData.Row = Row     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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
	Call InitData()
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

Sub txtprov_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtprov_dt.Action = 7
        frm1.txtprov_dt.focus
    End If
End Sub

Sub txtprov_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub
'========================================================================================================
' Name : OpenProveDt
' Desc : 최근 지급일 POPUP
'========================================================================================================
Function OpenProveDt(iWhere)
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급일팝업"		
	arrParam(1) = "지급일"
	arrParam(2) = "HDF110T"
	arrParam(3) = "PROV_DT"
	arrParam(4) = frm1.txtprov_dt.text	

	arrRet = window.showModalDialog(HRAskPRAspName("StandardDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtprov_dt.text = arrRet(0)
	end if
	frm1.txtprov_dt.focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급/상여자동기표내역조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD WIDTH=100% HEIGHT=20 VALIGN=TOP>
       					<FIELDSET CLASS="CLSFLD">
       						<TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
									<TD	CLASS="TD5"	NOWRAP>지급일자</TD>
									<TD	CLASS="TD6"	NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtprov_dt name=txtprov_dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="지급일자" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProveDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProveDt(0)"></td>
									<TD	CLASS="TD5"	NOWRAP>지급구분</TD>
									<TD	CLASS="TD6"	NOWRAP><INPUT NAME="txtProv_type" MAXLENGTH="1"	 SIZE="10" ALT ="지급구분" TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript:	OpenCondAreaPopup(4)">
														   <INPUT NAME="txtProv_type_nm" MAXLENGTH="20"	SIZE="20" ALT ="지급구분명" tag="14"></TD>
              					</TR>
					            <TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd" MAXLENGTH="10" SIZE=10 ALT ="사업장코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" onclick="vbscript: OpenCondAreaPopup('2')">
												           <INPUT NAME="txtBizAreaNm" MAXLENGTH="50" SIZE=20 ALT ="사업장명" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtAccntCd" MAXLENGTH="10" SIZE=10 ALT ="계정코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAccntCd" align=top TYPE="BUTTON" onclick="vbscript: OpenCondAreaPopup('3') ">
												           <INPUT NAME="txtAccntNm" MAXLENGTH="50" SIZE=20 ALT ="계정코드명" tag="14X"></TD>
              					</TR>
					            <TR>
              						<TD CLASS="TD5" NOWRAP>수당/공제구분</TD>
	                   				<TD CLASS="TD6"><SELECT NAME="cboAllow_kind" ALT="수당/공제구분" CLASS ="cbonormal" TAG="11"><OPTION VALUE=""></OPTION></SELECT></TD>
					                <TD CLASS="TD5" NOWRAP>수당/공제코드</TD>
					                <TD CLASS="TD6" NOWRAP><INPUT NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10 TYPE="Text"  ALT ="수당코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAllowCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
					                                       <INPUT NAME="txtAllow_nm" MAXLENGTH=20 SIZE=20 TYPE="Text"  ALT ="수당코드명" tag="14XXXU"></TD>
              					</TR>
					            <TR>
						  			<TD CLASS=TD5 NOWRAP>조회구분</TD>
				        	    	<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rbo_type" ID="rbo_type" VALUE="1" CLASS="RADIO" TAG="11" onclick=rbo_type1_OnClick() CHECKED><LABEL FOR="rbo_type1">계정별</LABEL>&nbsp;
				        	                             <INPUT TYPE="RADIO" NAME="rbo_type" ID="rbo_type" VALUE="2" CLASS="RADIO" TAG="11" onclick=rbo_type2_OnClick() ><LABEL FOR="rbo_type2">부서별</LABEL>&nbsp;
				        	                             <INPUT TYPE="RADIO" NAME="rbo_type" ID="rbo_type" VALUE="3" CLASS="RADIO" TAG="11" onclick=rbo_type3_OnClick() ><LABEL FOR="rbo_type3">개인별</LABEL></TD>
									<TD	CLASS="TD5"	NOWRAP>전표생성일자</TD>
									<TD	CLASS="TD6"	NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtTrans_dt name=txtTrans_dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="전표생성일자" tag="14X" VIEWASTEXT></OBJECT>');</SCRIPT>
              					</TR>
                            </TABLE>
						</FIELDSET>						        
				</TR>		   
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
		         	<TD WIDTH=100% HEIGHT=* valign=top>
		                <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4> 
					                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					          	</TD> 
					        </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>수당액(차변) 합계</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계" id=OBJECT3></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>공제액(대변) 합계</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계" id=OBJECT3></OBJECT>');</SCRIPT></TD>
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
	<TR >
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

