<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 
*  3. Program ID           	: h3002ma1
*  4. Program Name         	: 조직개편후 일괄발령 등록 
*  5. Program Desc         	: 조직개편후 일괄발령 등록,변경,삭제 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/30
*  8. Modified date(Last)  	: 2003/06/10
*  9. Modifier (First)     	: TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      	: Lee SiNa
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const BIZ_PGM_ID      = "h3002mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h3002mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                           '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                          'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow

Dim C_NAME
Dim C_EMP_NO
Dim C_PAY_GRD1
Dim C_PAY_GRD1_NM
Dim C_PAY_GRD2
Dim C_ROLL_PSTN
Dim C_ROLL_PSTN_NM
Dim C_DEPT_CD
Dim C_DEPT_CD_NM
Dim C_CHNG_DEPT_CD
Dim C_CHNG_DEPT_CD_NM
Dim C_CHNG_DEPT_CD_POP
Dim C_CHNG_CD
Dim C_CHNG_CD_NM
Dim C_FUNC_CD
Dim C_ROLE_CD
Dim C_COMP_CD
Dim C_SECT_CD
Dim C_WK_AREA_CD
Dim C_FLAG

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_NAME				= 1
	 C_EMP_NO			= 2
	 C_PAY_GRD1			= 3
	 C_PAY_GRD1_NM		= 4
	 C_PAY_GRD2			= 5
	 C_ROLL_PSTN		= 6
	 C_ROLL_PSTN_NM		= 7
	 C_DEPT_CD			= 8
	 C_DEPT_CD_NM		= 9
	 C_CHNG_DEPT_CD		= 10
	 C_CHNG_DEPT_CD_NM	= 11
	 C_CHNG_DEPT_CD_POP = 12
	 C_CHNG_CD			= 13
	 C_CHNG_CD_NM		= 14
	 C_FUNC_CD			= 15
	 C_ROLE_CD			= 16
	 C_COMP_CD			= 17
	 C_SECT_CD			= 18
	 C_WK_AREA_CD		= 19
	 C_FLAG				= 20	  
End Sub
'========================================================================================================

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
	lgOldRow = 0

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
               
	Dim strYear,strMonth,strDay
	Dim orgDt
	Dim IntRetCD

	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtStd_dt.focus 		

    IntRetCD =  CommonQueryRs(" ORGDT "," horg_abs "," 1=1 order by ORGDT desc" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If IntRetCD = True  Then
		orgDt =Trim(Replace(lgF0,Chr(11),""))
		frm1.txtStd_dt.Year = mid(orgDt,1,4) 		 '년월일 default value setting
		frm1.txtStd_dt.Month = mid(orgDt,5,2) 
		frm1.txtStd_dt.Day = mid(orgDt,7,2)        
    Else
		frm1.txtStd_dt.Year = strYear 		 '년월일 default value setting
		frm1.txtStd_dt.Month = strMonth 
		frm1.txtStd_dt.Day = strDay    
    End If

	frm1.txtGazet_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtGazet_dt.Month = strMonth 
	frm1.txtGazet_dt.Day = strDay
End Sub	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    lgKeyStream       = Trim(Frm1.txtDept_cd.value) & parent.gColSep      
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtGazet_dt.Text) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtDept_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtDept_nm.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtchng_dept_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtchng_dept_nm.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtChng_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtChng_nm.value) & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_FLAG + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0	
 		ggoSpread.ClearSpreadData  
		      
       Call  GetSpreadColumnPos("A")

             ggoSpread.SSSetEdit     C_NAME, "성명", 13,,,30,2
             ggoSpread.SSSetEdit     C_EMP_NO, "사번", 15,,,13,2
             ggoSpread.SSSetEdit     C_PAY_GRD1, "", 20
             ggoSpread.SSSetEdit     C_PAY_GRD1_NM, "급호", 15,,,50,2
             ggoSpread.SSSetEdit     C_PAY_GRD2, "호봉", 8,,,4,2
             ggoSpread.SSSetEdit     C_ROLL_PSTN, "", 15
             ggoSpread.SSSetEdit     C_ROLL_PSTN_NM, "직위", 15,,,50,2
             ggoSpread.SSSetEdit     C_DEPT_CD, "", 20
             ggoSpread.SSSetEdit     C_DEPT_CD_NM, "현재부서", 15,,,40,2
             ggoSpread.SSSetEdit     C_CHNG_DEPT_CD, "", 12
             ggoSpread.SSSetEdit     C_CHNG_DEPT_CD_NM, "발령부서", 15,,,40,2
             ggoSpread.SSSetButton   C_CHNG_DEPT_CD_POP    
             ggoSpread.SSSetEdit     C_CHNG_CD, "", 15
             ggoSpread.SSSetEdit     C_CHNG_CD_NM, "변동사유", 18,,,50,2
             ggoSpread.SSSetEdit     C_FUNC_CD,   "", 15
             ggoSpread.SSSetEdit     C_ROLE_CD,   "", 15
             ggoSpread.SSSetEdit     C_COMP_CD,   "", 10
             ggoSpread.SSSetEdit     C_SECT_CD,   "", 10
             ggoSpread.SSSetEdit     C_WK_AREA_CD,   "", 10
             ggoSpread.SSSetEdit     C_FLAG,   "", 10
             
             Call ggoSpread.SSSetColHidden(C_PAY_GRD1		,  C_PAY_GRD1		, True)
             Call ggoSpread.SSSetColHidden(C_ROLL_PSTN		,  C_ROLL_PSTN		, True)
             Call ggoSpread.SSSetColHidden(C_DEPT_CD		,  C_DEPT_CD		, True)
             Call ggoSpread.SSSetColHidden(C_CHNG_DEPT_CD	,  C_CHNG_DEPT_CD	, True)
             Call ggoSpread.SSSetColHidden(C_CHNG_CD		,  C_CHNG_CD		, True)
             Call ggoSpread.SSSetColHidden(C_FUNC_CD		,  C_FUNC_CD		, True)
             Call ggoSpread.SSSetColHidden(C_ROLE_CD		,  C_ROLE_CD		, True)
             Call ggoSpread.SSSetColHidden(C_COMP_CD		,  C_COMP_CD		, True)
             Call ggoSpread.SSSetColHidden(C_SECT_CD		,  C_SECT_CD		, True)
             Call ggoSpread.SSSetColHidden(C_WK_AREA_CD		,  C_WK_AREA_CD		, True)
             Call ggoSpread.SSSetColHidden(C_FLAG			,  C_FLAG			, True)             
             
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()

    With frm1
    
    .vspdData.ReDraw = False
         ggoSpread.SpreadLock    C_NAME			, -1, C_NAME
         ggoSpread.SpreadLock    C_EMP_NO		, -1, C_EMP_NO
         ggoSpread.SpreadLock    C_PAY_GRD1		, -1, C_PAY_GRD1
         ggoSpread.SpreadLock    C_PAY_GRD1_NM	, -1, C_PAY_GRD1_NM
         ggoSpread.SpreadLock    C_PAY_GRD2		, -1, C_PAY_GRD2
         ggoSpread.SpreadLock    C_ROLL_PSTN	, -1, C_ROLL_PSTN
         ggoSpread.SpreadLock    C_ROLL_PSTN_NM	, -1, C_ROLL_PSTN_NM
         ggoSpread.SpreadLock    C_DEPT_CD		, -1, C_DEPT_CD
         ggoSpread.SpreadLock    C_DEPT_CD_NM	, -1, C_DEPT_CD_NM
         ggoSpread.SpreadLock    C_CHNG_CD_NM	, -1, C_CHNG_CD_NM
         ggoSpread.SpreadLock    C_CHNG_DEPT_CD_NM	, -1, C_CHNG_DEPT_CD_NM
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1            
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1

    .vspdData.ReDraw = False
         ggoSpread.SSSetProtected	C_DEPT_CD_NM		, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_CHNG_DEPT_CD_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_NAME				, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_EMP_NO			, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_PAY_GRD1_NM		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_PAY_GRD2			, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_ROLL_PSTN_NM		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_DEPT_CD_NM		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_CHNG_CD_NM		, pvStartRow, pvEndRow
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
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_NAME				= iCurColumnPos(1)
			C_EMP_NO			= iCurColumnPos(2)
			C_PAY_GRD1			= iCurColumnPos(3)
			C_PAY_GRD1_NM		= iCurColumnPos(4)
			C_PAY_GRD2			= iCurColumnPos(5)
			C_ROLL_PSTN			= iCurColumnPos(6)
			C_ROLL_PSTN_NM		= iCurColumnPos(7)
			C_DEPT_CD			= iCurColumnPos(8)
			C_DEPT_CD_NM		= iCurColumnPos(9)
			C_CHNG_DEPT_CD		= iCurColumnPos(10)
			C_CHNG_DEPT_CD_NM	= iCurColumnPos(11)
			C_CHNG_DEPT_CD_POP	= iCurColumnPos(12)
			C_CHNG_CD			= iCurColumnPos(13)
			C_CHNG_CD_NM		= iCurColumnPos(14)
			C_FUNC_CD			= iCurColumnPos(15)
			C_ROLE_CD			= iCurColumnPos(16)
			C_COMP_CD			= iCurColumnPos(17)
			C_SECT_CD			= iCurColumnPos(18)
			C_WK_AREA_CD		= iCurColumnPos(19)
			C_FLAG				= iCurColumnPos(20)	          
            
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

	Call  ggoOper.FormatDate(frm1.txtStd_dt,  parent.gDateFormat, 1)
	Call  ggoOper.FormatDate(frm1.txtGazet_dt,  parent.gDateFormat, 1)
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetDefaultVal
	Call SetToolbar("1000100100001111")												'⊙: Set ToolBar
    
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
' Desc : 저장후 쿼리실행 함수(Condition 영역에서 Enter버튼 액션을 받지 않기 위함.
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    ggoSpread.ClearSpreadData
   
End Function

'========================================================================================================
' Name : FncQuery1
' Desc : 자동입력 버튼으로 실행되는 쿼리 
'========================================================================================================
Function FncQuery1()
    Dim IntRetCD 
    
    FncQuery1 = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	

    If txtDept_cd_OnChange() Then
        Exit Function
    End If
    If txtchng_dept_cd_OnChange() Then
       Exit Function
    End If
    If txtChng_cd_OnChange() Then
        Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    
	If DbQuery1 = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
    
    FncQuery1 = True																'☜: Processing is OK

End Function

'========================================================================================================
' Name : FncQuery2
' Desc : 저장후 쿼리실행 함수(Condition 영역에서 Enter버튼 액션을 받지 않기 위함.
'========================================================================================================
Function FncQuery2()
    Dim IntRetCD 
    
    FncQuery2 = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    With frm1
    
    .vspdData.ReDraw = False         
         
         ggoSpread.SpreadLock    C_NAME			, -1, C_NAME
         ggoSpread.SpreadLock    C_EMP_NO		, -1, C_EMP_NO
         ggoSpread.SpreadLock    C_PAY_GRD1		, -1, C_PAY_GRD1
         ggoSpread.SpreadLock    C_PAY_GRD1_NM	, -1, C_PAY_GRD1_NM
         ggoSpread.SpreadLock    C_PAY_GRD2		, -1, C_PAY_GRD2
         ggoSpread.SpreadLock    C_ROLL_PSTN	, -1, C_ROLL_PSTN
         ggoSpread.SpreadLock    C_ROLL_PSTN_NM	, -1, C_ROLL_PSTN_NM
         ggoSpread.SpreadLock    C_DEPT_CD		, -1, C_DEPT_CD
         ggoSpread.SpreadLock    C_CHNG_CD_NM	, -1, C_CHNG_CD_NM
         
         
    .vspdData.ReDraw = True

    End With


    Call MakeKeyStream("X")
    
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
      
    FncQuery2 = True																'☜: Processing is OK
   
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1000100100001111")												'⊙: Set ToolBar
    Call SetDefaultVal
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd
    
    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
	Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
    
    FncDelete=  True                                                              '☜: Processing is OK
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
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    
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

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
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
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1            
        
       .vspdData.ReDraw = True
    End With

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
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	 ggoSpread.Source = frm1.vspdData 
    	lDelRows =  ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call SetDefaultVal
    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call SetDefaultVal
    Call InitVariables														 '⊙: Initializes local global variables

     If LayerShowHide(1) = False then
    	Exit Function 
    End if
     
    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction
    
	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncNext = True                                                               '☜: Processing is OK
	
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
' Desc : This function is called by FncQuery1
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbQuery1
' Desc : This function is called by FncQuery1
'========================================================================================================
Function DbQuery1()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery1 = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal      & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal      & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal      & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery1 = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
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
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) = False then
    		Exit Function 
    	End if
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text

               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                          strVal = strVal & "C" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & Trim(.txtGazet_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD_NM        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                  
                    .vspdData.Col = C_FUNC_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COMP_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SECT_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_AREA_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & Trim(.txtGazet_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & Trim(.txtGazet_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If LayerShowHide(1) = False then
    		Exit Function 
    	End if
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

    Dim  lRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Frm1.txtStd_dt.focus()

	Call SetToolbar("1000100100001111")												'⊙: Set ToolBar

    frm1.vspdData.ReDraw = false
     ggoSpread.Source = frm1.vspdData
     ggoSpread.SpreadLock    C_NAME, -1, C_NAME
     ggoSpread.SpreadLock    C_EMP_NO, -1, C_EMP_NO
     ggoSpread.SpreadLock    C_PAY_GRD1, -1, C_PAY_GRD1
     ggoSpread.SpreadLock    C_PAY_GRD1_NM, -1, C_PAY_GRD1_NM
     ggoSpread.SpreadLock    C_PAY_GRD2, -1, C_PAY_GRD2
     ggoSpread.SpreadLock    C_ROLL_PSTN, -1, C_ROLL_PSTN
     ggoSpread.SpreadLock    C_ROLL_PSTN_NM, -1, C_ROLL_PSTN_NM
     ggoSpread.SpreadLock    C_DEPT_CD, -1, C_DEPT_CD
     ggoSpread.SpreadLock    C_DEPT_CD_NM, -1, C_DEPT_CD_NM
     ggoSpread.SpreadLock    C_CHNG_DEPT_CD, -1, C_CHNG_DEPT_CD     
     ggoSpread.SpreadLock    C_CHNG_DEPT_CD_NM, -1, C_CHNG_DEPT_CD_NM
     ggoSpread.SpreadLock    C_CHNG_DEPT_CD_POP, -1, C_CHNG_DEPT_CD_POP
	 ggoSpread.SpreadLock    C_CHNG_CD, -1, C_CHNG_CD
     ggoSpread.SpreadLock    C_CHNG_CD_NM, -1, C_CHNG_CD_NM     
     
    frm1.vspdData.ReDraw = True

    Call  ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful(for MB2)
'========================================================================================================
Function DbQueryOk1()

    Dim  lRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    Frm1.txtStd_dt.focus()

	Call SetToolbar("1000100100001111")												'⊙: Set ToolBar

    With Frm1                                                                     'Fetch the data to handle by batch.  
        .vspdData.ReDraw = false
          ggoSpread.Source = .vspdData
         
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = C_FLAG
            If .vspdData.Text = "_i_" Then
                .vspdData.Col = 0
                .vspdData.Text =  ggoSpread.InsertFlag
            Else
                .vspdData.Col = 0
                .vspdData.Text =  ggoSpread.UpdateFlag
            End If
       Next

          ggoSpread.SSSetRequired    C_CHNG_DEPT_CD_NM, -1, -1              
		  ggoSpread.SpreadLock		 C_DEPT_CD_NM	  , -1, -1
          ggoSpread.SpreadUnLock     C_CHNG_DEPT_CD_POP, -1, C_CHNG_DEPT_CD_POP
        .vspdData.ReDraw = True
    End With    
    
	Frm1.btnCb_autoisrt.disabled = False
	ggoSpread.ClearSpreadData "T"

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData  
    Call FncQuery2()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
End Function

'========================================================================================================
' Name : FncOpenPopup()
' Desc : Code PopUp at condition area 
'========================================================================================================
Function FncOpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	
	Select Case iWhere 
	    Case "1"
	    
	        arrParam(0) = "부서코드 팝업"			                                     ' 팝업 명칭 
	    	arrParam(1) = "b_acct_dept"						                                 ' TABLE 명칭 
	    	arrParam(2) = frm1.txtDept_cd.value                           	        		 ' Code Condition
	    	arrParam(3) = ""'frm1.txtDept_nm.value				                             ' Name Cindition
	    	arrParam(4) = " org_change_dt = (select max(org_change_dt) from b_acct_dept "    ' Where Condition
	    	arrParam(4) = arrParam(4) & "  where org_change_dt <  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"
	    	arrParam(5) = "부서코드" 			                                ' TextBox 명칭 
	
	    	arrField(0) = "DEPT_CD"						                        	' Field명(0)
	    	arrField(1) = "DEPT_NM"    					    	                    ' Field명(1)
    
	    	arrHeader(0) = "부서코드"	   		    	                        ' Header명(0)
	    	arrHeader(1) = "부서명"	    		                                ' Header명(1)
	    Case "2"

	        arrParam(0) = "부서코드 팝업"			                                    ' 팝업 명칭 
	    	arrParam(1) = "b_acct_dept"						                                ' TABLE 명칭 
	    	arrParam(2) = frm1.txtchng_dept_cd.value                                    	' Code Condition
	    	arrParam(3) = ""'frm1.txtchng_dept_nm.value                    		        	' Name Cindition
	    	arrParam(4) = " org_change_dt = (select min(org_change_dt) from b_acct_dept "   ' Where Condition
	    	arrParam(4) = arrParam(4) & "  where org_change_dt >=  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"
	    	arrParam(5) = "부서코드" 			                                ' TextBox 명칭 
	    	arrField(0) = "DEPT_CD"					                    	    	' Field명(0)
	    	arrField(1) = "DEPT_NM"    			                    		    	' Field명(1)
	    	arrHeader(0) = "부서코드"	   		    	                        ' Header명(0)
	    	arrHeader(1) = "부서명"	    		                                ' Header명(1)
	    Case "3"

	        arrParam(0) = "발령코드조회 팝업"	    ' 팝업 명칭 
	        arrParam(1) = "B_Minor"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtChng_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtChng_nm.value			' Name Cindition
	        arrParam(4) = " Major_cd=" & FilterVar("H0029", "''", "S") & " "			' Where Condition
	        arrParam(5) = "발령코드"			    ' TextBox 명칭 
	
            arrField(0) = "Minor_cd"					' Field명(0)
            arrField(1) = "Minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "발령코드"				' Header명(0)
            arrHeader(1) = "발령사유명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "1"
		        frm1.txtDept_cd.focus	
		    Case "2"
		        frm1.txtChng_dept_cd.focus
		    Case "3"
		        frm1.txtChng_cd.focus
        End Select	
		Exit Function
	Else
		Call SubSetOpenPop(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SubSetOpenPop()
'	Description : setting popUp at condition area
'=======================================================================================================
Sub SubSetOpenPop(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtDept_cd.value = arrRet(0)
		        .txtDept_nm.value = arrRet(1)	
		        .txtDept_cd.focus	
		    Case "2"
		        .txtChng_dept_cd.value = arrRet(0)
		        .txtChng_dept_nm.value = arrRet(1)		
		        .txtChng_dept_cd.focus
		    Case "3"
		        .txtChng_cd.value = arrRet(0)
		        .txtChng_nm.value = arrRet(1)		
		        .txtChng_cd.focus
        End Select
	End With
End Sub

'======================================================================================================
'	Name : OpenEmp()
'	Description : Employee PopUp
'======================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
		If frm1.txtEmp_no.value="" Then
    	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
    	Else
    	    arrParam(1) = ""                		' Name Cindition
        End If    	
	Else 'spread
	    frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	    frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent ,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtEmp_no.focus
		Else 'spread
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'======================================================================================================
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else 'spread
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case C_CHNG_DEPT_CD_POP

	        arrParam(0) = "부서코드 팝업"			                    ' 팝업 명칭 
	    	arrParam(1) = "b_acct_dept"						                ' TABLE 명칭 
	    	arrParam(2) = ""            			                        ' Code Condition
	    	arrParam(3) =  strCode								            ' Name Cindition
	    	
	    	arrParam(4) = " org_change_dt = (select min(org_change_dt) from b_acct_dept "   ' Where Condition
	    	arrParam(4) = arrParam(4) & "  where org_change_dt >=  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"
	    	arrParam(5) = "부서코드" 			            ' TextBox 명칭 
	    	arrField(0) = "dept_cd"						    	' Field명(0)
	    	arrField(1) = "dept_nm"    					    	' Field명(1)
    
	    	arrHeader(0) = "부서코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "부서명"	    		            ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_CHNG_DEPT_CD_NM
		frm1.vspdData.action =0
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_CHNG_DEPT_CD_POP
		        .vspdData.Col = C_CHNG_DEPT_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_CHNG_DEPT_CD_NM
		    	.vspdData.text = arrRet(1)   
		    	.vspdData.action =0
        End Select

	End With

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strCode
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col-1
    strCode = frm1.vspdData.text
	Select Case Col
	    Case C_CHNG_DEPT_CD_POP
                    Call OpenCode(strCode, C_CHNG_DEPT_CD_POP, Row)
    End Select    
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    Dim iDx
    Dim strWhere
    Dim IntRetCD
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	Select Case Col
	    Case C_CHNG_DEPT_CD_NM
            If Frm1.vspdData.Text = "" Then
                    Frm1.vspdData.Text = ""
    	            frm1.vspdData.Col = C_CHNG_DEPT_CD
                    frm1.vspdData.Text=""
	        Else
	    	        strWhere = " org_change_dt = (select min(org_change_dt) from b_acct_dept "   ' Where Condition
	    	        strWhere = strWhere & "  where org_change_dt >=  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"
					'strWhere = strWhere & "  where org_change_dt >= Convert(DateTime,'" & Trim(frm1.txtStd_dt.Text) & "'))"

                    IntRetCD =  CommonQueryRs(" dept_cd,dept_nm "," b_acct_dept "," dept_nm =  " & FilterVar(frm1.vspdData.Text, "''", "S") & " And " & strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    If IntRetCD=False And Trim(frm1.vspdData.Text)<>""  Then
                        Call  DisplayMsgBox("800054","X","X","X")                                 '☜ : 등록되지 않은 코드입니다.
                        frm1.vspdData.Text=""
                    ElseIf  parent.CountStrings(lgF0, Chr(11) ) > 1 Then        ' 같은명일 경우 pop up
                        Call OpenCode(frm1.vspdData.Text, C_CHNG_DEPT_CD_POP, Row)
                    Else
    	                    frm1.vspdData.Col = C_CHNG_DEPT_CD
                            frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
                    End If
            End If
            
    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
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
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
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


Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button), txtGazet_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStd_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtStd_dt.Action = 7
        frm1.txtStd_dt.focus
    End If
End Sub

Sub txtGazet_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtGazet_dt.Action = 7
        frm1.txtGazet_dt.focus
    End If
End Sub


'==========================================================================================
'   Event Name : btnCb_autoisrt_OnClick()
'   Event Desc : 자동입력 
'==========================================================================================
Sub btnCb_autoisrt_OnClick()
    Dim IntRetCD
    IntRetCD = FncQuery1()
    If IntRetCD = True Then
   	   Frm1.btnCb_autoisrt.disabled = True
    End If
End Sub
'======================================================================================================
'   Event Name : txtDept_cd_OnChange
'   Event Desc : 발령전부서가 변경될 경우 
'=======================================================================================================
Function txtDept_cd_OnChange()
    Dim IntRetCd
    Dim strWhere
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    If  frm1.txtDept_cd.value = "" Then
        frm1.txtDept_cd.Value=""
        frm1.txtDept_nm.Value=""
    Else
        strWhere = " dept_cd= " & FilterVar(frm1.txtDept_cd.Value, "''", "S") & " And org_change_dt = (select max(org_change_dt) from b_acct_dept "
		strWhere = strWhere & "  where org_change_dt <  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"

        IntRetCD =  CommonQueryRs(" dept_cd,dept_nm "," b_acct_dept ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtDept_cd.Value)<>""  Then
            Call  DisplayMsgBox("800438","X","X","X")                         '☜ : 조직개편일자이전으로 부서가 없습니다.
            frm1.txtDept_nm.Value=""
            frm1.txtDept_cd.focus
            txtDept_cd_OnChange = True
        ElseIf parent.CountStrings(lgF0, Chr(11) ) > 1 Then                         ' 같은명일 경우 pop up
            Call FncOpenPopup(1)
            txtDept_cd_OnChange = True
        Else
            frm1.txtDept_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
     End If
End Function
'======================================================================================================
'   Event Name : txtchng_dept_cd_OnChange
'   Event Desc : 발령후부서가 변경될 경우 
'=======================================================================================================
Function txtchng_dept_cd_OnChange()
    Dim IntRetCd
    Dim strWhere
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    If  frm1.txtchng_dept_cd.value = "" Then
        frm1.txtchng_dept_cd.Value=""
        frm1.txtchng_dept_nm.Value=""
    Else
        strWhere = " dept_cd= " & FilterVar(frm1.txtchng_dept_cd.Value, "''", "S") & " And org_change_dt = (select min(org_change_dt) from b_acct_dept "
		strWhere = strWhere & "  where org_change_dt >=  " & FilterVar(UNIConvDate(frm1.txtStd_dt.Text), "''", "S") & ")"
        IntRetCD =  CommonQueryRs(" dept_cd,dept_nm "," b_acct_dept ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtchng_dept_cd.Value)<>""  Then
            Call  DisplayMsgBox("800437","X","X","X")                         '☜ : 조직개편일자이후로 부서가 없습니다.
            frm1.txtchng_dept_nm.Value=""
            frm1.txtchng_dept_cd.focus
            txtchng_dept_cd_OnChange = True
        ElseIf parent.CountStrings(lgF0, Chr(11) ) > 1 Then    ' 동명이 나올경우 popup을 띄어줌 
            Call FncOpenPopup(2)
        Else
            frm1.txtchng_dept_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
     End If
End Function
'======================================================================================================
'   Event Name : txtChng_cd_OnChange
'   Event Desc : 변동사유가 변경될 경우 
'=======================================================================================================
Function txtChng_cd_OnChange()
    Dim IntRetCd
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    If  frm1.txtChng_cd.value = "" Then
        frm1.txtChng_cd.Value=""
        frm1.txtChng_nm.Value=""
    Else
        IntRetCD =  CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0029", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtChng_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtChng_cd.Value)<>""  Then
            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtChng_nm.Value=""
            frm1.txtChng_cd.focus
             txtChng_cd_OnChange = True           
        ElseIf  parent.CountStrings(lgF0, Chr(11) ) > 1 Then                         ' 같은명일 경우 pop up
            Call FncOpenPopup(3)
            txtChng_cd_OnChange = True
        Else
            frm1.txtChng_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
     End If
End Function


'==========================================================================================
'   Event Name : txtStd_dt_Keypress()
'   Event Desc : 
'==========================================================================================
Sub txtStd_dt_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub

Sub txtGazet_dt_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub

'========================================================================================================
' Name : OpenOrgId
' Desc : 기준일 POPUP
'========================================================================================================
Function OpenOrgId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "부서개편ID 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "horg_abs"					<%' TABLE 명칭 %>

	arrParam(2) = ""						<%' Code Condition%>
	
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "부서개편일자"					<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ORGDT"					<%' Field명(0)%>
    arrField(1) = "ORGNM"					<%' Field명(1)%>
    
    arrHeader(0) = "부서개편일자"				<%' Header명(0)%>
    arrHeader(1) = "부서개편명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtStd_dt.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtStd_dt.text = UNIDateClientFormat(arrRet(0))
	End If	
	
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %> ></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>조직개편후일괄발령</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>조직개편일</TD>
							    	<TD CLASS=TD6 NOWRAP>
							    	<script language =javascript src='./js/h3002ma1_txtStd_dt_txtStd_dt.js'></script>
							    	<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPromoteDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()">
							    	</TD>
								<TD CLASS=TD5 NOWRAP>발령일</TD>
								<TD CLASS=TD6 NOWRAP>
							    	<script language =javascript src='./js/h3002ma1_txtGazet_dt_txtGazet_dt.js'></script>
							    	</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발령전부서</TD>
			                    				<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtDept_cd" SIZE=10  MAXLENGTH=10  ALT ="발령전부서" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(1)">
								    <INPUT NAME="txtDept_nm" SIZE=20  MAXLENGTH=40  ALT ="발령전부서명" tag="14XXXU">
								</TD>
								<TD CLASS=TD5 NOWRAP>발령후부서</TD>
			                    				<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtchng_dept_cd" SIZE=10  MAXLENGTH=10 ALT ="발령후부서" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(2)">
								    <INPUT NAME="txtchng_dept_nm" SIZE=20  MAXLENGTH=40  ALT ="발령후부서명" tag="14XXXU">
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>변동사유</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_cd" SIZE=10  MAXLENGTH=10  ALT ="변동사유" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup('3')">
								                <INPUT NAME="txtChng_nm" SIZE=20  MAXLENGTH=50  ALT ="변동사유명" tag="14XXXU">
								<TD CLASS=TD5 NOWRAP></TD>
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
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%">
									<script language =javascript src='./js/h3002ma1_vaSpread1_vspdData.js'></script>
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
	<TR HEIGHT="20">
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN">자동입력</BUTTON></TD>
				    <TD WIDTH=* Align=RIGHT></TD>
				    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
