<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h3001ma1
*  4. Program Name         : 인사변동등록 
*  5. Program Desc         : 인사변동등록 조회,등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h3001mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow
Dim lsInternal_cd
Dim lsGetsvrDate

Dim C_GAZET_DT
Dim C_GAZET_CD
Dim C_GAZET_NM_POP
Dim C_GAZET_NM
Dim C_DEPT_CD
Dim C_DEPT_NM_POP
Dim C_DEPT_NM
Dim C_SECT_CD
Dim C_SECT_NM_POP
Dim C_SECT_NM
Dim C_ROLE_CD
Dim C_ROLE_NM_POP
Dim C_ROLE_NM
Dim C_ROLL_PSTN_CD
Dim C_ROLL_PSTN_NM_POP
Dim C_ROLL_PSTN_NM
Dim C_PAY_GRD1_CD
Dim C_PAY_GRD1_NM_POP
Dim C_PAY_GRD1_NM
Dim C_PAY_GRD2
Dim C_FUNC_CD
Dim C_FUNC_NM_POP
Dim C_FUNC_NM
Dim C_GAZET_RESN

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_GAZET_DT			= 1										'Column constant for Spread Sheet 
	 C_GAZET_CD			= 2 
	 C_GAZET_NM_POP		= 3
	 C_GAZET_NM			= 4
	 C_DEPT_CD			= 5
	 C_DEPT_NM_POP		= 6
	 C_DEPT_NM			= 7
	 C_SECT_CD			= 8
	 C_SECT_NM_POP		= 9
	 C_SECT_NM			= 10
	 C_ROLE_CD			= 11
	 C_ROLE_NM_POP		= 12
	 C_ROLE_NM			= 13
	 C_ROLL_PSTN_CD		= 14
	 C_ROLL_PSTN_NM_POP = 15
	 C_ROLL_PSTN_NM		= 16
	 C_PAY_GRD1_CD		= 17
	 C_PAY_GRD1_NM_POP	= 18
	 C_PAY_GRD1_NM		= 19
	 C_PAY_GRD2			= 20
	 C_FUNC_CD			= 21
	 C_FUNC_NM_POP		= 22
	 C_FUNC_NM			= 23
	 C_GAZET_RESN		= 24 
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
	lsGetsvrDate = "<%=GetsvrDate%>"
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
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		 WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    lgKeyStream       = Trim(frm1.txtEmp_no.value) & parent.gColSep       'You Must append one character( parent.gColSep)
    If  lsInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    Else
        lgKeyStream = lgKeyStream & lsInternal_cd & parent.gColSep
    End If
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
        .MaxCols = C_GAZET_RESN + 1											
	    .Col = .MaxCols														
        .ColHidden = True
        
        .MaxRows = 0			
		ggoSpread.ClearSpreadData        
		Call  GetSpreadColumnPos("A")
        
             ggoSpread.SSSetDate     C_GAZET_DT,         "발령일자", 10,2,  parent.gDateFormat
             ggoSpread.SSSetEdit     C_GAZET_CD,         "발령코드", 10
             ggoSpread.SSSetButton   C_GAZET_NM_POP
             ggoSpread.SSSetEdit     C_GAZET_NM,         "발령", 15,,,20                 
             ggoSpread.SSSetEdit     C_DEPT_CD,          "부서코드", 10
             ggoSpread.SSSetButton   C_DEPT_NM_POP
             ggoSpread.SSSetEdit     C_DEPT_NM,          "부서명", 18,,,40             
             ggoSpread.SSSetEdit     C_SECT_CD,          "근무구역코드", 10
             ggoSpread.SSSetButton   C_SECT_NM_POP    
             ggoSpread.SSSetEdit     C_SECT_NM,          "근무구역", 18,,,50             
             ggoSpread.SSSetEdit     C_ROLE_CD,          "직책코드", 10
             ggoSpread.SSSetButton   C_ROLE_NM_POP    
             ggoSpread.SSSetEdit     C_ROLE_NM,          "직책", 15,,,50             
             ggoSpread.SSSetEdit     C_ROLL_PSTN_CD,     "직위코드", 10
             ggoSpread.SSSetButton   C_ROLL_PSTN_NM_POP
             ggoSpread.SSSetEdit     C_ROLL_PSTN_NM,     "직위", 10,,,50                 
             ggoSpread.SSSetEdit     C_PAY_GRD1_CD,      "급호코드", 10
             ggoSpread.SSSetButton   C_PAY_GRD1_NM_POP
             ggoSpread.SSSetEdit     C_PAY_GRD1_NM,      "급호", 10,,,50                 
             ggoSpread.SSSetEdit     C_PAY_GRD2,         "호봉", 10,,,3,2
             ggoSpread.SSSetEdit     C_FUNC_CD,          "담당업무코드", 10
             ggoSpread.SSSetButton   C_FUNC_NM_POP
             ggoSpread.SSSetEdit     C_FUNC_NM,          "담당업무", 18,,,50                 
             ggoSpread.SSSetEdit     C_GAZET_RESN,       "발령사유", 20,,,40
             
             Call ggoSpread.MakePairsColumn(C_GAZET_CD		,  C_GAZET_NM_POP)
             Call ggoSpread.MakePairsColumn(C_DEPT_CD		,  C_DEPT_NM_POP)
             Call ggoSpread.MakePairsColumn(C_SECT_CD		,  C_SECT_NM_POP)
             Call ggoSpread.MakePairsColumn(C_ROLE_CD		,  C_ROLE_NM_POP)
             Call ggoSpread.MakePairsColumn(C_ROLL_PSTN_CD	,  C_ROLL_PSTN_NM_POP)
             Call ggoSpread.MakePairsColumn(C_PAY_GRD1_CD	,  C_PAY_GRD1_NM_POP)
             Call ggoSpread.MakePairsColumn(C_FUNC_CD		,  C_FUNC_NM_POP)

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
             ggoSpread.SpreadLock       C_GAZET_DT		, -1, C_GAZET_DT
             ggoSpread.SpreadLock       C_GAZET_CD		, -1, C_GAZET_CD
             ggoSpread.SpreadLock       C_GAZET_NM_POP	, -1, C_GAZET_NM_POP
             ggoSpread.SSSetRequired	C_DEPT_CD		, -1
             ggoSpread.SSSetRequired	C_SECT_CD		, -1
             ggoSpread.SSSetRequired	C_ROLE_CD		, -1
             ggoSpread.SSSetRequired	C_ROLL_PSTN_Cd	, -1
             ggoSpread.SSSetRequired	C_PAY_GRD1_CD	, -1
             ggoSpread.SSSetRequired	C_PAY_GRD2		, -1
             ggoSpread.SSSetRequired	C_FUNC_CD		, -1
            
             ggoSpread.SSSetProtected	C_GAZET_NM		, -1
             ggoSpread.SSSetProtected	C_DEPT_NM		, -1
             ggoSpread.SSSetProtected	C_SECT_NM		, -1
             ggoSpread.SSSetProtected	C_ROLE_NM		, -1
             ggoSpread.SSSetProtected	C_ROLL_PSTN_NM	, -1
             ggoSpread.SSSetProtected	C_PAY_GRD1_NM	, -1
             ggoSpread.SSSetProtected	C_FUNC_NM		, -1
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
             ggoSpread.SSSetRequired	C_GAZET_DT		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired    C_GAZET_CD		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_DEPT_CD		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_SECT_CD		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_ROLE_CD		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_ROLL_PSTN_CD	, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_PAY_GRD1_CD	, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_PAY_GRD2		, pvStartRow, pvEndRow
             ggoSpread.SSSetRequired	C_FUNC_CD		, pvStartRow, pvEndRow
            
            
             ggoSpread.SSSetProtected   C_GAZET_NM		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_DEPT_NM		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_SECT_NM		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_ROLE_NM		, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_ROLL_PSTN_NM	, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_PAY_GRD1_NM	, pvStartRow, pvEndRow
             ggoSpread.SSSetProtected	C_FUNC_NM		, pvStartRow, pvEndRow

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
            
            C_GAZET_DT			= iCurColumnPos(1)
			C_GAZET_CD			= iCurColumnPos(2)
			C_GAZET_NM_POP		= iCurColumnPos(3)
			C_GAZET_NM			= iCurColumnPos(4)
			C_DEPT_CD			= iCurColumnPos(5)
			C_DEPT_NM_POP		= iCurColumnPos(6)
			C_DEPT_NM			= iCurColumnPos(7)
			C_SECT_CD			= iCurColumnPos(8)
			C_SECT_NM_POP		= iCurColumnPos(9)
			C_SECT_NM			= iCurColumnPos(10)
			C_ROLE_CD			= iCurColumnPos(11)
			C_ROLE_NM_POP		= iCurColumnPos(12)
			C_ROLE_NM			= iCurColumnPos(13)
			C_ROLL_PSTN_CD		= iCurColumnPos(14)
			C_ROLL_PSTN_NM_POP	= iCurColumnPos(15)
			C_ROLL_PSTN_NM		= iCurColumnPos(16)
			C_PAY_GRD1_CD		= iCurColumnPos(17)
			C_PAY_GRD1_NM_POP	= iCurColumnPos(18)
			C_PAY_GRD1_NM		= iCurColumnPos(19)
			C_PAY_GRD2			= iCurColumnPos(20)
			C_FUNC_CD			= iCurColumnPos(21)
			C_FUNC_NM_POP		= iCurColumnPos(22)
			C_FUNC_NM			= iCurColumnPos(23)
			C_GAZET_RESN		= iCurColumnPos(24)            
            
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

    Call InitSpreadSheet                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    
    frm1.txtEmp_no.focus 
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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False                                                            '☜: Processing is NG
    
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

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    Call MakeKeyStream("X")
	Call DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True																'☜: Processing is OK
   

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
    
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
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
    Dim strtDt
    Dim strtDt2
    Dim lRow
    Dim lCount, lMaxDelRow
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
    
    lCount     = 0
	lMaxDelRow = 0
	With Frm1
		If .vspdData.MaxRows>1 Then					
			 .vspdData.Row = 1
   			 .vspdData.Col = C_GAZET_DT
 			  strtDt =  UniConvDateToYYYYMMDD(.vspdData.text, parent.gDateFormat,"")
        End If
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case  ggoSpread.InsertFlag,  ggoSpread.Updateflag
					.vspdData.Row = lRow
   	                .vspdData.Col = C_GAZET_DT
                    strtDt2 =  UniConvDateToYYYYMMDD(.vspdData.text, parent.gDateFormat,"")
                    If strtDt2 <> "" Then
                        If strtDt > strtDt2 then
	                        Call  DisplayMsgBox("972001","X","발령일자","최근발령일자")	'발령일자 은(는) 최근발령일자보다 커야합니다.
                            .vspdData.Text = ""
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if 
                    End if 
                    .vspdData.Col = C_GAZET_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","발령코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if 
                    
                    .vspdData.Col = C_DEPT_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","부서코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if 
					
					.vspdData.Col = C_SECT_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","근무구역코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if 
					
                    .vspdData.Col = C_ROLE_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","직책코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if 
					
					.vspdData.Col = C_ROLL_PSTN_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","직위코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if
					
					.vspdData.Col = C_PAY_GRD1_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","급호코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if
					
					.vspdData.Col = C_FUNC_NM
					if .vspdData.Text = "" then
						Call  DisplayMsgBox("970000","X","담당업무코드","X")
						.vspdData.focus
						 .vspdData.Action = 0
       					exit function
					end if
					
                Case  ggoSpread.deleteFlag    
                    lCount = lCount + 1
                    lMaxDelRow = lRow 
                    
            End Select
        Next
	End With
    
    if lCount <> lMaxDelRow then
			Call  DisplayMsgBox("800495","X","X","X") 
			exit function
	End if
		
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
'''	With Frm1
'        .vspdData.Row  = frm1.vspdData.ActiveRow
'        .vspdData.Col  = C_GAZET_NM
'        .vspdData.Text = ""
'        
'        .vspdData.ReDraw = True
'		.vspdData.focus
		
'	End With

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
	Dim strWhere
    Dim strDept_cd
    Dim strSect_cd
    Dim strRole_cd
    Dim strRoll_pstn
    Dim strPay_grd1    
	Dim imRow
	Dim iRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    If frm1.txtEmp_no.Value = "" Then
        Call  DisplayMsgBox("205152","X","사번","X")                  '☆:사번을 먼저 입력하세요."
        frm1.txtEmp_no.focus
		Exit Function
	End if
 
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
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1    
        
        .vspdData.Row = iRow
        .vspdData.Col = C_GAZET_DT
		.vspdData.Text =  UniConvDateAToB(lsGetsvrDate, parent.gServerDateFormat, parent.gDateFormat)
        strWhere = "  a.emp_no =  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & ""
        strWhere = strWhere & "  And a.internal_cd  = c.internal_cd "
        strWhere = strWhere & "  And c.org_change_dt = (SELECT MAX(org_change_dt) "
        strWhere = strWhere & "  From b_acct_dept Where org_change_dt <= GETDATE()) "
        strWhere = strWhere & "  And b.minor_cd = a.pay_grd1 And b.major_cd = " & FilterVar("H0001", "''", "S") & " "

        Call  CommonQueryRs(" Distinct(a.dept_cd), a.sect_cd, a.role_cd,  a.roll_pstn, a.pay_grd1 "," haa010t a, b_minor b, b_acct_dept c ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        .vspdData.Row = iRow
        strDept_cd   = Trim(Replace(lgF0,Chr(11),""))
        strSect_cd   = Trim(Replace(lgF1,Chr(11),""))
        strRole_cd   = Trim(Replace(lgF2,Chr(11),""))
        strRoll_pstn = Trim(Replace(lgF3,Chr(11),""))
        strPay_grd1  = Trim(Replace(lgF4,Chr(11),""))
    
        .vspdData.Col = C_DEPT_CD
        .vspdData.Text = strDept_cd
        .vspdData.Col = C_SECT_CD
        .vspdData.Text = strSect_cd
        .vspdData.Col = C_SECT_NM
        .vspdData.Text =  FuncCodeName(1,"H0035",strSect_cd)
        .vspdData.Col = C_ROLE_CD
        .vspdData.Text = strRole_cd
        .vspdData.Col = C_ROLE_NM
        .vspdData.Text =  FuncCodeName(1,"H0026",strRole_cd)
        .vspdData.Col = C_ROLL_PSTN_CD
        .vspdData.Text = strRoll_pstn
        .vspdData.Col = C_ROLL_PSTN_NM
        .vspdData.Text =  FuncCodeName(1,"H0002",strRoll_pstn)
        .vspdData.Col = C_PAY_GRD1_CD
        .vspdData.Text = strPay_grd1
         strWhere = strWhere & "  And c.dept_cd=  " & FilterVar(strDept_cd , "''", "S") & ""
         Call  CommonQueryRs("Distinct(a.pay_grd2), a.func_cd, c.dept_nm,b.minor_nm "," haa010t a, b_minor b, b_acct_dept c ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        .vspdData.Row = iRow
        .vspdData.Col = C_PAY_GRD2  
        .vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
        .vspdData.Col = C_FUNC_CD
        .vspdData.Text = Trim(Replace(lgF1,Chr(11),""))
        .vspdData.Col = C_DEPT_NM
        .vspdData.Text = Trim(Replace(lgF2,Chr(11),""))
        .vspdData.Col = C_PAY_GRD1_NM
        .vspdData.Text = Trim(Replace(lgF3,Chr(11),""))
        .vspdData.Col = C_FUNC_NM
        .vspdData.Text =  FuncCodeName(1,"H0004",Trim(Replace(lgF1,Chr(11),"")))
        Next
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
    Dim IDelCheck
    Dim IDx
    If Frm1.vspdData.MaxRows < 1 then   
       Call  DisplayMsgBox("900002","X","X","X")                                        '☜: Please do Display first. 
       Exit function
	End if	
    Frm1.vspdData.Col = 0
    IDelCheck = False
    
    For IDx = 1 To Frm1.vspdData.SelBlockRow-1 Step 1
        Frm1.vspdData.Row=IDx
        If Frm1.vspdData.Text<> ggoSpread.DeleteFlag Then
            IDelCheck = True
        End If
    Next
    
    If IDelCheck OR Frm1.vspdData.MaxRows=Frm1.vspdData.SelBlockRow Then                 '현재 선택된 row가 첫번째 row가 아니거나 row 개수가 1개라면 
        Call  DisplayMsgBox("800431","X","X","X")                                       '수정/삭제가 불가합니다. 
        Exit Function
    End If
   
    
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
    Call InitVariables														 '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

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
    Call InitVariables														 '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

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
	Call resok()
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
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
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
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1
	
	With Frm1
    
          if .vspddata.Text =  ggoSpread.DeleteFlag   then
          end if                             

       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                      strVal = strVal & "C" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_GAZET_DT      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_GAZET_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SECT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_GAZET_RESN    : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strVal = strVal & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_GAZET_DT      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_GAZET_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SECT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_GAZET_RESN    : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                      strDel = strDel & "D" & parent.gColSep
                                                      strDel = strDel & lRow & parent.gColSep
                                                      strDel = strDel & Trim(.txtEmp_no.value) & parent.gColSep
                    .vspdData.Col = C_GAZET_DT	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_GAZET_CD	    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
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
	Dim strVal
	Dim iRow
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    frm1.txtEmp_no.focus 

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar

	If frm1.vspdData.MaxRows>=2 Then

		For iRow=2 To frm1.vspdData.MaxRows
			With frm1
			    .vspdData.ReDraw = False			         
			          ggoSpread.SpreadLock       1, iRow, frm1.vspdData.MaxCols
			    .vspdData.ReDraw = True
			End With
		Next
	End If

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement
    frm1.vspdData.focus
End Function

Function Resok()
Dim iRow	

	If frm1.vspdData.MaxRows>=2 Then
		For iRow=2 To frm1.vspdData.MaxRows
			With frm1
			    .vspdData.ReDraw = False
			         ggoSpread.SpreadLock       1, iRow, frm1.vspdData.MaxCols
			    .vspdData.ReDraw = True
			End With
		Next
	End If
	
    Call  ggoOper.LockField(Document, "Q")
    
  
End function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
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
	    Case C_GAZET_NM_POP
	        arrParam(0) = "발령코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                          			' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0029", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "발령코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "발령코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "발령코드명"	    		        ' Header명(1)
	    Case C_SECT_NM_POP
	        arrParam(0) = "근무구역코드 팝업"			    ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                  		        	' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0035", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "근무구역코드" 			        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "근무구역코드"	   		    	' Header명(0)
	    	arrHeader(1) = "근무구역명"	    		        ' Header명(1)
	    Case C_ROLE_NM_POP
	        arrParam(0) = "직책코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                   			        ' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0026", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직책코드코드" 			        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직책코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "직책코드명"	    		        ' Header명(1)
	    Case C_ROLL_PSTN_NM_POP
	        arrParam(0) = "직위코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                   			        ' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직위코드코드" 			        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직위코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "직위코드명"	    		        ' Header명(1)
	    Case C_PAY_GRD1_NM_POP
	        arrParam(0) = "급호코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                           			' Code Condition
	    	arrParam(3) = ""								' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "급호코드" 		    	        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "호봉코드"	   		        	' Header명(0)
	    	arrHeader(1) = "호봉코드명"	    		        ' Header명(1)
	    Case C_FUNC_NM_POP
	        arrParam(0) = "담당업무코드 팝업"			    ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = strCode                           			' Code Condition
	    	arrParam(3) = ""	    						' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0004", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "담당업무코드" 			        ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "담당업무코드"	   		    	' Header명(0)
	    	arrHeader(1) = "담당업무명"	    		        ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then

		With frm1

			Select Case iWhere
			    Case C_GAZET_NM_POP
			        .vspdData.Col = C_GAZET_CD
			    	.vspdData.action =0
			    Case C_SECT_NM_POP
			        .vspdData.Col = C_SECT_CD
					.vspdData.action =0
			    Case C_ROLE_NM_POP
					.vspdData.Col = C_ROLE_CD
			    	.vspdData.action =0
			    Case C_ROLL_PSTN_NM_POP
			        .vspdData.Col = C_ROLL_PSTN_CD
			    	.vspdData.action =0
			    Case C_PAY_GRD1_NM_POP
			        .vspdData.Col = C_PAY_GRD1_CD
			    	.vspdData.action =0		    	
			    Case C_FUNC_NM_POP
			        .vspdData.Col = C_FUNC_CD
			    	.vspdData.action =0
		    End Select
		End With	
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
		    Case C_GAZET_NM_POP
		    	.vspdData.Col = C_GAZET_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_GAZET_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0
		    Case C_SECT_NM_POP
		    	.vspdData.Col = C_SECT_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_SECT_CD
		    	.vspdData.text = arrRet(0) 
				.vspdData.action =0
		    Case C_ROLE_NM_POP
		        .vspdData.Col = C_ROLE_NM
		    	.vspdData.text = arrRet(1)   
				.vspdData.Col = C_ROLE_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0
		    Case C_ROLL_PSTN_NM_POP
		    	.vspdData.Col = C_ROLL_PSTN_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_ROLL_PSTN_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0
		    Case C_PAY_GRD1_NM_POP
		    	.vspdData.Col = C_PAY_GRD1_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_PAY_GRD1_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0		    	
		    Case C_FUNC_NM_POP
		    	.vspdData.Col = C_FUNC_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_FUNC_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0
        End Select

	End With

End Function
'----------------------------------------  OpenDeptDt()  ------------------------------------------
'	Name : OpenDeptDt()
'	Description : 특정일자 입력받은 Dept PopUp
'---------------------------------------------------------------------------------------------------
Function OpenDeptDt(iWhere,strDate,TargetObj,TargetObj1, ByVal Row)

	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = TargetObj.Value          	    		' Code Condition%>
        If strDate="X" Then
        	arrParam(1) = ""                                ' 현재날짜!!!
        Else
        	arrParam(1) = strDate                           ' 특정 Date값을 parameter(1)로 넘긴다!!!
        End If    
	Else 'spread
		frm1.vspdData.Col = C_Dept_cd
        arrParam(0) = frm1.vspdData.Text                    'Code Condition
        
        frm1.vspdData.Col = C_GAZET_DT
        arrParam(1) = frm1.vspdData.Text                    ' 특정 Date값을 parameter(1)로 넘긴다!!!
	End If
        arrParam(2) = lgUsrIntCd                            ' 특정 권한값을 parameter(2)로 넘긴다!!!
		arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent ,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_DEPT_CD
		frm1.vspdData.action =0
		Exit Function
	Else
		Call SetDeptDt(arrRet, iWhere, TargetObj,TargetObj1, Row)
	End If	
			
End Function

'------------------------------------------  SetDeptDt()  ---------------------------------------------
'	Name : SetDeptDt()
'	Description : Dept Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetDeptDt(Byval arrRet, Byval iWhere, Byval TargetObj, Byval TargetObj1, ByVal Row)
		
		If iWhere = 0 Then 'TextBox(Condition)
			TargetObj.Value = arrRet(0)
			TargetObj1.Value = arrRet(1)
		Else 'spread
        	With frm1
			.vspdData.Col = C_DEPT_NM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
			 ggoSpread.UpdateRow Row
        	End With
		End If
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp(iWhere)
			
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""							' Name Cindition
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""							' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(Byval arrRet, Byval iWhere)
    Dim strVal
    
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtDept_nm.value = arrRet(2)
			.txtRoll_pstn.value = arrRet(3)
			.txtEntr_dt.text = arrRet(5)
			.txtPay_grd.value = arrRet(4)
			.txtEmp_no.focus
           
		Else 'spread
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strDate
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col-1
	Select Case Col
	    Case C_GAZET_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_GAZET_NM_POP, Row)
	    Case C_DEPT_NM_POP
                    Call OpenDeptDt(1,"X","X","X", Row)
'                   Call OpenCode(frm1.vspdData.Text, C_DEPT_NM_POP, Row)
	    Case C_SECT_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_SECT_NM_POP, Row)
	    Case C_ROLE_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_ROLE_NM_POP, Row)
	    Case C_ROLL_PSTN_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_ROLL_PSTN_NM_POP, Row)
	    Case C_PAY_GRD1_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_PAY_GRD1_NM_POP, Row)
	    Case C_FUNC_NM_POP
                    Call OpenCode(frm1.vspdData.Text, C_FUNC_NM_POP, Row)
    End Select    
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    Dim iDx
    Dim apply_strt_dt
    Dim IntRetCD
    Dim strWhere
    Dim strDept_nm, strInternal_cd
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	Select Case Col
	    Case C_PAY_GRD2
	      
	        frm1.vspdData.Col = C_GAZET_DT
            apply_strt_dt =  UNIConvDateCompanyToDB(frm1.vspdData.Text,  parent.gDateFormat)

	        frm1.vspdData.Col = C_PAY_GRD2
            IntRetCD =  CommonQueryRs(" pay_grd2 "," hdf010t "," pay_grd2 =  " & FilterVar(frm1.vspdData.Text, "''", "S") & " And apply_strt_dt = (SELECT MAX(apply_strt_dt) FROM hdf010t WHERE apply_strt_dt <=  " & FilterVar(apply_strt_dt , "''", "S") & ")",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD=False Then
                Call  DisplayMsgBox("800057","X","X","X")                         '☜ : 기본급테이블에 존재하지 않는 호봉입니다.
                frm1.vspdData.Text=""
            End If
            
		Case C_GAZET_CD
			
			IntRetCD =  CommonQueryRs(" MINOR_NM "," B_Minor "," Major_cd = " & FilterVar("H0029", "''", "S") & " and minor_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            
            frm1.vspdData.Col = C_Gazet_NM
            If IntRetCD=False Then
				Call  DisplayMsgBox("970000","X","발령코드","X")
                frm1.vspdData.Text=""
            else
				frm1.vspdData.Text = replace(lgF0,chr(11),"")
            End If
            
		Case C_DEPT_CD
			
			IntRetCd =  FuncDeptName(frm1.vspdData.text,"",lgUsrIntCd,strDept_nm,strInternal_cd)
			frm1.vspdData.col = C_DEPT_NM
			if  IntRetCd = -1 then
				Call  DisplayMsgBox("970000","X","부서코드","X")
				frm1.vspdData.text = ""
			else
				frm1.vspdData.text = strDept_nm
			end if
    
		CASE C_SECT_CD
		
			intRetCD =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			frm1.vspdData.col = C_SECT_NM
			if intRetCd = false then
				Call  DisplayMsgBox("970000","X","근무구역코드","X")
				frm1.vspdData.text = ""
			Else
				frm1.vspdData.text = Replace(lgF0, Chr(11), "")
			End If
		CASE C_ROLE_CD
		
			intRetCd =   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0026", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			frm1.vspdData.col = C_ROLE_NM
            if intRetCd = false then
				Call  DisplayMsgBox("970000","X","직책코드","X")
				frm1.vspdData.text = ""
			Else
				frm1.vspdData.text = Replace(lgF0, Chr(11), "")
			End If
		CASE C_ROLL_PSTN_CD
		
			intRetCD =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			frm1.vspdData.col = C_ROLL_PSTN_NM
			
			if intRetCd = false then
				Call  DisplayMsgBox("970000","X","직위코드","X")
				frm1.vspdData.text = ""
			Else
				frm1.vspdData.text = Replace(lgF0, Chr(11), "")
			End If
		CASE C_PAY_GRD1_CD
		
			intRetCD =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			frm1.vspdData.col = C_PAY_GRD1_NM
			
			if intRetCd = false then
				Call  DisplayMsgBox("970000","X","급호코드","X")
				frm1.vspdData.text = ""
			Else
				frm1.vspdData.text = Replace(lgF0, Chr(11), "")
			End If
		CASE C_FUNC_CD
		
			intRetCD =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0004", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			frm1.vspdData.col = C_Func_NM
			
			if intRetCd = false then
				Call  DisplayMsgBox("970000","X","담당업무코드","X")
				frm1.vspdData.text = ""
			Else
				frm1.vspdData.text = Replace(lgF0, Chr(11), "")
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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
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

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
        frm1.txtName.value = ""
        frm1.txtDept_nm.value = ""
        frm1.txtRoll_pstn.value = ""
        frm1.txtPay_grd.value = ""
        frm1.txtEntr_dt.Text = ""
    Else
    
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    If  IntRetCd < 0 then
			strVal = "../../../CShared/image/default_picture.jpg"
			Frm1.imgPhoto.src = strVal	    
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            frm1.txtName.value = ""
            frm1.txtDept_nm.value = ""
            frm1.txtRoll_pstn.value = ""
            frm1.txtPay_grd.value = ""
            frm1.txtEntr_dt.Text = ""

			ggoSpread.Source = Frm1.vspdData    
			ggoSpread.ClearSpreadData 
           
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true

        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2

            frm1.txtEntr_dt.Text =  UNIDateClientFormat(strEntr_dt)

			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    		if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if

    		Frm1.imgPhoto.src = strVal
        End if 
    End if
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>인사변동등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=7%>
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			                <TR HEIGHT=69>
			                    <TD>
                                    <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=60 HEIGHT=69 HSPACE=10 VSPACE=0 BORDER=1>
			                    </TD>
			                </TR>
			            </TABLE>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp('0')">
			    	        	   	    		      <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14"></TD>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDept_nm" SIZE=20  MAXLENGTH=40 ALT="부서명" TYPE="Text"  tag="14"></TD>
			            	</TR>
			            	<TR>	
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtRoll_pstn" SIZE=20  MAXLENGTH=50  ALT="직위" TYPE="Text"  tag="14"></TD>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
							    <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h3001ma1_txtEntr_dt_txtEntr_dt.js'></script></TD>
			            	</TR>
			            	<TR>	
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPay_grd" SIZE=20  MAXLENGTH=50 ALT="급호" TYPE="Text"  tag="14"></TD>
			            		<TD CLASS="TD5" NOWRAP></TD>
			            		<TD CLASS="TD6" NOWRAP></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h3001ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">인사마스타</a></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
