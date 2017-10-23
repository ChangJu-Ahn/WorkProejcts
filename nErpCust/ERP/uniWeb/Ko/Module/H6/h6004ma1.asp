<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h6004ma1
*  4. Program Name         : h6004ma1
*  5. Program Desc         : 급여관리/고정수당조회및등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const BIZ_PGM_ID = "h6004mb1.asp"                                    'Biz Logic ASP 
Const BIZ_PGM_ID1 = "h6004mb2.asp"                                      'Biz Logic ASP  
Const TAB1 = 1
Const TAB2 = 2

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
Dim lsConcd
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)
 
Dim C_EMP_NO
Dim C_EMP_NO_POP
Dim C_NAME
Dim C_DEPT_CD
Dim C_ALLOW_CD
Dim C_ALLOW_CD_POP
Dim C_ALLOW_NM
Dim C_ALLOW_AMT
Dim C_APPLY_YYMM_DT
Dim C_REVOKE_YYMM_DT														

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
    C_EMP_NO  =       1
    C_EMP_NO_POP =    2
    C_NAME  =         3
    C_DEPT_CD =       4
    C_ALLOW_CD =      5
    C_ALLOW_CD_POP =  6
    C_ALLOW_NM =      7
    C_ALLOW_AMT =     8
    C_APPLY_YYMM_DT = 9
    C_REVOKE_YYMM_DT = 10
end sub
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
	lsInternal_cd     = ""
	
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtApply_yymm_dt.Year = strYear 
	frm1.txtApply_yymm_dt.Month = strMonth 

	frm1.txtRevoke_yymm_dt.Year = strYear 
	frm1.txtRevoke_yymm_dt.Month = strMonth 
	
	frm1.txtValidDt.Year = strYear 		 '년월일 default value setting
	frm1.txtValidDt.Month = strMonth 
	frm1.txtValidDt.Day = strDay	
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
    Dim strMin
    Dim strMax
   
    Call FuncGetTermDept(lgUsrIntCd,frm1.txtValidDt.text,strMin,strMax)
    
    lgKeyStream       = Frm1.txtEmp_no.Value & parent.gColSep   '0                                       'You Must append one character(parent.gColSep)
    if  Frm1.txtFr_Dept_cd.Value = "" then
        lgKeyStream = lgKeyStream & strMin & parent.gColSep
    else
	    lgKeyStream = lgKeyStream & Frm1.txtFr_internal_cd.Value & parent.gColSep '1
    end if
	
    if  Frm1.txtTo_Dept_cd.Value = "" then
        lgKeyStream = lgKeyStream & strMax & parent.gColSep
    else
	    lgKeyStream = lgKeyStream & Frm1.txtTo_internal_cd.Value & parent.gColSep  '2
    end if	
	lgKeyStream  = lgKeyStream & Frm1.txtAllow_cd.value & parent.gColSep			'3
	lgKeyStream  = lgKeyStream & Frm1.txtFr_Dept_cd.value & parent.gColSep			'4
	lgKeyStream  = lgKeyStream & Frm1.txtTo_Dept_cd.value & parent.gColSep			'5
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtapply_yymm_dt.text) & parent.gColSep    '6
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtrevoke_yymm_dt.text) & parent.gColSep   '7
    lgKeyStream  = lgKeyStream & Trim(Frm1.cboFix_yn.value) & Parent.gColSep    '8
        lgKeyStream  = lgKeyStream & Trim(frm1.txtValidDt.text) & parent.gColSep   '9
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    iCodeArr = "Y" & Chr(11) & "N" & Chr(11)
    iNameArr = "YES" & Chr(11) & "NO" & Chr(11)

    Call SetCombo2(frm1.cboFix_yn,iCodeArr, iNameArr,Chr(11)) 
End Sub	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
'	name: Tab Click
'	desc: Tab Click시 필요한 기능을 수행한다.
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	frm1.cboFix_yn.selectedIndex = 1
	 ProtectTag(frm1.cboFix_yn)    
	Call SetToolbar("1100111100111111")		
	frm1.btnCb_autoisrt.disabled = false	 	 
End Function
'-------------------------------------------------------------------------------------------------------
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ReleaseTag(frm1.cboFix_yn)   
	Call SetToolbar("1100101100011111")		
	frm1.btnCb_autoisrt.disabled = true  
End Function

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	
	Call initSpreadPosVariables()  
        
	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
         
	   .ReDraw = false
       .MaxCols   = C_REVOKE_YYMM_DT + 1                                                     ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                                  ' ☜:☜:

       .MaxRows = 0
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData
		        
		Call GetSpreadColumnPos("A")       

        ggoSpread.SSSetEdit     C_EMP_NO,        "사번",       13,,, 13,2
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetEdit     C_NAME,          "성명",       17,,, 30,2
        ggoSpread.SSSetEdit     C_DEPT_CD,       "부서",       22,,, 40,2
        ggoSpread.SSSetEdit     C_ALLOW_CD,      "수당코드",    8,,, 3,2
        ggoSpread.SSSetButton   C_ALLOW_CD_POP
        ggoSpread.SSSetEdit     C_ALLOW_NM,      "수당코드명", 18,,,20,2
        ggoSpread.SSSetFloat    C_ALLOW_AMT   ,  "수당액" ,    15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetMask	   C_APPLY_YYMM_DT,	"적용년월",   9,2, strMaskYM
		ggoSpread.SSSetMask	   C_REVOKE_YYMM_DT,"해제년월",   9,2, strMaskYM

        Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)    'sbk
        Call ggoSpread.MakePairsColumn(C_ALLOW_CD,C_ALLOW_CD_POP)    'sbk

	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
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
            
            C_EMP_NO		= iCurColumnPos(1)
            C_EMP_NO_POP	= iCurColumnPos(2)
            C_NAME			= iCurColumnPos(3)
            C_DEPT_CD		= iCurColumnPos(4)
            C_ALLOW_CD		= iCurColumnPos(5)
            C_ALLOW_CD_POP	= iCurColumnPos(6)
            C_ALLOW_NM		= iCurColumnPos(7)
            C_ALLOW_AMT		= iCurColumnPos(8)
            C_APPLY_YYMM_DT = iCurColumnPos(9)
            C_REVOKE_YYMM_DT= iCurColumnPos(10)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	With frm1
	.vspdData.ReDraw = False
	  ggoSpread.SpreadLock      C_EMP_NO, -1, C_EMP_NO, -1
	  ggoSpread.SpreadLock      C_EMP_NO_POP, -1, C_EMP_NO_POP, -1
	  ggoSpread.SpreadLock      C_NAME, -1, C_NAME, -1
	  ggoSpread.SpreadLock      C_DEPT_CD, -1, C_DEPT_CD, -1
	  ggoSpread.SpreadLock      C_ALLOW_CD, -1, C_ALLOW_CD, -1
	  ggoSpread.SpreadLock      C_ALLOW_NM, -1, C_ALLOW_NM, -1
	  ggoSpread.SpreadLock      C_ALLOW_CD_POP, -1, C_ALLOW_CD_POP, -1
	  ggoSpread.SSSetRequired   C_ALLOW_AMT, -1, -1
	  ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
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
      ggoSpread.SSSetProtected    C_NAME , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DEPT_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired     C_EMP_NO , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired     C_ALLOW_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected     C_ALLOW_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired     C_ALLOW_AMT, pvStartRow, pvEndRow
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
    
	Call AppendNumberPlace("6", "9", "0")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call ggoOper.FormatDate(frm1.txtApply_yymm_dt, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtRevoke_yymm_dt, parent.gDateFormat, 2)
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
   
    Call SetDefaultVal
    Call SetToolbar("1100111100101111")										        '버튼 툴바 제어 
	Call DisableBtnByAuth()											'⊙: Set ToolBar
    frm1.txtEmp_no.focus
    	
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
    Call InitComboBox  
    gSelframeFlg = TAB1
    frm1.cboFix_yn.selectedIndex = 1
	ProtectTag(frm1.cboFix_yn)      
	Call CookiePage (0)                                                             '☜: Check Cookie
    
End Sub
'========================================================================================================
' Name : DisableBtnByAuth  버튼 권한 설정(화면권한이 조회일 경우 버튼 비활성화)
'========================================================================================================
Sub DisableBtnByAuth()

	Dim strAuth
	
	strAuth = ""
	Err.Clear

	'gStrRequestMenuID : uni2kCM.inc file에서 정의된 프로그램을 요청한 메뉴의 ID
	If UCase(Left(gStrRequestMenuID, 1)) <> "Z" Then
		Call parent.uni2kMenu.Restore("BizMenu")
	Else
		Call parent.uni2kMenu.Restore("System")
	End If
	strAuth = parent.uni2kMenu.MenuItemAuthority(gStrRequestMenuID)
		
	Err.Clear

	Select Case strAuth
		Case "N" , "Q" ,"E"              ' None
			frm1.btnCb_autoisrt.disabled = true
	End Select   
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
    Dim strStartDept
    Dim strEndDept
    Dim strFrYymm
    Dim strToYymm
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If (frm1.txtApply_yymm_dt.Text = "") Then                       '년월의 값이 없으면 주는 기본값정의와 메시지 체크 
        strFrYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")
    Else
        strFrYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtapply_yymm_dt.Year, Right("0" & frm1.txtapply_yymm_dt.month , 2), "01")
    End if 
    
    If (frm1.txtRevoke_yymm_dt.Text = "") Then
        strToYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, "2500", "12", "31")
    Else
        strToYYMM = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtrevoke_yymm_dt.Year, Right("0" & frm1.txtrevoke_yymm_dt.month , 2), "01")
    End if 
         
    If CompareDateByFormat(strFrYymm,strToYymm,frm1.txtapply_yymm_dt.Alt,frm1.txtrevoke_yymm_dt.Alt,"970025",parent.gDateFormat,parent.gComDateType,True) = False Then
        frm1.txtapply_yymm_dt.focus
        Set gActiveElement = document.activeElement

        Exit Function
    End if 

    If txtEmp_no_Onchange() Then          'enter key 로 조회시 사원를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtAllow_cd_Onchange() Then          'enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtFr_dept_cd_Onchange() Then        'enter key 로 조회시 시작부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    If txtTo_dept_cd_Onchange()  Then        'enter key 로 조회시 종료부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Fr_dept_cd = frm1.txtFr_dept_cd.value
    To_dept_cd = frm1.txtTo_dept_cd.value
   
    If Fr_dept_cd = "" then    
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,frm1.txtValidDt.text, rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept				
		frm1.txtFr_dept_nm.value = ""
    End If	
	
    If To_dept_cd   = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,frm1.txtValidDt.text, rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
    End If  
    
    strStartDept =  frm1.txtFr_internal_cd.value
    strEndDept   = frm1.txtTo_internal_cd.value
    
    If (strStartDept <> "") AND (strEndDept <>"") Then       
        If strStartDept > strEndDept  then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_internal_cd.value = ""
            frm1.txtTo_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF 

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
 
    FncQuery = True                                                              '☜: Processing is OK
															'☜: Processing is OK

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
    
   	 Dim strApplyDt
   	 Dim strApplyDt2
   	 Dim strRevokeDt
   	 Dim strRevokeDt2
   	 Dim strAllowAmt
   	 Dim lRow

	 With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
           
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_NAME

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                        Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
				
					.vspdData.Col = C_ALLOW_NM
		
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                        Call DisplayMsgBox("800145","X","X","X")
						Exit Function
					end if
									
   	                .vspdData.Col = C_APPLY_YYMM_DT
   					If Trim(.vspdData.Text) <> parent.gComDateType AND Trim(.vspdData.Text) <> "" Then
						strApplyDt = UNIGetLastDay(.vspdData.Text, parent.gDateFormatYYYYMM)
						if Trim(StrApplyDt) = "" then
							Call DisplayMsgBox("200006","X","X","X")	
						    .vspdData.focus
						    Set gActiveElement = document.activeElement
						    Exit Function
						end if
					End If
						
   	                .vspdData.Col = C_REVOKE_YYMM_DT
   					If Trim(.vspdData.Text) <> parent.gComDateType AND Trim(.vspdData.Text) <> "" Then
						strRevokeDt = UNIGetLastDay(.vspdData.Text, parent.gDateFormatYYYYMM)

						if Trim(strRevokeDt) = "" then
							Call DisplayMsgBox("200006","X","X","X")	
						    .vspdData.focus
						    Set gActiveElement = document.activeElement
						    Exit Function
						end if
					End If
					
                    If Trim(.vspdData.Text) = parent.gComDateType Then
                    Else
                        If CompareDateByFormat(strApplyDt,strRevokeDt,"적용년월","해제년월","970023",parent.gDateFormat,parent.gComDateType,True) = False then
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_APPLY_YYMM_DT
                            .vspdData.focus
                            Set gActiveElement = document.activeElement
                            Exit Function
                        Else
                        End if 
                    End if  
                    
            End Select
        Next
	End With
 
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_ALLOW_AMT
                    strAllowAmt = UNICDbl(.vspdData.Text)
                  '  If strAllowAmt = "" or strAllowAmt = 0 Then   '2004.07.14
					If strAllowAmt = "" Then
	                     Call DisplayMsgBox("800379","X","X","X")	'수당금액은 입력필수항목입니다.
	                     .vspdData.Row = lRow
  	                     .vspdData.Col = C_ALLOW_AMT
                         .vspdData.focus
                         Set gActiveElement = document.activeElement
                         Exit Function
                    End if  
            End Select
        Next
	End With
    
    Call MakeKeyStream("X")
    Call DisableToolBar(parent.TBC_SAVE)
	If DBSave=False Then
	   Call RestoreToolBar()
	   Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
    
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
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
	
    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables			    											 '⊙: Initializes local global variables

    If LayerShowHide(1)=False Then
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
	
    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables							    							 '⊙: Initializes local global variables

    If LayerShowHide(1)=False Then
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
    Call parent.FncFind( parent.C_MULTI, False)                                          '☜:화면 유형, Tab 유무 
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
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
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1)=False Then
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
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel

    Dim iColSep, iRowSep
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
 	Dim iFormLimitByte						'102399byte
 	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
 	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
 	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size
    
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
 	
 	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
 	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
     
     '102399byte
     iFormLimitByte = parent.C_FORM_LIMIT_BYTE
     
     '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
 	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				
 
 	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
 	
 	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
    DbSave = False                                                          
    
	If LayerShowHide(1)=False Then
		Exit Function
	End If

    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
					strVal = ""               
                                                            strVal = strVal & "C" & parent.gColSep 'array(0)
                                                            strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_NAME	              : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD    	      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ALLOW_CD	          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ALLOW_AMT	          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_YYMM_DT	      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REVOKE_YYMM_DT      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 
               Case ggoSpread.UpdateFlag                                      '☜: Update
					strVal = ""                              
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ALLOW_CD	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ALLOW_AMT	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_YYMM_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REVOKE_YYMM_DT    : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
				    strDel = ""
                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ALLOW_CD          : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'삭제시 key만								
                    lGrpCnt = lGrpCnt + 1
           End Select
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If

			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   

			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)

			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
           
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	

	End With

    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                                 '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If

    Call DisableToolBar(parent.TBC_DELETE)
	If DBDelete=False Then
	   Call RestoreToolBar()
	   Exit Function
	End If
    
    FncDelete = True                                                        '⊙: Processing is OK
                                                            '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode = parent.OPMD_UMODE    

    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()

	Call DisableBtnByAuth()	

	If gSelframeFlg = TAB1 Then
		frm1.cboFix_yn.selectedIndex = 1
		ProtectTag(frm1.cboFix_yn)   
		ggoSpread.SpreadUnLock		C_APPLY_YYMM_DT,	-1, C_APPLY_YYMM_DT,	-1
		ggoSpread.SpreadUnLock		C_REVOKE_YYMM_DT,	-1, C_REVOKE_YYMM_DT,	-1
		ggoSpread.SpreadUnLock		C_ALLOW_AMT,		-1, C_ALLOW_AMT,		-1		
		ggoSpread.SSSetRequired		C_ALLOW_AMT,	-1, -1
		Call SetToolbar("1100111100111111")			
	Else
		ReleaseTag(frm1.cboFix_yn) 
		ggoSpread.Source = frm1.vspdData
'		ggoSpread.SpreadLockWithOddEvenRowColor()	
		ggoSpread.SpreadLock      C_ALLOW_AMT,		-1, C_ALLOW_AMT,		-1
		ggoSpread.SpreadLock      C_APPLY_YYMM_DT,	-1, C_APPLY_YYMM_DT,	-1
		ggoSpread.SpreadLock      C_REVOKE_YYMM_DT, -1, C_REVOKE_YYMM_DT,	-1
 	
		Call SetToolbar("1100101100011111")					
	End If
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
	Call RemovedivTextArea    
    Call InitVariables															'⊙: Initializes local global variables

	If DBQuery=False Then
	   Call RestoreToolBar()
	   Exit Function
	End If
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
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_ALLOW_CD_POP
	        arrParam(0) = "수당코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
            frm1.vspdData.col = C_ALLOW_CD
	        arrParam(2) = frm1.vspdData.Text               ' Code Condition
            frm1.vspdData.col = C_ALLOW_NM
	        arrParam(3) = ""'frm1.vspdData.Text				' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND CALCU_TYPE = " & FilterVar("N", "''", "S") & "  "				' Where Condition
	        arrParam(5) = "수당코드"			    ' TextBox 명칭 
	
            arrField(0) = "allow_cd"					' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당코드"				' Header명(0)
            arrHeader(1) = "수당코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_ALLOW_CD
		frm1.vspdData.action =0
		Exit Function
	Else
		With frm1

			Select Case iWhere
			    Case C_ALLOW_CD_POP
			    	.vspdData.Col = C_ALLOW_NM
			    	.vspdData.text = arrRet(1)
			        .vspdData.Col = C_ALLOW_CD
			    	.vspdData.text = arrRet(0) 
			    	.vspdData.action =0
		    End Select

		End With
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

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

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	End If
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		With frm1
			If iWhere = 0 Then
				.txtEmp_no.value = arrRet(0)
				.txtName.value = arrRet(1)
				.txtEmp_no.focus
			Else
				.vspdData.Col = C_NAME
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_DEPT_CD
				.vspdData.Text = arrRet(2)
				.vspdData.Col = C_EMP_NO
				.vspdData.Text = arrRet(0)
				.vspdData.action =0
			End If
		End With
	End If	
			
End Function
 
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
	        arrParam(3) = ""'frm1.txtAllow_nm.value				' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND CALCU_TYPE = " & FilterVar("N", "''", "S") & "  " ' Where Condition
	        arrParam(5) = "수당코드"			    ' TextBox 명칭 
	
            arrField(0) = "allow_cd"					' Field명(0)
            arrField(1) = "allow_nm"				    ' Field명(1)
    
            arrHeader(0) = "수당코드"				' Header명(0)
            arrHeader(1) = "수당코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtAllow_cd.focus
		Exit Function
	Else
		With Frm1
			Select Case iWhere
			    Case "1"
			        .txtAllow_cd.value = arrRet(0)
			        .txtAllow_nm.value = arrRet(1)		
			        .txtAllow_cd.focus
		    End Select
		End With
	End If	
	
End Function
 
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
	arrParam(1) = Frm1.txtValidDt.Text
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
        End Select	
		Exit Function
	Else
		With frm1
			Select Case iWhere
			     Case "0"
		           .txtFr_dept_cd.value = arrRet(0)
		           .txtFr_dept_nm.value = arrRet(1)
		           .txtFr_internal_cd.value = arrRet(2)
		           .txtFr_dept_cd.focus
		         Case "1"  
		           .txtTo_dept_cd.value = arrRet(0)
		           .txtTo_dept_nm.value = arrRet(1) 
		           .txtTo_internal_cd.value = arrRet(2) 
		           .txtTo_dept_cd.focus
		    End Select
		End With
	End If				
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_EMP_NO_POP
                    Call OpenEmptName("1")
	    Case C_ALLOW_CD_POP
                    Call OpenCode("", C_ALLOW_CD_POP, Row)
    End Select
End Sub

'======================================================================================================
'	Name : AutoInsertButtonClicked()
'	Description : h4007mb2.asp 로 가는 Condition........일괄등록...........
'=======================================================================================================
Sub AutoInsertButtonClicked(Byval ButtonDown)
    Dim strVal
    Dim IntRetCD
    
    Dim strEmpNo
    Dim strFrInternalCd
    Dim strToInternalCd
    Dim strDelete 
    DIM strWhere
    
    Dim strFrYymm, strToYymm
        
    Dim strMin
    Dim strMax
   
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit sub
		End If
    End If
    
    If Trim(Frm1.txtAllow_cd.Value) = "" then                         '☜: 자동입력에서 수당코드는 필수입력사항이다..
        Call DisplayMsgBox("970021", "X","수당코드","x")
        Frm1.txtAllow_cd.focus
       Exit sub 
    Else
    End if 
    
    ' 기간이 비어있을때 비교시 에러 수정 2002.09.09 이석민 
    If (frm1.txtApply_yymm_dt.Text = "") Then                       '년월의 값이 없으면 주는 기본값정의와 메시지 체크 
        strFrYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "02")
    Else
        strFrYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtapply_yymm_dt.Year, Right("0" & frm1.txtapply_yymm_dt.month , 2), "01")
    End if 
    
    If (frm1.txtRevoke_yymm_dt.Text = "") Then
        strToYymm = UniConvYYYYMMDDToDate(parent.gDateFormat, "2500", "12", "31")
    Else
        strToYYMM = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtrevoke_yymm_dt.Year, Right("0" & frm1.txtrevoke_yymm_dt.month , 2), "01")
    End if 
         
    If CompareDateByFormat(strFrYymm,strToYymm,frm1.txtapply_yymm_dt.Alt,frm1.txtrevoke_yymm_dt.Alt,"970025",parent.gDateFormat,parent.gComDateType,True) = False Then
        frm1.txtapply_yymm_dt.focus
        Set gActiveElement = document.activeElement

        Exit sub
    End if 
    
    strEmpNo = frm1.txtEmp_no.value
    strFrInternalCd = frm1.txtFr_internal_cd.value
    strToInternalCd = frm1.txtTo_internal_cd.value
    
    If strEmpNo = "" then
        strEmpNo = "%"
    End if
    
    Call FuncGetTermDept(lgUsrIntCd,frm1.txtValidDt.text,strMin,strMax)
    
    If strFrInternalCd = "" then
        strFrInternalCd = strMin
    End if
    If strToInternalCd = "" then
        strToInternalCd = strMax
    End if
    
    If (strFrInternalCd = "") AND (strToInternalCd = "") Then       
    Else
        If strFrInternalCd = "" Then 
            strFrInternalCd = "0"
        End If
        If strToInternalCd = "" Then
            strToInternalCd = "zzzzzzzzzz"
        End If
        If strFrInternalCd > strToInternalCd then
	        Call DisplayMsgBox("800359","X","X","X")	'시작부서보다 작은값입니다.
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit sub
        Else
        End if 
    End if 
    
    strWhere = "  a.emp_no = b.emp_no AND a.emp_no = c.emp_no AND a.prov_type = " & FilterVar("Y", "''", "S") & "  "
    strWhere = strWhere & " AND a.emp_no LIKE " & FilterVar(strEmpNo, "''", "S")
    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(frm1.txtValidDt.text,"'%'", "S") & ") >= " & FilterVar(strFrInternalCd , "''", "S")
    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(frm1.txtValidDt.text,"'%'", "S") & ") <= " & FilterVar(strToInternalCd, "''", "S") 
    strWhere = strWhere & " AND b.allow_cd LIKE " & FilterVar(frm1.txtAllow_cd.value, "''", "S")
    
    Call CommonQueryRs(" COUNT(*) AS counts ", " hdf020t a, hdf030t b , haa010t c  ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Trim(Replace(lgF0,Chr(11),"")) = 0 Then
			strDelete = "Normal"
        Else
            IntRetCD = DisplayMsgBox("800502", 35,"X","X")	    '이미 생성된 자료가 있습니다.?
            If IntRetCD = vbCancel Then
	         '  	call FncQuery()
	           	Exit Sub
	        
	        ELSEif IntRetCD = vbYes then
				strDelete = "Del"
			else
				strDelete = "Add"
            End If    
	    END IF
	
	frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    
    lgKeyStream        = frm1.txtEmp_no.Value & parent.gColSep    '0
    
    if  Frm1.txtFr_Dept_cd.Value = "" then
        lgKeyStream    = lgKeyStream & strMin & parent.gColSep    '1
    else
	    lgKeyStream    = lgKeyStream & Frm1.txtFr_internal_cd.Value & parent.gColSep    '1
    end if
	
    if  Frm1.txtTo_Dept_cd.Value = "" then
        lgKeyStream   = lgKeyStream & strMax & parent.gColSep    '2
    else
	    lgKeyStream   = lgKeyStream & Frm1.txtTo_internal_cd.Value & parent.gColSep    '2
    end if	
    strFrYymm = UniConvYYYYMMDDToDate(parent.gDateFormatYYYYMM, frm1.txtapply_yymm_dt.Year,Right("0" & frm1.txtapply_yymm_dt.month , 2),"01")
    strToYymm = UniConvYYYYMMDDToDate(parent.gDateFormatYYYYMM, frm1.txtrevoke_yymm_dt.Year,Right("0" & frm1.txtrevoke_yymm_dt.month , 2),"01")       
	
    lgKeyStream       = lgKeyStream & Frm1.txtAllow_cd.Value & parent.gColSep     '3
	lgKeyStream       = lgKeyStream & Frm1.txtAllow_amt.Text & parent.gColSep    '4
    lgKeyStream       = lgKeyStream & Trim(strFrYymm) & parent.gColSep            '5
    lgKeyStream       = lgKeyStream & Trim(strToYymm) & parent.gColSep            '6
	lgKeyStream       = lgKeyStream & strDelete & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtAllow_nm.Value & parent.gColSep
    lgKeyStream       = lgKeyStream & frm1.txtValidDt.text & parent.gColSep    

    With Frm1
    	strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                          'mb2 자동입력......						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_EMP_NO
            iDx = Trim(Frm1.vspdData.Text)
   	        Frm1.vspdData.Col = C_EMP_NO
    
            If Frm1.vspdData.Text = "" Then
  	            Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.Text = ""
  	            Frm1.vspdData.Col = C_DEPT_CD
                Frm1.vspdData.Text = ""
            Else
	            IntRetCd = FuncGetEmpInf2(iDx,lgUsrIntCd,strName,strDept_nm, strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	            If  IntRetCd < 0 then
	                If  IntRetCd = -1 then
                		Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                    Else
                        Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                    End if
  	                Frm1.vspdData.Col = C_NAME
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_DEPT_CD
                    Frm1.vspdData.Text = ""
                    vspdData_Change = true
                Else
		       	    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.Text = strName
		       	    
		       	    Frm1.vspdData.Col = C_DEPT_CD
		       	    Frm1.vspdData.Text = strDept_nm
                End if 
            End if 
         Case  C_ALLOW_CD
            iDx = Trim(Frm1.vspdData.Text)
            If Trim(Frm1.vspdData.Text) = "" Then
  	            Frm1.vspdData.Col = C_ALLOW_NM
                Frm1.vspdData.Text = ""
            Else
                IntRetCd = CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  and ALLOW_CD =  " & FilterVar(iDx , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If IntRetCd = false then
  	                Frm1.vspdData.Col = C_ALLOW_NM
                    Frm1.vspdData.Text = ""
			        Call DisplayMsgBox("800145","X","X","X")	'수당정보에 등록되지 않은 코드입니다.
                    vspdData_Change = true			        
                Else
		       	    Frm1.vspdData.Col = C_ALLOW_NM
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
                End if 
            End if 
            
    End Select    

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function


'======================================================================================================
'	Name : DBAutoQueryOk()
'	Description : h4007mb2.asp 이후 Query OK해 줌 
'=======================================================================================================
Sub DBAutoQueryOk()
    Dim lRow
	Dim intIndex
	Dim daytimeVal 
    With Frm1
        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
        
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            .vspdData.Text = ggoSpread.InsertFlag

			.vspdData.Col = C_ALLOW_CD
			.vspdData.Text = frm1.txtAllow_cd.value 
		
			.vspdData.Col = C_ALLOW_NM
			.vspdData.Text = frm1.txtAllow_nm.value
					
			.vspdData.col = C_ALLOW_AMT
			.vspdData.Text = frm1.txtAllow_amt.Text
		    
			.vspdData.col = C_APPLY_YYMM_DT
			.vspdData.text = frm1.txtApply_yymm_dt.text
		
			.vspdData.col = C_REVOKE_YYMM_DT
			.vspdData.text = frm1.txtRevoke_yymm_dt.text 			
             
       Next
    
        .vspdData.ReDraw = TRUE
        
         Call InitData()        '수당액 합계를 구한다.........
         ggoSpread.ClearSpreadData "T"
    End With 
  
    Set gActiveElement = document.ActiveElement   
End Sub

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
	
	With Frm1.VspdData
 
    End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
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
        re
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
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1111110111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
       frm1.vspdData.Row = Row
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

'=======================================
'   Event Name :txtApply_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtApply_yymm_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtApply_yymm_dt.Action = 7
        frm1.txtApply_yymm_dt.focus
    End If
End Sub

Sub txtApply_yymm_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

'=======================================
'   Event Name : txtRevoke_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtRevoke_yymm_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRevoke_yymm_dt.Action = 7
        frm1.txtRevoke_yymm_dt.focus
    End If
End Sub

Sub txtRevoke_yymm_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

Sub txtValidDt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        Frm1.txtValidDt.Action = 7
        Frm1.txtValidDt.focus
    End If
End Sub

Sub txtValidDt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
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
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)

	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
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
        IntRetCd = CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  AND CALCU_TYPE = " & FilterVar("N", "''", "S") & "  and ALLOW_CD =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800145","X","X","X")	'수당정보에 등록되지 않은 코드입니다.
			 frm1.txtAllow_nm.value = ""
             frm1.txtAllow_cd.focus
            Set gActiveElement = document.ActiveElement
            txtAllow_cd_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function



'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value, frm1.txtValidDt.text,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement
       	    txtFr_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtFr_dept_nm.value = strDept_nm
            frm1.txtFr_internal_cd.value = lsInternal_cd
        end if        
    End if  
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value, frm1.txtValidDt.text,lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement
       	    txtTo_dept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtTo_dept_nm.value = strDept_nm
            frm1.txtTo_internal_cd.value = lsInternal_cd
        end if
    End if  
End Function

'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split(parent.gDateFormat,parent.gComDateType)
	
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고정수당조회및등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>숨은고정수당조회및삭제</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>		
								     <TD CLASS=TD5 NOWRAP>사원</TD>
				     				    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
								                          <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
			    				     <TD CLASS="TD5" NOWRAP>고정구분</TD>
			    				     <TD CLASS="TD6" NOWRAP><SELECT Name="cboFix_yn" ALT="고정구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
								</TR>
								<TR>
				    				 <TD CLASS=TD5 NOWRAP>부서코드</TD>              
								     <TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
								                          <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
								                          <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=10 MAXLENGTH=30  tag="14XXXU">&nbsp;~&nbsp;
				    									  <INPUT NAME="txtTo_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(1)">
								                          <INPUT NAME="txtTo_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
								                          <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=10 MAXLENGTH=30  tag="14XXXU"></TD>
			    					<TD CLASS="TD5" NOWRAP>부서기준일자</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtValidDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="부서기준일자"></OBJECT>');</SCRIPT>
								                          
								</TR>
								<TR>
									    <TD CLASS="TD5" NOWRAP>수당코드</TD>
									    <TD CLASS="TD6" NOWRAP><INPUT NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10 TYPE="Text"  ALT ="수당코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
									                           <INPUT NAME="txtAllow_nm" MAXLENGTH=20 SIZE=20 TYPE="Text"  ALT ="수당코드명" tag="14XXXU"></TD>
								     <TD CLASS="TD5" NOWRAP>수당금액</TD>
									    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtAllow_amt name=txtAllow_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="11X2" ALT="수당금액"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									    <TD CLASS="TD5" NOWRAP>적용년월</TD>
									    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtApply_yymm_dt" CLASS=FPDTYYYYMM tag="11X1" Title="FPDATETIME" ALT="적용년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									    <TD CLASS="TD5" NOWRAP>해제년월</TD>
									    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtRevoke_yymm_dt" CLASS=FPDTYYYYMM tag="11X1" Title="FPDATETIME" ALT="해제년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
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
					    	<TABLE <%=LR_SPACE_TYPE_20%> >
					    		<TR>
					    			<TD HEIGHT="100%">
					    				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					    			</TD>
					    		</TR>
					    	</TABLE>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO"></DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO"></DIV>
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
	                <TD width=100><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: AutoInsertButtonClicked('1')">자동입력</BUTTON></TD>
	                <TD WIDTH=*>&nbsp;</TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
 <P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

