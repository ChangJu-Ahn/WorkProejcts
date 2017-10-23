<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 정기호봉상승등록 
*  3. Program ID           : H3004ma1
*  4. Program Name         : H3004ma1
*  5. Program Desc         : 근무이력관리/정기호봉상승등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/25
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H3006mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_ID1 = "H3006mb2.asp"
Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_NAME
Dim C_EMP_NO_POP 
Dim C_EMP_NO 
Dim C_DEPT_CD 
Dim C_DEPT_NM
Dim C_PAY_GRD1 
Dim C_PAY_GRD1_NM 
Dim C_PAY_GRD2 
Dim C_ROLL_PSTN
Dim C_ROLL_PSTN_NM
Dim C_OCPT_TYPE 
Dim C_OCPT_TYPE_NM
Dim C_FUNC_CD
Dim C_FUNC_CD_NM
Dim C_ROLE_CD 
Dim C_ROLE_CD_NM
Dim C_ENTR_DT
Dim C_RESENT_PROMOTE_DT
Dim C_CHNG_DEPT_CD
Dim C_CHNG_PAY_GRD1
Dim C_CHNG_PAY_GRD2
Dim C_CHNG_ROLL_PSTN
Dim C_CHNG_OCPT_TYPE
Dim C_CHNG_FUNC_CD
Dim C_CHNG_ROLE_CD
Dim C_PROMOTE_DT
Dim C_CHNG_CD 
Dim C_CHNG_CD_NM 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	C_NAME = 1
	C_EMP_NO_POP = 2
	C_EMP_NO = 3
	C_DEPT_CD = 4
	C_DEPT_NM = 5
	C_PAY_GRD1 = 6
	C_PAY_GRD1_NM = 7
	C_PAY_GRD2 = 8 
	C_ROLL_PSTN = 9
	C_ROLL_PSTN_NM = 10
	C_OCPT_TYPE = 11
	C_OCPT_TYPE_NM = 12
	C_FUNC_CD = 13
	C_FUNC_CD_NM = 14
	C_ROLE_CD = 15
	C_ROLE_CD_NM = 16
	C_ENTR_DT = 17
	C_RESENT_PROMOTE_DT = 18
	C_CHNG_DEPT_CD = 19
	C_CHNG_PAY_GRD1 = 20
	C_CHNG_PAY_GRD2 = 21
	C_CHNG_ROLL_PSTN = 22
	C_CHNG_OCPT_TYPE = 23
	C_CHNG_FUNC_CD = 24
	C_CHNG_ROLE_CD = 25
	C_PROMOTE_DT = 26
	C_CHNG_CD = 27
	C_CHNG_CD_NM = 28

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
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	frm1.txtChng_cd.value = "98"
	frm1.txtChng_cd_nm.value = "정기호봉상승"
	frm1.txtResent_promote_dt.text =  UniConvDateAToB("<%=GetSvrDate%>",  parent.gServerDateFormat,  parent.gDateFormat)
	frm1.txtPro_dt.text = frm1.txtResent_promote_dt.text
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
		
		Call FncQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
   
    lgKeyStream       = Frm1.txtEmp_no.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtName.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Call  CommonQueryRs(" minor_cd, minor_nm "," b_minor "," major_cd = " & FilterVar("H0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.txtPay_grd1, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim iDx
    With frm1
		For intRow = 1 To .vspdData.MaxRows	'C_PAY_GRD2 C_CHNG_PAY_GRD2		
			.vspdData.Row = intRow
			.vspdData.Col = C_PAY_GRD2
			iDx = .vspdData.Text
			.vspdData.Col = C_CHNG_PAY_GRD2
			If  Trim(.vspdData.Text) = iDx Then
                  ggoSpread.Source = .vspdData
			    .vspdData.Col = -1 
			    .vspdData.Col2 = -1
			    .vspdData.ForeColor = RGB(255,0,0)
			End If
		Next
    End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	

         ggoSpread.Source = frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_CHNG_CD_NM + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>

	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0	
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 
		Call GetSpreadColumnPos("A")

         ggoSpread.SSSetEdit   C_NAME,             "성명",      08,,,08,2
         ggoSpread.SSSetButton C_EMP_NO_POP
         ggoSpread.SSSetEdit   C_EMP_NO,           "사번",      13,,,13,2
         ggoSpread.SSSetEdit   C_DEPT_CD,          "부서",      05,,,15,2
         ggoSpread.SSSetEdit   C_DEPT_NM,          "부서",      10,,,30,2
         ggoSpread.SSSetEdit   C_PAY_GRD1,         "급호",      05,,,15,2
         ggoSpread.SSSetEdit   C_PAY_GRD1_NM,      "급호",      10,,,20,2
         ggoSpread.SSSetEdit   C_PAY_GRD2,         "호봉",      06,,,15,2
         ggoSpread.SSSetEdit   C_ROLL_PSTN,        "직위",      05,,,15,2
         ggoSpread.SSSetEdit   C_ROLL_PSTN_NM,     "직위",      10,,,20,2
         ggoSpread.SSSetEdit   C_OCPT_TYPE,        "직종",      05,,,15,2
         ggoSpread.SSSetEdit   C_OCPT_TYPE_NM,     "직종",      10,,,20,2
         ggoSpread.SSSetEdit   C_FUNC_CD,          "직무",      05,,,15,2
         ggoSpread.SSSetEdit   C_FUNC_CD_NM,       "직무",      10,,,20,2
         ggoSpread.SSSetEdit   C_ROLE_CD,          "직책",      05,,,15,2
         ggoSpread.SSSetEdit   C_ROLE_CD_NM,       "직책",      10,,,20,2
         ggoSpread.SSSetDate   C_ENTR_DT,          "입사일",    10,2,  parent.gDateFormat
         ggoSpread.SSSetDate   C_RESENT_PROMOTE_DT,"최근승급일",10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit   C_CHNG_DEPT_CD,     "변동부서",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD1,    "변동급호",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD2,    "변동호봉",  10,,,3,2
         ggoSpread.SSSetEdit   C_CHNG_ROLL_PSTN,   "변동직위",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_OCPT_TYPE,   "변동직종",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_FUNC_CD,     "변동직무",  05,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_ROLE_CD,     "변동직책",  05,,,15,2
         ggoSpread.SSSetDate   C_PROMOTE_DT,       "승급예정일",10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit   C_CHNG_CD,          "변동사유",  10,,,15,2
         ggoSpread.SSSetEdit   C_CHNG_CD_NM,       "변동사유",  15,,,40,2

       call ggoSpread.MakePairsColumn(C_EMP_NO_POP,C_EMP_NO)
                       
        Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)	
        Call ggoSpread.SSSetColHidden(C_PAY_GRD1,C_PAY_GRD1,True)	
        Call ggoSpread.SSSetColHidden(C_ROLL_PSTN,C_ROLL_PSTN,True)	
        Call ggoSpread.SSSetColHidden(C_OCPT_TYPE,C_OCPT_TYPE,True)
        Call ggoSpread.SSSetColHidden(C_FUNC_CD,C_FUNC_CD,True)	
        Call ggoSpread.SSSetColHidden(C_ROLE_CD,C_ROLE_CD,True)	
        Call ggoSpread.SSSetColHidden(C_EMP_NO_POP,C_EMP_NO_POP,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_DEPT_CD,C_CHNG_DEPT_CD,True)
        Call ggoSpread.SSSetColHidden(C_CHNG_PAY_GRD1,C_CHNG_PAY_GRD1,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_ROLL_PSTN,C_CHNG_ROLL_PSTN,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_OCPT_TYPE,C_CHNG_OCPT_TYPE,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_FUNC_CD,C_CHNG_FUNC_CD,True)
        Call ggoSpread.SSSetColHidden(C_CHNG_ROLE_CD,C_CHNG_ROLE_CD,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_CD,C_CHNG_CD,True)

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
            
			C_NAME = iCurColumnPos(1)
			C_EMP_NO_POP = iCurColumnPos(2)
			C_EMP_NO = iCurColumnPos(3)
			C_DEPT_CD = iCurColumnPos(4)
			C_DEPT_NM = iCurColumnPos(5)
			C_PAY_GRD1 = iCurColumnPos(6)
			C_PAY_GRD1_NM = iCurColumnPos(7)
			C_PAY_GRD2 = iCurColumnPos(8)
			C_ROLL_PSTN = iCurColumnPos(9)
			C_ROLL_PSTN_NM = iCurColumnPos(10)
			C_OCPT_TYPE = iCurColumnPos(11)
			C_OCPT_TYPE_NM = iCurColumnPos(12)
			C_FUNC_CD = iCurColumnPos(13)
			C_FUNC_CD_NM = iCurColumnPos(14)
			C_ROLE_CD = iCurColumnPos(15)
			C_ROLE_CD_NM = iCurColumnPos(16)
			C_ENTR_DT = iCurColumnPos(17)
			C_RESENT_PROMOTE_DT = iCurColumnPos(18)
			C_CHNG_DEPT_CD = iCurColumnPos(19)
			C_CHNG_PAY_GRD1 = iCurColumnPos(20)
			C_CHNG_PAY_GRD2 = iCurColumnPos(21)
			C_CHNG_ROLL_PSTN = iCurColumnPos(22)
			C_CHNG_OCPT_TYPE = iCurColumnPos(23)
			C_CHNG_FUNC_CD = iCurColumnPos(24)
			C_CHNG_ROLE_CD = iCurColumnPos(25)
			C_PROMOTE_DT = iCurColumnPos(26)
			C_CHNG_CD = iCurColumnPos(27)
			C_CHNG_CD_NM = iCurColumnPos(28)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
     ggoSpread.SpreadLock C_NAME, -1,C_NAME
     ggoSpread.SpreadLock C_EMP_NO_POP, -1,C_EMP_NO_POP     
     ggoSpread.SpreadLock C_EMP_NO, -1,C_EMP_NO
     ggoSpread.SpreadLock C_DEPT_CD, -1,C_DEPT_CD     
     ggoSpread.SpreadLock C_DEPT_NM, -1,C_DEPT_NM     
     ggoSpread.SpreadLock C_PAY_GRD1, -1,C_PAY_GRD1    
     ggoSpread.SpreadLock C_PAY_GRD1_NM, -1,C_PAY_GRD1_NM     
     ggoSpread.SpreadLock C_PAY_GRD2, -1,C_PAY_GRD2 
     ggoSpread.SpreadLock C_ROLL_PSTN, -1,C_ROLL_PSTN    
     ggoSpread.SpreadLock C_ROLL_PSTN_NM, -1,C_ROLL_PSTN_NM     
     ggoSpread.SpreadLock C_OCPT_TYPE, -1,C_OCPT_TYPE     
     ggoSpread.SpreadLock C_OCPT_TYPE_NM, -1,C_OCPT_TYPE_NM    
     ggoSpread.SpreadLock C_FUNC_CD, -1,C_FUNC_CD     
     ggoSpread.SpreadLock C_FUNC_CD_NM, -1,C_FUNC_CD_NM 
     ggoSpread.SpreadLock C_ROLE_CD, -1,C_ROLE_CD        
     ggoSpread.SpreadLock C_ROLE_CD_NM, -1,C_ROLE_CD_NM       
     ggoSpread.SpreadLock C_ENTR_DT, -1,C_ENTR_DT                              
     ggoSpread.SpreadLock C_RESENT_PROMOTE_DT, -1,C_RESENT_PROMOTE_DT                                                             
     ggoSpread.SpreadLock C_CHNG_DEPT_CD, -1,C_CHNG_DEPT_CD                              
     ggoSpread.SpreadLock C_CHNG_PAY_GRD1, -1,C_CHNG_PAY_GRD1                                                             
     ggoSpread.SpreadLock C_CHNG_PAY_GRD2, -1,C_CHNG_PAY_GRD2                              
     ggoSpread.SpreadLock C_CHNG_ROLL_PSTN, -1,C_CHNG_ROLL_PSTN                                                             
     ggoSpread.SpreadLock C_CHNG_OCPT_TYPE, -1,C_CHNG_OCPT_TYPE                                                             
     ggoSpread.SpreadLock C_CHNG_FUNC_CD, -1,C_CHNG_FUNC_CD 
     ggoSpread.SpreadLock C_CHNG_CD_NM, -1,C_CHNG_CD_NM                                                                  
     ggoSpread.SpreadLock C_CHNG_CD, -1,C_CHNG_CD
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
     ggoSpread.SSSetProtected	C_NAME, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_EMP_NO, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD1, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD2, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_ROLL_PSTN, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_OCPT_TYPE, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_FUNC_CD, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_CHNG_ROLE_CD, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired		C_PROMOTE_DT, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	C_CHNG_CD, pvStartRow, pvEndRow
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
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1000100100001111")										        '버튼 툴바 제어 
    frm1.txtPro_dt.focus
    
    Call InitComboBox
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

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
                                                                 '☜: Query db data
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
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
    Dim lRow
    Dim strAdmi_dt
    Dim strGrudt_dt

    Dim strPay_grd1
    Dim strPay_grd2
    Dim strPromote_dt
    Dim strSQL
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If  Trim(frm1.txtResent_promote_dt.Text) <> "" and Trim(frm1.txtResent_promote_dt.Text) = Trim(frm1.txtPro_dt.Text) then
        '변동일은 최근승급일보다 커야합니다.
        Call  DisplayMsgBox("800396","X","X","X")
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit Function
    ElseIf   CompareDateByFormat(frm1.txtResent_promote_dt.Text,frm1.txtPro_dt.Text,frm1.txtResent_promote_dt.Alt,frm1.txtPro_dt.Alt,"800396", parent.gDateFormat, parent.gComDateType,True) = False THEN
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit Function
    END IF

    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
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
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
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
         ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
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
    Call SetSpreadLock
    Call SetSpreadColor(1,Frm1.vspdData.MaxRows)
End Sub
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

    If LayerShowHide(1)=False Then
		Exit Function
	End If
	
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
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

	Dim strRes_no

    DbSave = False                                                          
    
    If LayerShowHide(1)=False Then
		Exit Function
	End If


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
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCPT_TYPE         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESENT_PROMOTE_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD1	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLL_PSTN	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_OCPT_TYPE    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_FUNC_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLE_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PAY_GRD2	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLL_PSTN	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_OCPT_TYPE         : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_FUNC_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ROLE_CD           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RESENT_PROMOTE_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_DEPT_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD1	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_PAY_GRD2	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLL_PSTN	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_OCPT_TYPE    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_FUNC_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_ROLE_CD      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT: strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CD   : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
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
	If DBDelete=False Then
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
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
    Call SetToolbar("1000100100001111")										        '버튼 툴바 제어 

	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

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
		If iWhere = 0 Then
			frm1.txtDept_cd.focus
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
		If iWhere = 0 Then
			.txtDept_cd.value = arrRet(0)
			.txtDept_cd_Nm.value = arrRet(1)
			.txtDept_cd.focus
		End If
	End With
End Function

'========================================================================================================
' Name : OpenPromoteDt
' Desc : 최근 승급일 POPUP
'========================================================================================================
Function OpenPromoteDt(iWhere)
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtResent_promote_dt.text			<%' 조건부에서 누른 경우 Code Condition%>
	arrParam(1) = 1 'haa010t
	arrRet = window.showModalDialog(HRAskPRAspName("PromoteDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtResent_promote_dt.text = arrRet(0)
	end if
		frm1.txtResent_promote_dt.focus	
End Function

'======================================================================================================
'   Event Name : btnAuto_OnClick
'   Event Desc : 자동입력버튼 
'=======================================================================================================
Sub btnAuto_OnClick()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd

    If  frm1.txtResent_promote_dt.Text = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtResent_promote_dt.Alt,"X")
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if

    If  frm1.txtPro_dt.Text = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtPro_dt.Alt,"X")
        frm1.txtPro_dt.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if

    If  frm1.txtChng_cd.value = "" then
        Call  DisplayMsgBox("970021","X",frm1.txtChng_cd.Alt,"X")
        frm1.txtChng_cd.focus
        Set gActiveElement = document.ActiveElement
        exit sub        
    end if   

    If  Trim(frm1.txtResent_promote_dt.Text) <> "" and Trim(frm1.txtResent_promote_dt.Text) = Trim(frm1.txtPro_dt.Text) then
        '변동일은 최근승급일보다 커야합니다.
        Call  DisplayMsgBox("800396","X","X","X")
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit Sub
    ElseIf   CompareDateByFormat(frm1.txtResent_promote_dt.Text,frm1.txtPro_dt.Text,frm1.txtResent_promote_dt.Alt,frm1.txtPro_dt.Alt,"800396", parent.gDateFormat, parent.gComDateType,True) = False THEN
        frm1.txtResent_promote_dt.focus
        Set gActiveElement = document.ActiveElement
        exit Sub
    END IF

	ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData 
    if  frm1.txtDept_cd.value = "" then
        strInternal_cd = ""
    else
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            exit sub
        end if
    end if

    lgKeyStream = Frm1.txtResent_promote_dt.Text & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtPay_grd1.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtPay_grd2.Value & parent.gColSep
    
    if  strInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
        lgKeyStream = lgKeyStream & strInternal_cd & parent.gColSep
    end if
    lgKeyStream = lgKeyStream & Frm1.txtPro_dt.Text & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
	
	If AutoDbQuery=False Then
	   Call  RestoreToolBar()
	   Exit Sub
	End If

End Sub
'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function AutoDbQuery() 

    AutoDbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

    If LayerShowHide(1)=False Then
		Exit Function
	End If
	
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With

		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    AutoDbQuery = True
    
End Function
'========================================================================================================
' Function Name : DbAutoQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbAutoQueryOk()													     

    Dim lRow	

    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field

    
     ggoSpread.Source = frm1.vspdData

	frm1.vspddata.ReDraw = false
     ggoSpread.SSSetProtected	C_NAME, -1
     ggoSpread.SSSetProtected	C_EMP_NO, -1
     ggoSpread.SSSetProtected	C_PAY_GRD2, -1
     ggoSpread.SSSetProtected	C_CHNG_CD, -1

     ggoSpread.SpreadUnLock C_CHNG_DEPT_CD, -1, C_PROMOTE_DT

     ggoSpread.SSSetRequired		C_CHNG_PAY_GRD2, -1

     ggoSpread.SSSetRequired		C_PROMOTE_DT, -1

    frm1.vspdData.Row = -1
    frm1.vspdData.Col = C_PROMOTE_DT
    frm1.vspdData.text = frm1.txtPro_dt.text


    frm1.vspdData.Col = C_CHNG_CD
    frm1.vspdData.text = frm1.txtChng_cd.value
    frm1.vspdData.Col = C_CHNG_CD_NM
    frm1.vspdData.text = frm1.txtChng_cd_nm.value    
    
    For lRow = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row = lRow
        frm1.vspdData.Col = 0
        frm1.vspdData.text =  ggoSpread.InsertFlag
    Next
    
    frm1.vspddata.ReDraw = true
    ggoSpread.ClearSpreadData "T"    
    Call InitData()
    Call SetToolbar("1000100100001111")										        '버튼 툴바 제어 

End Function
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
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
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
	End With

   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000110111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    frm1.vspdData.Row = Row
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    
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
			If AutoDBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'========================================================================================================
' Name : txtPro_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtPro_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtPro_dt.Action = 7 
        frm1.txtPro_dt.focus
    End If
    lgBlnFlgChgValue = True    
End Sub

Sub txtPro_dt_Keypress(Key)
    lgBlnFlgChgValue = True    
End Sub

'========================================================================================================
' Name : txtResent_promote_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtResent_promote_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtResent_promote_dt.Action = 7 
        frm1.txtResent_promote_dt.focus
    End If
    lgBlnFlgChgValue = True
End Sub


Sub txtResent_promote_dt_Keypress(Key)
    lgBlnFlgChgValue = True    
End Sub


Sub txtDept_cd_OnChange()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd

    if  frm1.txtDept_cd.value <> "" then    
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtDept_cd_nm.value = ""
        else
            frm1.txtDept_cd_nm.value = strDept_nm
        end if
       lgBlnFlgChgValue = True    
	Else
			frm1.txtDept_cd_nm.value = ""      
    end if
			
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>정기호봉상승등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=100%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>급호</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_grd1" ALT="급호" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT>&nbsp;<INPUT NAME="txtPay_grd2" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN:left" tag="11"></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(0)">&nbsp;<INPUT NAME="txtDept_cd_nm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최근승급일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h3006ma1_txtResent_promote_dt_txtResent_promote_dt.js'></script>
								<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPromoteDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPromoteDt(0)"></TD>
								<TD CLASS=TD5 NOWRAP>변동일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h3006ma1_txtPro_dt_txtPro_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>변동사유</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_cd" ALT="변동사유" TYPE="Text" MAXLENGTH=2 SiZE=5 tag=14XXXU>&nbsp;<INPUT NAME="txtChng_cd_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
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
									<script language =javascript src='./js/h3006ma1_vaSpread1_vspdData.js'></script>
								</TD>
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
	<TR HEIGHT="30">
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnAuto" CLASS="CLSMBTN">자동입력</BUTTON></TD>
				    <TD WIDTH=* Align=RIGHT></TD>
				    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

