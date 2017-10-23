<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h3008ma1
*  4. Program Name         : 입/퇴사자 조회 
*  5. Program Desc         : 입/퇴사자 조회 
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
Const BIZ_PGM_ID      = "h3008mb1.asp"						           '☆: Biz Logic ASP Name
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

Dim C_EMP_NO										'Column Dimant for Spread Sheet 
Dim C_NAME 														
Dim C_DEPT_CD 															
Dim C_ROLL_PSTN 														
Dim C_ENTR_DT 	
Dim C_RESENT_PROMOTE_DT														
Dim C_OCPT_TYPE														
Dim C_PAY_GRD1 													
Dim C_PAY_GRD2 													
Dim C_SEX_CD 													
Dim C_RETIRE_DT 															
Dim C_RETIRE_RESN 														

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
	C_EMP_NO = 1											'Column ant for Spread Sheet 
	C_NAME = 2
	C_DEPT_CD = 3
	C_ROLL_PSTN = 4
	C_ENTR_DT = 5
	C_RESENT_PROMOTE_DT = 6
	C_OCPT_TYPE = 7
	C_PAY_GRD1 = 8
	C_PAY_GRD2 = 9
	C_SEX_CD = 10
	C_RETIRE_DT = 11
	C_RETIRE_RESN = 12
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

	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
    
    frm1.txtFrom_dt.focus 
	frm1.txtFrom_dt.Year=strYear
	frm1.txtFrom_dt.Month=strMonth
	frm1.txtFrom_dt.Day="01"
	
	frm1.txtTo_dt.Year=strYear
	frm1.txtTo_dt.Month=strMonth
	frm1.txtTo_dt.Day=strDay
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
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
    
    lgKeyStream       = Trim(frm1.txtFrom_dt.text) & parent.gColSep           'You Must append one character( parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(frm1.txtTo_dt.text) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtFr_internal_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtTo_internal_cd.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtPay_grd1.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & gSelframeFlg & parent.gColSep
    If  lsInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    Else
        lgKeyStream = lgKeyStream & lsInternal_cd & parent.gColSep
    End If
    lgKeyStream = lgKeyStream & StrDt & parent.gColSep
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
End Function
'-------------------------------------------------------------------------------------------------------
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
End Function

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
       .MaxCols   = C_RETIRE_RESN + 1                                                     ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                                  ' ☜:☜:

       .MaxRows = 0
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData
		        
		Call GetSpreadColumnPos("A")       

         ggoSpread.SSSetEdit C_EMP_NO,       "사번", 10
         ggoSpread.SSSetEdit C_NAME,         "성명", 15
         ggoSpread.SSSetEdit C_DEPT_CD,      "부서명", 20
         ggoSpread.SSSetEdit C_ROLL_PSTN,    "직위", 10
         ggoSpread.SSSetDate C_ENTR_DT,      "입사일", 10,2,  parent.gDateFormat
         ggoSpread.SSSetDate C_RESENT_PROMOTE_DT, "최근승급일", 12,2,  parent.gDateFormat
         ggoSpread.SSSetEdit C_OCPT_TYPE,    "직종", 15
         ggoSpread.SSSetEdit C_PAY_GRD1,     "급호", 10
         ggoSpread.SSSetEdit C_PAY_GRD2,     "호봉", 8
         ggoSpread.SSSetEdit C_SEX_CD,       "성별", 8
         ggoSpread.SSSetDate C_RETIRE_DT,    "퇴사일", 10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit C_RETIRE_RESN,  "퇴직사유",18

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
            
			C_EMP_NO = iCurColumnPos(1)										'Column ant for Spread Sheet 
			C_NAME = iCurColumnPos(2)										
			C_DEPT_CD = iCurColumnPos(3)														
			C_ROLL_PSTN = iCurColumnPos(4)														
			C_ENTR_DT = iCurColumnPos(5)													
			C_RESENT_PROMOTE_DT = iCurColumnPos(6)													
			C_OCPT_TYPE = iCurColumnPos(7)														
			C_PAY_GRD1 = iCurColumnPos(8)														
			C_PAY_GRD2 = iCurColumnPos(9)														
			C_SEX_CD = iCurColumnPos(10)													
			C_RETIRE_DT = iCurColumnPos(11)														
			C_RETIRE_RESN = iCurColumnPos(12)
    End Select    
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
		
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	gSelframeFlg = TAB1
	Call changeTabs(TAB1)
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
    
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
    Dim IntRetCD , rFrDept ,rToDept
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    If  ValidDateCheck(frm1.txtFrom_dt, frm1.txtTo_dt)=False Then
        Exit Function
    End If

    If txtPay_grd1_OnChange() Then
        Exit Function
    End If

    If txtFr_dept_cd_OnChange() Then
        Exit Function
    End If
    If txtTo_dept_cd_OnChange() Then
        Exit Function
    End If
	
	Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
	
    If frm1.txtFr_internal_cd.value="" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If frm1.txtTo_internal_cd.value = "" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  

    If Trim(frm1.txtFr_internal_cd.value) > Trim(frm1.txtTo_internal_cd.value) then
	    Call  DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
        frm1.txtFr_internal_cd.value = ""
        frm1.txtFr_dept_cd.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End IF 

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
		                                                                        'Required Field Check 시 Error발생하면 Tab이동후 이동한 tab page번호를 
		                                                                        'gSelframeFlg(tab page Flag)에게 넘겨줍니다.
		If gPageNo > 0 Then
			gSelframeFlg = gPageNo
		End If
       Exit Function
    End If
    
    Call MakeKeyStream("X")

    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncQuery = True																'☜: Processing is OK

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
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

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
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
	Frm1.vspdData.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call FncQuery()
	Call ClickTab1()
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
' Name : FncOpenPopup
' Desc : developer describe this line 
'========================================================================================================
Function FncOpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   
	
	Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
	
	IsOpenPop = True
	Select Case iWhere
	    Case "1"
	        arrParam(0) = "부서코드 팝업"			                                     ' 팝업 명칭 
	    	arrParam(1) = "b_acct_dept"						                                 ' TABLE 명칭 
	    	arrParam(2) = frm1.txtFr_dept_cd.value                         	        		 ' Code Condition
	    	arrParam(3) = ""'frm1.txtFr_dept_nm.value 	    		                             ' Name Cindition
	    	arrParam(4) = " org_change_dt = (select max(org_change_dt) from b_acct_dept "    ' Where Condition
	    	arrParam(4) = arrParam(4) & " where org_change_dt <= StrDt)"
	    	arrParam(5) = "부서코드" 			                                         ' TextBox 명칭 
	            
	    	arrField(0) = "DEPT_CD"						                                     ' Field명(0)
	    	arrField(1) = "DEPT_NM"    					    	                             ' Field명(1)
	    	arrField(2) = "INTERNAL_CD"    					                                 ' Field명(1)
    
	    	arrHeader(0) = "부서코드"	   		    	                                 ' Header명(0)
	    	arrHeader(1) = "부서명"	    		                                         ' Header명(1)

	    Case "2"
	        
	        arrParam(0) = "부서코드 팝업"			                                     ' 팝업 명칭 
	    	arrParam(1) = "b_acct_dept"						                                 ' TABLE 명칭 
	    	arrParam(2) = frm1.txtTo_dept_cd.value                                       	 ' Code Condition
	    	arrParam(3) = ""'frm1.txtTo_dept_nm.value	                    		        	 ' Name Cindition
	    	arrParam(4) = " org_change_dt = (select max(org_change_dt) from b_acct_dept "    ' Where Condition
	    	arrParam(4) = arrParam(4) & " where org_change_dt <= StrDt)"
	    	arrParam(5) = "부서코드" 			                                         ' TextBox 명칭 
	
	    	arrField(0) = "DEPT_CD"					                    	                 ' Field명(0)
	    	arrField(1) = "DEPT_NM"    			                    		    	         ' Field명(1)
	    	arrField(2) = "INTERNAL_CD"    			                   		    	         ' Field명(1)

	    	arrHeader(0) = "부서코드"	   	        	    	                         ' Header명(0)
	    	arrHeader(1) = "부서명"	    		                                         ' Header명(1)

	    Case "3"
	        
            arrParam(0) = "급호조회 팝업"   ' 팝업 명칭 
            arrParam(1) = "B_MINOR"       ' TABLE 명칭 
            arrParam(2) = frm1.txtPay_grd1.value         ' Code Condition
            arrParam(3) = ""     ' Name Cindition
            arrParam(4) = " MAJOR_CD=" & FilterVar("H0001", "''", "S") & ""      ' Where Condition
            arrParam(5) = "급호코드"       ' TextBox 명칭 
 
            arrField(0) = "MINOR_CD"     ' Field명(0)
            arrField(1) = "MINOR_NM"        ' Field명(1)
    
            arrHeader(0) = "급호코드"    ' Header명(0)
            arrHeader(1) = "급호명"           ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "1"
		        frm1.txtFr_dept_cd.focus	
		    Case "2"
		        frm1.txtTo_dept_cd.focus
		    Case "3"
		        frm1.txtPay_grd1.focus
        End Select	
		Exit Function
	Else
		Call SubSetOpenPop(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SubSetOpenPop()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetOpenPop(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtFr_dept_cd.value = arrRet(0)
		        .txtFr_dept_nm.value = arrRet(1)		
		        .txtFr_internal_cd.value = arrRet(2)	
		        .txtFr_dept_cd.focus	
		    Case "2"
		        .txtTo_dept_cd.value = arrRet(0)
		        .txtTo_dept_nm.value = arrRet(1)		
		        .txtTo_internal_cd.value = arrRet(2)		
		        .txtTo_dept_cd.focus
            Case "3"
                .txtPay_grd1.value = arrRet(0)
                .txtPay_grd1_nm.value = arrRet(1)  
                .txtPay_grd1.focus
        End Select
	End With
End Sub
'----------------------------------------  OpenDeptDt()  ------------------------------------------
'	Name : OpenDeptDt()
'	Description : 특정일자 입력받은 Dept PopUp
'---------------------------------------------------------------------------------------------------
Function OpenDeptDt(iWhere,strDate,TargetObj,TargetObj1)
	Dim arrRet
	Dim arrParam(2)
	'Dim StrDt
        'StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = TargetObj.Value          	    		' Code Condition%>
        If strDate="X" Then
        	arrParam(1) = frm1.txtFrom_dt.text                                ' 현재날짜!!!
        Else
        	arrParam(1) = frm1.txtFrom_dt.text                           ' 특정 Date값을 parameter(1)로 넘긴다!!!
        End If    
	Else 'spread
		arrParam(0) = StrDt            		    	        ' Code Condition%>
        frm1.vspdData.Col = C_GAZET_DT
        arrParam(1) = frm1.vspdData.Text                    ' 특정 Date값을 parameter(1)로 넘긴다!!!
	End If
        arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			TargetObj.focus
		Else 'spread
			frm1.vspdData.Col = C_DEPT_CD
			frm1.vspdData.action =0
		End If
	
		Exit Function
	Else
		Call SetDeptDt(arrRet, iWhere, TargetObj,TargetObj1)
	End If	
			
End Function

'------------------------------------------  SetDeptDt()  ---------------------------------------------
'	Name : SetDeptDt()
'	Description : Dept Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetDeptDt(Byval arrRet, Byval iWhere, Byval TargetObj, Byval TargetObj1)
		
		If iWhere = 0 Then 'TextBox(Condition)
			TargetObj.Value = arrRet(0)
			TargetObj1.Value = arrRet(1)
			TargetObj.focus
		Else 'spread
        	With frm1
			.vspdData.Col = C_DEPT_NM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
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
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

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


Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
End Sub
'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    'Dim StrDt
    'StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtFr_dept_cd.value , frm1.txtFrom_dt.text, lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtFrom_dt.Year, Right("0" & frm1.txtFrom_dt.month , 2), "01")
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtTo_dept_cd.value , frm1.txtFrom_dt.text, lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'======================================================================================================
'   Event Name : txtPay_grd1_OnChange
'   Event Desc : 직급코드가 변경될 경우 
'=======================================================================================================
Function txtPay_grd1_OnChange()
    Dim IntRetCd

    If Trim(frm1.txtPay_grd1.value) = "" Then
        frm1.txtPay_grd1_nm.Value = ""
    Else
        IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtPay_grd1.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtPay_grd1.Value)<>""  Then
            frm1.txtPay_grd1_nm.Value=""
            Call DisplayMsgBox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다.
			txtPay_grd1_OnChange = true
        Else
            frm1.txtPay_grd1_nm.Value=Trim(Replace(lgF0,Chr(11),""))
        
        End If
    End If
End Function

'=======================================================================================================
'   Event Name : txtFrom_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrom_dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrom_dt.Action = 7
		Call SetFocusToDocument("M")     
        frm1.txtFrom_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTo_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtTo_dt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTo_dt.Action = 7
		Call SetFocusToDocument("M")         
        frm1.txtTo_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFrom_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtFrom_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTo_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtTo_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입사자조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>퇴사자조회</font></td>
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
								<TD CLASS=TD5 NOWRAP>조회기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD 
									name=txtFrom_dt classid=<%=gCLSIDFPDT%> ALT="시작조회기간" tag="11X1" VIEWASTEXT></OBJECT>
	                                &nbsp;~&nbsp;
	                                <OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD 
	                                name=txtTo_dt classid=<%=gCLSIDFPDT%>ALT="종료조회기간" tag="11x1" VIEWASTEXT></OBJECT>
								</TD>
								<TD CLASS=TD5 NOWRAP>부서코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd"  SIZE=10  MAXLENGTH=10  ALT ="시작부서코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: Call OpenDeptDt(0,'X',frm1.txtFr_dept_cd,frm1.txtFr_dept_nm)">
								                     <INPUT NAME="txtFr_dept_nm"  SIZE=20  MAXLENGTH=40  ALT ="시작부서명"   tag="14XXXU">
								                     <INPUT NAME="txtFr_internal_cd" TYPE="HIDDEN" MAXLENGTH="20" SIZE=20  ALT ="시작부서명"   tag="14XXXU">부터
							</TR>
							<TR>	
               					<TD CLASS="TD5" NOWRAP>급호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd1" ALT="급호" TYPE="Text" MAXLENGTH=10 SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:FncOpenPopup('3')">&nbsp;<INPUT NAME="txtPay_grd1_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
								</TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTo_dept_cd"  SIZE=10  MAXLENGTH=10  ALT ="종료부서코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: Call OpenDeptDt(0,'X',frm1.txtTo_dept_cd,frm1.txtTo_dept_nm)">
								                     <INPUT NAME="txtTo_dept_nm"  SIZE=20  MAXLENGTH=40  ALT ="종료부서명"   tag="14XXXU">
								                     <INPUT NAME="txtTo_internal_cd" TYPE="HIDDEN" MAXLENGTH="20" SIZE=20  ALT ="종료부서명"   tag="14XXXU">까지</TD>
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
					    				<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread>
					    					<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
					    				</OBJECT>
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

