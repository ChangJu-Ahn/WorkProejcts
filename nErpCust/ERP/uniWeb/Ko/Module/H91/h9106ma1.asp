<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h9106ma1
*  4. Program Name         : h9106ma1
*  5. Program Desc         : 연말정산관리/연말정산/현물지급급여등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/07
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H9106mb1.asp"                                      'Biz Logic ASP 
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
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)
Dim lsInternal_cd

DIM C_EMP_NO                  
DIM C_EMP_NO_POP              
DIM C_NAME                    
DIM C_NAME_POP                
DIM C_HFA080T_PAY_YYMM_DT     
DIM C_HFA080T_PAY_TEXT        
DIM C_HFA080T_PAY_TOT_AMT     
DIM C_HFA080T_INCOME_TAX_AMT  
DIM C_HFA080T_RES_TAX_AMT     
DIM C_HFA080T_MED_INSUR_AMT   
DIM C_HFA080T_EMP_INSUR_AMT   

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
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    Call ggoOper.FormatDate(frm1.txtFrom_dt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtTo_dt, parent.gDateFormat, 1)
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H",  "NOCOOKIE","MA") %>
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
Sub MakeKeyStream(pRow)
    lgKeyStream = Frm1.txtEmp_no.Value & parent.gColSep						'0    
    lgKeyStream = lgKeyStream & frm1.txtFrom_dt.Text & parent.gColSep			'1You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & frm1.txtTo_dt.Text & parent.gColSep			'2You Must append one character(parent.gColSep)    
    lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep					'3    
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

sub InitSpreadPosVariables()
	C_EMP_NO                  = 1
	C_EMP_NO_POP              = 2
	C_NAME                    = 3
	C_NAME_POP                = 4
	C_HFA080T_PAY_YYMM_DT     = 5
	C_HFA080T_PAY_TEXT        = 6
	C_HFA080T_PAY_TOT_AMT     = 7
	C_HFA080T_INCOME_TAX_AMT  = 8
	C_HFA080T_RES_TAX_AMT     = 9
	C_HFA080T_MED_INSUR_AMT   = 10
	C_HFA080T_EMP_INSUR_AMT   = 11
end sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	call InitSpreadPosVariables
	With frm1.vspdData
	
        .MaxCols = C_HFA080T_EMP_INSUR_AMT + 1										
	    .Col = .MaxCols	
        .ColHidden = True

	 
        .MaxRows = 0
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021126",, parent.gAllowDragDropSpread
		Call GetSpreadColumnPos("A")
	    
	    .ReDraw = false
	    
        ggoSpread.SSSetEdit     C_NAME,     "성명",          15,,, 30,2
        ggoSpread.SSSetButton   C_NAME_POP
        ggoSpread.SSSetEdit     C_EMP_NO,     "사번",        12,,, 13,2
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetDate     C_HFA080T_PAY_YYMM_DT,       "현물지급일",     14,2, parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetEdit     C_HFA080T_PAY_TEXT,          "지급내역",       18,,, 50
        ggoSpread.SSSetFloat    C_HFA080T_PAY_TOT_AMT,       "지급액" ,        10,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_HFA080T_INCOME_TAX_AMT,    "소득세" ,        10,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_HFA080T_RES_TAX_AMT,       "주민세" ,        10,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_HFA080T_MED_INSUR_AMT,     "의료보험" ,      12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_HFA080T_EMP_INSUR_AMT,     "고용보험" ,      12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		
		call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)
		call ggoSpread.SSSetColHidden(C_NAME_POP,C_NAME_POP,true)
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
    
      ggoSpread.SpreadLock      C_NAME , -1, C_NAME
      ggoSpread.SpreadLock      C_NAME_POP , -1, C_NAME_POP
      ggoSpread.SpreadLock      C_EMP_NO , -1, C_EMP_NO
      ggoSpread.SpreadLock      C_EMP_NO_POP , -1, C_EMP_NO_POP
      ggoSpread.SpreadLock      C_HFA080T_PAY_YYMM_DT , -1, C_HFA080T_PAY_YYMM_DT
      ggoSpread.SSSetProtected  .vspdData.MaxCols   , -1, -1
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
      ggoSpread.SSSetRequired    C_EMP_NO , pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_HFA080T_PAY_YYMM_DT , pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    
    End With
End Sub


'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_EMP_NO                  = iCurColumnPos(1)
			C_EMP_NO_POP              = iCurColumnPos(2)
			C_NAME                    = iCurColumnPos(3)
			C_NAME_POP                = iCurColumnPos(4)
			C_HFA080T_PAY_YYMM_DT     = iCurColumnPos(5)
			C_HFA080T_PAY_TEXT        = iCurColumnPos(6)
			C_HFA080T_PAY_TOT_AMT     = iCurColumnPos(7)
			C_HFA080T_INCOME_TAX_AMT  = iCurColumnPos(8)
			C_HFA080T_RES_TAX_AMT     = iCurColumnPos(9)
			C_HFA080T_MED_INSUR_AMT   = iCurColumnPos(10)
			C_HFA080T_EMP_INSUR_AMT   = iCurColumnPos(11)
    End Select    
End Sub

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
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'======================================================================================================
' Function Name : vspdData_ScriptLeaveCell
' Function Desc :
'======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call SetToolbar("1100111100101111")										        '버튼 툴바 제어 
    
    frm1.txtEmp_no.Focus

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
    Dim strFromDt
    Dim strToDt
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if


    strFromDt = frm1.txtFrom_dt.Text
    strToDt   = frm1.txtTo_dt.Text
    
    If (strFromDt = "") AND (strToDt = "") Then
        strFromDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"1900","01","01")
        strToDt   = UniConvYYYYMMDDToDate(parent.gDateFormat,"2500","12","31")
    Else
        If strFromDt = "" Then
            strFromDt =  UniConvYYYYMMDDToDate(parent.gDateFormat,"1900","01","01")
        End if
        If strToDt = "" Then
            strToDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"2500","12","31")
        End if 
        
        
        If CompareDateByFormat(strFromDt, strToDt, frm1.txtFrom_dt.Alt, frm1.txtTo_dt.Alt,"970025",parent.gDateFormat,parent.gComDateType,True) = False Then
			Exit Function
        End IF 
        
    End if 

    ggoSpread.source = frm1.vspdData
    ggoSpread.ClearSpreadData	

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If
       
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
    Dim IntRetCD ,lRow
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgbox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
 	For lRow = 1 To  frm1.vspdData.MaxRows
		With Frm1		
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				.vspdData.Col = C_NAME
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("800048","X","X","X")
					.vspdData.Col = C_EMP_NO
					.vspddata.focus
       	            exit function
				 end if 
            end if
		end with         
	Next   

    Call MakeKeyStream("X")
    If DbSave = False Then
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
	
	With Frm1.VspdData
           .Col  = C_NAME
           .Row  = .ActiveRow
           .Text = ""
           
           .Col  = C_EMP_NO
           .Row  = .ActiveRow
           .Text = ""
           
           .Col  = C_HFA080T_PAY_YYMM_DT
           .Row  = .ActiveRow
           .Text = ""
   
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
    
	Call SetToolbar("110011110011111")	
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal PvRowCnt) 
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = false
	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowCount()
		if imRow = "" then
			Exit function
		end if
	end if

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
 
    if Err.number = 0 then
		FncInsertRow = true
	end if
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

	 if LayerShowHide(1) = false then
	    Exit Function
	end if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
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
    
    if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & .txtEmp_no.value & parent.gColSep
                                                  strVal = strVal & .txtFrom_dt.Text & parent.gColSep
                                                  strVal = strVal & .txtTo_dt.Text & parent.gColSep
                    .vspdData.Col = C_NAME	                    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_YYMM_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_TEXT	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_TOT_AMT       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_INCOME_TAX_AMT	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_RES_TAX_AMT	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_MED_INSUR_AMT     : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_EMP_INSUR_AMT     : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                                                      strVal = strVal & .txtEmp_no.value & parent.gColSep
                                                      strVal = strVal & .txtFrom_dt.Text & parent.gColSep
                                                      strVal = strVal & .txtTo_dt.Text & parent.gColSep
                    .vspdData.Col = C_NAME	                    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_YYMM_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_TEXT	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_TOT_AMT       : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_INCOME_TAX_AMT	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_RES_TAX_AMT	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_MED_INSUR_AMT     : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_EMP_INSUR_AMT     : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HFA080T_PAY_YYMM_DT	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.source = frm1.vspdData
    ggoSpread.ClearSpreadData	
    
    Call InitVariables															'⊙: Initializes local global variables
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DBQuery = false Then
        Call RestoreToolBar()
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
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================

Function SetCode(Byval arrRet, Byval iWhere)
Dim strRel
	With frm1

		Select Case iWhere
		    Case C_FAMILY_NAME_POP
		        .vspdData.Col = C_FAMILY_NAME
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_FAMILY_REL
		    	.vspdData.text = arrRet(1)
                Call CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0023", "''", "S") & " and minor_cd =  " & FilterVar(arrRet(1), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		    	.vspdData.Col = C_FAMILY_REL_NM
		    	.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
        End Select

	End With

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
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  

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
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
    Dim arrACompNo
    Dim strACompNo
    Dim arrBuff
    Dim sumAComp
    Dim intTenMinusSum
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
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMP_NO
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.value = ""
            Else
	            IntRetCd = FuncGetEmpInf2(iDx,lgUsrIntCd,strName,strDept_nm, strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	            if  IntRetCd < 0 then
	                if  IntRetCd = -1 then
                		Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                    else
                        Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                    end if
  	                Frm1.vspdData.Col = C_NAME
                    Frm1.vspdData.value = ""
                    Frm1.vspdData.Col = C_EMP_NO
                    Frm1.vspddata.focus
                Else
		       	    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.value = strName
		       	    Frm1.vspdData.Col = C_EMP_NO
		       	     Frm1.vspddata.focus
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
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	
    Dim strName
    Dim strRel
    Dim strRelCd

	With frm1.vspdData
    
	End With
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
	Call SetPopupMenuItemInf("1101111111")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 컬럼을 클릭할 경우 발생 
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

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_NAME_POP
                    Call OpenEmptName("1")
	    Case C_EMP_NO_POP
                    Call OpenEmptName("1")
    End Select 
    
End Sub

'=======================================
'   Event Name :txtFrom_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtFrom_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtFrom_dt.Action = 7
        frm1.txtFrom_dt.focus
    End If
End Sub

'=======================================
'   Event Name :txtTo_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtTo_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtTo_dt.Action = 7
        frm1.txtTo_dt.focus
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

'=======================================================================================================
'   Event Name : txtFrom_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtFrom_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub

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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><IMG src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>현물지급급여등록</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><IMG src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
			    	    		                <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>지급기간</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h9106ma1_fpDateTime_txtFrom_dt.js'></script>&nbsp;~&nbsp;
								                     <script language =javascript src='./js/h9106ma1_fpDateTime1_txtTo_dt.js'></script></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h9106ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

