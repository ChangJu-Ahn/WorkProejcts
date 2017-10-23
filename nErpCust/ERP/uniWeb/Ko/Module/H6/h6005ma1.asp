<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h6005ma1
*  4. Program Name         : h6005ma1
*  5. Program Desc         : 급여관리/학자금지원등록 
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
Const BIZ_PGM_ID = "H6005mb1.asp"                                      'Biz Logic ASP 
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

Dim C_PROV_DT         														'Spread Sheet의 Column별 상수 
Dim C_FAMILY_NAME
Dim C_FAMILY_NAME_POP
Dim C_FAMILY_REL
Dim C_FAMILY_REL_NM
Dim C_PROV_AMT
Dim C_PAYROLL_UPDATE_CD
Dim C_PAYROLL_UPDATE 
Dim C_PAYROLL_YYMM_DT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_PROV_DT         = 1															<%'Spread Sheet의 Column별 상수 %>
    C_FAMILY_NAME     = 2
    C_FAMILY_NAME_POP = 3
    C_FAMILY_REL      = 4
    C_FAMILY_REL_NM   = 5
    C_PROV_AMT        = 6
    C_PAYROLL_UPDATE_CD  = 7
    C_PAYROLL_UPDATE  = 8
    C_PAYROLL_YYMM_DT = 9

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
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
    lgKeyStream = Frm1.txtEmp_no.Value & Parent.gColSep                                           'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtName.Value & Parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0114", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = chr(11) & lgF0
    iNameArr = chr(11) & lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_PAYROLL_UPDATE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_PAYROLL_UPDATE         ''''''''DB에서 불러 gread에서 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_PAYROLL_UPDATE_CD
			intIndex = .value
			.col = C_PAYROLL_UPDATE
			.value = intindex
		Next	
	End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM

	Call initSpreadPosVariables()   'sbk 
	
	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	
	
	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
	
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	    .ReDraw = false

        .MaxCols = C_PAYROLL_YYMM_DT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True

        .MaxRows = 0
        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A") 'sbk

        ggoSpread.SSSetDate     C_PROV_DT,          "지급일"  ,      20,2, Parent.gDateFormat   'Lock->Unlock/ Date
        ggoSpread.SSSetEdit     C_FAMILY_NAME,       "가족성명",      21,,, 30,2
        ggoSpread.SSSetButton   C_FAMILY_NAME_POP
		ggoSpread.SSSetEdit     C_FAMILY_REL,       "가족관계",      10,,, 30,2
		ggoSpread.SSSetEdit     C_FAMILY_REL_NM,    "가족관계",     20,,, 30,2    
        ggoSpread.SSSetFloat    C_PROV_AMT,         "지급액" ,       20,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetCombo    C_PAYROLL_UPDATE_CD,   "payrollupdate",  10
		ggoSpread.SSSetCombo    C_PAYROLL_UPDATE,   "급여반영여부",  15
		ggoSpread.SSSetMask	   C_PAYROLL_YYMM_DT,	"반영년월",      17,2, strMaskYM

	    .ReDraw = true

        Call ggoSpread.MakePairsColumn(C_FAMILY_NAME,C_FAMILY_NAME_POP)    'sbk

        Call ggoSpread.SSSetColHidden(C_FAMILY_REL,C_FAMILY_REL,True)
        Call ggoSpread.SSSetColHidden(C_PAYROLL_UPDATE_CD,C_PAYROLL_UPDATE_CD,True)
	
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
    ggoSpread.SpreadLock    C_PROV_DT, -1, C_PROV_DT, -1 
    ggoSpread.SpreadLock    C_FAMILY_NAME, -1, C_FAMILY_NAME, -1 
    ggoSpread.SpreadLock    C_FAMILY_REL, -1, C_FAMILY_REL, -1 
    ggoSpread.SpreadLock    C_FAMILY_NAME_POP, -1, C_FAMILY_NAME_POP, -1 
    ggoSpread.SpreadLock    C_PAYROLL_YYMM_DT, -1, C_PAYROLL_YYMM_DT, -1 
    ggoSpread.SpreadLock	C_FAMILY_REL_NM, -1, C_FAMILY_REL_NM, -1 
    ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetRequired    C_PROV_DT, pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_FAMILY_NAME, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_FAMILY_REL, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_FAMILY_REL_NM, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_PAYROLL_UPDATE_CD, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_PAYROLL_UPDATE, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_PAYROLL_YYMM_DT, pvStartRow, pvEndRow
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
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
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

            C_PROV_DT         = iCurColumnPos(1)
            C_FAMILY_NAME     = iCurColumnPos(2)
            C_FAMILY_NAME_POP = iCurColumnPos(3)
            C_FAMILY_REL      = iCurColumnPos(4)
            C_FAMILY_REL_NM   = iCurColumnPos(5)
            C_PROV_AMT        = iCurColumnPos(6)
            C_PAYROLL_UPDATE_CD = iCurColumnPos(7)
            C_PAYROLL_UPDATE  = iCurColumnPos(8)
            C_PAYROLL_YYMM_DT = iCurColumnPos(9)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
   
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
    
    frm1.txtEmp_no.Focus()
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Call InitSpreadSheet                                                            'Setup the Spread sheet
        Exit Function
    End if
    
    ggoSpread.ClearSpreadData

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreTooBar()
        Exit Function
    End If                                                                   '☜: Query db data

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
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    if  frm1.txtEmp_no.value = "" then
        Frm1.txtEmp_no.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    end if

   	 Dim strEntr_dt
   	 Dim strProv_dt
   	 Dim lRow

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
        
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_FAMILY_REL

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
			            Call DisplayMsgBox("970000","X","가족성명","X")             '☜ : 등록되지 않은 코드입니다.
						Exit Function
					end if
					
                    strEntr_dt = UniConvDateToYYYYMMDD(.txtEntr_dt.Text,Parent.gDateFormat,"")
                    
   	                .vspdData.Col = C_PROV_DT
                    strProv_dt = UniConvDateToYYYYMMDD(.vspdData.Text,Parent.gDateFormat,"")
                    
                    If .vspdData.Text = "" Then
                    Else
                    	If strEntr_dt >= strProv_dt Then                        
                            Call DisplayMsgBox("970022","X","지급일","입사일")	'지급일은 입사일보다 커야합니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_PROV_DT
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                       End if 
                    End if  
            End Select
        Next
	End With

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
	
           .Col  = C_PROV_DT
           .Text = ""
           .Col  = C_FAMILY_NAME
           .Text = ""
           .Col  = C_FAMILY_REL
           .Text = ""
           .Col  = C_FAMILY_REL_NM
           .Text = ""

           .Col  = C_PAYROLL_UPDATE_CD
           .Text = ""
           .Col  = C_PAYROLL_UPDATE
           .Text = ""
           .Col  = C_PAYROLL_YYMM_DT
           .Text = ""
   
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
    Call initData()
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
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
    Call InitComboBox
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
	
	if LayerShowHide(1) = False then
		Exit Function
	end if	
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
    Call InitComboBox  
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
	Dim strRes_no
	Dim strYear,strMonth,strDay

    DbSave = False                                                          
    
    if LayerShowHide(1) = False then
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
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                                                  strVal = strVal & .txtEmp_no.value & Parent.gColSep
                    .vspdData.Col = C_PROV_DT	         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FAMILY_NAME	     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FAMILY_REL         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_PROV_AMT	         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_PAYROLL_UPDATE_CD	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep                    
                    .vspdData.Col = C_PAYROLL_YYMM_DT    : Call lgConvDateAndFormatDate(.vspdData.Text,Parent.gComDateType,strYear,strMonth,strDay)                    
														   strVal = strVal & Trim(strYear & strMonth) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & Parent.gColSep
                                                      strVal = strVal & lRow & Parent.gColSep
                                                      strVal = strVal & .txtEmp_no.value & Parent.gColSep
                    .vspdData.Col = C_FAMILY_REL	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_PROV_AMT	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_PAYROLL_UPDATE_CD  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_PROV_DT	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FAMILY_NAME     : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                                                  strDel = strDel & .txtEmp_no.value & Parent.gColSep
                    .vspdData.Col = C_PROV_DT	    : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FAMILY_REL	: strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_FAMILY_NAME	: strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = Parent.UID_M0002
       .txtUpdtUserId.value  = Parent.gUsrID
       .txtInsrtUserId.value = Parent.gUsrID
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    If DbDelete= False Then
        Exit Function
    End If
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	Dim strVal
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100111100111111")	
	ggoSpread.SSSetProtected   C_PROV_AMT, -1, -1
	ggoSpread.SSSetProtected   C_PAYROLL_UPDATE, -1, -1							
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call InitVariables
    Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreTooBar()
        Exit Function
    End If           															'⊙: Initializes local global variables
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
'	Name : OpenCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
        
	    Case C_FAMILY_NAME_POP

	        arrParam(0) = "가족성명 팝업"			' 팝업 명칭 
	        arrParam(1) = " haa020t a (nolock), b_major c (nolock), b_minor d (nolock) "				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtEmp_no.value               		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = " c.major_cd = " & FilterVar("H0023", "''", "S") & "  and c.major_cd = d.major_cd   and d.minor_cd = a.rel_cd and a.emp_no = " &FilterVar(frm1.txtEmp_no.value, "''", "S")  ' Where Condition
	        arrParam(5) = "사번"			    ' TextBox 명칭 
	
            arrField(0) =  "HH" & parent.gcolsep & " a.emp_no "					' Field명(0)
            arrField(1) =  "ED21" & parent.gcolsep & " a.family_nm "					' Field명(0)
            arrField(2) =  "ED22" & parent.gcolsep & " d.minor_nm "				        ' Field명(1)
    
            arrHeader(0) = "사번"				' Header명(0)
            arrHeader(1) = "가족성명"				' Header명(0)
            arrHeader(2) = "가족관계"			    ' Header명(1)
	    	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
    If arrRet(0) = "" Then
		frm1.vspdData.Col = C_FAMILY_NAME
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

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
		    	.vspdData.text = arrRet(1) 
		    	.vspdData.Col = C_FAMILY_REL_NM
		    	.vspdData.text = arrRet(2)
                Call CommonQueryRs(" minor_cd "," b_minor "," major_cd=" & FilterVar("H0023", "''", "S") & " and minor_nm =  " & FilterVar(arrRet(2), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		    	.vspdData.Col = C_FAMILY_REL
		    	.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
		        .vspdData.Col = C_FAMILY_NAME
		        .vspdData.action =0		    	
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
			frm1.vspdData.actino =0
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
    Dim strVal
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtDept_cd.value = arrRet(2)
			.txtRoll_pstn.value = arrRet(3)
			.txtEntr_dt.Text = arrRet(5)
			.txtPay_grd.value = arrRet(4)


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
                          '☜: Key
            .imgPhoto.src = strVal
            .txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.actino =0
	
		End If
	End With
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim strRel
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_FAMILY_NAME       '가족명 
            iDx = Trim(Frm1.vspdData.Text)
   	        Frm1.vspdData.Col = C_FAMILY_NAME
    
            If Frm1.vspdData.Text = "" Then
  	            Frm1.vspdData.Col = C_FAMILY_REL
                Frm1.vspdData.Text = ""
  	            Frm1.vspdData.Col = C_FAMILY_REL_NM
                Frm1.vspdData.Text = ""
            Else
                strRel = CommonQueryRs(" REL_CD "," HAA020T "," FAMILY_NM =  " & FilterVar(iDx , "''", "S") & " AND EMP_NO=  " & FilterVar(frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If strRel = false then
	        		Call DisplayMsgBox("800201","X","X","X")	'가족사항에 등록되지 않은 가족입니다.
  	                Frm1.vspdData.Col = C_FAMILY_REL
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FAMILY_REL_NM
                    Frm1.vspdData.Text = ""
                    vspdData_Change = true
                Else
		       	    Frm1.vspdData.Col = C_FAMILY_REL
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
                    Call CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0023", "''", "S") & " and minor_cd = '" &Trim(Replace(lgF0,Chr(11),"")) & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		       	    Frm1.vspdData.Col = C_FAMILY_REL_NM
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
                End if 
            End if 
         Case  C_PAYROLL_UPDATE     ' 반영여부 
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_PAYROLL_UPDATE_CD
            Frm1.vspdData.value = iDx
    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

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
		.Row = Row
    
        Select Case Col
            Case C_FAMILY_REL
                .Col = Col
                intIndex = .value
				.Col = C_FAMILY_REL_NM
				.value = intIndex
				
            Case C_PAYROLL_UPDATE
                .Col = Col
                intIndex = .value
				.Col = C_PAYROLL_UPDATE_CD
				.value = intIndex
				
            Case C_FAMILY_REL_NM
                .Col = Col
                intIndex = .value
				.Col = C_FAMILY_REL
				.value = intIndex

   	            .Col = C_FAMILY_REL
                strRelCd = .Text
    	        .Col = C_FAMILY_NAME
                strName = .Text
                If .value = "" Then
                Else
                     strRel = CommonQueryRs(" REL_CD "," HAA020T "," FAMILY_NM =  " & FilterVar(strName , "''", "S") & " AND EMP_NO=  " & FilterVar(frm1.txtEmp_no.value , "''", "S") & " AND REL_CD=  " & FilterVar(strRelCd , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
                    If strRel = false then
	            		Call DisplayMsgBox("800201","X","X","X")	'가족사항에 등록되지 않은 가족입니다.
  	                    .Col = C_FAMILY_REL
                        .Text = ""
  	                    .Col = C_FAMILY_REL_NM
                        .Text = ""
                    Else
		           	    .Col = C_FAMILY_REL
		           	    .text = Trim(Replace(lgF0,Chr(11),""))
                        Call CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0023", "''", "S") & " and minor_cd = '" &Trim(Replace(lgF0,Chr(11),"")) & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		           	    .Col = C_FAMILY_REL_NM
		           	    .text = Trim(Replace(lgF0,Chr(11),""))
                    End if 
				End if
				
		End Select
	End With
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
' Function Name : vspdData_ScriptLeaveCell
' Function Desc :
'======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
    Dim strYear,strMonth,strDay
    Dim TempDate,ChkDate : ChkDate = False

    With frm1.vspdData
		If Col <> NewCol And NewCol > 0 Then
		
			If Col = C_PAYROLL_YYMM_DT Then
				.Row = Row
				.Col = Col
			
				If .Text <> "" Then
					TempDate = lgConvDateAndFormatDate(.Text,Parent.gComDateType,strYear,strMonth,strDay)    
					ChkDate = CheckDateFormat(Trim(TempDate),Parent.gDateFormat)				    				    	
				    If ChkDate = False And Not Trim(.Text) = Parent.gComDateType And IsDate(strYear & Parent.gServerDateType & strMonth  & Parent.gServerDateType & strDay) = False Then					
						Call DisplayMsgBox("140318","X","X","X")	'년월을 올바로 입력하세요.
						.Text = ""
					End If
				End If
			End If
		
		End If
    End With

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
    
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
                Case C_FAMILY_NAME_POP
				    .Col = C_FAMILY_NAME
				    .Row = Row
                    Call OpenCode("", C_FAMILY_NAME_POP, Row)
			End Select
		End If
	End With
    
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

'========================================================================================================
'   Event Name : txtEmp_no_change                                    '<==인사마스터에 있는 사원인지 확인 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strVal
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
		frm1.txtDept_cd.value = ""
		frm1.txtRoll_pstn.value = ""
		frm1.txtEntr_dt.text = ""
		frm1.txtPay_grd.value = ""
		Frm1.imgPhoto.src = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    
	    if  IntRetCd < 0 then
			strVal = "../../../CShared/image/default_picture.jpg"
			Frm1.imgPhoto.src = strVal	    
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
		    frm1.txtName.value = ""
		    frm1.txtDept_cd.value = ""
		    frm1.txtRoll_pstn.value = ""
		    frm1.txtEntr_dt.text = ""
		    frm1.txtPay_grd.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
            frm1.txtDept_cd.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text = UNIDateClientFormat(strEntr_dt)
            
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

'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split(Parent.gDateFormat,Parent.gComDateType)
	
	If Parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & Parent.gComDateType
	Else
		lgStrComDateType = Parent.gComDateType
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

'========================================================================================================
' Function Name : lgConvDateAndFormatDate()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function lgConvDateAndFormatDate(Byval strDate,strDateType,strYear,strMonth,strDay)

	Dim i,ArrType,ArrDate,ArrgType,strTempDate,strgDate,TempYear,TempMonth
	strYear = "" : strMonth = "" : strDay = "" : strgDate = "" : TempYear = "" : TempMonth = ""
	
	If Trim(strDate) = "" Then
		lgConvDateAndFormatDate=""
		Exit Function
	End If

    ArrType = Split(Parent.gDateFormatYYYYMM,strDateType)
    ArrDate = Split(Trim(strDate),strDateType)
    If IsArray(ArrType) And IsArray(ArrDate) Then
		For i=0 To Ubound(ArrType)
			If Instr(UCase(ArrType(i)),"Y") Then
				TempYear = ArrDate(i)
				If Len(Trim(ArrType(i))) <= 2 Then
					strYear = ConvertYYToYYYY(TempYear)
				Else
					strYear = TempYear
				End If
			ElseIf Instr(UCase(ArrType(i)),"M") Then
				TempMonth = ArrDate(i)
				If Len(Trim(ArrType(i))) >= 3 Then
					strMonth = ConvertMMMToMM(TempMonth)
				Else
					strMonth = TempMonth
				End If
			End If
		Next
		strDay = "01"	
	End If

    ArrgType = Split(Parent.gDateFormat,strDateType)
    If IsArray(ArrgType) Then
		ReDim strTempDate(Ubound(ArrgType))
		For i=0 To Ubound(ArrgType)
			If Instr(UCase(ArrgType(i)),"Y") Then
				strTempDate(i) = TempYear
			ElseIf Instr(UCase(ArrgType(i)),"M") Then
				strTempDate(i) = TempMonth
			ElseIf Instr(UCase(ArrgType(i)),"D") Then
				strTempDate(i) = strDay
			End If
			If i < Ubound(ArrgType) Then
				strTempDate(i) = strTempDate(i) & strDateType
			End If
		Next
	End If
	
	For i = 0 To Ubound(strTempDate)
		strgDate = strgDate & strTempDate(i)
	Next
	
	lgConvDateAndFormatDate = strgDate

End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>학자금지원등록</font></td>
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
    	            <TD HEIGHT=20 WIDTH=7%>
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			                <TR HEIGHT=69>
			                    <TD>
                                    <img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=60 HEIGHT=69 HSPACE=10 VSPACE=0 BORDER=1>
			                    </TD>
			                </TR>
			            </TABLE>
    	            </TD>
    	            <TD HEIGHT=20 WIDTH=93%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사번</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_cd" ALT="부서명" TYPE="Text" SiZE=20  tag="14XXXU"></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15  tag="14XXXU"></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>입사일</TD>
			            		<TD CLASS="TD6"><script language =javascript src='./js/h6005ma1_fpDateTime2_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15  tag="14XXXU"></TD>
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
									<script language =javascript src='./js/h6005ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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

