<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h4010ma1
*  4. Program Name         : h4010ma1
*  5. Program Desc         : 근태관리/기간별근태조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/
*  8. Modified date(Last)  : 2003/06/11
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h4010mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const C_SHEETMAXROWS1   = 21                                           '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lsInternal_cd
Dim lgStrPrevKey1
Dim topleftOK

Dim C_DILIG_CD1                                                 'Column Dimant for Spread Sheet 
Dim C_DILIG_NM1  
Dim C_DILIG_CNT1  
Dim C_DILIG_HR1     
Dim C_DILIG_MN1    

Dim C_DEPT_CD   
Dim C_DEPT_NM         
Dim C_HAA010T_EMP_NO    
Dim C_HAA010T_NAME       
Dim C_HCA060T_DILIG_DT   
Dim C_HCA010T_DILIG_CD                                             'Column Dimant for Spread Sheet 
Dim C_HCA010T_DILIG_NM 
Dim C_HCA060T_DILIG_HH    
Dim C_HCA060T_DILIG_MM 

'========================================================================================================
' Name : initSpreadPosVariables(spd)	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables(spd) 
	if spd="A" or spd="ALL" then
			C_DEPT_CD             = 1
			C_DEPT_NM             = 2
			C_HAA010T_EMP_NO      = 3
			C_HAA010T_NAME        = 4
			C_HCA060T_DILIG_DT    = 5
			C_HCA010T_DILIG_CD    = 6 
			C_HCA010T_DILIG_NM    = 7
			C_HCA060T_DILIG_HH    = 8
			C_HCA060T_DILIG_MM    = 9
	end if
	if spd="B" or spd="ALL" then
			C_DILIG_CD1     = 1 
			C_DILIG_NM1     = 2
			C_DILIG_CNT1    = 3
			C_DILIG_HR1     = 4
			C_DILIG_MN1     = 5

	end if
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgStrPrevKey       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey1	   = ""                                     '⊙: initializes Previous Key Index
    lgSortKey          = 1                                      '⊙: initializes sort direction
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
	
	frm1.txtDilig_dt_1_dt.Focus			'년월 default value setting
	frm1.txtDilig_dt_1_dt.Year = strYear 
	frm1.txtDilig_dt_1_dt.Month = strMonth 
	frm1.txtDilig_dt_1_dt.Day = "01"
	frm1.txtDilig_dt_2_dt.Year = strYear 
	frm1.txtDilig_dt_2_dt.Month = strMonth 
	frm1.txtDilig_dt_2_dt.Day = strDay 
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
   
    Dim strDiligDt1, strDiligDt2
    Dim rdoWk_yesno
    Dim rdoDilig_type

    If frm1.rdoWk_yesno(0).checked Then
        rdoWk_yesno = "1"
    ElseIf frm1.rdoWk_yesno(1).checked Then
        rdoWk_yesno = "2"
    End If
    
    If frm1.rdoDilig_type(0).checked Then
        rdoDilig_type =   "1"
    ElseIf frm1.rdoDilig_type(1).checked Then
        rdoDilig_type = "2"
    Else
        rdoDilig_type = "3"
    End If
                           
    If Trim(frm1.txtDilig_dt_1_dt.Text) = "" or IsNull(Trim(frm1.txtDilig_dt_1_dt.Text)) Then
        strDiligDt1 = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")
    Else 
        strDiligDt1 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDilig_dt_1_dt.Year, Right("0" & frm1.txtDilig_dt_1_dt.Month,2), Right("0" & frm1.txtDilig_dt_1_dt.Day,2))
    End If

    If Trim(frm1.txtDilig_dt_2_dt.Text) = "" or IsNull(Trim(frm1.txtDilig_dt_2_dt.Text)) Then
        strDiligDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, "3000", "12", "31")
    Else 
        strDiligDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDilig_dt_2_dt.Year, Right("0" & frm1.txtDilig_dt_2_dt.Month,2), Right("0" & frm1.txtDilig_dt_2_dt.Day,2))
    End If

    lgKeyStream       = strDiligDt1 & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream       = lgKeyStream & strDiligDt2                    & parent.gColSep
    lgKeyStream       = lgKeyStream & rdoWk_yesno                    & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtDept_cd.Value          & parent.gColSep
    if  Frm1.txtDept_cd.Value = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
        lgKeyStream = lgKeyStream & Frm1.txtInternal_cd.Value & parent.gColSep
    end if
    lgKeyStream       = lgKeyStream & Frm1.txtDilig_cd.Value         & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtName.Value             & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value           & parent.gColSep
    lgKeyStream       = lgKeyStream & rdoDilig_type                  & parent.gColSep
    
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
Sub InitSpreadSheet(strSPD)

	Call initSpreadPosVariables(strSPD)  

	if (strSPD = "A" or strSPD = "ALL") then	
		With Frm1.vspdData
		    ggoSpread.Source = Frm1.vspdData

		    ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false
		   .MaxCols = C_HCA060T_DILIG_MM + 1                                                 '☜:☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                              '☜:☜: Hide maxcols
		   .ColHidden = True                                                            '☜:☜:

		   .MaxRows = 0
    
			Call GetSpreadColumnPos("A")  
	
		   Call AppendNumberPlace("6","3","0")

		    ggoSpread.SSSetEdit     C_DEPT_CD,             "code",            5,,, 13,2
		    ggoSpread.SSSetEdit     C_DEPT_NM,             "부서",        20,,, 40,2
		    ggoSpread.SSSetEdit     C_HAA010T_EMP_NO,      "사번",        13,,, 13,2
		    ggoSpread.SSSetEdit     C_HAA010T_NAME,        "성명",        15,,, 30,2
		    ggoSpread.SSSetDate     C_HCA060T_DILIG_DT,    "일자",        10,2, parent.gDateFormat   'Lock->Unlock/ Date
		    ggoSpread.SSSetEdit     C_HCA010T_DILIG_CD    , "code" ,          5,,,2,2
		    ggoSpread.SSSetEdit     C_HCA010T_DILIG_NM    , "근태코드" ,  15,,,20,2
		    ggoSpread.SSSetFloat    C_HCA060T_DILIG_HH    , "시간" ,      7, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat    C_HCA060T_DILIG_MM    , "분"   ,      7, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		   
		   Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)
		   Call ggoSpread.SSSetColHidden(C_HCA010T_DILIG_CD,C_HCA010T_DILIG_CD,True)
	
		   .ReDraw = true
	
		   lgActiveSpd = "M"
		   Call SetSpreadLock 
    
		End With
    End if
    
   	if (strSPD = "B" or strSPD = "ALL") then

		With Frm1.vspdData1
		    ggoSpread.Source = frm1.vspdData1
		    ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false
		   .MaxCols   = C_DILIG_MN1 + 1                                                      ' ☜:☜: Add 1 to Maxcols
		   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
		   .ColHidden = True                                                            ' ☜:☜:
		   
		   .MaxRows = 0	
			Call GetSpreadColumnPos("B")  
		   Call AppendNumberPlace("6","3","0")

		   ggoSpread.SSSetEdit   C_DILIG_CD1    , "code" ,      5,,,2,2
		   ggoSpread.SSSetEdit   C_DILIG_NM1    , "근태코드" ,  15,,,20,2
		   ggoSpread.SSSetFloat  C_DILIG_CNT1   , "횟수" ,      8, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetFloat  C_DILIG_HR1    , "시간" ,      8, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		   ggoSpread.SSSetFloat  C_DILIG_MN1    , "분"   ,      7, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		   Call ggoSpread.SSSetColHidden(C_DILIG_CD1,C_DILIG_CD1,True)

		   .ReDraw = true  
	
			lgActiveSpd = "S"
			Call SetSpreadLock 
	    
		End With
    End if
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
            
			C_DEPT_CD             = iCurColumnPos(1)
			C_DEPT_NM             = iCurColumnPos(2)
			C_HAA010T_EMP_NO      = iCurColumnPos(3)
			C_HAA010T_NAME        = iCurColumnPos(4)
			C_HCA060T_DILIG_DT    = iCurColumnPos(5)
			C_HCA010T_DILIG_CD    = iCurColumnPos(6)
			C_HCA010T_DILIG_NM    = iCurColumnPos(7)
			C_HCA060T_DILIG_HH    = iCurColumnPos(8)
			C_HCA060T_DILIG_MM    = iCurColumnPos(9)
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DILIG_CD1     = iCurColumnPos(1) 
			C_DILIG_NM1     = iCurColumnPos(2)
			C_DILIG_CNT1    = iCurColumnPos(3)
			C_DILIG_HR1     = iCurColumnPos(4)
			C_DILIG_MN1     = iCurColumnPos(5)
		
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
        Case  "S"
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
    End Select  
             
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With Frm1
             ggoSpread.Source = .vspdData
             .vspdData.ReDraw = False
                ggoSpread.SSSetProtected   C_DEPT_CD , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_DEPT_NM , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HAA010T_EMP_NO , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HAA010T_NAME , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HCA060T_DILIG_DT , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HCA010T_DILIG_CD , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HCA010T_DILIG_NM , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HCA060T_DILIG_HH , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_HCA060T_DILIG_MM , pvStartRow, pvEndRow
            .vspdData.ReDraw = True
    End With

    With Frm1
             ggoSpread.Source = .vspdData1
             .vspdData1.ReDraw = False
                ggoSpread.SSSetProtected   C_DILIG_CD1 , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_DILIG_NM1 , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_DILIG_CNT1 , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_DILIG_HR1 , pvStartRow, pvEndRow
                ggoSpread.SSSetProtected   C_DILIG_MN1 , pvStartRow, pvEndRow
            .vspdData1.ReDraw = True
    End With
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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 ' go to 
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
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    
    Call ggoOper.FormatDate(frm1.txtDilig_dt_1_dt, parent.gDateFormat, 1)
    Call ggoOper.FormatDate(frm1.txtDilig_dt_2_dt, parent.gDateFormat, 1)
        
    Call InitSpreadSheet("ALL")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    
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
    Dim strDiligDt1
    Dim strDiligDt2
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")		
	Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If Trim(frm1.txtDilig_dt_1_dt.Text) = "" or IsNull(Trim(frm1.txtDilig_dt_1_dt.Text)) Then
        strDiligDt1 = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")
    Else 
        strDiligDt1 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDilig_dt_1_dt.Year, Right("0" & frm1.txtDilig_dt_1_dt.Month,2), Right("0" & frm1.txtDilig_dt_1_dt.Day,2))
    End If

    If Trim(frm1.txtDilig_dt_2_dt.Text) = "" or IsNull(Trim(frm1.txtDilig_dt_2_dt.Text)) Then
        strDiligDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, "3000", "12", "31")
    Else 
        strDiligDt2 = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtDilig_dt_2_dt.Year, Right("0" & frm1.txtDilig_dt_2_dt.Month,2), Right("0" & frm1.txtDilig_dt_2_dt.Day,2))
    End If
   
    If CompareDateByFormat(strDiligDt1,strDiligDt2,frm1.txtDilig_dt_1_dt.Alt,frm1.txtDilig_dt_2_dt.Alt,"970023",parent.gDateFormat,parent.gComDateType,True) = False Then
       frm1.txtDilig_dt_1_dt.focus
       Set gActiveElement = document.activeElement
       Exit Function
    End if  
    
    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if

    If txtDept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtDilig_cd_Onchange() Then        'enter key 로 조회시 근태코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    lgCurrentSpd = "M"
    Call MakeKeyStream("X")

	Call DisableToolBar(parent.TBC_QUERY)
	topleftOK = false    
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
       
    FncQuery = True																'☜: Processing is OK
   
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

    ggoSpread.Source = Frm1.vspdData1

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = Frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
	Call DisableToolBar(parent.TBC_SAVE)
	If DbSAVE = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
    FncSave = True                                                               '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False    
     If Frm1.vspdData1.MaxRows < 1 Then
        Exit Function
     End If
    
     With Frm1
          If .vspdData1.ActiveRow > 0 Then
             .vspdData1.ReDraw = False
		
              ggoSpread.Source = .vspdData1	
              ggoSpread.CopyRow
              SetSpreadColor   .vspdData1.ActiveRow, .ActiveRow
              .vspdData1.Col  = 1
              .vspdData1.Text = ""
             .vspdData1.ReDraw = True
             .vspdData1.Focus
         End If
    End With

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 

    ggoSpread.Source = Frm1.vspdData1	
    ggoSpread.EditUndo  

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If

     With Frm1
              .vspdData1.ReDraw = False
              .vspdData1.Focus
               ggoSpread.Source = .vspdData1
               ggoSpread.InsertRow ,imRow
               SetSpreadColor .vspdData1.ActiveRow, .vspdData.ActiveRow + imRow - 1
              .vspdData1.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData1.MaxRows < 1 then
       Exit function
    End if	

    With Frm1.vspdData1 
              .Focus
              ggoSpread.Source = frm1.vspdData1 
              lDelRows = ggoSpread.DeleteRow
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
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
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
	ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

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
   
    select case gActiveSpdSheet.id
		case "vaSpread"
			Call InitSpreadSheet("A")      
		case "vaSpread1"
			Call InitSpreadSheet("B")      		
	end select      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
	If lgCurrentSpd = "M" Then
       strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
    Else
       strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
       strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData1.MaxRows         '☜: Max fetched data
    End If   
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
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
		
  	With Frm1
		.txtMode.value      = parent.UID_M0002                                            '☜: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                     strVal = strVal & "C" & parent.gColSep                    '0
                                                     strVal = strVal & lRow & parent.gColSep                    '1
                    .vspdData1.Col = C_MinorCd     : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep   '3
                    .vspdData1.Col = C_MinorNm     : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep   '4
                    .vspdData1.Col = C_MinorTypeCd : strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U" & parent.gColSep
                                                     strVal = strVal & lRow & parent.gColSep
                    .vspdData1.Col = C_MinorCd     : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                    .vspdData1.Col = C_MinorNm     : strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                    .vspdData1.Col = C_MinorTypeCd : strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                     strDel = strDel & "D" & parent.gColSep
                                                     strDel = strDel & lRow & parent.gColSep
                    .vspdData1.Col = C_MinorCd     : strDel = strDel & Trim(.vspdData1.Text) & parent.gRowSep
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
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '☜: Processing is OK
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strGrid2Data

 	If lgCurrentSpd = "M" Then
		If frm1.vspdData1.MaxRows < 1 then
	        lgCurrentSpd       = "S"
            Call InitData()
	        Call MakeKeyStream(1)

			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Function
			End If																'☜: Query db data
        End if
    Else
       Call InitData()
	End If
	
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData1
    Frm1.vspdData1.MaxRows = 0
    lgCurrentSpd = "S"

	Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
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
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
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
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
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
	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtDilig_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtDilig_nm.value		' Name Cindition
	        arrParam(4) = ""							' Where Condition
	        arrParam(5) = "근태코드"			    ' TextBox 명칭 
	
            arrField(0) = "dilig_cd"					' Field명(0)
            arrField(1) = "dilig_nm"				    ' Field명(1)
    
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDilig_cd.focus
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
		        .txtDilig_cd.value = arrRet(0)
		        .txtDilig_nm.value = arrRet(1)
		        .txtDilig_cd.focus
        End Select
	End With
End Sub

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
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	End If
    arrParam(1) = frm1.txtDilig_dt_2_dt.Text
	arrParam(2) = lgUsrIntCd                           ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
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
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               .txtInternal_cd.value = arrRet(2)
               .txtDept_cd.focus
        End Select
	End With
End Function       		

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
    lgActiveSpd      = "M"
End Sub
'========================================================================================================
'   Event Name : vspdData1_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_OnFocus()
    lgActiveSpd      = "S"
End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")    

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
'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================

Sub vspdData1_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SP1C"   
    Set gActiveSpdSheet = frm1.vspdData1
    
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData1.Row = Row   
	Call SetPopupMenuItemInf("0000111111")    
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
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
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
'   Event Name : vspdData1_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData1.MaxRows = 0 then
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
'   Event Name : vspdDat1a_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData1_GotFocus()

    ggoSpread.Source = Frm1.vspdData1
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
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
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
'   Event Name : vspdData1_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
     End If
End Sub    


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub
'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
             
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
          Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)

    Dim iDx
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col
             
    If Frm1.vspdData1.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData1.text) < UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
          Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
End Sub
'========================================================================================================
'   Event Name : vspdData1_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData1_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)

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
			lgCurrentSpd = "M"	
			Call MakeKeyStream("X")				
			Call DisableToolBar(Parent.TBC_QUERY)
			topleftOK = true			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub


'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			lgCurrentSpd = "S"
			Call MakeKeyStream("X")				
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
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
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtDept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtDept_cd.value = "" Then
		frm1.txtDept_nm.value = ""
		frm1.txtInternal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,UNIConvDate(frm1.txtDilig_dt_2_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        If  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '부서코드정보에 등록되지 않은 코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
		    frm1.txtDept_nm.value = ""
		    frm1.txtInternal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtDept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtDept_cd_Onchange = true
            Exit Function      
        Else
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtInternal_cd.value = lsInternal_cd
        End if
        
    End if  
    
End Function


'========================================================================================================
'   Event Name : txtDilig_cd_change
'   Event Desc :
'========================================================================================================
Function txtDilig_cd_Onchange()
    Dim IntRetCd
    
    If frm1.txtDilig_cd.value = "" Then
		frm1.txtDilig_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800099","X","X","X")	'근태코드정보에 등록되지 않은 코드입니다.
		    frm1.txtDilig_nm.value = ""
            frm1.txtDilig_cd.focus
            Set gActiveElement = document.ActiveElement 
            
            txtDilig_cd_Onchange = true
            Exit Function      
        Else
            Call CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			frm1.txtDilig_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End Function
'=======================================
'   Event Name : txtDilig_dt_1_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtDilig_dt_1_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDilig_dt_1_dt.Action = 7
        frm1.txtDilig_dt_1_dt.focus
    End If
End Sub

'=======================================
'   Event Name : txtDilig_dt_2_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtDilig_dt_2_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtDilig_dt_2_dt.Action = 7
        frm1.txtDilig_dt_2_dt.focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtDilig_dt_1_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDilig_dt_1_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

Sub txtDilig_dt_2_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기간별근태조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
              					<TD CLASS="TD5" NOWRAP>근태기간</TD>
	                   			<TD CLASS="TD6"><script language =javascript src='./js/h4010ma1_fpDateTime1_txtDilig_dt_1_dt.js'></script>&nbsp;~&nbsp;
	                   			                <script language =javascript src='./js/h4010ma1_fpDateTime2_txtDilig_dt_2_dt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>기간중퇴사자포함여부</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoWk_yesno ID=rdoWk_yesno1  tag="11"><LABEL FOR=rdoWk_yesno1>포함</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoWk_yesno ID=rdoWk_yesno2  tag="11" Checked><LABEL FOR=rdoWk_yesno2>미포함</LABEL></TD>
							</TR>
							<TR>
				    	        <TD CLASS=TD5 NOWRAP>부서코드</TD>              
			                    <TD CLASS=TD6 NOWRAP><INPUT id=txtDept_cd NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                                         <INPUT id=txtDept_nm NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
						                             <INPUT id=txtInternal_cd NAME="txtInternal_cd" ALT="내부코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14XXXU"></TD>
				    	        <TD CLASS=TD5 NOWRAP>근태코드</TD>              
			                    <TD CLASS=TD6 NOWRAP><INPUT id=txtDilig_cd NAME="txtDilig_cd" ALT="근태코드" TYPE="Text" SiZE=3 MAXLENGTH=2  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
			                                         <INPUT id=txtDilig_nm NAME="txtDilig_nm" ALT="근태코드명" TYPE="Text" SiZE=15 MAXLENGTH=20  tag="14XXXU"></TD>
							</TR>
							<TR>
				                <TD CLASS=TD5 NOWRAP>사원</TD>
				     	        <TD CLASS=TD6 NOWRAP><INPUT id=txtEmp_no NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
				     	                             <INPUT id=txtName NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>근태구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoDilig_type ID=rdoDilig_type1  tag="11" Checked><LABEL FOR=rdoDilig_type1>근태</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoDilig_type ID=rdoDilig_type2  tag="11"><LABEL FOR=rdoDilig_type2>잔업</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoDilig_type ID=rdoDilig_type3  tag="11"><LABEL FOR=rdoDilig_type3>전체</LABEL></TD>
						    </TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
	
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				
				<TR >
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
            			    <TR >
            					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><script language =javascript src='./js/h4010ma1_vaSpread_vspdData.js'></script></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
				            		<TABLE WIDTH=380 HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><script language =javascript src='./js/h4010ma1_vaSpread1_vspdData1.js'></script></TD>
					            		</TR>
					            	</TABLE>
					            </TD>
                            </TR>  
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
	
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

