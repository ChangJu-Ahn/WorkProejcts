<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 
*  3. Program ID           	: h4007ma1
*  4. Program Name         	: h4007ma1
*  5. Program Desc         	: 근태관리/근태일괄등록 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/
*  8. Modified date(Last)  	: 2003/06/11
*  9. Modifier (First)     	: mok young bin
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment             	:
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
Const BIZ_PGM_ID = "h4007mb1.asp"                                    'Biz Logic ASP 
Const BIZ_PGM_ID1 = "h4007mb2.asp"                                      'Biz Logic ASP  
Const C_SHEETMAXROWS    =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          
Dim lsInternal_cd

Dim C_NAME 
Dim C_NAME_POP 
Dim C_EMP_NO  
Dim C_EMP_NO_POP 
Dim C_DEPT_CD 
Dim C_WK_TYPE_CD
Dim C_WK_TYPE 
Dim C_DILIG_DT 
Dim C_DILIG_CD
Dim C_DILIG_NM
Dim C_DILIG_POP 
Dim C_DILIG_CNT
Dim C_DILIG_HH 
Dim C_DILIG_MM 
Dim C_DAY_TIME
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_NAME  =         1
	C_NAME_POP =      2
	C_EMP_NO  =       3
	C_EMP_NO_POP =    4
	C_DEPT_CD =       5
	C_WK_TYPE_CD =    6
	C_WK_TYPE =       7
	C_DILIG_DT =      8
	C_DILIG_CD =      9
	C_DILIG_NM =     10
	C_DILIG_POP =    11
	C_DILIG_CNT =    12
	C_DILIG_HH =     13
	C_DILIG_MM =     14
	C_DAY_TIME =     15
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
	
	frm1.txtDilig_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtDilig_dt.Month = strMonth 
	frm1.txtDilig_dt.Day = strDay
	
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
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtDilig_dt.Text & parent.gColSep                                           'You Must append one character(parent.gColSep)
    if  Frm1.txtDept_cd.Value = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    else
        lgKeyStream = lgKeyStream & Frm1.txtInternal_cd.Value & parent.gColSep
    end if
	lgKeyStream       = lgKeyStream & Frm1.cboWk_type.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtDilig_cd.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtDilig_nm.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtCnt.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtHh.value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtMm.value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_WK_TYPE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_WK_TYPE         ''''''''DB에서 불러 gread에서 

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboWk_type,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서 
End Sub

'========================================================================================================
' Name : InitSpreadComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitSpreadComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_WK_TYPE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_WK_TYPE         ''''''''DB에서 불러 gread에서 
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
		    ' Combo 일경우 ********************
			.Row = intRow
			.Col = C_WK_TYPE_CD
			intIndex = .value
			.col = C_WK_TYPE
			.value = intindex
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

       .MaxCols = C_DAY_TIME + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       .MaxRows = 0
		Call GetSpreadColumnPos("A") 	
	
       Call AppendNumberPlace("6", "2", "0")
	   Call AppendNumberPlace("7", "1", "0")

        ggoSpread.SSSetEdit     C_NAME,       "성명",          20,,, 30,2
        ggoSpread.SSSetButton   C_NAME_POP
        ggoSpread.SSSetEdit     C_EMP_NO,     "사번",          13,,, 13,2
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetEdit     C_DEPT_CD,    "부서",           20,,, 40,2
        ggoSpread.SSSetCombo    C_WK_TYPE_CD, "근무조코드",     10
        ggoSpread.SSSetCombo    C_WK_TYPE,    "근무조구분",     15
        ggoSpread.SSSetDate     C_DILIG_DT,   "근태일",         10,2, parent.gDateFormat
        ggoSpread.SSSetEdit     C_DILIG_CD,   "근태코드",       10,,, 2,2
        ggoSpread.SSSetEdit     C_DILIG_NM,   "근태",           15,,,20,2
        ggoSpread.SSSetButton   C_DILIG_POP
        ggoSpread.SSSetFloat    C_DILIG_CNT   , "횟수" ,        6, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_DILIG_HH    , "근태시간" ,    8, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat    C_DILIG_MM    , "근태분"   ,    8, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
        ggoSpread.SSSetEdit     C_DAY_TIME    , "일.반",        10,,, 2,2
		
	    Call ggoSpread.SSSetColHidden(C_NAME_POP,C_NAME_POP,True)	
	    Call ggoSpread.SSSetColHidden(C_EMP_NO_POP,C_EMP_NO_POP,True)	
	    Call ggoSpread.SSSetColHidden(C_DILIG_POP,C_DILIG_POP,True)	
	    Call ggoSpread.SSSetColHidden(C_WK_TYPE_CD,C_WK_TYPE_CD,True)	
	    Call ggoSpread.SSSetColHidden(C_DAY_TIME,C_DAY_TIME,True)	
	    Call ggoSpread.SSSetColHidden(C_DILIG_CD,C_DILIG_CD,True)	
	    			    	    			    
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
      ggoSpread.SpreadLock		C_NAME_POP,-1,C_NAME_POP
      ggoSpread.SpreadLock		C_EMP_NO,-1,C_EMP_NO
      ggoSpread.SpreadLock		C_EMP_NO_POP,-1,C_EMP_NO_POP
      ggoSpread.SpreadLock		C_DEPT_CD,-1,C_DEPT_CD
      ggoSpread.SpreadLock		C_WK_TYPE_CD,-1,C_WK_TYPE_CD
      ggoSpread.SpreadLock		C_WK_TYPE,-1,C_WK_TYPE      
      ggoSpread.SpreadLock		C_DILIG_DT,-1,C_DILIG_DT
      ggoSpread.SpreadLock		C_DILIG_CD,-1,C_DILIG_CD
      ggoSpread.SpreadLock		C_DILIG_NM,-1,C_DILIG_NM
      ggoSpread.SpreadLock		C_DILIG_POP,-1,C_DILIG_POP      
      ggoSpread.SpreadLock		C_DAY_TIME ,-1,C_DAY_TIME
      ggoSpread.SSSetRequired    C_DILIG_CNT,  -1, C_DILIG_CNT
      ggoSpread.SSSetRequired    C_DILIG_HH,   -1, C_DILIG_HH
      ggoSpread.SSSetRequired    C_DILIG_MM,   -1, C_DILIG_MM
	  ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1       
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
      ggoSpread.SSSetProtected    C_NAME, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_NAME_POP,   pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_EMP_NO,		pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_EMP_NO_POP, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DEPT_CD,    pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_WK_TYPE_CD, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_WK_TYPE,    pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DILIG_DT,   pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DILIG_CD,   pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DILIG_NM,   pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DILIG_POP,  pvStartRow, pvEndRow
      ggoSpread.SSSetProtected    C_DAY_TIME,   pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_CNT,  pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_HH,   pvStartRow, pvEndRow
      ggoSpread.SSSetRequired    C_DILIG_MM,   pvStartRow, pvEndRow

    .vspdData.ReDraw = True
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
			C_NAME  =         iCurColumnPos(1)
			C_NAME_POP =      iCurColumnPos(2)
			C_EMP_NO  =       iCurColumnPos(3)
			C_EMP_NO_POP =    iCurColumnPos(4)
			C_DEPT_CD =       iCurColumnPos(5)
			C_WK_TYPE_CD =    iCurColumnPos(6)
			C_WK_TYPE =       iCurColumnPos(7)
			C_DILIG_DT =      iCurColumnPos(8)
			C_DILIG_CD =      iCurColumnPos(9)
			C_DILIG_NM =      iCurColumnPos(10)
			C_DILIG_POP =     iCurColumnPos(11)
			C_DILIG_CNT =     iCurColumnPos(12)
			C_DILIG_HH =      iCurColumnPos(13)
			C_DILIG_MM =      iCurColumnPos(14)
			C_DAY_TIME =      iCurColumnPos(15)
            
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
	Call AppendNumberPlace("6", "2", "0")
	Call AppendNumberPlace("7", "1", "0")
    Call AppendNumberRange("8","0","59")
    	
 '   Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
 
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100100000101111")										        '버튼 툴바 제어 
    
    frm1.txtDilig_cd.focus 
	Call CookiePage (0)
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

    If txtDept_cd_Onchange() Then        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    If txtDilig_cd_Onchange() Then        'enter key 로 조회시 근태코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
    Call SetSpreadLock                                   '자동입력때 풀어준 부분을 다시 조회할때 Lock시킴 
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True                                                              '☜: Processing is OK
    
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
    Dim strDiligCd
    Dim strDiligDt
    Dim strGridEmpNo
    Dim strGridDiligDt
    Dim strGridDiligCd
    Dim arrSplitEmpNo
    Dim arrSplitCount
    Dim index
    Dim lRow

    Dim  strSQL , strCloseDt , strReturn_value
    
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
    
    strDiligCd = frm1.txtDilig_cd.value
    strDiligDt = UNIConvDate(frm1.txtDilig_dt.Text)
    
   IntRetCD = CommonQueryRs(" COUNT(emp_no), EMP_NO "," HCA060T "," DILIG_DT =  " & FilterVar(strDiligDt , "''", "S") & " AND DILIG_CD =  " & FilterVar(strDiligCd , "''", "S") & " GROUP BY EMP_NO ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
   If IntRetCD = true then 

        arrSplitCount = Split(lgF0,Chr(11))
        arrSplitEmpNo = Split(lgF1,Chr(11))
        For index = 0 To   ubound(arrSplitCount) 
            With Frm1
                For lRow = 1 To .vspdData.MaxRows
                    
                    .vspdData.Row = lRow
                    .vspdData.Col = 0
                    Select Case .vspdData.Text
                        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                            
                            .vspdData.Col = C_EMP_NO
                            strGridEmpNo = .vspdData.Text
                            
                            Select Case frm1.vspdData.Text 
                                   Case ggoSpread.InsertFlag      
                                
                            If arrSplitEmpNo(index) =  strGridEmpNo Then
                                Call DisplayMsgBox("971002","X",strGridEmpNo,"X")	'.....이미 존재합니다.
	                            Exit Function                                    '바로 return한다....자동입력을 멈춘다.
                            Else
                            End if
							End Select 		
                    End Select
                Next
            End With
        Next 
    End if


    '------------------------------------------------------------------------------------------------   
   
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & "  AND pay_gubun = " & FilterVar("Z", "''", "S") & "  AND PAY_TYPE = " & FilterVar("#", "''", "S") & " "
    IntRetCD = CommonQueryRs(" close_type, convert(char(10),close_dt,20), emp_no "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  IntRetCd = False Then
        strReturn_value = "Y"
    Else
	 
		strCloseDt = UniConvDateToYYYYMMDD(Replace(lgF1, Chr(11), ""),parent.gServerDateFormat,"")
		strDiligDt = UniConvDateToYYYYMMDD(frm1.txtDilig_dt.text,parent.gDateFormat,"")
	
        Select Case Replace(lgF0, Chr(11), "")
            Case "1"    '마감형태 : 정상 
                If strCloseDt <= strDiligDt Then
                    strReturn_value = "Y"
                Else
                    strReturn_value = "N"
                End If

            Case "2"    '마감형태 : 마감 
                If strCloseDt < strDiligDt Then
                    strReturn_value = "Y"
                Else
                    strReturn_value = "N"
                End If
                
        End Select
    End If
    If  strReturn_value = "N" Then
        Call DisplayMsgBox("800291","X","X","X")
        Exit Function
    End If

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Call MakeKeyStream("X")
    
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False  
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
    Call initData()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.col = C_Dilig_dt
        .vspdData.Text = frm1.txtDilig_dt.text	    
	    .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    Call InitSpreadSheet()      
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status
    
	If LayerShowHide(1) = False then
    	Exit Function 
    End if
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
	
 	Call RemovedivTextArea     
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
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if

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
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD    	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CNT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_HH	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_MM      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 
               
               Case ggoSpread.UpdateFlag                                      '☜: Update
					strVal = ""               
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_DT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD    	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CNT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_HH	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_MM      : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      '☜: Delete
				    strDel = ""
                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_DT	        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD          : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'삭제시 key만								
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
'	   .txtSpread.value      = strDel & strVal
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
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
        Call RestoreToolBar()
        Exit Function
    End If
    
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
   
	Call SetToolbar("1100101100111111")									
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
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
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
'	Name : OpenCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_NAME_POP
	        arrParam(0) = "인사조회 팝업"			' 팝업 명칭 
	        arrParam(1) = "HAA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = ""							' Where Condition
	        arrParam(5) = "사번"			    ' TextBox 명칭 
	
            arrField(0) = "name"					' Field명(0)
            arrField(1) = "emp_no"				    ' Field명(1)
            arrField(2) = "dept_cd"					' Field명(2)
            arrField(3) = "dept_nm"					' Field명(3)
            arrField(4) = "pay_grd1"				' Field명(4)
            arrField(5) = "pay_grd2"				' Field명(5)
            arrField(6) = "internal_cd"				    ' Field명(6)
    
            arrHeader(0) = "성명"				' Header명(0)
            arrHeader(1) = "사번"			    ' Header명(1)
            arrHeader(2) = "부서코드"			' Header명(2)
            arrHeader(3) = "부서명"				' Header명(3)
            arrHeader(4) = "직위"			    ' Header명(4)
            arrHeader(5) = "급호"			    ' Header명(5)
            arrHeader(6) = "내부코드"			' Header명(6)
            
	    Case C_EMP_NO_POP
	        arrParam(0) = "인사조회 팝업"			' 팝업 명칭 
	        arrParam(1) = "HAA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""	    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = ""							' Where Condition
	        arrParam(5) = "사번"			    ' TextBox 명칭 
	
            arrField(0) = "name"					' Field명(0)
            arrField(1) = "emp_no"				    ' Field명(1)
            arrField(2) = "dept_cd"					' Field명(2)
            arrField(3) = "dept_nm"					' Field명(3)
            arrField(4) = "pay_grd1"				' Field명(4)
            arrField(5) = "pay_grd2"				' Field명(5)
            arrField(6) = "internal_cd"				    ' Field명(6)
    
            arrHeader(0) = "성명"				' Header명(0)
            arrHeader(1) = "사번"			    ' Header명(1)
            arrHeader(2) = "부서코드"			' Header명(2)
            arrHeader(3) = "부서명"				' Header명(3)
            arrHeader(4) = "직위"			    ' Header명(4)
            arrHeader(5) = "급호"			    ' Header명(5)
            arrHeader(6) = "내부코드"			' Header명(6)
            
	    Case C_DILIG_POP

	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""                		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
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
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_NAME_POP
		        .vspdData.Col = C_NAME
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_EMP_NO
		    	.vspdData.text = arrRet(1) 
		        .vspdData.Col = C_DEPT_CD
		    	.vspdData.text = arrRet(2) 
		    Case C_EMP_NO_POP
		        .vspdData.Col = C_NAME
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_EMP_NO
		    	.vspdData.text = arrRet(1) 
		        .vspdData.Col = C_DEPT_CD
		    	.vspdData.text = arrRet(2) 
        End Select

	End With

End Function


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
   
        Select Case Col
            Case C_WK_TYPE
                .Col = Col
                intIndex = .Value
				.Col = C_WK_TYPE_CD
				.Value = intIndex
            Case C_WK_TYPE_CD
                .Col = Col
                intIndex = .Value
				.Col = C_WK_TYPE
				.Value = intIndex
				
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
            arrField(2) = "day_time"				    ' Field명(2)
    
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
            arrHeader(2) = "일수/시간/반차"			' Header명(2)

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
		        .txtDay_time.value = arrRet(2)
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
    	arrParam(1) = frm1.txtDilig_dt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

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
'                        5.5 Tag Event
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_NAME_POP
                    Call OpenCode("", C_NAME_POP, Row)
	    Case C_EMP_NO_POP
                    Call OpenCode("", C_EMP_NO_POP, Row)
	    Case C_DILIG_POP
                    Call OpenCode("", C_DILIG_POP, Row)
    End Select    
End Sub

'======================================================================================================
'	Name : ButtonClicked()
'	Description : h4007mb2.asp 로 가는 Condition........일괄등록...........
'=======================================================================================================

Sub ButtonClicked(Byval ButtonDown)
	Call BtnDisabled(1)
    Dim strKeyStream
    Dim strVal
    Dim IntRetCD
    Dim strEmpNo
    Dim strInternalCd
    Dim strWkType
    Dim strWhere
    
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Call BtnDisabled(0)
			Exit sub
		End If
    End If
    
	ggoSpread.ClearSpreadData      

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Call BtnDisabled(0)
       Exit Sub
    End If
    
	If frm1.txtDay_time.value="1" then                   '일수/시간/반차에 따라 시:분 입력이 결정되므로 메세지 띄어줌...
	    frm1.txtHh.value = "0"                           '일수일경우 시간,분이 없다...(DAY_TIME=1)인 경우 
	    frm1.txtMm.value = "0"
	    If CInt(frm1.txtCnt.value)=0 Then
	        Call DisplayMsgBox("800441","X","X","X")	'해당근태는 근태횟수를 입력해야합니다.
	        frm1.txtCnt.focus
            Set gActiveElement = document.ActiveElement
            Call BtnDisabled(0)
            Exit Sub
        End if
	else
	    If CInt(frm1.txtHh.value)=0 AND CInt(frm1.txtMm.value)=0 Then
	        Call DisplayMsgBox("800432","X","X","X")	'해당근태는 근태시간을 입력해야합니다.
	        frm1.txtHh.focus
            Set gActiveElement = document.ActiveElement
            Call BtnDisabled(0)
            Exit Sub
        End if
	end if
'----------------   
    strInternalCd = frm1.txtInternal_cd.value
    strWkType = frm1.cboWk_type.value
    If strEmpNo = "" then
        strEmpNo = "%"
    End if
   
    If strWkType = "" then
        strWkType = "%"
    End if

    strWhere = "a.emp_no = c.emp_no AND a.emp_no = b.emp_no AND b.retire_dt IS Null AND c.emp_no = d.emp_no AND c.chang_dt = d.chang_dt "
    strWhere = strWhere & " AND a.emp_no LIKE  " & FilterVar(strEmpNo, "''", "S") & " AND c.wk_type LIKE  " & FilterVar(strWkType, "''", "S") & ""
    strWhere = strWhere & " AND a.emp_no not in (SELECT emp_no FROM HCA060T WHERE dilig_dt =  " & FilterVar(UNIConvDate(frm1.txtDilig_dt.Text), "''", "S") & ""
    strWhere = strWhere &               " AND dilig_cd =  " & FilterVar(frm1.txtDilig_cd.Value , "''", "S") & ")"

    If strInternalCd = "" then
        strInternalCd = lgUsrIntCd
        strWhere = strWhere & " AND a.internal_cd LIKE  " & FilterVar(strInternalCd & "%", "''", "S") & ""
        Call CommonQueryRs(" Count(a.emp_no) "," hdf020t a, hca040t c, hca041t d, haa010t b ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Else
        strWhere = strWhere & " AND a.internal_cd =  " & FilterVar(strInternalCd , "''", "S") & ""
        Call CommonQueryRs(" Count(a.emp_no) "," hdf020t a, hca040t c, hca041t d, haa010t b ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    End if

	If Trim(Replace(lgF0,Chr(11),"")) = 0 then
'        Call DisplayMsgBox("800065","X","X","X")	'자동 입력할 사원이 없습니다.
'	    Call BtnDisabled(0)
'	    Exit Sub                                    '바로 return한다....자동입력을 멈춘다.
    End if
'------------------
	frm1.vspdData.MaxRows = 0
    strKeyStream       = frm1.txtDilig_dt.text & parent.gColSep '0
	strKeyStream       = strKeyStream & Frm1.txtInternal_cd.Value & parent.gColSep '1
	strKeyStream       = strKeyStream & Frm1.cboWk_type.Value & parent.gColSep     '2
    strKeyStream       = strKeyStream & Frm1.txtDilig_cd.Value & parent.gColSep    '3
	strKeyStream       = strKeyStream & Frm1.txtDilig_nm.Value & parent.gColSep    '4
	strKeyStream       = strKeyStream & Frm1.txtCnt.Value & parent.gColSep         '5
	strKeyStream       = strKeyStream & Frm1.txtHh.Value & parent.gColSep         '6
	strKeyStream       = strKeyStream & Frm1.txtMm.Value & parent.gColSep         '7
	strKeyStream       = strKeyStream & Frm1.txtDept_cd.Value & parent.gColSep    '8
	strKeyStream       = strKeyStream & Frm1.txtDilig_nm.Value & parent.gColSep    '8

    With Frm1
    	strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                          'mb2 자동입력......						         
        strVal = strVal     & "&txtKeyStream="       & strKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
	Call BtnDisabled(0)
    Call SetToolbar("1100100100101111")	
End Sub

'======================================================================================================
'	Name : DBAutoQueryOk()
'	Description : h4007mb2.asp 이후 Query OK해 줌 
'=======================================================================================================
Sub DBAutoQueryOk()

    Dim lRow
	Dim intIndex
	Dim daytimeVal 
    Dim IntRetCD 
    
	
    With Frm1
        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
        
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
            
			.vspdData.Col = C_WK_TYPE_CD
			intIndex = .vspdData.value
			.vspdData.col = C_WK_TYPE
			.vspdData.value = intindex 
        Next
        .vspdData.ReDraw = TRUE
    ggoSpread.ClearSpreadData "T"            
    End With    
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    


'========================================================================================================
'   Event Name : txtDept_cd_change
'   Event Desc :
'========================================================================================================
Function txtDept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtDept_cd.value = "" Then
		frm1.txtDept_nm.value = ""
		frm1.txtInternal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,UNIConvDate(frm1.txtDilig_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
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
        else
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtInternal_cd.value = lsInternal_cd
        end if

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
		frm1.txtDay_time.value = ""
    Else
        IntRetCd = CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800099","X","X","X")	'근태코드정보에 등록되지 않은 코드입니다.
		    frm1.txtDilig_nm.value = ""
		    frm1.txtDay_time.value = ""
            frm1.txtDilig_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtDilig_cd_Onchange = true
            Exit Function      
        Else
            Call CommonQueryRs(" DILIG_NM,DAY_TIME  "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			frm1.txtDilig_nm.value = Trim(Replace(lgF0,Chr(11),""))
		    frm1.txtDay_time.value = Trim(Replace(lgF1,Chr(11),""))
        End if 
    End if  
    
End Function

'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtDilig_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDilig_dt.Action = 7
        frm1.txtDilig_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDilig_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDilig_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtCnt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtCnt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtHh_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtHh_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtMm_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>근태일괄등록</font></td>
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
			     		 <TD CLASS=TD5 NOWRAP>근태일</TD>       
				    	 <TD CLASS=TD6 ><script language =javascript src='./js/h4007ma1_txtDilig_dt_txtDilig_dt.js'></script></TD>
				    	 <TD CLASS=TD5 NOWRAP>근태코드</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDilig_cd" ALT="근태코드" TYPE="Text" SiZE=3 MAXLENGTH=2  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
			                                  <INPUT NAME="txtDilig_nm" ALT="근태코드명" TYPE="Text" SiZE=15 MAXLENGTH=20  tag="14">
			                                  <INPUT NAME="txtDay_time" ALT="일수/시간/반차" TYPE="hidden" SiZE=3 MAXLENGTH=1  tag="14"></TD>
			           </TR>
		               <TR>
				    	 <TD CLASS=TD5 NOWRAP>부서</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                                  <INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14">
			                                  <INPUT NAME="txtInternal_cd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"></TD>
			             <TD CLASS=TD5 NOWRAP>근무조</TD>
						 <TD CLASS="TD6" NOWRAP><SELECT NAME="cboWk_type" ALT="근무조" CLASS ="cbonormal" TAG="11N"><OPTION VALUE=""></OPTION></SELECT></TD>
					   </TR>
		               <TR>
					     <TD CLASS="TD5" NOWRAP>횟수</TD>
						 <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h4007ma1_txtCnt_txtCnt.js'></script></TD>
					     <TD CLASS="TD5" NOWRAP>시간</TD>
						 <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h4007ma1_txtHh_txtHh.js'></script>&nbsp;:&nbsp;
						                        <script language =javascript src='./js/h4007ma1_txtMm_txtMm.js'></script></TD>
					   </TR>
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h4007ma1_vaSpread_vspdData.js'></script>
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
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: ButtonClicked('1')" flag=1>일괄등록</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
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
