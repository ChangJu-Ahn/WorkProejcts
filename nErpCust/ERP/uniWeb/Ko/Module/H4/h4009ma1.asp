<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 
*  3. Program ID           	: h4009ma1
*  4. Program Name         	: h4009ma1
*  5. Program Desc         	: 근태관리/월근태조회및조정 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/
*  8. Modified date(Last)  	: 2003/06/11
*  9. Modifier (First)     	: mok young bin
* 10. Modifier (Last)      	: Lee SiNa
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
Const BIZ_PGM_ID      = "h4009mb1.asp"						           '☆: Biz Logic ASP Name
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
Dim lsInternal_cd
Dim gDecimal_day
Dim gDecimal_time
Dim lgStrPrevKey1
Dim topleftOK
Dim gSpreadFlg

Dim C_HCA010T_DILIG_CD                                             'Column Dimant for Spread Sheet 
Dim C_HCA010T_DILIG_POP  
Dim C_HCA010T_DILIG_NM    
Dim C_HCA070T_DILIG_CNT    
Dim C_HCA070T_DILIG_HH    
Dim C_HCA070T_DILIG_MM    
Dim C_DAY_TIME            
Dim C_BAS_MARGIR         
Dim C_WK_DAY              
Dim C_ATTEND_DAY          

Dim C_HCA010T_DILIG_CD1                                             'Column Dimant for Spread Sheet 
Dim C_HCA010T_DILIG_POP1  
Dim C_HCA010T_DILIG_NM1    
Dim C_HCA070T_DILIG_CNT1    
Dim C_HCA070T_DILIG_HH1    
Dim C_HCA070T_DILIG_MM1    
Dim C_DAY_TIME1            
Dim C_BAS_MARGIR1         
Dim C_WK_DAY1              
Dim C_ATTEND_DAY1  

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
Sub initSpreadPosVariables(spd) 
	if spd="A" or spd="ALL" then
		C_HCA010T_DILIG_CD     = 1                                                  'Column ant for Spread Sheet 
		C_HCA010T_DILIG_POP    = 2
		C_HCA010T_DILIG_NM     = 3
		C_HCA070T_DILIG_CNT    = 4
		C_HCA070T_DILIG_HH     = 5
		C_HCA070T_DILIG_MM     = 6
		C_DAY_TIME             = 7
		C_BAS_MARGIR           = 8
		C_WK_DAY               = 9
		C_ATTEND_DAY           = 10
	end if
	if spd="B" or spd="ALL" then	
		C_HCA010T_DILIG_CD1     = 1                                                  'Column ant for Spread Sheet 
		C_HCA010T_DILIG_POP1    = 2
		C_HCA010T_DILIG_NM1     = 3
		C_HCA070T_DILIG_CNT1    = 4
		C_HCA070T_DILIG_HH1     = 5
		C_HCA070T_DILIG_MM1     = 6
		C_DAY_TIME1             = 7
		C_BAS_MARGIR1           = 8
		C_WK_DAY1               = 9
		C_ATTEND_DAY1           = 10
	end if	
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
    lgStrPrevKey1      = ""                                      '⊙: initializes Previous Key    
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gSpreadFlg		  = 1
		
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
	
	frm1.txtDilig_month_dt.focus 		
	frm1.txtDilig_month_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtDilig_month_dt.Month = strMonth 
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
    lgKeyStream       = frm1.txtDilig_month_dt.year & Right("0" & frm1.txtDilig_month_dt.month, 2)  & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtName.Value & parent.gColSep
    lgKeyStream       = lgKeyStream & Frm1.txtEmp_No.Value & parent.gColSep
End Sub        
	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
    Dim lRow
    Dim daytimeVal
    
    lgBlnFlgChgValue = false
    
    With Frm1
        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
    
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
	        .vspdData.Col = C_DAY_TIME
	         daytimeVal = .vspdData.Text

	        If daytimeVal = "1" then
                 ggoSpread.SpreadLock C_HCA070T_DILIG_HH, lRow, C_HCA070T_DILIG_HH, lRow
                 ggoSpread.SpreadLock C_HCA070T_DILIG_MM, lRow, C_HCA070T_DILIG_MM, lRow
                 ggoSpread.SSSetRequired  C_HCA070T_DILIG_CNT , lRow, lRow
	        ElseIf daytimeVal = "2" then
                 ggoSpread.SSSetRequired  C_HCA070T_DILIG_HH , lRow, lRow
                 ggoSpread.SSSetRequired  C_HCA070T_DILIG_MM , lRow, lRow
            end if
            ggoSpread.SpreadLock C_HCA010T_DILIG_POP, lRow, C_HCA010T_DILIG_POP, lRow
       Next
        .vspdData.ReDraw = TRUE
    End With    
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(strSPD)

	Call initSpreadPosVariables(strSPD)  
	if (strSPD = "A" or strSPD = "ALL") then	
		
		With frm1.vspdData
		    ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

		   .ReDraw = false
		   .MaxCols   = C_ATTEND_DAY + 1                                                      ' ☜:☜: Add 1 to Maxcols
		   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
		   .ColHidden = True
		   .MaxRows = 0
		   Call GetSpreadColumnPos("A")  
	
			Call AppendNumberPlace("6","3","0")	
		   ggoSpread.SSSetEdit   C_HCA010T_DILIG_CD    , "코드" ,      7,,,2,2
		   ggoSpread.SSSetButton C_HCA010T_DILIG_POP
		   ggoSpread.SSSetEdit   C_HCA010T_DILIG_NM    , "근태" ,      14,,,20,2
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_CNT   , "횟수" ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_HH    , "시간" ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_MM    , "분"   ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z",0,59
		   ggoSpread.SSSetEdit   C_DAY_TIME            , "DAY_TIME" ,      15,,,1,2
		   ggoSpread.SSSetEdit   C_BAS_MARGIR          , "BAS_MARGIR" ,      15,,,1,2
		   ggoSpread.SSSetEdit   C_WK_DAY              , "WK_DAY" ,          15,,,1,2
		   ggoSpread.SSSetEdit   C_ATTEND_DAY          , "ATTEND_DAY" ,      15,,,1,2

			call ggoSpread.MakePairsColumn(C_HCA010T_DILIG_CD,C_HCA010T_DILIG_POP)	
		    Call ggoSpread.SSSetColHidden(C_DAY_TIME,C_DAY_TIME,True)	
		    Call ggoSpread.SSSetColHidden(C_BAS_MARGIR,C_BAS_MARGIR,True)
		    Call ggoSpread.SSSetColHidden(C_WK_DAY,C_WK_DAY,True)
		    Call ggoSpread.SSSetColHidden(C_ATTEND_DAY,C_ATTEND_DAY,True)

 		   .ReDraw = true  
			lgActiveSpd = "A"
		   Call SetSpreadLock 
    
		End With
    End if
    
   	if (strSPD = "B" or strSPD = "ALL") then
    
		With frm1.vspdData1
	
		    ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

		   .ReDraw = false
		   .MaxCols   = C_ATTEND_DAY1 + 1                                                      ' ☜:☜: Add 1 to Maxcols
		   .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
		   .ColHidden = True
		   
		   .MaxRows = 0
		   Call GetSpreadColumnPos("B")  	
		   ggoSpread.SSSetEdit   C_HCA010T_DILIG_CD1    , "코드" ,      7,,,2,2
		   ggoSpread.SSSetButton C_HCA010T_DILIG_POP1
		   ggoSpread.SSSetEdit   C_HCA010T_DILIG_NM1    , "근태" ,      14,,,20,2
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_CNT1   , "횟수" ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_HH1    , "시간" ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		   ggoSpread.SSSetFloat  C_HCA070T_DILIG_MM1    , "분"   ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z",0,59
		   ggoSpread.SSSetEdit   C_DAY_TIME1            , "DAY_TIME" ,      15,,,1,2
		   ggoSpread.SSSetEdit   C_BAS_MARGIR1          , "BAS_MARGIR" ,      15,,,1,2
		   ggoSpread.SSSetEdit   C_WK_DAY1              , "WK_DAY" ,          15,,,1,2
		   ggoSpread.SSSetEdit   C_ATTEND_DAY1          , "ATTEND_DAY" ,      15,,,1,2

		   call ggoSpread.MakePairsColumn(C_HCA010T_DILIG_CD1,C_HCA010T_DILIG_POP1)
		   Call ggoSpread.SSSetColHidden(C_DAY_TIME1,C_DAY_TIME1,True)
		   Call ggoSpread.SSSetColHidden(C_BAS_MARGIR1,C_BAS_MARGIR1,True)
		   Call ggoSpread.SSSetColHidden(C_WK_DAY1,C_WK_DAY1,True)
		   Call ggoSpread.SSSetColHidden(C_ATTEND_DAY1,C_ATTEND_DAY1,True)
		   
 		   .ReDraw = true  
			lgActiveSpd = "B"	
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
            
			C_HCA010T_DILIG_CD     = iCurColumnPos(1)
			C_HCA010T_DILIG_POP    = iCurColumnPos(2)
			C_HCA010T_DILIG_NM     = iCurColumnPos(3)
			C_HCA070T_DILIG_CNT    = iCurColumnPos(4)
			C_HCA070T_DILIG_HH     = iCurColumnPos(5)
			C_HCA070T_DILIG_MM     = iCurColumnPos(6)
			C_DAY_TIME             = iCurColumnPos(7)
			C_BAS_MARGIR           = iCurColumnPos(8)
			C_WK_DAY               = iCurColumnPos(9)
			C_ATTEND_DAY           = iCurColumnPos(10)
   
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_HCA010T_DILIG_CD1     = iCurColumnPos(1)
			C_HCA010T_DILIG_POP1    = iCurColumnPos(2)
			C_HCA010T_DILIG_NM1     = iCurColumnPos(3)
			C_HCA070T_DILIG_CNT1    = iCurColumnPos(4)
			C_HCA070T_DILIG_HH1     = iCurColumnPos(5)
			C_HCA070T_DILIG_MM1     = iCurColumnPos(6)
			C_DAY_TIME1             = iCurColumnPos(7)
			C_BAS_MARGIR1           = iCurColumnPos(8)
			C_WK_DAY1               = iCurColumnPos(9)
			C_ATTEND_DAY1           = iCurColumnPos(10)
    End Select        
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "A"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "A"
			With frm1
				.vspdData.ReDraw = False
				ggoSpread.source = .vspddata
				ggoSpread.SpreadLock      C_HCA010T_DILIG_CD, -1, C_HCA010T_DILIG_CD
				ggoSpread.SpreadLock      C_HCA010T_DILIG_NM, -1, C_HCA010T_DILIG_NM
				ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 				
				.vspdData.ReDraw = True
			End With
        Case  "B"
        
			With frm1
				.vspdData1.ReDraw = False
				ggoSpread.source = .vspddata1
				ggoSpread.SpreadLock      -1,-1,-1
				ggoSpread.SSSetProtected	.vspdData1.MaxCols,-1,-1 				
				.vspdData1.ReDraw = True
			End With
		 End Select 
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetRequired    C_HCA010T_DILIG_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_HCA010T_DILIG_NM , pvStartRow, pvEndRow
    .vspdData.ReDraw = True          
    
    .vspdData1.ReDraw = False
      ggoSpread.SSSetRequired    C_HCA010T_DILIG_CD1 , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_HCA010T_DILIG_NM1 , pvStartRow, pvEndRow
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
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData1.Col = iDx
              Frm1.vspdData1.Row = iRow
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
    
	call get_decimal()
	Call AppendNumberPlace("6", "3", gDecimal_time) 'time
	Call AppendNumberPlace("7", "2", "0")
	Call AppendNumberPlace("8", "3", "0")
	Call AppendNumberPlace("9", "2", gDecimal_day) 'day
	
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatNumber(frm1.txtTot_day,31,0,false)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	
    Call ggoOper.FormatDate(frm1.txtDilig_month_dt, parent.gDateFormat, 2)
    
    Call InitSpreadSheet("ALL")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1110100000001111")												'⊙: Set ToolBar
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
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    ggoSpread.Source = Frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If  
    	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
 
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
 
    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if

    Call MakeKeyStream("X")
    gSpreadFlg = "1"
	frm1.txtPrevNext.value = ""    
	topleftOK = false    
	If DbQuery = False Then
		Call RestoreToolBar()
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
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1111100000001111")							                 '⊙: Set ToolBar
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    Call DisableToolBar(parent.TBC_DELETE)
	If DbDelete = False Then
		Call RestoreToolBar()
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
    Dim intCnt
    Dim intHh
    Dim intMM
    Dim strDayTime
    Dim lRow

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear        
                                                                
    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
	else
		ggoSpread.Source = frm1.vspdData1	

		If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		else
		    IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
		    Exit Function
		End If
    End If
        
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
    if cdbl(frm1.txtTot_day.text) < cdbl(frm1.txtSun_day.text) + cdbl(frm1.txtHol_day.text) then
		
		IntRetCD = DisplayMsgBox("800484","X",frm1.txtAttend_day.alt,"X")                           '⊙: No data changed!!
        Exit Function
    end if
    
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag  

                    .vspdData.Col = C_HCA010T_DILIG_NM
                    If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                        Call DisplayMsgBox("800099","X","X","X")
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if

   	                .vspdData.Col = C_HCA070T_DILIG_CNT
                    intCnt = .vspdData.value
   	                .vspdData.Col = C_HCA070T_DILIG_HH
                    intHh = .vspdData.value
                    .vspdData.Col = C_HCA070T_DILIG_MM
                    intMM = .vspdData.value
   	                .vspdData.Col = C_DAY_TIME
                    strDayTime = .vspdData.value
                    IF (strDayTime="1" or strDayTime="3") and intCnt=0 THEN 
                        Call DisplayMsgBox("970021","X","횟수","X")	     '횟수는 입력필수 항목입니다 
	                    Exit Function                                            '바로 return한다 
                    END IF
                   
                    IF strDayTime="2" and intHh=0 and intMM=0 THEN 
                        Call DisplayMsgBox("970021","X","시간, 분","X")	     '시간는 입력필수 항목입니다 
	                    Exit Function                                            '바로 return한다 
                    END IF
                    	 
            End Select
        Next
	End With

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
    
	With Frm1.VspdData
	    
		If .ActiveRow > 0 Then
			.ReDraw = False
		
			ggoSpread.Source = frm1.vspdData	
			ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, .ActiveRow

           .Col  = C_HCA010T_DILIG_CD
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_HCA010T_DILIG_NM
           .Row  = .ActiveRow
           .Text = ""

			.ReDraw = True
			.focus
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
	Call SetToolbar("1110111100111111")
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
    Call SetToolbar("1110111100111111")
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

  
    Call MakeKeyStream("X")
	frm1.txtPrevNext.value = "P"
	topleftOK = false 
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
	
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														 '⊙: Initializes local global variables

    Call MakeKeyStream("X")
	topleftOK = false 
	frm1.txtPrevNext.value = "N"	
	
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncNext = True                                                               '☜: Processing is OK
	
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
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&gSpreadFlg="       & gSpreadFlg                      '☜: Next key tag
	if gSpreadFlg = "1" then
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
	elseif gSpreadFlg = "2" then
		strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
	end if	   
    
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	strVal = strVal     & "&dayPoint="           & gdecimal_day
	strVal = strVal     & "&timePoint="          & gdecimal_time
    strVal = strVal     & "&txtPrevNext="        & frm1.txtPrevNext.value
	
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
	Dim strVal
	Dim strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) = False then
    		Exit Function 
    End if
		
	Dim SundayValue
	Dim NonWeekdayValue
	
	SundayValue     = frm1.txtSun_day.value
	NonWeekdayValue = frm1.txtNon_week_day.value
	
	If Cint(NonWeekdayValue) > Cint(SundayValue) Then                             '휴일 = 무휴일 + 주휴일  이기 때문에....
		Call DisplayMsgBox("800433","X","X","X")	'무휴일은 일요일보다 클수 없습니다.
	    Call LayerShowHide(0)
	    frm1.txtNon_week_day.focus
	    Exit Function
	Else
	End if
	
	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With


    strVal = ""
    strDel = ""
    lGrpCnt = 1
	
	Dim strYear
    Dim strMonth
	Dim strDilig_month_dt

	strYear = frm1.txtDilig_month_dt.year
    strMonth = frm1.txtDilig_month_dt.month
    
    If len(strMonth) = 1 then
        strMonth = "0" & strMonth
    End if

	strDilig_month_dt = strYear & strMonth

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & strDilig_month_dt & parent.gColSep
                                                  strVal = strVal & Trim(frm1.txtEmp_No.Value) & parent.gColSep
                    .vspdData.Col = C_HCA010T_DILIG_CD  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_CNT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_HH  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_MM      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DAY_TIME              : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BAS_MARGIR            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_DAY                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ATTEND_DAY            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                     strVal = strVal & Trim(frm1.txtAttend_day.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
              Case ggoSpread.UpdateFlag                                      '☜: Update
			  
												  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & strDilig_month_dt & parent.gColSep
                                                  strVal = strVal & Trim(frm1.txtEmp_No.Value) & parent.gColSep
                    .vspdData.Col = C_HCA010T_DILIG_CD  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_CNT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_HH  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HCA070T_DILIG_MM      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DAY_TIME              : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BAS_MARGIR            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_DAY                : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ATTEND_DAY            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                                     strVal = strVal & Trim(frm1.txtAttend_day.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & strDilig_month_dt & parent.gColSep
                                                  strDel = strDel & Trim(frm1.txtEmp_No.Value) & parent.gColSep
                    .vspdData.Col = C_HCA010T_DILIG_CD   : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep	
                    .vspdData.Col = C_HCA070T_DILIG_CNT   : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep	
                    .vspdData.Col = C_HCA070T_DILIG_HH    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep	
                    .vspdData.Col = C_HCA070T_DILIG_MM    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep	
                    .vspdData.Col = C_DAY_TIME            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BAS_MARGIR          : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_DAY              : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ATTEND_DAY          : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                                                   strDel = strDel & Trim(frm1.txtAttend_day.Text) & parent.gRowSep
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
    strVal = strVal     & "&txtKeyStream="    & lgKeyStream                   '☜: Query Key
 
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim lRow
	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	Call SetToolbar("111111111111111")												'⊙: Set ToolBar

    Call InitData()
    Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus

End Function
	
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
           
	    Case C_HCA010T_DILIG_POP

	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""                		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = " dilig_type = " & FilterVar("2", "''", "S") & " "		    ' Where Condition
	        arrParam(5) = "근태코드"			    ' TextBox 명칭 
	
            arrField(0) = "dilig_cd"					' Field명(0)
            arrField(1) = "dilig_nm"				    ' Field명(1)
            arrField(2) = "day_time"				    ' Field명(2)
            arrField(3) = "bas_margir"				    ' Field명(3)
            arrField(4) = "wk_day"				        ' Field명(4)
            arrField(5) = "attend_day"				    ' Field명(5)
            
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
            arrHeader(2) = "DAY_TIME"			        ' Header명(2)
            arrHeader(3) = "BAS_MARGIR"			        ' Header명(3)
            arrHeader(4) = "WK_DAY"			            ' Header명(4)
            arrHeader(5) = "ATTEND_DAY"			        ' Header명(5)
	    	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_HCA010T_DILIG_CD
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere, Row)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================

Function SetCode(arrRet, iWhere, Row)
	With frm1

		Select Case iWhere
		    Case C_HCA010T_DILIG_POP
		        .vspdData.Row = Row
		        .vspdData.Col = C_HCA010T_DILIG_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_HCA010T_DILIG_NM
		    	.vspdData.text = arrRet(1) 
		        .vspdData.Row = Row
		        .vspdData.Col = C_DAY_TIME
		    	.vspdData.text = arrRet(2) 
		    	.vspdData.Col = C_BAS_MARGIR
		    	.vspdData.text = arrRet(3) 
		    	.vspdData.Col = C_WK_DAY
		    	.vspdData.text = arrRet(4) 
		        .vspdData.Col = C_ATTEND_DAY
		    	.vspdData.text = arrRet(5) 
             
                .vspdData.ReDraw = false
                 ggoSpread.Source = .vspdData
		    	     If arrRet(2) = "1" then   'day_time이 1인 경우 "시간"과 "분"을 입력받을수 없게 해준다.
                          ggoSpread.SpreadLock C_HCA070T_DILIG_HH, Row, C_HCA070T_DILIG_HH, Row
                          ggoSpread.SpreadLock C_HCA070T_DILIG_MM, Row, C_HCA070T_DILIG_MM, Row
                          ggoSpread.SSSetRequired  C_HCA070T_DILIG_CNT , Row, Row
	                 else                     'day_time이 1이 아닌 경우 "시간"과 "분"을 입력받을수 있게 해준다 
                          ggoSpread.SpreadUnLock C_HCA070T_DILIG_HH, Row, C_HCA070T_DILIG_HH, Row
                          ggoSpread.SpreadUnLock C_HCA070T_DILIG_MM, Row, C_HCA070T_DILIG_MM, Row
                          ggoSpread.SpreadUnLock C_HCA070T_DILIG_CNT, Row, C_HCA070T_DILIG_CNT, Row
                          ggoSpread.SSSetRequired  C_HCA070T_DILIG_HH , Row, Row
                          ggoSpread.SSSetRequired  C_HCA070T_DILIG_MM , Row, Row
                     end if
                 .vspdData.ReDraw = TRUE
		        .vspdData.Col = C_HCA010T_DILIG_CD
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_HCA010T_DILIG_POP
                Call OpenCode("", C_HCA010T_DILIG_POP, Row)
			End Select
		End If
    
	End With
End Sub
'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		If Row > 0 Then
			Select Case Col
			Case C_HCA010T_DILIG_POP1
                Call OpenCode("", C_HCA010T_DILIG_POP1, Row)
			End Select
		End If
    
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_HCA010T_DILIG_CD
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_HCA010T_DILIG_CD
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_HCA010T_DILIG_NM
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DAY_TIME
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_BAS_MARGIR
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_WK_DAY
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_ATTEND_DAY
                Frm1.vspdData.value = ""
            Else
                IntRetCd = CommonQueryRs(" DILIG_NM, DAY_TIME, BAS_MARGIR, WK_DAY, ATTEND_DAY "," HCA010T "," DILIG_CD =  " & FilterVar(iDx , "''", "S") & " AND DILIG_TYPE = " & FilterVar("2", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If IntRetCd = false then
			        Call DisplayMsgBox("800099","X","X","X")	'해당근태는 근태코드정보에 존재하지 않습니다.
  	                Frm1.vspdData.Col = C_HCA010T_DILIG_NM
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_DAY_TIME
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_BAS_MARGIR
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_WK_DAY
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_ATTEND_DAY
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_HH
                    Frm1.vspdData.value = 0
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_MM
                    Frm1.vspdData.value = 0
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_CNT
                    Frm1.vspdData.value = 0
                Else
		       	    Frm1.vspdData.Col = C_HCA010T_DILIG_NM
		       	    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
		       	    Frm1.vspdData.Col = C_DAY_TIME
		       	    Frm1.vspdData.value = Trim(Replace(lgF1,Chr(11),""))  
		       	    Frm1.vspdData.Col = C_BAS_MARGIR
		       	    Frm1.vspdData.value = Trim(Replace(lgF2,Chr(11),""))
		       	    Frm1.vspdData.Col = C_WK_DAY
		       	    Frm1.vspdData.value = Trim(Replace(lgF3,Chr(11),""))
		       	    Frm1.vspdData.Col = C_ATTEND_DAY
		       	    Frm1.vspdData.value = Trim(Replace(lgF4,Chr(11),""))
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_HH
                    Frm1.vspdData.value = 0
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_MM
                    Frm1.vspdData.value = 0
  	                Frm1.vspdData.Col = C_HCA070T_DILIG_CNT
                    Frm1.vspdData.value = 0
	                With frm1
                        .vspdData.ReDraw = false
                          ggoSpread.Source = .vspdData
	                          If Trim(Replace(lgF1,Chr(11),"")) = "1" then   'day_time이 1인 경우 "시간"과 "분"을 입력받을수 없게 해준다.
                                  ggoSpread.SpreadLock C_HCA070T_DILIG_HH, Row, C_HCA070T_DILIG_HH, Row
                                  ggoSpread.SpreadLock C_HCA070T_DILIG_MM, Row, C_HCA070T_DILIG_MM, Row
                                  ggoSpread.SSSetRequired  C_HCA070T_DILIG_CNT , Row, Row
	                          else                                          'day_time이 1이 아닌 경우 "시간"과 "분"을 입력받을수 있게 해준다 
                                  ggoSpread.SpreadUnLock C_HCA070T_DILIG_HH, Row, C_HCA070T_DILIG_HH, Row
                                  ggoSpread.SpreadUnLock C_HCA070T_DILIG_MM, Row, C_HCA070T_DILIG_MM, Row
                                  ggoSpread.SpreadUnLock C_HCA070T_DILIG_CNT, Row, C_HCA070T_DILIG_CNT, Row
                                  ggoSpread.SSSetRequired  C_HCA070T_DILIG_HH , Row, Row
                                  ggoSpread.SSSetRequired  C_HCA070T_DILIG_MM , Row, Row
                              end if
                         .vspdData.ReDraw = TRUE
	                 End With
		       	    
                End if 
            End if 
    End Select
             
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
	gSpreadFlg = 1
    
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
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111") 	   
    gMouseClickStatus = "SP1C" 
	gSpreadFlg = 2
	
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
'   Event Name : vspdData_DblClick
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
'   Event Name : vspdData1_GotFocus
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

'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtDilig_month_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDilig_month_dt.Action = 7
        frm1.txtDilig_month_dt.focus
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
			
			topleftOK = true
			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
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
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

Sub txtTot_day_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSun_day_Change()
    lgBlnFlgChgValue = True
    
    Dim intWeekHol
    Dim intNonWeek
    Dim intSunday
    intSunday = frm1.txtSun_day.value
    intNonWeek = frm1.txtNon_week_day.value
    frm1.txtWeek_hol_day.value = Cint(intSunday) - Cint(intNonWeek)
    frm1.txtHol_day.focus
End Sub

Sub txtHol_day_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtWeek_hol_day_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtMargir_day_count_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtAttend_day_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtMargir_time_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtWk_day_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtWork_day_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtNon_week_day_Change()
    lgBlnFlgChgValue = True
    
    Dim intWeekHol
    Dim intNonWeek
    Dim intSunday
	intSunday  = frm1.txtSun_day.value
	intNonWeek = frm1.txtNon_week_day.value
	
	If cint(intNonWeek) > Cint(intSunday) Then                             '휴일 = 무휴일 + 주휴일  이기 때문에....
		Call DisplayMsgBox("800433","X","X","X")	'무휴일은 일요일보다 클수 없습니다.
	    frm1.txtNon_week_day.value = ""
	    frm1.txtNon_week_day.focus
	    Exit Sub
	Else
	End if

    frm1.txtWeek_hol_day.value = Cint(intSunday) - Cint(intNonWeek)
    frm1.txtWk_day.focus
End Sub


Sub txtWk_time_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDilig_month_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDilig_month_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub


'=======================================================================================================
'   Event Name : txtDilig_month_dt_Keypress(Key)
'   Event Desc : 총근무일수 / 시간 끝전 처리 
'=======================================================================================================
sub get_decimal()
dim intRetCd
	gDecimal_day  = 0
	gDecimal_time = 0
	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("1", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_day  = Trim(Replace(lgF0,Chr(11),""))
	end if

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("2", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_time = Trim(Replace(lgF0,Chr(11),""))
	end if

end sub


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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>월근태조회및조정</font></td>
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
								<TD CLASS="TD5" NOWRAP>근태월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h4009ma1_txtDilig_month_dt_txtDilig_month_dt.js'></script></TD>
				                			<TD CLASS=TD5 NOWRAP>사원</TD>
				     	        			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
				     	                             		<INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14XXXU"></TD>
						    </TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
			    </TR>
				<TR>
					<TD WIDTH=100% HEIGHT=30% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
            			    <TR >
            					<TD WIDTH="100%" VALIGN=TOP>
            					    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>월근무현황</LEGEND>
	            					<TABLE WIDTH="100%" HEIGHT=100% CELLSPACING=0>
        						        <TR>
        						            <TD CLASS="TD5" NOWRAP>총일수</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtTot_day_txtTot_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>주휴일</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtWeek_hol_day_txtWeek_hol_day.js'></script></TD>
						        	    </TR>
        						        <TR>
              							    <TD CLASS="TD5" NOWRAP>일요일</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtSun_day_txtSun_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>무휴일</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtNon_week_day_txtNon_week_day.js'></script></TD>
						        	    </TR>
        						        <TR>
              							    <TD CLASS="TD5" NOWRAP>휴일</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtHol_day_txtHol_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>차감일수</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtMargir_day_count_txtMargir_day_count.js'></script></TD>
						        	    </TR>
        						        <TR>
              							    <TD CLASS="TD5" NOWRAP>출근일수</TD>
	                   						<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtAttend_day_txtAttend_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>차감시간</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtMargir_time_txtMargir_time.js'></script></TD>
						        	    </TR>
        						        <TR>
					        	        	<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						                	<TD CLASS="TD6"></TD>
					        	        	<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						                	<TD CLASS="TD6"></TD>
						        	    </TR>
        						        <TR>
              							    <TD CLASS="TD5" NOWRAP><HR></TD>
	                   						<TD CLASS="TD6"><HR></TD>
					        	        	<TD CLASS="TD5" NOWRAP><HR></TD>
	                   						<TD CLASS="TD6"><HR></TD>
						        	    </TR>
        						        <TR>
	                   						<TD CLASS="TD5" NOWRAP>총근무일수</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtWk_day_txtWk_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>총근무시간</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtWk_time_txtWk_time.js'></script></TD>
						        	    </TR>
        						        <TR>
              							    <TD CLASS="TD5" NOWRAP>지급일수</TD>
						                	<TD CLASS="TD6"><script language =javascript src='./js/h4009ma1_txtWork_day_txtWork_day.js'></script></TD>
					        	        	<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						                	<TD CLASS="TD6"></TD>
						        	    </TR>						        	  
					            	</TABLE>
					            	</FIELDSET>
		            			</TD>
		            		</tr>
		            	</table>
		            </TD>
		        </TR>
		        <TR>
					<TD WIDTH=100% HEIGHT=70% VALIGN=TOP>
						<FIELDSET CLASS="CLSFLD" ><LEGEND ALIGN=LEFT>근태현황</LEGEND>
						<TABLE WIDTH=100% HEIGHT=250 CELLSPACING=0>
				            <TR>
								<TD CLASS="TD5" NOWRAP>시간외근무</TD>
								<TD CLASS="TD6"></TD>
								<TD CLASS="TD5" NOWRAP>근태</TD>
								<TD CLASS="TD6"></TD>
	                   		</TR>	
				            <TR>
							    <TD HEIGHT="100%" width ="50%" colspan=2><script language =javascript src='./js/h4009ma1_vaSpread_vspdData.js'></script></TD>
							    <TD HEIGHT="100%" width="50%" colspan=2><script language =javascript src='./js/h4009ma1_vaSpread1_vspdData1.js'></script></TD>
                            </TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2x"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
