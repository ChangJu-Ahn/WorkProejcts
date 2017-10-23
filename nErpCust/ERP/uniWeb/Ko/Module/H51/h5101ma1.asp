<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h3008ma1
*  4. Program Name         : 표준보수월액 등록 
*  5. Program Desc         : 표준보수월액 조회,등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/01/02
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : chcho
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const BIZ_PGM_ID      = "h5101mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h5101mb2.asp"						           '☆: Biz Logic ASP Name

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

Dim C_GRADE
Dim C_STD_STRT_AMT
Dim C_STD_END_AMT
Dim C_STD_AMT
Dim C_COM_INSUR_AMT
Dim C_INSUR_RATE
Dim C_INSUR_AREA
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_GRADE			= 1										'Column constant for Spread Sheet 
	 C_STD_STRT_AMT		= 2
	 C_STD_END_AMT		= 3
	 C_STD_AMT			= 4
	 C_COM_INSUR_AMT	= 5
	 C_INSUR_RATE		= 6
	 C_INSUR_AREA		= 7  
End Sub
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

	gblnWinEvent        = False
	lgBlnFlawChgFlg     = False
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
    lgKeyStream       = Trim(frm1.txtInsur_type.value) & parent.gColSep       'You Must append one character( parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(frm1.txtInsur_area.value) & parent.gColSep
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
     Call  SetCombo(frm1.txtInsur_type,"1","건강보험")
     Call  SetCombo(frm1.txtInsur_type,"2","국민연금")

     Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("h0038", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
     iCodeArr = lgF0
     iNameArr = lgF1
     Call  SetCombo2(frm1.txtInsur_area,iCodeArr, iNameArr,Chr(11))
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
        .MaxCols = C_INSUR_AREA + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	
        Call  AppendNumberPlace("7","2","2")
	    Call  GetSpreadColumnPos("A")
         
         ggoSpread.SSSetEdit     C_GRADE,         "등급", 12,,,4,2
         ggoSpread.SSSetFloat    C_STD_STRT_AMT,  "시작표준보수월액" ,22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat    C_STD_END_AMT,   "종료표준보수월액" ,22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat    C_STD_AMT,       "표준보수월액"     ,22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat    C_COM_INSUR_AMT, "보험료"           ,22, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat    C_INSUR_RATE,    "보험율"           ,15,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetEdit     C_INSUR_AREA,    "보험지역", 13
         
         Call ggoSpread.SSSetColHidden(C_INSUR_AREA,  C_INSUR_AREA, True)

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
         ggoSpread.SpreadLock    C_GRADE		, -1, C_GRADE
         ggoSpread.SpreadLock    C_COM_INSUR_AMT, -1, C_COM_INSUR_AMT
         ggoSpread.SpreadLock    C_INSUR_RATE	, -1, C_INSUR_RATE
         ggoSpread.SpreadLock    C_INSUR_AREA	, -1, C_INSUR_AREA
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
         ggoSpread.SSSetRequired	C_GRADE			, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_STD_STRT_AMT	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_STD_END_AMT	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_STD_AMT		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_COM_INSUR_AMT	, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_INSUR_RATE	, pvStartRow, pvEndRow
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
            
            C_GRADE			= iCurColumnPos(1)										'Column constant for Spread Sheet 
			C_STD_STRT_AMT	= iCurColumnPos(2)
			C_STD_END_AMT	= iCurColumnPos(3)
			C_STD_AMT		= iCurColumnPos(4)
			C_COM_INSUR_AMT	= iCurColumnPos(5)
			C_INSUR_RATE	= iCurColumnPos(6)
			C_INSUR_AREA	= iCurColumnPos(7)                         
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() event
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call  AppendNumberPlace("6","2","2")
		
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
    
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitData()
    
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    
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
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    ggoSpread.ClearSpreadData
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True																'☜: Processing is OK

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
	
	ggoSpread.ClearSpreadData 									'⊙: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery1 = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery1 = True																'☜: Processing is OK

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
    Dim iRow
    Dim intGrade
    Dim intStd_strt_amt, intStd_strt_amt2
    Dim intStd_end_amt, intStd_end_amt2
    Dim intStd_amt
    Dim intInsur_rate
    Dim txtAmt
	dim strWhere

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

    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case  ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                .Col = C_STD_STRT_AMT
				intStd_strt_amt =  UNICDbl(.text)
				.Col = C_STD_END_AMT
                intStd_end_amt =  UNICDbl(.text)
   	            .Col = C_INSUR_RATE
                intInsur_rate =  UNICDbl(.text)
                .Col = C_STD_AMT
                intStd_amt =  UNICDbl(.text)
                If intStd_strt_amt > intStd_end_amt  Then
                    Call  DisplayMsgBox("800105","X","X","X")                         '표준보수월액은 시작표준보수월액보다 커야 합니다.
  	                    .Col = C_STD_STRT_AMT
  	                    
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                    Exit Function
		        ElseIf  intStd_strt_amt > intStd_amt Then
                    Call  DisplayMsgBox("800108","X","X","X")                         '표준보수월액은 시작표준보수월액보다 커야 합니다.
  	                    .Col = C_STD_STRT_AMT
  	                    
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                    Exit Function
		        ElseIf  intStd_amt > intStd_end_amt  Then
                    Call  DisplayMsgBox("800109","X","X","X")                         '표준보수월액은 종료표준보수월액보다 작아야 합니다.
  	                    .Col = C_STD_AMT
  	                    
  	                    .Action=0
                        Set gActiveElement = document.activeElement
                    Exit Function
                Else
					.Col = C_Grade
					intGrade = 	right("00" & .Text, 3) '20040116

					strWhere = " Grade = ( select max(grade) from hdb010t where"
					strWhere = strWhere & " INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") 
					strWhere = strWhere & " AND INSUR_AREA = " & FilterVar(frm1.txtInsur_area.value, "''", "S") 
					strWhere = strWhere & " AND right(" & FilterVar("00", "''", "S") & "+Grade,3) < " & FilterVar(intGrade, "''", "S") & ")" '20040116
					strWhere = strWhere & " AND INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") 
					strWhere = strWhere & " AND INSUR_AREA = " & FilterVar(frm1.txtInsur_area.value, "''", "S") 
					intRetCd =  CommonQueryRs(" STD_END_AMT ", " HDB010T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					if (intRetCd = true And replace(lgf0,chr(11),"") <> "") then
						intStd_end_amt2 = cdbl(replace(lgf0,chr(11),""))
					elseif  replace(lgf0,chr(11),"")  = "" then
						intStd_end_amt2 = 0
					end if
					
					strWhere = " Grade = (select MIN(grade) from hdb010t where"
					strWhere = strWhere & " INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") 
					strWhere = strWhere & " AND INSUR_AREA = " & FilterVar(frm1.txtInsur_area.value, "''", "S") 
					strWhere = strWhere & " AND right(" & FilterVar("00", "''", "S") & "+Grade,3) > " & FilterVar(intGrade, "''", "S") & ")" '20040116
					strWhere = strWhere & " AND INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") 
					strWhere = strWhere & " AND INSUR_AREA = " & FilterVar(frm1.txtInsur_area.value, "''", "S") 
					intRetCd =  CommonQueryRs(" STD_STRT_AMT ", "HDB010T ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					if (intRetCd = true And replace(lgf0,chr(11),"") <> "") then
						intStd_strt_amt2 = cdbl(replace(lgf0,chr(11),""))
					elseif  replace(lgf0,chr(11),"")  = "" then
						intStd_strt_amt2 = 9999999999999.99
					end if
					
					if not (intStd_strt_amt >= intStd_end_amt2  and intStd_end_amt <= intStd_strt_amt2 ) then
						Call DisplayMsgBox("800497","x","x","x")
						.Action=0
						Exit Function
					end if
					'2003-08-27 by lsn
					'건강보험공단에서 제공하는 표준보수월액표에 오차 생겨, 맞추기 위해 올림부분고침 
                  '   txtAmt = math.floor((math.floor(intStd_amt * (intInsur_rate/2) / 100 , 0) /10),0) * 10 * 2 			
					 txtAmt = math.floor(( intStd_amt * ( intInsur_rate /2/100)) /10, 0) * 10 * 2
					
   	                .Col = C_COM_INSUR_AMT
                    .Text = UNIFormatNumber(txtAmt,ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
					
		        End If
          End Select
        Next
    End With
    
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

	With Frm1.VspdData           
           
           .Row  = .ActiveRow
           .Col  = C_GRADE
           .Text = ""
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
	Dim iRow

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
        For iRow = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
        
			.vspdData.Row = iRow
			.vspdData.Col = C_INSUR_RATE
			.vspdData.text = .txtInsur_rate
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

	if LayerShowHide(1) = false then
	   exit Function
	end if

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

	if LayerShowHide(1) = false then
	exit Function
	end if

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
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
	exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Name : DbQuery1
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery1()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery1 = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
	exit Function
	end if

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
		
	if LayerShowHide(1) = false then
	exit Function
	end if
		
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
                                                          strVal = strVal & Trim(.txtInsur_type.value) & parent.gColSep
                                                          strVal = strVal & Trim(.txtInsur_area.value) & parent.gColSep
                    .vspdData.Col = C_GRADE             : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_STRT_AMT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_END_AMT       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_AMT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COM_INSUR_AMT     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INSUR_RATE        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                                                          strVal = strVal & Trim(.txtInsur_type.value) & parent.gColSep
                                                          strVal = strVal & Trim(.txtInsur_area.value) & parent.gColSep
                    .vspdData.Col = C_GRADE             : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_STRT_AMT	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_END_AMT       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_STD_AMT           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COM_INSUR_AMT     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_INSUR_RATE        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                                                          strDel = strDel & Trim(.txtInsur_type.value) & parent.gColSep
                                                          strDel = strDel & Trim(.txtInsur_area.value) & parent.gColSep
                     .vspdData.Col = C_GRADE 	        : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
		
	if LayerShowHide(1) = false then
	exit Function
	end if
		
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
	Call SetToolbar("1100111100111111")											 '⊙: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus    
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful자동입력 버튼으로 실행되는 쿼리시 
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    Call btnCb_control_OnClick("DbQuery1")
    Call  ggoOper.LockField(Document, "Q")
    
    ggoSpread.ClearSpreadData "T"
    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call MainQuery
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call MainNew	
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
    Dim intStd_strt_amt
    Dim intStd_end_amt
    Dim intStd_amt
    Dim intInsur_rate
    Dim txtAmt    

    Select Case Col
         Case  C_STD_AMT,C_STD_STRT_AMT,C_STD_END_AMT
        	With frm1.vspdData
                .Row = Row
                .Col = C_STD_STRT_AMT
                intStd_strt_amt =  UNICDbl(.Text)
                .Col = C_STD_END_AMT
                intStd_end_amt =  UNICDbl(.Text)
   	            .Col = C_INSUR_RATE
                intInsur_rate =  UNICDbl(.Text)
                .Col = C_STD_AMT
                intStd_amt =  UNICDbl(.Text)
                '2003-08-27 by lsn
				'건강보험공단에서 제공하는 표준보수월액표에 오차 생겨, 맞추기 위해 올림부분고침 
            '    txtAmt = math.floor((math.floor(intStd_amt * (intInsur_rate/2) / 100 , 0) /10),0) * 10 * 2
				txtAmt = math.floor((intStd_amt * ( intInsur_rate /2/100)) /10, 0) * 10 * 2 
				
				
   	            Frm1.vspdData.Col = C_COM_INSUR_AMT
                Frm1.vspdData.Text =  UNIFormatNumber(txtAmt, ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
           End With
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

'==========================================================================================
'   Event Name : btnCb_control_OnClick()
'   Event Desc : 전체조정 
'==========================================================================================

Sub btnCb_control_OnClick(ByVal iWhere)
	Dim intRow
	Dim IntRetCD
	Dim txtAmt
    
    If frm1.vspdData.ActiveRow = 0 Then
        IntRetCD =  DisplayMsgBox("800107",  parent.VB_YES_NO,"X","X")	    '조정할 자료가 없습니다. 자료를 생성하시겠습니까?
        If IntRetCD = vbNo Then
	       	Exit Sub
	    Else
            Call FncQuery1()
	    End If
	Else
        If frm1.txtInsur_rate.value=0 Then
	        Call  DisplayMsgBox("800220","X","X","X")	            '보험율을  입력하십시오.
	        frm1.txtInsur_rate.focus
	        Exit Sub
        End If        

    	With frm1.vspdData
    		For intRow = 1 To .MaxRows			
    			.Row = intRow
    			.Col = 0
   			    If (.Text =  ggoSpread.UpdateFlag) Or (.Text =  ggoSpread.InsertFlag) Then
   			        Call FncQuery1()
   			        Exit Sub
   			    End If
    		Next 
    		For intRow = 1 To .MaxRows			
    			.Row = intRow
    			.Col = C_STD_AMT
                txtAmt = math.floor(( UNICDbl(.Text) * ( UNICDbl(frm1.txtInsur_rate.Text)/2/100)) /10, 0) * 10 * 2    			

    			.Col = C_COM_INSUR_AMT
                .Text =  UNIFormatNumber(txtAmt, ggAmtOfMoney.DecPoint,-2,0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)

    			.Col = C_INSUR_RATE
    			.Text = frm1.txtInsur_rate.Text
    			.Col = C_INSUR_AREA
    			If .Text = frm1.txtInsur_area.Value Then
    			    .Col = 0
   			        .Text =  ggoSpread.UpdateFlag
   			    Else
    			    .Col = 0
   			        .Text =  ggoSpread.InsertFlag
   			    End If
    		Next	
    	End With
    End If
    frm1.txtInsur_area.disabled = True
    frm1.txtInsur_Type.disabled = True
     ggoOper.SetReqAttr frm1.txtInsur_rate, "Q"
End Sub
'=======================================================================================================
'   Event Name : txtInsur_rate_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtInsur_rate_Keypress(Key)
    If Key = 13 Then
       Call MainQuery
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" dir=ltr>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>표준보수월액등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>보험구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="txtInsur_type" ALT="보험구분" CLASS=cboNormal TAG="12"></OPTION></SELECT>
								<TD CLASS=TD5 NOWRAP>보험지역</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="txtInsur_area" ALT="보험지역" CLASS=cboNormal TAG="12"></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>보험율</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5101ma1_txtInsur_rate_txtInsur_rate.js'></script></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h5101ma1_vspdData_vspdData.js'></script>
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
				    <TD><BUTTON NAME="btnCb_control" CLASS="CLSMBTN">전체조정</BUTTON></TD>
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
