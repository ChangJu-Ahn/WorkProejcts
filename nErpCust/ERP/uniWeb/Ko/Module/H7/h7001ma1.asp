<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 상여기준등록 
*  3. Program ID           : H7001ma1
*  4. Program Name         : H7001ma1
*  5. Program Desc         : 상여관리/상여기준등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/16
*  8. Modified date(Last)  : 2001/05/16
*  9. Modifier (First)     : YBI
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
Const BIZ_PGM_ID      = "H7001mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row

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
Dim lgStrPrevKey1
Dim gSpreadFlg
Dim topleftOK

Dim C_DILIG
Dim C_DILIG_POP
Dim C_DILIG_NM
Dim C_DILIG_CNT

Dim C_DILIG_STRT
Dim C_BAR       
Dim C_DILIG_END 
Dim C_MINUS_RATE
Dim C_MINUS_AMT 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_DILIG = 1                                                  'Column constant for Spread Sheet 
        C_DILIG_POP = 2
        C_DILIG_NM = 3
        C_DILIG_CNT = 4
    ElseIf pvSpdNo = "B" Then
        C_DILIG_STRT    = 1
        C_BAR           = 2
        C_DILIG_END     = 3
        C_MINUS_RATE    = 4
        C_MINUS_AMT     = 5
    End If
    
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
    lgStrPrevKey1      = ""                                      '⊙: initializes Previous Key    
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gSpreadFlg		  = 1
			
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
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
Sub MakeKeyStream(pOpt)
    lgKeyStream = "000" & Parent.gColSep       'You Must append one character(Parent.gColSep)
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0089", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtday_calcu, lgF0, lgF1, Chr(11))

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0090", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtcalcu_bas_dd, lgF0, lgF1, Chr(11))

    iCodeArr = "" & vbTab & "-1" & vbTab & "-2" & vbTab & "-3" & vbTab & "-4" & vbTab & "-5" & vbTab & "-6" & vbTab & "-7" & vbTab & "-8" & vbTab & "-9" & vbTab & "-10" & vbTab & "-11" & vbTab & "-12" & vbTab
    iNameArr = "" & vbTab & "-01월" & vbTab & "-02월" & vbTab & "-03월" & vbTab & "-04월" & vbTab & "-05월" & vbTab & "-06월" & vbTab & "-07월" & vbTab & "-08월" & vbTab & "-09월" & vbTab & "-10월" & vbTab & "-11월" & vbTab & "-12월" & vbTab
    Call SetCombo2(frm1.txtCrt_strt_mm, iCodeArr, iNameArr, vbTab)
    iCodeArr = "" & vbTab & "-1" & vbTab & "-2" & vbTab & "-3" & vbTab & "-4" & vbTab & "-5" & vbTab & "-6" & vbTab & "-7" & vbTab & "-8" & vbTab & "-9" & vbTab & "-10" & vbTab & "-11" & vbTab & "-12" & vbTab
    iNameArr = "" & vbTab & "-01월" & vbTab & "-02월" & vbTab & "-03월" & vbTab & "-04월" & vbTab & "-05월" & vbTab & "-06월" & vbTab & "-07월" & vbTab & "-08월" & vbTab & "-09월" & vbTab & "-10월" & vbTab & "-11월" & vbTab & "-12월" & vbTab
    Call SetCombo2(frm1.txtCrt_end_mm, iCodeArr, iNameArr, vbTab)
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
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","3","2")

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

	    With frm1.vspdData
	
            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols   = C_DILIG_CNT + 1                                                 ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                         ' ☜:☜: Hide maxcols
           .ColHidden = True                                                             ' ☜:☜:
           
           .MaxRows = 0
           ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk
               
           ggoSpread.SSSetEdit   C_DILIG,      "근태코드",   10,,,2,2
           ggoSpread.SSSetButton C_DILIG_POP
           ggoSpread.SSSetEdit   C_DILIG_NM,   "근태코드명", 20,,,20
           ggoSpread.SSSetFloat  C_DILIG_CNT,  "횟수",       10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

           Call ggoSpread.MakePairsColumn(C_DILIG,C_DILIG_POP)    'sbk

	       .ReDraw = true
	
           Call SetSpreadLock 
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

 	    With frm1.vspdData1
	
            ggoSpread.Source = frm1.vspdData1
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

            .ReDraw = false

	    	.MaxCols = C_MINUS_AMT + 1
	    	.Col = .MaxCols				'☜: 공통콘트롤 사용 Hidden Column
	    	.ColHidden = True

	    	.MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 'sbk

            ggoSpread.SSSetFloat    C_DILIG_STRT, "근태 시작회수",       15,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
            ggoSpread.SSSetEdit     C_BAR,          "" , 2,2
            ggoSpread.SSSetFloat    C_DILIG_END,    "근태 종료회수",     15,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
            ggoSpread.SSSetFloat    C_MINUS_RATE,   "차감율",   10,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
            ggoSpread.SSSetFloat    C_MINUS_AMT,    "차감금액", 18,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

            Call ggoSpread.MakePairsColumn(C_DILIG_STRT,C_DILIG_END)    'sbk
            
	    	.ReDraw = True
            
            Call SetSpreadLock1 
        End with
    End If
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData

        .ReDraw = False
        ggoSpread.SpreadLock      C_DILIG, -1, C_DILIG, -1
        ggoSpread.SpreadLock      C_DILIG_NM, -1, C_DILIG_NM, -1
        ggoSpread.SpreadLock      C_DILIG_POP, -1, C_DILIG_POP, -1
        ggoSpread.SSSetRequired   C_DILIG_CNT, -1, -1
        ggoSpread.SSSetProtected   .MaxCols   , -1, -1
        .ReDraw = True

    End With

End Sub

Sub SetSpreadLock1()

	With frm1.vspdData1
	
        ggoSpread.Source = frm1.vspdData1

        .ReDraw = False
        ggoSpread.SpreadLock      C_DILIG_STRT, -1, C_DILIG_STRT, -1
        ggoSpread.SpreadLock      C_BAR, -1, C_BAR, -1
        ggoSpread.SpreadLock      C_DILIG_END, -1, C_DILIG_END, -1
        ggoSpread.SSSetProtected  .MaxCols   , -1, -1
        .ReDraw = True

    End With

End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
    
        .ReDraw = False
        ggoSpread.SSSetRequired    C_DILIG, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_DILIG_NM, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired    C_DILIG_CNT, pvStartRow, pvEndRow
        .ReDraw = True
    
    End With

End Sub

Sub SetSpreadColor1(ByVal pvStartRow,ByVal pvEndRow)

	With frm1.vspdData1
	
        ggoSpread.Source = frm1.vspdData1
    
        .ReDraw = False
        ggoSpread.SSSetRequired    C_DILIG_STRT, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected   C_BAR, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired    C_DILIG_END, pvStartRow, pvEndRow
        .ReDraw = True
    
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

            C_DILIG      = iCurColumnPos(1)
            C_DILIG_POP  = iCurColumnPos(2)
            C_DILIG_NM   = iCurColumnPos(3)
            C_DILIG_CNT  = iCurColumnPos(4)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_DILIG_STRT = iCurColumnPos(1)
            C_BAR        = iCurColumnPos(2)
            C_DILIG_END  = iCurColumnPos(3)
            C_MINUS_RATE = iCurColumnPos(4)
            C_MINUS_AMT  = iCurColumnPos(5)
            
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
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call AppendNumberPlace("6","1","0")
    Call AppendNumberPlace("7","2","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    Call MakeKeyStream("X")
	topleftOK = false
	frm1.txtPrevNext.value = ""
	
	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreTooBar()
        Exit Function
    End If  															'☜: Query db data

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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                                       '☜: Clear Condition Field

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = Frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim intDilig_strt
    Dim intDilig_end
    Dim lRow

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData        ' Multi-1
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        ggoSpread.Source = frm1.vspdData1   ' Multi-2
        If  ggoSpread.SSCheckChange = False Then
            if  lgBlnFlgChgValue = False then
                IntRetCD = DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
                Exit Function
            end if
        End If
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData    ' Multi-1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	With Frm1
       For lRow = 1 To .vspdData.MaxRows
  
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text = ggoSpread.InsertFlag OR .vspdData.Text = ggoSpread.UpdateFlag then
				.vspdData.Col = C_DILIG_NM

				If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					Call DisplayMsgBox("970000", "X","근태코드","x")
					Exit Function
				end if     
                .vspdData.Col = C_DILIG_CNT
                if  Cint(.vspdData.Text) > 0 then
                else
                    call DisplayMsgBox("970021", "x","횟수","x")
                    .vspdData.Action = 0 ' go to
                    exit function
                end if
            end if
        next

    end with


	ggoSpread.Source = frm1.vspdData1    ' Multi-2
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	With Frm1
       For lRow = 1 To .vspdData1.MaxRows
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
           if   .vspdData1.Text = ggoSpread.InsertFlag  then
					
                .vspdData1.Col = C_DILIG_STRT
                if  Cint(.vspdData1.Text) > 0 then
                else
                    call DisplayMsgBox("970021", "x","횟수","x")
                    .vspdData1.Action = 0 ' go to
                    exit function
                end if
                intDilig_strt = Cint(.vspdData1.Text)
                .vspdData1.Col = C_DILIG_END
                if  Cint(.vspdData1.Text) > 0 then
                else
                    call DisplayMsgBox("970021", "x","횟수","x")
                    .vspdData1.Action = 0 ' go to
                    exit function
                end if
				intDilig_end = Cint(.vspdData1.Text)
				if  intDilig_strt > intDilig_end then
				    call DisplayMsgBox("800171", "x","x","x")
				    .vspdData1.Action = 0 ' go to
				    exit function
				end if

				if Period_check(intDilig_strt, intDilig_end) = false then
					call DisplayMsgBox ("229901","X","X", "X")
					exit function
				end if
            end if
        next

    end with

    if  frm1.txtcrt_strt_dd.value <> "" then
        if  Cint(frm1.txtcrt_strt_dd.value) >= 0 AND Cint(frm1.txtcrt_strt_dd.value) <= 31 then
        else
            call DisplayMsgBox("970027", "x","일자입력","x")
            frm1.txtcrt_strt_dd.focus
            exit function
        end if
    end if

    if  frm1.txtcrt_end_dd.value <> "" then
        if  Cint(frm1.txtcrt_end_dd.value) >= 0 AND Cint(frm1.txtcrt_end_dd.value) <= 31 then
        else
            call DisplayMsgBox("970027", "x","일자입력","x")
            frm1.txtcrt_end_dd.focus
            exit function
        end if
    end if

    if  frm1.txtcrt_strt_mm.value <> "" and frm1.txtcrt_end_mm.value <> "" then
        if  Cint(frm1.txtcrt_strt_mm.value) > Cint(frm1.txtcrt_end_mm.value) then
            call DisplayMsgBox("970027", "x","계산기준기간","x")
            frm1.txtcrt_strt_mm.focus
            exit function
        end if
    end if

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

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
      
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
            If Frm1.vspdData.MaxRows < 1 Then
                Exit Function
            End If
    
	        With Frm1.vspdData
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData	
	        		ggoSpread.CopyRow
                    SetSpreadColor .ActiveRow, .ActiveRow
                    .col = C_DILIG
                    .Text = ""
                    .col = C_DILIG_NM
                    .Text = ""

	        		.ReDraw = True
	        		.focus
	        	End If
	        End With
        Case  Else

            If Frm1.vspdData1.MaxRows < 1 Then
                Exit Function
            End If
    
	        With Frm1.vspdData1
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData1
	        		ggoSpread.CopyRow
                    SetSpreadColor1 .ActiveRow, .ActiveRow

                    .col = C_DILIG_STRT
                    .Text = ""
                    .col = C_DILIG_END
                    .Text = "0"
    
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With

    End Select 

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
      
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
            ggoSpread.Source = frm1.vspdData
            ggoSpread.EditUndo
        Case  Else
            ggoSpread.Source = frm1.vspdData1
            ggoSpread.EditUndo
    End Select 

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
      
    Dim imRow, iCnt

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

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
                  With Frm1
                       .vspdData.ReDraw = False
                       .vspdData.Focus
                        ggoSpread.Source = .vspdData
                        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
                        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
                          
                        For iCnt = 1 To imRow 
                            .vspdData.Row = .vspdData.ActiveRow + iCnt - 1
                            
                            .vspdData.Col = C_BAR
                            .vspdData.Text = "~"
                        Next
                        
                       .vspdData.ReDraw = True
                  End With
        Case  Else
                  With Frm1
                       .vspdData1.ReDraw = False
                       .vspdData1.Focus
                        ggoSpread.Source = .vspdData1
                        ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
                        SetSpreadColor1 .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1

                        For iCnt = 1 To imRow 
                            .vspdData1.Row = .vspdData1.ActiveRow + iCnt - 1
                            
                            .vspdData1.Col = C_BAR
                            .vspdData1.Text = "~"
                        Next
                        
                       .vspdData1.ReDraw = True
                  End With
    End Select

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

    If Trim(lgActiveSpd) = "" Then
        lgActiveSpd = "M"
    End If
       
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
            If Frm1.vspdData.MaxRows < 1 then
                Exit function
            End if	
            With Frm1.vspdData 
               	.focus
               	ggoSpread.Source = frm1.vspdData 
              	lDelRows = ggoSpread.DeleteRow
            End With
       Case  "S"
            If Frm1.vspdData1.MaxRows < 1 then
                Exit function
            End if	
            With Frm1.vspdData1 
               	.focus
               	ggoSpread.Source = frm1.vspdData1 
               	lDelRows = ggoSpread.DeleteRow
            End With

   End Select

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
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, True)
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
 
    Select Case gActiveSpdSheet.id
   
		Case "vaSpread"
			Call InitSpreadSheet("A")      
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		
	End Select 
    
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
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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

	if LayerShowHide(1) = False then
		Exit Function
	end if	

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgSpreadFlg="       & gSpreadFlg    

    strVal = strVal     & "&topleftOK="       & topleftOK                   '☜: Query Key    

	if gSpreadFlg = "1" then
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
	else
		strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag        
	end if
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
		
    If frm1.txtCrt_strt_dd.value <> "" Then
        If frm1.txtCrt_strt_dd.value < 0 Or frm1.txtCrt_strt_dd.value > 31 Then
            call DisplayMsgBox("970027", "X","일자", "X")
            Frm1.txtCrt_strt_dd.focus 
            Set gActiveElement = document.ActiveElement   
            Exit Function
        End If
    end if

    If frm1.txtCrt_end_dd.value <> "" Then
        If frm1.txtCrt_end_dd.value < 0 Or frm1.txtCrt_end_dd.value > 31 Then
            call DisplayMsgBox("970027", "X","일자", "X")
            Frm1.txtCrt_end_dd.focus 
            Set gActiveElement = document.ActiveElement   
            Exit Function
        End If
    end if

	if LayerShowHide(1) = False then
		Exit Function
	end if	

'** Multi-1
	With frm1
		.txtMode.value        = Parent.UID_M0002                                  '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
         lgKeyStream = lgKeyStream & "1" & Parent.gColSep       ' Multi-1
        .txtKeyStream.Value   = lgKeyStream                                '☜: Save Key

        strVal = ""
        strDel = ""
        lGrpCnt = 1

        ggoSpread.Source = frm1.vspdData
        For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                   strVal = strVal & "C" & Parent.gColSep
                                                   strVal = strVal & lRow & Parent.gColSep
                                                   strVal = strVal & "000" & Parent.gColSep
                    .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                   strVal = strVal & "U" & Parent.gColSep
                                                   strVal = strVal & lRow & Parent.gColSep
                                                   strVal = strVal & "000" & Parent.gColSep
                    .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                   strDel = strDel & "D" & Parent.gColSep
                                                   strDel = strDel & lRow & Parent.gColSep
                                                   strDel = strDel & "000" & Parent.gColSep
                    .vspdData.Col = C_DILIG 	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
        Next
	End With
    Frm1.txtMaxRows.value = lGrpCnt-1	
	Frm1.txtSpread.value = strDel & strVal

'** Multi-2
	With frm1
		.txtMode.value        = Parent.UID_M0002                                  '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
         lgKeyStream = lgKeyStream & "2" & Parent.gColSep       ' Multi-2
        .txtKeyStream.Value   = lgKeyStream                                '☜: Save Key

        strVal = ""
        strDel = ""
        lGrpCnt = 1

        ggoSpread.Source = frm1.vspdData1
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0

           Select Case .vspdData1.Text
               Case ggoSpread.InsertFlag                                      '☜: Insert
					
                                                    strVal = strVal & "C" & Parent.gColSep
                                                    strVal = strVal & lRow & Parent.gColSep
                    .vspdData1.Col = C_DILIG_STRT : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_DILIG_END  : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MINUS_RATE : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MINUS_AMT  : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
					
                                                    strVal = strVal & "U" & Parent.gColSep
                                                    strVal = strVal & lRow & Parent.gColSep
                    .vspdData1.Col = C_DILIG_STRT : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_DILIG_END  : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MINUS_RATE : strVal = strVal & Trim(.vspdData1.Text) & Parent.gColSep
                    .vspdData1.Col = C_MINUS_AMT  : strVal = strVal & Trim(.vspdData1.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                    strDel = strDel & "D" & Parent.gColSep
                                                    strDel = strDel & lRow & Parent.gColSep
                    .vspdData1.Col = C_DILIG_STRT : strDel = strDel & Trim(.vspdData1.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	End With
    Frm1.txtMaxRows1.value = lGrpCnt-1	
	Frm1.txtSpread1.value = strDel & strVal

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
	
	if LayerShowHide(1) = False then
		Exit Function
	end if		
		
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtAllow_cd=" & Trim(frm1.txtAllow_cd.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    Call InitData()
    ggoSpread.Source = Frm1.vspdData
    Call ggoOper.LockField(Document, "Q")
    ggoSpread.Source = Frm1.vspdData1
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = False
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData

    ggoSpread.Source = Frm1.vspdData1
    Frm1.vspdData1.MaxRows = 0
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
End Function
'========================================================================================================
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
		
	strRet = window.showModalDialog("s1413pa1.asp", "", _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
       Exit Sub
	Else
       Call SetCollateralNo(strRet)
	End If	
End Sub

'======================================================================================================
' Name : OpenDiligPopup
' Desc : OpenDiligPopup Reference Popup
'======================================================================================================
Function OpenDiligPopup(Byval strCode, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    ggoSpread.Source = frm1.vspdData

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		arrParam(1) = "HCA010T"         ' TABLE 명칭 
		arrParam(2) = strCode								           ' Code Condition
		arrParam(3) = ""									           ' Name Cindition
		arrParam(4) = ""		                                       ' Where Condition
		arrParam(5) = "근태코드"	         			           ' TextBox 명칭 
	
		arrField(0) = "dilig_cd"			    			           ' Field명(0)
		arrField(1) = "dilig_nm"		    				           ' Field명(1)
    
		arrHeader(0) = "근태코드"						           ' Header명(0)
		arrHeader(1) = "근태코드명"						           ' Header명(1)

	arrParam(0) = arrParam(5)							               ' 팝업 명칭 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_DILIG
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetDiligPopUp(arrRet)
	End If	
	
End Function

'========================================================================================================
' Name : SetDiligPopUp()
' Desc : OpenSalesPlanPopup에서 Return되는 값 setting
'========================================================================================================
Function SetDiligPopUp(Byval arrRet)

	With frm1
		.vspdData.Col = C_DILIG
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_DILIG_NM
		.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_DILIG, .vspdData.Row)		<% ' 변경이 읽어났다고 알려줌 %>
		.vspdData.Col = C_DILIG
		.vspdData.action =0
	End With
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : SetCollateralNo()
' Desc : developer describe this line 
'========================================================================================================
Function SetCollateralNo(arrRet)
	frm1.txtGlNo.Value = arrRet
End Function

'========================================================================================================
' Name : SetBizPartner()
' Desc : developer describe this line 
'========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtCustomer.Value = arrRet(0)
	frm1.txtCustomerNm.Value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : SetCurrency
' Desc : developer describe this line 
'========================================================================================================
Function SetCurrency(arrRet)
	frm1.txtCurrency.Value = arrRet(0)
	lgBlnFlgChgValue = True
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

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"
	gSpreadFlg = 1
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
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
	gSpreadFlg = 2
    gMouseClickStatus = "SP1C"

    Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If

End Sub

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

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_DILIG_POP
				.Col = Col - 1
				.Row = Row
				Call OpenDiligPopup(.Text, Row)
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        case C_DILIG
            IntRetCd = CommonQueryRs(" dilig_nm "," HCA010T "," dilig_cd =  " & FilterVar(Frm1.vspdData.Text , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If  IntRetCd = false then
				Call DisplayMsgBox("970000", "X","근태코드","x")
                Frm1.vspdData.Col = C_DILIG_NM
		    	Frm1.vspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                vspdData_Change = true
            Else
                Frm1.vspdData.Col = C_DILIG_NM
		    	Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
            End if 
                   
    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Function

'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

   	If Frm1.vspdData1.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData1.text) < UNICDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
    End If
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	topleftOK = true	
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
'===============================================================================================
' Function : Period_check
' Description : 근태 회수 구간이 이미 존재하면 false
'===============================================================================================
Function Period_Check(Scnt, Ecnt)
	
	Dim txtOld
	Dim txtMid
	
	txtOld = frm1.txtPeriod.value 

	Period_check = false
	
	if len(txtOld) > Ecnt  then
		txtMid = Mid(txtOld,Scnt, Ecnt - Scnt)  '대상 구획만 잘라낸다 '

		if Len(Replace(txtMid, "0", "")) = (Ecnt - Scnt) then ' 그 구획 안에 1이 하나라도 있으면 false
			Period_check = true
		end if
	else

		if Len(txtOld) < Scnt then
			Period_check = true
		end if
	end if	
	
End function

Sub txtCrt_strt_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtCrt_end_mm_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtCrt_strt_dd_OnChange()

    if  isnumeric(frm1.txtCrt_strt_dd.value) then
        frm1.txtCrt_strt_dd.value = Right("0" & frm1.txtCrt_strt_dd.value, 2)
    else
        frm1.txtCrt_strt_dd.value = "00"
    end if
    lgBlnFlgChgValue = True

End Sub

Sub txtCrt_end_dd_OnChange()

    if  isnumeric(frm1.txtCrt_end_dd.value) then
        frm1.txtCrt_end_dd.value = Right("0" & frm1.txtCrt_end_dd.value, 2)
    else
        frm1.txtCrt_end_dd.value = "00"
    end if
    lgBlnFlgChgValue = True

End Sub

Sub txtday_calcu_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtcalcu_bas_dd_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtday_calcu_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub txtcalcu_bas_dd_OnChange()
    lgBlnFlgChgValue = True
End Sub

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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여기준등록</font></td>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>

            					<TD WIDTH="43%" HEIGHT=100%>
                                    <script language =javascript src='./js/h7001ma1_vaSpread_vspdData.js'></script>
		            			</TD>
            					<TD WIDTH="57%" HEIGHT=100%>
                                    <script language =javascript src='./js/h7001ma1_vaSpread1_vspdData1.js'></script>
		            			</TD>
						    </TR>
						</TABLE>
                    </TD>
                </TR>
                <TR> 
					<TD WIDTH=100% HEIGHT=100 VALIGN=TOP>
                        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>상여계산기준</LEGEND>
					        <TABLE WIDTH=100% HEIGHT="100%" CELLSPACING=0>
	 						    <TR>
	                   				<TD CLASS="TD656">
	                   				    계산기준기간&nbsp;<SELECT NAME="txtCrt_strt_mm" ALT="계산기준기간월" STYLE="WIDTH:90px" TAG="21N"></SELECT>&nbsp; 월 
                                        <INPUT TYPE=TEXT NAME="txtCrt_strt_dd" SIZE=05 MAXLENGTH=2 tag=21XXXU ALT="계산기준기간일"> 일&nbsp;~&nbsp;

	                   				    <SELECT NAME="txtCrt_end_mm" ALT="계산기준기간월" STYLE="WIDTH:90px" TAG="21N"></SELECT>&nbsp; 월 
                                        <INPUT TYPE=TEXT NAME="txtCrt_end_dd" SIZE=05 MAXLENGTH=2 tag=21XXXU ALT="계산기준기간일"> 일<BR>
                                        <BR>
	                   				    일할계산방법  상여금 / &nbsp;
	                   				    <SELECT NAME="txtday_calcu" ALT="일할계산방법" STYLE="WIDTH:150px" TAG="21N"><OPTION Value=""></OPTION></SELECT>&nbsp;*&nbsp;
	                   				    <SELECT NAME="txtcalcu_bas_dd" ALT="일할계산방법" STYLE="WIDTH:150px" TAG="21N"><OPTION Value=""></OPTION></SELECT>
                                        <BR>
	                   				</TD>
	                   			</TR>
						    </TABLE>
					    </FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtPeriod"      TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows1" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


