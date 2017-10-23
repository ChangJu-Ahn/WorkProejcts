<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 연차조회및조정 
*  3. Program ID           : H1a03ma1
*  4. Program Name         : H1a03ma1
*  5. Program Desc         : 연차조회및조정 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/29
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Hwang Jeong-won
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID      = "h9203mb1.asp"	
Const C_SHEETMAXROWS    = 15	                                         '☜: Visble row

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim gSpreadFlg
Dim topleftOK
Dim lgStrPrevKey1
Dim lsInternal_cd
Dim lgSpreadChange
Dim lgSpreadChange1

Dim C_ALLOW_NM  
Dim C_ALLOW_AMT   

Dim C_ALLOW_NM_S
Dim C_ALLOW_AMT_S  

Dim C_DILIG_CD
Dim C_DILIG_NM
Dim C_DILIG_TYPE
Dim C_DILIG_CNT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
		C_ALLOW_NM		= 1
		C_ALLOW_AMT		= 2

		C_ALLOW_NM_S    = 1
		C_ALLOW_AMT_S   = 2
    End If

    If pvSpdNo = "B" Then
		C_DILIG_CD		= 1
		C_DILIG_NM		= 2
		C_DILIG_TYPE	= 3
		C_DILIG_CNT		= 4
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
	lsInternal_cd     = ""
	
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
	lgSpreadChange    = False
	lgSpreadChange1   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay
	lgBlnFlgChgValue = False
	
   	frm1.txtYymm.focus
    Call ggoOper.FormatDate(frm1.txtYymm, parent.gDateFormat, 2)	
    Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType,strYear,strMonth,strDay)	
    frm1.txtYymm.Year	= strYear
    frm1.txtYymm.Month	= strMonth
    frm1.txtYymm.Day	= strDay
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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	
    lgKeyStream	= frm1.txtYymm.Year & Right("0" & frm1.txtYymm.Month,2) & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.cboYearType.Value & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep       'You Must append one character(parent.gColSep)
   
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'H0046' AND MINOR_CD <> '3' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    iCodeArr = lgF0
    iNameArr = lgF1

    Call SetCombo2(frm1.cboYearType,iCodeArr, iNameArr,Chr(11))

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
'	frm1.vspdData1S.Col = C_ALLOW_AMT_S
'	frm1.vspdData1S.Row = 1
'	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData,C_ALLOW_AMT,1,frm1.vspdData.MaxRows,False,4,5,"V")
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   'sbk 

	    With frm1.vspdData
            ggoSpread.Source = Frm1.vspdData
            ggoSpread.Spreadinit "V20060811",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_ALLOW_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                             ' ☜:☜: Hide maxcols
           .ColHidden = True
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk
           
			ggoSpread.SSSetEdit   C_ALLOW_NM  , "수당코드"    ,20                             
			ggoSpread.SSSetFloat  C_ALLOW_AMT , "수당액",      22, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

	       .ReDraw = true

           Call SetSpreadLock("A")
   
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	Call initSpreadPosVariables("B")   'sbk 

	    With frm1.vspdData1
            ggoSpread.Source = Frm1.vspdData1

            ggoSpread.Spreadinit "V20060811",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_DILIG_CNT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                              ' ☜:☜: Hide maxcols
           .ColHidden = True
           
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("B") 'sbk
           
			Call AppendNumberPlace("6","3","0")	
		   ggoSpread.SSSetEdit   C_DILIG_CD    ,	"연차발생근태" ,      10,,,2,2
		   ggoSpread.SSSetEdit   C_DILIG_NM    ,	"연차발생근태명" ,      14,,,20,2
		   ggoSpread.SSSetEdit   C_DILIG_TYPE    ,	"근태타입" ,  10,,,20,2
		   ggoSpread.SSSetFloat  C_DILIG_CNT   ,	"횟수" ,      10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

	       .ReDraw = true

           Call SetSpreadLock("B")

        End With
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1

        If pvSpdNo = "A" Then
            ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
        End If

        If pvSpdNo = "B" Then

            ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
        End If
      
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

            C_ALLOW_NM	= iCurColumnPos(1)
            C_ALLOW_AMT = iCurColumnPos(2)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_DILIG_CD		= iCurColumnPos(1)
			C_DILIG_NM		= iCurColumnPos(2)
			C_DILIG_TYPE	= iCurColumnPos(3)
			C_DILIG_CNT		= iCurColumnPos(4)	
    End Select    
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	
	Call ggoOper.LockField(Document, "N")		

    Call AppendNumberPlace("6","3","2")
    Call AppendNumberPlace("7","2","0")
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtYymm, Parent.gDateFormat, 2)

    Call InitSpreadSheet("")                                                           'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    
	Call SetToolbar("1100100111011111")												'⊙: Set ToolBar
  
    Call InitComboBox
  
    Call cboYearType_onchange()
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
    Dim iDx
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	Else
		ggoSpread.Source = Frm1.vspdData1
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
    End If    
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange()  then
       Exit Function
    End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    Call MakeKeyStream("X")
	topleftOK = false
	frm1.txtPrevNext.value = ""
		
	Call DisableToolBar(Parent.TBC_QUERY)
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
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
	Call SetToolbar("1100111111011111")							                 '⊙: Set ToolBar
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
	Call DisableToolBar(Parent.TBC_DELETE)
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
    Dim lFlag
  	Dim strDate 
        
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
        
    If lgBlnFlgChgValue = False Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			ggoSpread.Source = frm1.vspdData1
			If ggoSpread.SSCheckChange = False Then
				IntRetCD = DisplayMsgBox("900001","X","X","X")                   '⊙: No data changed!!
				Exit Function
			End If
		End If        
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

	strDate		= UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtyymm.Year,Right("0" & frm1.txtyymm.Month,2),"01")	

    IF  FuncAuthority("@", UniConvDateToYYYYMMDD(strDate,parent.gDateFormat,""), parent.gUsrID) = "N" THEN
        Call DisplayMsgbox("800304","X","X","X")         '연월차 마감처리된 지급월 입니다.
        Exit Function
    END IF   
	

    lFlag = DisplayMsgbox("800439",35,"X","X")  'VB_YES_NO_CANCEL

    If lFlag = vbYes Then   '6
		frm1.txtTaxFlag.value = "Y"
	Elseif lFlag = vbNo Then  '7
		frm1.txtTaxFlag.value = "N"
	Else
		Call FncQuery()
		Exit Function     
	End If	

    Call MakeKeyStream("X")
	Call DisableToolBar(Parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
	If gSpreadFlg = 1 Then
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo
	Else
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.EditUndo
	End If
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

    Call MakeKeyStream("X")
	topleftOK = false
	frm1.txtPrevNext.value = "P"
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
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

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgSpreadFlg="       & gSpreadFlg
	if gSpreadFlg = "1" then
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
	else
		strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
	end if	
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & frm1.txtPrevNext.value
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
		
	With frm1
		.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
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
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	Call SetToolbar("1100111111011111")												'⊙: Set ToolBar
	Call ggoOper.LockField(Document, "Q")
	Set gActiveElement = document.ActiveElement   
	lgBlnFlgChgValue = False
	frm1.vspdData.focus
	lgIntFlgMode = Parent.OPMD_UMODE
    Call cboYearType_onchange()
End Function

Function DBQueryFail()

  	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtYymm.focus
 
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar

    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    frm1.vspdData.focus
    Call cboYearType_onchange()
    
   ' Call CheckLogic(frm1.cboYearType.value)
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
	Call MainNew()	
End Function
'========================================================================================================
' Function Name : FuncAuthority
' Function Desc : 시스템마감체크 
'========================================================================================================
Function FuncAuthority(Pay_gubun, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"    
    IntRetCD = CommonQueryRs("close_type, close_dt, emp_no","hda270t","org_cd='1'  and pay_gubun = 'Z' and pay_type= " & FilterVar(Pay_gubun, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  IntRetCD = false then
        strRet = "Y"
    else
        SELECT CASE Replace(lgF0, Chr(11), "")
        	CASE "1" '마감형태 : 정상 
        	    IF  UniConvDateToYYYYMMDD(Replace(lgF1,Chr(11),""),parent.gServerDateFormat,"") <= Pay_yymmdd THEN 
        	        strRet = "Y"
        		ELSE
        	        strRet = "N" 
        		END IF
           CASE "2" '마감형태 : 마감 
        	    IF  UniConvDateToYYYYMMDD(Replace(lgF1,Chr(11),""),parent.gServerDateFormat,"") < Pay_yymmdd THEN 
        	        strRet = "Y" 
        		ELSE
        	        strRet = "N" 
        	    END IF
        END SELECT
        
    end if

    FuncAuthority = strRet

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
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
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)

	Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field

	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value   = arrRet(0)
			.txtName.value     = arrRet(1)
			.txtDept_nm.value  = arrRet(2)
			.txtRollPstn.value = arrRet(3)
			.txtPay_grd.value  = arrRet(4)
			.txtEntr_dt.text   = arrRet(5)
		Else 'spread
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
		End If

		'Set gActiveElement = document.ActiveElement

		lgBlnFlgChgValue = False
	End With
End Sub
 
'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
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
    Dim srtGroup_entr_dt

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    	frm1.txtDept_nm.value  = ""
		frm1.txtRollPstn.value = ""
		frm1.txtPay_grd.value  = ""
		frm1.txtEntr_dt.text   = ""

	    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
        Call initData()
	Else
	
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                              strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)	    
	    Call CommonQueryRs("GROUP_ENTR_DT","HAA010T"," EMP_NO =  " & FilterVar(frm1.txtEmp_no.value , "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    
	    strEntr_dt			= UNIDateClientFormat(strEntr_dt)
		srtGroup_entr_dt	= UNIDateClientFormat(Replace(lgF0, Chr(11), ""))
		
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
			frm1.txtEmp_no.value = ""
			frm1.txtDept_nm.value  = ""
			frm1.txtRollPstn.value = ""
			frm1.txtPay_grd.value  = ""
			frm1.txtEntr_dt.text   = ""
			frm1.txtEmp_no.focus

		    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
		    Call initData()

		    Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
    		frm1.txtDept_nm.value  = strDept_nm
			frm1.txtRollPstn.value = strRoll_pstn
			frm1.txtPay_grd.value  = strPay_grd1 & "-" & strPay_grd2
			
			If srtGroup_entr_dt = "" or srtGroup_entr_dt = "X" Then 
				frm1.txtEntr_dt.text   = strEntr_dt
			Else
				frm1.txtEntr_dt.text   = srtGroup_entr_dt
			End If
        End if 
    End if  

End Function 

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method            (중복된 체크로직을 Function 에 담음..)
'=======================================================================================================
Function CheckLogic(iWhere)											
	With Frm1
	    If  iWhere = 1 Then    '연차  
			.txtYearSaveTot.value = int(.txtYearSave.value) + int(.txtYearPart.value) + int(.txtYearBonus.value)

			If int(.txtYearSaveTot.value) > int(.txtMaxYearSave.value) Then
				.txtYearSaveTot.value = .txtMaxYearSave.value
			End If
			
			.txtYearCnt.value = int(.txtYearSaveTot.value)  - .txtYearUse.value	    
			.txtYearAmt.value = int(.txtYearCnt.value) * .txtBasAmt.value


			.txtTotAmt.value = .txtYearAmt.value
			.txtRealProvAmt.value =   .txtYearAmt.value - .txtIncomeTaxAmt.value -.txtResTaxAmt.value - .txtEmpInsurAmt.value
		ElseIF 	iWhere = 2 Then '월차 
		     if int(.txtMonthDutyCnt.value ) > int(.txtMonthUse.value ) Then
		     
				.txtMonthAmt.value = (int(.txtMonthSave.value) - .txtMonthDutyCnt.value) * .txtBasAmt.value
				.txtMonthCnt.value = int(.txtMonthSave.value) - .txtMonthDutyCnt.value
			 else
				.txtMonthAmt.value = (int(.txtMonthSave.value) - .txtMonthUse.value) * .txtBasAmt.value
				.txtMonthCnt.value = int(.txtMonthSave.value) - .txtMonthUse.value
			 end if	
			.txtTotAmt.value =  .txtMonthAmt.value
			.txtRealProvAmt.value =   .txtMonthAmt.value - .txtIncomeTaxAmt.value -.txtResTaxAmt.value - .txtEmpInsurAmt.value			 
		End If
	End With	
End Function
'--------------------------------------------------------------------------------------------------------
Sub txtMonthSave_Change()  '월차발생 
	IF frm1.cboYearType.value = "2" THEN
		Call CheckLogic(2)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtMonthSave.value= "0"
	END IF 
End Sub

Sub txtMonthUse_Change()  '월차사용 
	IF frm1.cboYearType.value = "2" THEN
		Call CheckLogic(2)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtMonthUse.value= "0"	
	END IF 	
End Sub

Sub txtMonthAmt_Change()  '월차수당 
	IF frm1.cboYearType.value = "2" THEN
'		frm1.txtIncomeTaxAmt.value = "0"
'		frm1.txtResTaxAmt.value = "0"
'		frm1.txtEmpInsurAmt.value = "0"		
		frm1.txtTotAmt.value = frm1.txtMonthAmt.value
		frm1.txtRealProvAmt.value =   frm1.txtMonthAmt.value - frm1.txtIncomeTaxAmt.value -frm1.txtResTaxAmt.value - frm1.txtEmpInsurAmt.value
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtMonthAmt.value= "0"	
	END IF 	
End Sub
'--------------------------------------------------------------------------------------------------------
Sub txtYearSave_Change()  '연차발생 
	IF frm1.cboYearType.value = "1" THEN
		Call CheckLogic(1)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtYearSave.value= "0"
	END IF 
End Sub


Sub txtYearPart_Change()  '근속가산 
	IF frm1.cboYearType.value = "1" THEN
		Call CheckLogic(1)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtYearPart.value= "0"
	END IF 
End Sub

Sub txtYearBonus_Change()  '연차분할 
	IF frm1.cboYearType.value = "1" THEN
		Call CheckLogic(1)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtYearBonus.value= "0"
	END IF 
End Sub

Sub txtYearUse_Change()  '연차사용 
	IF frm1.cboYearType.value = "1" THEN
		Call CheckLogic(1)
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtYearUse.value= "0"	
	END IF 	
End Sub

Sub txtYearAmt_Change()  '연차수당 
	IF frm1.cboYearType.value = "1" THEN
'		frm1.txtIncomeTaxAmt.value = "0"
'		frm1.txtResTaxAmt.value = "0"
'		frm1.txtEmpInsurAmt.value = "0"		
		frm1.txtTotAmt.value = frm1.txtYearAmt.value
		frm1.txtRealProvAmt.value =   frm1.txtYearAmt.value - frm1.txtIncomeTaxAmt.value -frm1.txtResTaxAmt.value - frm1.txtEmpInsurAmt.value
		lgBlnFlgChgValue = True
	ELSE
	    frm1.txtYearAmt.value= "0"	
	END IF 	
End Sub
'========================================================================================================
'   Event Name : cboYearType_onchange
'========================================================================================================
Sub cboYearType_onchange()
on error resume next

	if Frm1.cboYearType.value = 1 Then
		Call LockObjectField( Frm1.txtMonthSave,"P")
		Call LockObjectField( Frm1.txtMonthUse,"P")
		Call LockObjectField( Frm1.txtMonthAmt,"P")

		Call LockObjectField( Frm1.txtYearSave,"O")
		Call LockObjectField( Frm1.txtYearPart,"O")
		Call LockObjectField( Frm1.txtYearBonus,"O")
		Call LockObjectField( Frm1.txtYearUse,"O")
		Call LockObjectField( Frm1.txtYearAmt,"O")			
	else
		Call LockObjectField( Frm1.txtMonthSave,"O")
		Call LockObjectField( Frm1.txtMonthUse,"O")
		Call LockObjectField( Frm1.txtMonthAmt,"O")
			
		Call LockObjectField( Frm1.txtYearSave,"P")
		Call LockObjectField( Frm1.txtYearPart,"P")
		Call LockObjectField( Frm1.txtYearBonus,"P")
		Call LockObjectField( Frm1.txtYearUse,"P")
		Call LockObjectField( Frm1.txtYearAmt,"P")													
	end if
End Sub

'======================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================%>
Private Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData1 
	
    ggoSpread.Source = frm1.vspdData1
   
    If Row > 0 And Col = C_SUB_CD_POP Then
		    .Row = Row
		    .Col = C_SUB_CD

		    Call OpenAllow(1)        
    End If
    
    End With
End Sub

'========================================================================================================
' Function Name : vspdData_Click
' Function Desc : gSpreadFlg Setting
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

	gSpreadFlg = 1

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
' Function Name : vspdData1_Click
' Function Desc : gSpreadFlg Setting
'========================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP2C"   

    Set gActiveSpdSheet = frm1.vspdData1

	gSpreadFlg = 2
   
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
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
     End If
End Sub    

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================%>
Private Sub vspdData1_Change(ByVal Col , ByVal Row )
   Dim IntRetCD

   Frm1.vspdData1.Row = Row
   Frm1.vspdData1.Col = Col

   Select Case Col
'         Case  C_SUB_CD
 '          	IntRetCD = CommonQueryRs(" allow_cd,allow_nm "," hda010t "," pay_cd='*' And code_type='2' And allow_cd =  " & FilterVar(frm1.vspdData1.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'
 '          	If IntRetCD=False And Trim(frm1.vspdData1.Text)<>"" Then
  '              Call DisplayMsgBox("800176","X","X","X")             '☜ : 등록되지 않은 코드입니다.
   ' 	    	frm1.vspdData1.Col = C_SUB_CD_NM
    '       		frm1.vspdData1.Text=""
     '       Else
    '	    	frm1.vspdData1.Col = C_SUB_CD_NM
     '       	frm1.vspdData1.Text=Trim(Replace(lgF1,Chr(11),""))
      '     	End If
   End Select    

   If Frm1.vspdData1.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData1.text) < CDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
   End If
   
   ggoSpread.Source = frm1.vspdData1
   ggoSpread.UpdateRow Row
   lgSpreadChange1 = True

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
'=======================================================================================================
'   Event Name : txtYymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYymm.Action = 7
        frm1.txtYymm.focus
    End If
End Sub

Sub txtYymm_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연월차수당조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>정산년월</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="정산년도" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>	
									<TD CLASS=TD5>연월차구분</TD>
									<TD CLASS=TD6><SELECT NAME="cboYearType" STYLE="Width:100px;" tag="12" ALT="연월차구분"></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
						        <TD COLSPAN=4>
                                    
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
						            	<TR>
											<TD CLASS=TD5 NOWRAP>부서명</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_nm" MAXLENGTH="30" SIZE=30  ALT ="부서명" tag="14">&nbsp;</TD>
											<TD CLASS=TD5 NOWRAP>직  위</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRollPstn" MAXLENGTH="20" SIZE=20  ALT ="직위" tag="14">&nbsp;</TD>
										</TR>
						            	<TR>
											<TD CLASS=TD5 NOWRAP>급  호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd" MAXLENGTH="20" SIZE=20  ALT ="급호" tag="14">&nbsp;</TD>
											<TD CLASS=TD5 NOWRAP>입사일</TD>
											<TD CLASS=TD6 NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEntr_dt CLASSID=<%=gCLSIDFPDT%> ALT="입사일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
											</TD>
										</TR>
						            	<TR>
											<TD CLASS=TD5 NOWRAP>근속기간</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDutyYy" MAXLENGTH="2" SIZE=5  ALT ="근속년" tag="24">&nbsp;/&nbsp;
											                     <INPUT NAME="txtDutyMm" MAXLENGTH="2" SIZE=5  ALT ="근속월" tag="24">&nbsp;/&nbsp;
											                     <INPUT NAME="txtDutyDd" MAXLENGTH="2" SIZE=5  ALT ="근속일" tag="24"></TD>
					                        <TD CLASS=TD5 NOWRAP>기준금액</TD>
					                        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtBasAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="24X2Z" ALT="기준금액"></OBJECT>');</SCRIPT></TD>
			                        	</TR>
							        </TABLE>
							        
							    </TD>
							</TR>
						    <TR>
						        <TD COLSPAN=4>
                                    
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
										<TR>  
										    <TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;월 차 발 생</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthSave CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X7Z" ALT="월차발생"></OBJECT>');</SCRIPT>개</TD>
										    <TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;연 차 발 생</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearSave CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X7Z" ALT="연차발생"></OBJECT>');</SCRIPT>개</TD>	                            
                           				</TR>
										<TR>
										    <TD CLASS=TD5 NOWRAP>월 차 사 용</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthUse CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X6Z" ALT="월차사용"></OBJECT>');</SCRIPT>개</TD>
										    <TD CLASS=TD5 NOWRAP>근 속 가 산</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearPart CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X7Z" ALT="근속가산"></OBJECT>');</SCRIPT>개</TD>
	                       				</TR>
										<TR>
										    <TD CLASS=TD5 NOWRAP>월차 의무 사용</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthDutyCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="월차의무사용"></OBJECT>');</SCRIPT>개</TD>
										    <TD CLASS=TD5 NOWRAP>연 차 분 할</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearBonus CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="연차분할"></OBJECT>');</SCRIPT>개</TD>	                            
            							</TR>
										<TR>
										    <TD CLASS=TD5 NOWRAP>월 차 지 급</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="월차지급"></OBJECT>');</SCRIPT>개</TD>
										    <TD CLASS=TD5 NOWRAP>연차최대개수</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMaxYearSave CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="연차최대개수"></OBJECT>');</SCRIPT>개)</TD>
            							</TR>
										<TR>
										    <TD CLASS=TD5 NOWRAP>월 차 수 당</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="21X2Z" ALT="월차수당"></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP></TD>
										    <TD CLASS=TD6 NOWRAP><HR ALIGN= "LEFT" WIDTH=150></TD>
            							</TR>
										<TR>
										    <TD CLASS=TD5 NOWRAP></TD>
										    <TD CLASS=TD6 NOWRAP></TD>
										    <TD CLASS=TD5 NOWRAP>연 차 총 발 생</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearSaveTot CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="연차총발생"></OBJECT>');</SCRIPT>개</TD>	                            
										<TR>
											<TD CLASS=TD5>소  득  세</TD>
											<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtIncomeTaxAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="소득세" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP>연 차 사 용</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearUse CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X6Z" ALT="연차사용"></OBJECT>');</SCRIPT>개</TD>
            							</TR>
										<TR>
											<TD CLASS=TD5>주  민  세</TD>
											<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtResTaxAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="주민세" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP></TD>
										    <TD CLASS=TD6 NOWRAP><HR ALIGN= "LEFT" WIDTH=150></TD>
            							</TR>
										<TR>
											<TD CLASS=TD5>고 용 보 험</TD>
											<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtEmpInsurAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="고용보험" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP>연 차 지 급</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="연차지급"></OBJECT>');</SCRIPT>개</TD>	                            
            							</TR>
										<TR>
											<TD CLASS=TD5>총 지 급 액</TD>
											<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="총지급액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP>연 차 수 당</TD>
										    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="21X2Z" ALT="연차수당"></OBJECT>');</SCRIPT></TD>
            							</TR> 
										<TR>
											<TD CLASS=TD5>실 지 급 액</TD>
											<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRealProvAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="실지급액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										    <TD CLASS=TD5 NOWRAP></TD>
										    <TD CLASS=TD6 NOWRAP></TD>								
            							</TR>  
							        </TABLE>
							        
							    </TD>
							</TR>
							<TR>
								<TD HEIGHT="35%" WIDTH="50%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT="35%" WIDTH="50%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtsave_fund"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtincome_tax"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtres_tax"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtmed_insur"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtanut"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtemp_insur"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtTaxFlag"     TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows1" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

