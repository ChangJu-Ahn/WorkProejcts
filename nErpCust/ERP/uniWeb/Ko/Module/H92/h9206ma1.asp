<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single-Multi Sample
*  3. Program ID           : h9206ma1
*  4. Program Name         : h9206ma1
*  5. Program Desc         : Single-Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h9206mb1.asp"						           '☆: Biz Logic ASP Name
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
Dim lsInternal_cd
Dim topleftOK
Dim lgStrPrevKey1
Dim gSpreadFlg

Dim C_AllowNm  
Dim C_AllowAmt   
Dim C_AllowNm1
Dim C_AllowAmt1  

Dim C_AllowNm2  
Dim C_AllowAmt2   
Dim C_AllowNm21
Dim C_AllowAmt21  

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(spd)  
	if spd = "A" then
		C_AllowNm   = 1
		C_AllowAmt  = 2
	elseif spd= "B" then
		C_AllowNm1  = 1
		C_AllowAmt1 = 2
	elseif spd = "C" then
		C_AllowNm2    = 1
		C_AllowAmt2   = 2
	elseif spd= "D" then
		C_AllowNm21   = 1
		C_AllowAmt21  = 2
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
	gSpreadFlg		  = 1
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay

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
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream	= frm1.txtYymm.Year & Right("0" & frm1.txtYymm.Month,2) & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep       'You Must append one character(parent.gColSep)
End Sub        
	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	frm1.vspdData1S.Col = C_AllowAmt1
	frm1.vspdData1S.Row = 1
	frm1.vspdData1S.Text = FncSumSheet(frm1.vspdData,C_AllowAmt,1,frm1.vspdData.MaxRows,False,4,5,"V")
	
	frm1.vspdData2S.Col = C_AllowAmt21
	frm1.vspdData2S.Row = 1
	frm1.vspdData2S.Text = FncSumSheet(frm1.vspdData2,C_AllowAmt2,1,frm1.vspdData2.MaxRows,False,4,5,"V")

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(spd)

	Call initSpreadPosVariables(spd)  
	
	if spd ="A" then
		With frm1.vspdData

			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread    

		.ReDraw = false        
		.MaxCols   = C_AllowAmt + 1                                                  ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:
	       
		.MaxRows = 0

			Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit   C_AllowNm  , "연차기준수당코드"    ,26                             
		ggoSpread.SSSetFloat  C_AllowAmt , "연차기준수당액",      22, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.ReDraw = true
		   
		End With
	end if
	
	if spd ="B" then
		With frm1.vspdData1S
			ggoSpread.Source = Frm1.vspdData1S
			ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread
			
		.ReDraw = false        
		.MaxCols = C_AllowAmt1 + 1
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True
		.MaxRows = 1
			Call GetSpreadColumnPos("B") 
		.ScrollBars   = 0
		.DisplayColHeaders = False
		.OperationMode = 1

		.Col = 0 
		.Row = 1
		.Text = "합계"

			ggoSpread.SSSetEdit   C_AllowNm1    , "", 26                             
			ggoSpread.SSSetFloat  C_AllowAmt1   , "", 22, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","10"

		.ReDraw = true	
		End With
	end if

	if spd ="C" then
		With frm1.vspdData2

			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread    

		.ReDraw = false        
		.MaxCols   = C_AllowAmt2 + 1                                                  ' ☜:☜: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:
	       
		.MaxRows = 0

		Call GetSpreadColumnPos("C")  

		ggoSpread.SSSetEdit   C_AllowNm2  , "월차기준수당코드"    ,26                             
		ggoSpread.SSSetFloat  C_AllowAmt2 , "월차기준수당액",      22, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		.ReDraw = true
		   
		End With
	end if
	
	if spd ="D" then
		With frm1.vspdData2S
			ggoSpread.Source = Frm1.vspdData2S
			ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread
			
		.ReDraw = false        
		.MaxCols = C_AllowAmt21 + 1
		.Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
		.ColHidden = True
		.MaxRows = 1
		
		Call GetSpreadColumnPos("D") 
		
		.ScrollBars   = 0
		.DisplayColHeaders = False
		.OperationMode = 1

		.Col = 0 
		.Row = 1
		.Text = "합계"

			ggoSpread.SSSetEdit   C_AllowNm21    , "", 26                             
			ggoSpread.SSSetFloat  C_AllowAmt21   , "", 22, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","10"

		.ReDraw = true	
		End With
	end if

    Call SetSpreadLock 

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()

    ggoSpread.Source = Frm1.vspdData1S
    With frm1.vspdData1S
		.ReDraw = False
         ggoSpread.SpreadLock      -1,-1,-1
		 ggoSpread.SSSetProtected	.MaxCols,-1,-1                        
		.ReDraw = True
    End With

    ggoSpread.Source = Frm1.vspdData2
    ggoSpread.SpreadLockWithOddEvenRowColor()

    ggoSpread.Source = Frm1.vspdData2S
    With frm1.vspdData2S
		.ReDraw = False
         ggoSpread.SpreadLock      -1,-1,-1
		 ggoSpread.SSSetProtected	.MaxCols,-1,-1                        
		.ReDraw = True
    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected   C_AllowNm , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_AllowAmt , pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected   C_AllowNm2 , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_AllowAmt2 , pvStartRow, pvEndRow
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
			C_AllowNm    = iCurColumnPos(1)
			C_AllowAmt   = iCurColumnPos(2)
       Case "B"
            ggoSpread.Source = frm1.vspdData1S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_AllowNm1    = iCurColumnPos(1)
			C_AllowAmt1   = iCurColumnPos(2)			
       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_AllowNm2    = iCurColumnPos(1)
			C_AllowAmt2   = iCurColumnPos(2)
       Case "D"
            ggoSpread.Source = frm1.vspdData2S
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_AllowNm21    = iCurColumnPos(1)
			C_AllowAmt21   = iCurColumnPos(2)			
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
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call AppendNumberPlace("6","3","1")
    Call AppendNumberPlace("7","2","0")
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet("A")                                                            'Setup the Spread sheet
	Call InitSpreadSheet("B")
    Call InitSpreadSheet("C")                                                            'Setup the Spread sheet
	Call InitSpreadSheet("D")
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    

    Call SetDefaultVal
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
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
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables                                                           '⊙: Initializes local global variables

    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
       Exit Function
    End if

    Call MakeKeyStream("X")
	topleftOK = false
	
    If DbQuery = False Then  
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
       IntRetCD = DisplayMsgbox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
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
    Dim IntRetCd
	Dim strDate
	    
    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    

	strDate		= UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtyymm.Year,Right("0" & frm1.txtyymm.Month,2),"01")	

    IF  FuncAuthority("@", UniConvDateToYYYYMMDD(strDate,parent.gDateFormat,""), parent.gUsrID) = "N" THEN
        Call DisplayMsgbox("800304","X","X","X")         '연월차 마감처리된 지급월 입니다.
        Exit Function
    END IF
    
    If DbDelete = False Then
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
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgbox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
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
    
	With Frm1
	    
		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False
		
			ggoSpread.Source = frm1.vspdData	
			ggoSpread.CopyRow ,imRow
			SetSpreadColor frm1.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
			.vspdData.ReDraw = True
			.vspdData.focus
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
Function FncInsertRow() 
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow
       .vspdData.ReDraw = True
    End With
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
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

     if LayerShowHide(1) = false then
	    Exit Function
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    Call InitVariables														 '⊙: Initializes local global variables

     if LayerShowHide(1) = false then
	    Exit Function
	end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey1=" & lgStrPrevKey1             '☜: Next key tag
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
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub 
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD1" Then
		ggoSpread.Source = frm1.vspdData1S 
		Call ggoSpread.SaveSpreadColumnInf()
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD2" Then
		ggoSpread.Source = frm1.vspdData2S 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : m
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    If IsEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
    End If
    
    Select Case gActiveSpdSheet.id
		Case "vaSpread1"
			Call InitSpreadSheet("A")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData1S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("B")
		    ggoSpread.Source = frm1.vspdData1S 
	        Call ggoSpread.ReOrderingSpreadData()

'        	Call InitData("A")
		Case "vaSpread2"
			Call InitSpreadSheet("C")
            ggoSpread.Source = gActiveSpdSheet
        	Call ggoSpread.ReOrderingSpreadData()

		    ggoSpread.Source = frm1.vspdData2S 
            Call ggoSpread.RestoreSpreadInf()
            Call InitSpreadSheet("D")
		    ggoSpread.Source = frm1.vspdData2S 
	        Call ggoSpread.ReOrderingSpreadData()
	End Select 
	
	call InitData

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData1S.ColWidth(pvCol1) = frm1.vspdData.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData1s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    frm1.vspdData2S.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)
    ggoSpread.Source = frm1.vspdData2s
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

     if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgSpreadFlg="       & gSpreadFlg    
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
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	 if LayerShowHide(1) = false then
	    Exit Function
	end if
		
	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
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
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
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
	    Exit Function
	end if
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtYymm.focus 

	If frm1.txtRealProvAmt.text > 0 Then
		Call SetToolbar("1101000000011111")												'⊙: Set ToolBar
	Else	
		Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
	End If


    Call InitData()

    Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus

End Function

'========================================================================================================
' Function Name : DbQueryNo
' Function Desc : Called by MB Area when query operation is not successful
'========================================================================================================
Function DbQueryNo()
	
    Frm1.txtYymm.focus 

    Call InitData()

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Frm1.txtGlNo.value =  Frm1.txtLcNo.value  
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
	Call FncQuery()		
End Function

'========================================================================================================
' Function Name : FuncAuthority
' Function Desc : 시스템마감체크 
'========================================================================================================
Function FuncAuthority(Pay_gubun, Pay_yymmdd, Emp_no)

    Dim strRet
    Dim IntRetCD

    strRet = "N"    
    IntRetCD = CommonQueryRs("close_type, close_dt, emp_no","hda270t","org_cd=" & FilterVar("1", "''", "S") & "  and pay_gubun=" & FilterVar("Z", "''", "S") & "  and pay_type= " & FilterVar(Pay_gubun, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
		If iWhere = 0 Then
			.txtEmp_no.value   = arrRet(0)
			.txtName.value     = arrRet(1)
			.txtDept_nm.value  = arrRet(2)
			.txtRollPstn.value = arrRet(3)
			.txtPay_grd.value  = arrRet(4)
			.txtEntr_dt.text   = arrRet(5)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
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
	    Call CommonQueryRs("GROUP_ENTR_DT","HAA010T"," EMP_NO =  " & FilterVar(frm1.txtEmp_no.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    
	    strEntr_dt			= UNIDateClientFormat(strEntr_dt)
		srtGroup_entr_dt	= UNIDateClientFormat(Replace(lgF0, Chr(11), ""))
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000001111") 	    
	
    gMouseClickStatus = "SPC"  
    gSpreadFlg = 1	
    Set gActiveSpdSheet = frm1.vspdData
    ggoSpread.Source = frm1.vspdData        
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
Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000001111") 	    
	
    gMouseClickStatus = "SP2C"  	
	gSpreadFlg = 2
    Set gActiveSpdSheet = frm1.vspdData2
    ggoSpread.Source = frm1.vspdData2        
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData2.Row = Row       

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
Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
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
        frm1.vspdData1s.LeftCol=NewLeft
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

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2s.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

    ggoSpread.Source = frm1.vspdData1S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")

    ggoSpread.Source = frm1.vspdData2S
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")
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

<BODY SCROLL="no" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>연월차단독지급결과조회</font></td>
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
					<TD HEIGHT=30 WIDTH=100%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>정산년월</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="정산년도" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>	
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
								</TR>
							</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=14% valign=top>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
					        <TR>
					            <TD WIDTH="50%" valign=top>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>기본사항</LEGEND>
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
												<TD CLASS=TD5 NOWRAP>지급일</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtProv_dt CLASSID=<%=gCLSIDFPDT%> ALT="지급일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												</TD>
			                            	</TR>
						                </TABLE>
							        </FIELDSET>
            			        </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=30% valign=top>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
					        <TR>
					        <TD WIDTH=50% HEIGHT=40% valign=top>
                                <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>연차사항</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
	                                <TR>                        
	                                    <TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;연 차 발 생</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearSaveTot CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X7Z" ALT="연차발생"></OBJECT>');</SCRIPT>개</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>연 차 사 용</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearUse CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="연차사용"></OBJECT>');</SCRIPT>개</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>연 차 지 급</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="연차지급"></OBJECT>');</SCRIPT>개</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>기 준 금 액</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearBasAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2Z" ALT="연차기준금액"></OBJECT>');</SCRIPT>&nbsp;</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>연 차 수 당</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtYearAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2Z" ALT="연차수당"></OBJECT>');</SCRIPT>&nbsp;</TD>
            		        	    </TR>
	            	        		</TABLE>
				                </FIELDSET>
					        </TD>
					        <TD WIDTH=50% HEIGHT=40% valign=top>
                                <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>월차사항</LEGEND>
                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
					        	    <TR>  
	                                    <TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;월 차 발 생</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthSave CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X7Z" ALT="월차발생"></OBJECT>');</SCRIPT>개</TD>
                                   	</TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>월 차 사 용</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthUse CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="월차사용"></OBJECT>');</SCRIPT>개</TD>
	                               	</TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>월차 의무 사용</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthDutyCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="월차의무사용"></OBJECT>');</SCRIPT>개</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>월 차 지 급</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthCnt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X6Z" ALT="월차지급"></OBJECT>');</SCRIPT>개</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>기 준 금 액</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthBasAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2Z" ALT="월차기준금액"></OBJECT>');</SCRIPT>&nbsp;</TD>
            		        	    </TR>
	                                <TR>
	                                    <TD CLASS=TD5 NOWRAP>월 차 수 당</TD>
	                                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtMonthAmt CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X2Z" ALT="월차수당"></OBJECT>');</SCRIPT>&nbsp;</TD>
            		        	    </TR>
	            	        		</TABLE>
				                </FIELDSET>
					        </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=15% valign=top>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
					        <TR>
					            <TD WIDTH="50%" valign=top>
                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>지급액</LEGEND>
                                        <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
	                                    <TR>
							            	<TD CLASS=TD5>소  득  세</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtIncomeTaxAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="소득세" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5>국외근로비과세</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtNon_tax5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="국외근로비과세" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
            			                </TR>
	                                    <TR>
							            	<TD CLASS=TD5>주  민  세</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtResTaxAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="주민세" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5>과 세 금 액</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="과세금액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
            			                </TR>
	                                    <TR>
							            	<TD CLASS=TD5>고 용 보 험</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtEmpInsurAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="고용보험" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5>총 지 급 액</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="총지급액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
            			                </TR>
	                                    <TR>
							            	<TD CLASS=TD5>공 제 총 액</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSubTotAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="공제총액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
							            	<TD CLASS=TD5>실 지 급 액</TD>
							            	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRealProvAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="실지급액" tag="24X2Z" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;</TD>
            			                </TR>
						                </TABLE>
							        </FIELDSET>
            			        </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=* valign=top>
                         <TABLE <%=LR_SPACE_TYPE_60%>>
								<TR HEIGHT=80>
								        <TD WIDTH="50%" >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
									<TD WIDTH="50%" >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR HEIGHT=20>
									<TD WIDTH="50%" >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1S WIDTH=100% HEIGHT=100% tag="43" TITLE="SPREAD" id=vaSpread1S> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
									<TD WIDTH="50%" >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2S WIDTH=100% HEIGHT=100% tag="43" TITLE="SPREAD" id=vaSpread2S> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
					     </TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize  framespacing=0></IFRAME> 

		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

