
<%@ LANGUAGE="VBSCRIPT" %>

<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        : 
*  3. Program ID           : GA003MA1
*  4. Program Name         : 경영손익항목등록 
*  5. Program Desc         : 경영손익항목등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/18
*  8. Modified date(Last)  : 2001/12/18
*  9. Modifier (First)     : Lee Kang Yeong
* 10. Modifier (Last)      : Lee Tae Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID =  "GA003mb1.asp"												'Biz Logic ASP 

Dim C_ACCT_CD  
Dim C_ACCT_POP  
Dim C_ACCT_NM  																'Spread Sheet의 Column별 상수 
Dim C_GAIN_CD  
Dim C_GAIN_POP  
Dim C_GAIN_NM  
Dim C_GAIN_GRP  
Dim C_ACCT_TYPE  
Dim C_ACCT_TYPE_NM  
Dim C_DA_METHOD  
Dim C_DA_METHOD_NM  
Dim C_GAIN_YN  

'Const C_SHEETMAXROWS    = 100													'한 화면에 보여지는 최대갯수*1.5%>
'Const C_SHEETMAXROWS_D  = 100													'☆: Server에서 한번에 fetch할 최대 데이타 건수 

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lsConcd
Dim IsOpenPop          

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE												'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False													'⊙: Indicates that no value changed
	lgIntGrpCount     = 0														'⊙: Initializes Group View Size
    lgStrPrevKey      = ""														'⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""														'⊙: initializes Previous Key Index
    lgSortKey         = 1														'⊙: initializes sort direction
		
End Sub

'========================================================================================================
Sub SetDefaultVal()
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "G", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
Sub MakeKeyStream(pRow)
End Sub        

'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx 
    
    iCodeArr = "V" & vbTab & "F" & vbTab & "T"
    iNameArr = "변동비" & vbTab & "고정비" & vbTab & "관세환급"

    '미만/이하(이상,이하,미만,초과)
    ggoSpread.SetCombo iCodeArr, C_ACCT_TYPE
    ggoSpread.SetCombo iNameArr, C_ACCT_TYPE_NM
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("g1004", "''", "S") & " and (MINOR_CD = " & FilterVar("5", "''", "S") & " OR MINOR_CD = " & FilterVar("6", "''", "S") & ") ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
       iCodeArr = "" & vbTab & lgF0
       iNameArr = "" & vbTab & lgF1
    
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DA_METHOD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DA_METHOD_NM
  
    iNameArr = "" & vbTab & "Y" & vbTab & "N"
    ggoSpread.SetCombo iNameArr, C_GAIN_YN

End Sub

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_ACCT_TYPE
			intIndex = .value
			.col = C_ACCT_TYPE_NM
			.value = intindex	
			
			.Row = intRow
			.Col = C_DA_METHOD
			intIndex = .value
			.col = C_DA_METHOD_NM
			.value = intindex					
		Next	
	End With
End Sub

'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_ACCT_TYPE_NM
                .Col = Col
                intIndex = .Value												'COMBO의 VALUE값 
				.Col = C_ACCT_TYPE												'CODE값란으로 이동 
				.Value = intIndex												'CODE란의 값은 COMBO의 VALUE값이된다.
				
		    Case C_DA_METHOD_NM
                .Col = Col
                intIndex = .Value               
		        .Col = C_DA_METHOD                  
		        .Value = intIndex      
		End Select
	End With	

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_ACCT_CD		= 1
	C_ACCT_POP		= 2
	C_ACCT_NM		= 3																'Spread Sheet의 Column별 상수 
	C_GAIN_CD		= 4
	C_GAIN_POP		= 5
	C_GAIN_NM		= 6
	C_GAIN_GRP		= 7
	C_ACCT_TYPE		= 8
	C_ACCT_TYPE_NM	= 9
	C_DA_METHOD		= 10
	C_DA_METHOD_NM	= 11
	C_GAIN_YN		= 12
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
       .MaxCols = C_GAIN_YN + 1                                                 ' ☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                          ' ☜: Hide maxcols
       .ColHidden = True                                                            
    
        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021127", ,parent.gAllowDragDropSpread
		
		ggoSpread.ClearSpreadData

	   .ReDraw = false
		
       Call GetSpreadColumnPos("A")			   
	   

       ggoSpread.SSSetEdit    C_ACCT_CD    , "계정코드", 15 ,,,20,2
       ggoSpread.SSSetButton  C_ACCT_POP    
       ggoSpread.SSSetEdit    C_ACCT_NM    , "계정코드명", 20,,,40
       ggoSpread.SSSetEdit    C_GAIN_CD    , "손익항목코드", 15,,,4,2
       ggoSpread.SSSetButton  C_GAIN_POP    
       ggoSpread.SSSetEdit    C_GAIN_NM    , "손익항목", 18,,,30     
       ggoSpread.SSSetEdit   C_GAIN_GRP     , "대분류명" , 18,,,30
       ggoSpread.SSSetCombo   C_ACCT_TYPE       , "변/고정비코드", 15 , 0 
       ggoSpread.SSSetCombo   C_ACCT_TYPE_NM       , "변/고정비", 12 , 0
       ggoSpread.SSSetCombo   C_DA_METHOD       , "직과코드", 5 , 0 
       ggoSpread.SSSetCombo   C_DA_METHOD_NM       , "직과", 12 , 0
       ggoSpread.SSSetCombo   C_GAIN_YN    , "사용여부", 10 , 0 
       
       Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_POP)
       Call ggoSpread.MakePairsColumn(C_GAIN_CD,C_GAIN_POP)
              
       Call ggoSpread.SSSetColHidden(C_ACCT_TYPE,C_ACCT_TYPE,True)
	   Call ggoSpread.SSSetColHidden(C_DA_METHOD,C_DA_METHOD,True)			

       .ReDraw = true
	
       Call SetSpreadLock
    End With
End Sub

'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
     .vspdData.ReDraw = False
      ggoSpread.SpreadLock      C_ACCT_CD , -1, C_ACCT_CD
      ggoSpread.SpreadLock      C_ACCT_NM , -1, C_ACCT_NM
      ggoSpread.SSSetRequired   C_GAIN_CD , -1, C_GAIN_CD
      ggoSpread.SpreadLock      C_GAIN_NM , -1, C_GAIN_NM
      ggoSpread.SpreadLock	    C_GAIN_GRP  , -1, -1
      ggoSpread.SSSetRequired	C_ACCT_TYPE_NM, -1, -1
      ggoSpread.SSSetRequired	C_GAIN_YN, -1, -1
      ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
     .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
     .vspdData.ReDraw = False
      ggoSpread.SSSetRequired    C_ACCT_CD		,pvStartRow		,pvEndRow
      ggoSpread.SSSetProtected	C_ACCT_NM		,pvStartRow		,pvEndRow
      ggoSpread.SSSetRequired    C_GAIN_CD		,pvStartRow		,pvEndRow
      ggoSpread.SSSetProtected   C_GAIN_NM		,pvStartRow		,pvEndRow
      ggoSpread.SSSetProtected   C_GAIN_GRP		,pvStartRow		,pvEndRow
      ggoSpread.SSSetRequired    C_ACCT_TYPE_NM	,pvStartRow		,pvEndRow
      ggoSpread.SSSetRequired	C_GAIN_YN		,pvStartRow		,pvEndRow
     .vspdData.ReDraw = True
    End With
End Sub

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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ACCT_CD			= iCurColumnPos(1)
			C_ACCT_POP			= iCurColumnPos(2)
			C_ACCT_NM 			= iCurColumnPos(3)    
			C_GAIN_CD			= iCurColumnPos(4)
			C_GAIN_POP 			= iCurColumnPos(5)
			C_GAIN_NM 			= iCurColumnPos(6)
			C_GAIN_GRP  		= iCurColumnPos(7)
			C_ACCT_TYPE			= iCurColumnPos(8)
			C_ACCT_TYPE_NM     	= iCurColumnPos(9)    
			C_DA_METHOD  		= iCurColumnPos(10)
			C_DA_METHOD_NM   	= iCurColumnPos(11)
			C_GAIN_YN   		= iCurColumnPos(12)
    End Select    
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData
End Sub

'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
  
    Call InitSpreadSheet															'Setup the Spread sheet
    Call InitVariables																'Initializes local global variables
    
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 

	Call CookiePage (0)                                                             '☜: Check Cookie
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call SetDefaultVal
'    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    If DbQuery = False Then
        Exit Function
    End If
              
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

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

    If DbSave = False Then
        Exit Function
    End If
            
    FncSave = True                                                              '☜: Processing is OK
    
End Function

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
			SetSpreadColor .ActiveRow , .ActiveRow
            .Col = C_ACCT_NM
            .Text = ""
                                   
            .ReDraw = True
		    .Focus
		    .Action = 0 ' go to 
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

	With Frm1.VspdData
           .Col  = C_ACCT_CD
           .Row  = .ActiveRow
           .Text = ""
    End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo
    Call  initData()  
End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
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
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function

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
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
Function FncPrev() 
    On Error Resume Next													'☜: Protect system from crashing
End Function

'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
Function FncFind() 
'    Call parent.FncFind(Parent.C_MULTI, False)                                   '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")					'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                     '☜: Clear err status

	if LayerShowHide(1) = False then
	   Exit Function
	end if
	
	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
		strVal = strVal     & "&txtKeyStream="       & lgKeyStream                '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex          '☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)     '☜: Max fetched data at a time
    End With
    
  		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
	Call RunMyBizASP(MyBizASP, strVal)                                            '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

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
    Dim iColSep
    Dim iRowSep     
	
    DbSave = False   
                                                           
    if LayerShowHide(1) = False then
	   Exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
           
               Case ggoSpread.InsertFlag											'☜: Update추가 
													strVal = strVal & "C" & iColSep
													strVal = strVal & lRow & iColSep
                  
                    .vspdData.Col = C_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_GAIN_CD		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_GAIN_YN		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ACCT_TYPE		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_DA_METHOD		: strVal = strVal & Trim(.vspdData.Text) & iRowSep
                     lGrpCnt = lGrpCnt + 1
            
               Case ggoSpread.UpdateFlag											'☜: Update
													strVal = strVal & "U" & iColSep
													strVal = strVal & lRow & iColSep
                   
                   .vspdData.Col = C_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                   .vspdData.Col = C_GAIN_CD		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                   .vspdData.Col = C_GAIN_YN  		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                   .vspdData.Col = C_ACCT_TYPE		: strVal = strVal & Trim(.vspdData.Text) & iColSep
                   .vspdData.Col = C_DA_METHOD		: strVal = strVal & Trim(.vspdData.Text) & iRowSep
                    lGrpCnt = lGrpCnt + 1
        
               Case ggoSpread.DeleteFlag											'☜: Delete
                                                  strDel = strDel & "D" & iColSep
                                                  strDel = strDel & lRow & iColSep
                  
                   .vspdData.Col = C_ACCT_CD   : strDel = strDel & Trim(.vspdData.Text) & iRowSep	'삭제시 key만								
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
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False																'⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then												'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")									'☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")							'⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then															'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    Call DbDelete																	'☜: Delete db data
    
    FncDelete = True																'⊙: Processing is OK


End Function

'========================================================================================================
Function DbQueryOk()													     
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")											'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	
End Function

'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")											'⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    
    Call InitVariables																'⊙: Initializes local global variables
	call MainQuery()
End Function

'========================================================================================================
Function DbDeleteOk()

End Function

'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)  
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtAllow_cd.value = arrRet(0)
		        .txtAllow_nm.value = arrRet(1)		
'		 	
        End Select
	End With
End Sub

'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
	    Case C_GAIN_POP
	        arrParam(0) = "손익항목팝업"											' 팝업 명칭 
	    	arrParam(1) = "b_minor a, b_minor b, g_gain c"								' TABLE 명칭 
	    	arrParam(2) = strCode                   									' Code Condition
	    	arrParam(3) = ""															' Name Cindition
	    	arrParam(4) = " a.major_cd = " & FilterVar("g1006", "''", "S") & " and b.major_cd = " & FilterVar("g1005", "''", "S") & "  and c.gain_group = b.minor_cd and c.gain_cd = a.minor_cd and (c.gain_cd not in (" & FilterVar("A01", "''", "S") & "," & FilterVar("B01", "''", "S") & "," & FilterVar("Q01", "''", "S") & "," & FilterVar("Q02", "''", "S") & ") and not (c.gain_cd LIKE " & FilterVar("C%", "''", "S") & " and c.gain_cd <> " & FilterVar("C04", "''", "S") & ")) "                       ' Where Condition

	    	arrParam(5) = "손익항목" 												' TextBox 명칭 
	
	    	arrField(0) = "a.minor_cd"													' Field명(0)
	    	arrField(1) = "a.minor_nm"    												' Field명(1)
	    	arrField(2) = "b.minor_nm"    												' Field명(2)
    
	    	arrHeader(0) = "손익항목"	   		    								' Header명(0)
	    	arrHeader(1) = "손익항목명"	    										' Header명(1)
	    	arrHeader(2) = "대분류명"	    										' Header명(2)
	    	
	    Case C_ACCT_POP
	    
			ggoSpread.Source = frm1.vspdData                                   
	
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			If frm1.vspdData.Text <> ggoSpread.InsertFlag Then 
				IsOpenPop = False
				Exit Function  
			end if  

	        arrParam(0) = "계정코드팝업"											' 팝업 명칭 
	    	arrParam(1) = "A_ACCT"										' TABLE 명칭 
	    	arrParam(2) = strCode                   									' Code Condition
	    	arrParam(3) = ""															' Name Cindition
	    	arrParam(4) = "DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND temp_fg_3 LIKE " & FilterVar("G%", "''", "S") & ""               ' Where Condition

	    	arrParam(5) = "계정코드" 												' TextBox 명칭 
	
	    	arrField(0) = "ACCT_CD"													' Field명(0)
	    	arrField(1) = "ACCT_NM"    												' Field명(1)
    
	    	arrHeader(0) = "계정코드"	   		    								' Header명(0)
	    	arrHeader(1) = "계정코드명"	    										' Header명(1)	
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

'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_GAIN_POP
		        .vspdData.Col = C_GAIN_CD
		    	.vspdData.text = arrRet(0) 
		        .vspdData.Col = C_GAIN_NM
		    	.vspdData.text = arrRet(1)	
		    	.vspdData.Col = C_GAIN_GRP
		    	.vspdData.text = arrRet(2)    	 
		    Case C_ACCT_POP
		        .vspdData.Col = C_ACCT_CD
		    	.vspdData.text = arrRet(0) 
		        .vspdData.Col = C_ACCT_NM
		    	.vspdData.text = arrRet(1)	
        End Select

	End With

End Function

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD, EFlag
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
   	
   	EFlag = False

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col        	
    	    Case C_ACCT_CD
				IntRetCD= CommonQueryRs(" acct_nm ","a_acct"," ACCT_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
				    Call DisplayMsgBox("110100","X","X","X")                         '☜ : 계정코드가 존재하지않습니다 
				    frm1.vspdData.Text=""
				    frm1.vspdData.Col = C_ACCT_NM
				    frm1.vspdData.Text=""
				    frm1.vspdData.Col = Col
				    frm1.vspdData.Action=0
				    Set gActiveElement = document.activeElement  
					EFlag = True
				Else
    				frm1.vspdData.Col = C_ACCT_NM
					frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))

    			    frm1.vspdData.Col = C_ACCT_CD
    			    
    			    IntRetCD= CommonQueryRs(" ACCT_CD "," G_ACCT "," ACCT_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    			    If IntRetCD=True And Trim(frm1.vspdData.Text)<>"" Then
    					Call DisplayMsgBox("800446","X","X","X")
						frm1.vspdData.Col = C_ACCT_CD
						frm1.vspdData.Text=""
						frm1.vspdData.Col = C_ACCT_NM
						frm1.vspdData.Text=""
						frm1.vspdData.Col = Col
						frm1.vspdData.Action=0
						Set gActiveElement = document.activeElement  
						EFlag = True
					END IF	
				End If
            	      
            Case C_GAIN_CD
				IntRetCD= CommonQueryRs(" a.minor_nm, B.MINOR_NM "," b_minor a, b_minor b, g_gain c ","  a.major_cd = " & FilterVar("g1006", "''", "S") & " and b.major_cd = " & FilterVar("g1005", "''", "S") & "  and c.gain_group = b.minor_cd and (c.gain_cd not in (" & FilterVar("A01", "''", "S") & "," & FilterVar("B01", "''", "S") & "," & FilterVar("Q01", "''", "S") & "," & FilterVar("Q02", "''", "S") & ") and not (c.gain_cd LIKE " & FilterVar("C%", "''", "S") & " and c.gain_cd <> " & FilterVar("C04", "''", "S") & ")) and c.gain_cd = a.minor_cd and A.MINOR_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
				    Call DisplayMsgBox("GB0301","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
				    frm1.vspdData.Text=""
				    frm1.vspdData.Col = C_GAIN_NM
				    frm1.vspdData.Text=""
				    frm1.vspdData.Col = C_GAIN_GRP
				    frm1.vspdData.Text=""
				    frm1.vspdData.Col = Col
				    frm1.vspdData.Action=0
				    Set gActiveElement = document.activeElement  
					EFlag = True
				Else
    			    frm1.vspdData.Col = C_GAIN_NM
				    frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
				    frm1.vspdData.Col = C_GAIN_GRP
				    frm1.vspdData.Text=Trim(Replace(lgF1,Chr(11),""))
				End If
            
    End Select    

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
     '데이터 확인시 틀린데이터에 대해 undo 해준다.    
             
	Call CheckMinNumSpread(frm1.vspdData,Col,Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0
        
    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()				
	End If
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
     
     Call SetPopupMenuItemInf("1101111111")
     gMouseClickStatus = "SPC" 
     Set gActiveSpdSheet = frm1.vspdData
     
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
 
	Select Case Col
	    Case C_GAIN_POP
            Call OpenCode(Trim(frm1.vspdData.text), C_GAIN_POP, Row)
        Case C_ACCT_POP
            Call OpenCode(Trim(frm1.vspdData.text), C_ACCT_POP, Row)    
    End Select
		Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")      
End Sub

'========================================================================================================
Sub txtallow_cd_OnChange()
    Dim iDx
    Dim IntRetCd   
    
    IF frm1.txtallow_cd.value<>"" THEN
        IntRetCd = CommonQueryRs(" allow_nm "," HDA010T "," allow_cd =  " & FilterVar(frm1.txtallow_cd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false  Then
            Call DisplayMsgBox("800145","X","X","X")								'수당정보에 등록되지않은 코드입니다.
            frm1.txtallow_cd.value=""
            frm1.txtallow_nm.value=""
            frm1.txtAllow_cd.focus
        ELSE   '수당코드 
            frm1.txtallow_nm.value=Trim(Replace(lgF0,Chr(11),""))
        END IF
    ELSE
        frm1.txtallow_nm.value=""
    END IF  
End Sub 

'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End if
End Sub

</SCRIPT>
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>경영손익항목등록</font></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
		
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
					
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%" WIDTH=100%>
									<script language =javascript src='./js/ga003ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
					
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1"></IFRAME></TD>
	</TR>
		
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24" TABINDEX = "-1">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
	
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</BODY>
</HTML>
