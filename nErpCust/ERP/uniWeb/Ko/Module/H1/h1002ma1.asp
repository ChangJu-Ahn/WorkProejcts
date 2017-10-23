<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : H1002ma1
*  4. Program Name         : H1002ma1
*  5. Program Desc         : 근태코드등록
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/27
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : chcho 
* 10. Modifier (Last)      : Lee Sina
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "h1002mb1.asp"                                      '비지니스 로직 ASP명 
Const CookieSplit = 1233

Dim IsOpenPop  

Dim C_DILIG_CD										'Spread Sheet의 Column별 상수 
Dim C_DILIG_NM         
Dim C_DILIG_TYPE
Dim C_DILIG_TYPE_NM    
Dim C_DAY_TIME
Dim C_DAY_TIME_NM      
Dim C_BAS_MARGIR       
Dim C_WK_DAY										'Spread Sheet의 Column별 상수 
Dim C_ATTEND_DAY                      
Dim C_WEEK_CNT_APPLY   
Dim C_HOLIDAY_APPLY    
Dim C_SYS_FLAG         
Dim C_DILIG_SEQ

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
'========================================================================================================
' Name : InitSpreadPosVariables() 
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()
	C_DILIG_CD		 = 1         
	C_DILIG_NM       = 2   
	C_DILIG_TYPE     = 3   
	C_DILIG_TYPE_NM  = 4   
	C_DAY_TIME       = 5   
	C_DAY_TIME_NM    = 6  
	C_BAS_MARGIR     = 7  
	C_WK_DAY         = 8   
	C_ATTEND_DAY     = 9                 
	C_WEEK_CNT_APPLY = 10  
	C_HOLIDAY_APPLY  = 11  
	C_SYS_FLAG       = 12  
	C_DILIG_SEQ      = 13
End Sub
'========================================================================================================
' Name : InitVariables() 
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						 '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False									 '⊙: Indicates that no value changed
	lgIntGrpCount     = 0                                       '⊙: Initializes Group View Size
	lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
	lgSortKey         = 1                                       '⊙: initializes sort direction
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
' Name : CookiePage()
' Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    Call txtDilig_cd_OnChange()
    
    Dim rbo_sort
    IF frm1.rbo_sort(0).checked THEN
        rbo_sort="1"
    ELSEIF frm1.rbo_sort(1).checked THEN
        rbo_sort="2"
    ELSE
        rbo_sort="3"
    END IF
    
    lgKeyStream   = Frm1.txtDilig_cd.Value & Parent.gColSep 
    lgKeyStream   = lgKeyStream & Frm1.CboDilig_type.Value & Parent.gColSep        
    lgKeyStream   = lgKeyStream & rbo_sort & Parent.gColSep                   'You Must append one character( Parent.gColSep)
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()    
    Dim    iCodeArr
    Dim    iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0085", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call  SetCombo2(frm1.CboDilig_type,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서 
End Sub

Sub InitComboBox2()
    Dim    iCodeArr
    Dim    iNameArr
    
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0085", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DILIG_TYPE
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DILIG_TYPE_NM

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0086", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
     ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DAY_TIME
     ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DAY_TIME_NM
  
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_BAS_MARGIR
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_WK_DAY
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_ATTEND_DAY
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_WEEK_CNT_APPLY
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_HOLIDAY_APPLY
     ggoSpread.SetCombo "YES" & vbtab & "NO" , C_SYS_FLAG    
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
 
    With frm1.vspdData
		.Row = Row
		Select Case Col
		    Case C_DILIG_TYPE_NM
		        .Col = Col
		        intIndex = .Value 
				.Col = C_DILIG_TYPE
				.Value = intIndex
		  Case C_DAY_TIME_NM
		        .Col = Col
		        intIndex = .Value   
				.Col = C_DAY_TIME  
				.Value = intIndex 
		End Select
    End With

     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
	With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021128",,parent.gAllowDragDropSpread    

	    .ReDraw = false

        .MaxCols = C_DILIG_SEQ + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True

        .MaxRows = 0
        ggoSpread.ClearSpreadData

		Call  GetSpreadColumnPos("A")      

         ggoSpread.SSSetEdit  C_DILIG_CD         , "근태코드", 10,,, 2,2
         ggoSpread.SSSetEdit  C_DILIG_NM         , "근태코드명", 12,,, 20,2    
         ggoSpread.SSSetCombo C_DILIG_TYPE       , "근태구분CD" , 10, 0
         ggoSpread.SSSetCombo C_DILIG_TYPE_NM    , "근태구분" , 12, 0
         ggoSpread.SSSetCombo C_DAY_TIME         , "일수/시CD간", 10, 0 
         ggoSpread.SSSetCombo C_DAY_TIME_NM      , "일수/시간", 11, 0      
         ggoSpread.SSSetCombo C_BAS_MARGIR       , "기본급차감여부", 15, 0
         ggoSpread.SSSetCombo C_WK_DAY           , "근무일수포함여부",  17, 0 
         ggoSpread.SSSetCombo C_ATTEND_DAY       , "출근일수포함여부" , 17, 0
         ggoSpread.SSSetCombo C_WEEK_CNT_APPLY   , "주차적용여부",  14, 0
         ggoSpread.SSSetCombo C_HOLIDAY_APPLY    , "휴일적용여부",  14, 0
         ggoSpread.SSSetCombo C_SYS_FLAG         , "시스템코드여부",  14, 0
         ggoSpread.SSSetFloat C_DILIG_SEQ		, "순번",7,"6",ggStrIntegeralPart,ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,,"0","99"          
         
		Call ggoSpread.SSSetColHidden(C_DILIG_TYPE,C_DILIG_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_DAY_TIME,  C_DAY_TIME, True)
		
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

         ggoSpread.SpreadLock    C_DILIG_CD,		-1, C_DILIG_CD
         ggoSpread.SSSetRequired C_DILIG_NM,		-1, C_DILIG_NM
         ggoSpread.SSSetRequired C_DILIG_TYPE_NM,	-1, C_DILIG_TYPE_NM
         ggoSpread.SSSetRequired C_DAY_TIME_NM,		-1, C_DAY_TIME_NM
         ggoSpread.SSSetRequired C_BAS_MARGIR,		-1, C_BAS_MARGIR
         ggoSpread.SSSetRequired C_WK_DAY,			-1, C_WK_DAY	
         ggoSpread.SSSetRequired C_ATTEND_DAY,		-1, C_ATTEND_DAY
         ggoSpread.SSSetRequired C_WEEK_CNT_APPLY,	-1, C_WEEK_CNT_APPLY	
         ggoSpread.SSSetRequired C_HOLIDAY_APPLY,	-1, C_HOLIDAY_APPLY
         ggoSpread.SSSetRequired C_SYS_FLAG,			-1, C_SYS_FLAG
         ggoSpread.SSSetRequired C_DILIG_SEQ,		-1, C_DILIG_SEQ        
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
          ggoSpread.SSSetRequired  C_DILIG_CD		, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_DILIG_NM		, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_DILIG_TYPE_NM	, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_DAY_TIME_NM	, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_BAS_MARGIR		, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_WK_DAY			, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_ATTEND_DAY		, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_WEEK_CNT_APPLY	, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_HOLIDAY_APPLY	, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_SYS_FLAG		, pvStartRow, pvEndRow
          ggoSpread.SSSetRequired  C_DILIG_SEQ		, pvStartRow, pvEndRow
          ggoSpread.SSSetProtected	.vspdData.MaxCols, pvStartRow, pvEndRow
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
    iPosArr = Split(iPosArr, Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  Parent.UC_PROTECTED Then
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
                
            C_DILIG_CD		 = iCurColumnPos(1)         
			C_DILIG_NM       = iCurColumnPos(2)   
			C_DILIG_TYPE     = iCurColumnPos(3)   
			C_DILIG_TYPE_NM  = iCurColumnPos(4)   
			C_DAY_TIME       = iCurColumnPos(5)   
			C_DAY_TIME_NM    = iCurColumnPos(6)  
			C_BAS_MARGIR     = iCurColumnPos(7)  
			C_WK_DAY         = iCurColumnPos(8)   
			C_ATTEND_DAY     = iCurColumnPos(9)                 
			C_WEEK_CNT_APPLY = iCurColumnPos(10)  
			C_HOLIDAY_APPLY  = iCurColumnPos(11)  
			C_SYS_FLAG       = iCurColumnPos(12)  
			C_DILIG_SEQ      = iCurColumnPos(13)            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
	    
	Call  ggoOper.FormatField(Document, "1",ggStrIntegeralPart,ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")           '⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitComboBox
    Call InitComboBox2
    
    Call SetToolbar("1100000000001111")                  '버튼 툴바 제어 
    
    frm1.txtDilig_cd.focus 
    Call CookiePage(0) 
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
    FncQuery = False 
    Err.Clear

    ggoSpread.Source = frm1.vspdData
	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO,"x","x")      '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
   
	ggoSpread.ClearSpreadData
                 
    If Not chkField(Document, "1") Then
       Exit Function
    End If
  
	Call txtDilig_nm_Onchange()
    Call InitVariables
    Call MakeKeyStream("X")
    
    Call  DisableToolBar( Parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreTooBar()
        Exit Function
    End If
    FncQuery = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    FncSave = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

	Call  DisableToolBar(Parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
		Exit Function
	End If               
    FncSave = True
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
            .Col = C_DILIG_CD
            .Text = ""
            .Col = C_DILIG_NM
            .Text = ""
                                   
            .ReDraw = True
            .Col = C_DILIG_CD
			.Focus
			.Action = 0 ' go to 
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
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim imRow, iRow
	
    On Error Resume Next
    Err.Clear
 
    FncInsertRow = False 

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
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
			.vspdData.Row  = iRow		
            .vspdData.Col = C_BAS_MARGIR             
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If
                      
           .vspdData.Col = C_WK_DAY             
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If
                      
           .vspdData.Col = C_ATTEND_DAY
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If
           
           .vspdData.Col = C_WEEK_CNT_APPLY
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If
           
           .vspdData.Col = C_SYS_FLAG
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If
           
           .vspdData.Col = C_HOLIDAY_APPLY
           If .vspdData.Text = "" Then
              .vspdData.Text = "NO" 
           End If 
       Next 
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True
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
		.Row = .ActiveRow
		.Col = C_SYS_FLAG
		If .Text = "YES" Then
			Call  DisplayMsgBox("900031", "X", "X", "X")
		 Exit Function
		End If
     
		ggoSpread.Source = frm1.vspdData 
		lDelRows =  ggoSpread.DeleteRow
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
    Call parent.FncExport( Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO,"x","x")   '⊙: Data is changed.  Do you want to exit? 
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
    Call InitComboBox2
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal
    DbQuery = False
    Err.Clear

	if LayerShowHide(1) = false then
		exit Function
	end if
 
 	With Frm1
	 	strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001               
	    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
	    strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
	    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
	End With
	 
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
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
    Dim intCnt 
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                                        
                    .vspdData.Col = C_DILIG_CD          : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_NM         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_TYPE     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DAY_TIME         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_BAS_MARGIR     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_WK_DAY         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_ATTEND_DAY     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_WEEK_CNT_APPLY : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HOLIDAY_APPLY   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SYS_FLAG       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_SEQ         : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                                               
                    .vspdData.Col = C_DILIG_CD   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_NM         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_TYPE     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DAY_TIME         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_BAS_MARGIR     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_WK_DAY         : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_ATTEND_DAY     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_WEEK_CNT_APPLY : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_HOLIDAY_APPLY   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_SYS_FLAG       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_DILIG_SEQ         : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1


               Case  ggoSpread.DeleteFlag                                      '☜: Delete
' 메세지처리 2007.04.20  900020 이 데이타를 참조하고 있는 데이타가 있어서 삭제가 불가능합니다.
	      .vspdData.Col = C_DILIG_CD  

     	      If CommonQueryRs(" COUNT(*) "," hca060t", " dilig_cd = " & FilterVar( Trim(.vspdData.Text) , "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		intCnt = CInt(Replace(lgF0, Chr(11), ""))
	      end if

	      if intCnt > 0 then 
	 	Call LayerShowHide(0)
  		Call DisplayMsgbox("900020","X","X","X") 
       		Exit function
   	      end if

     	      If CommonQueryRs(" COUNT(*) "," hca070t", " dilig_cd = " & FilterVar( Trim(.vspdData.Text) , "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true Then
		intCnt = CInt(Replace(lgF0, Chr(11), ""))
	      end if

	      if intCnt > 0 then 
	 	Call LayerShowHide(0)
  		Call DisplayMsgbox("900020","X","X","X") 
       		Exit function
   	      end if

                   strDel = strDel & "D" & Parent.gColSep
                   strDel = strDel & lRow & Parent.gColSep
                   strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep        
                   lGrpCnt = lGrpCnt + 1

           End Select
       Next
		.txtMode.value        =  Parent.UID_M0002
		.txtUpdtUserId.value  =  Parent.gUsrID
		.txtInsrtUserId.value =  Parent.gUsrID
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
    FncDelete = False
    If lgIntFlgMode <>  Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
	IntRetCD =  DisplayMsgBox("900003",  Parent.VB_YES_NO,"X","X")              '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function 
	End If
    
	Call  DisableToolBar( Parent.TBC_DELETE)
	If DbQuery = False Then
		Call  RestoreToolBar()
	Exit Function
	End If             
    
    FncDelete = True
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()                  
	lgIntFlgMode =  Parent.OPMD_UMODE    
	Call  ggoOper.LockField(Document, "Q")          '⊙: Lock field
	Call SetToolbar("110011110011111")    
	frm1.vspdData.focus								
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	    
	Call InitVariables               '⊙: Initializes local global variables
	Call  DisableToolBar( Parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If                '☜: Query db data
End Function

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
			arrParam(0) = "근태코드 팝업"
			arrParam(1) = "HCA010T"
			arrParam(2) = frm1.txtdilig_cd.value
			arrParam(3) = ""						            ' Name Cindition
			arrParam(4) = ""                                    ' Where Condition
			arrParam(5) = "근태코드"                        ' TextBox 명칭 
				 
			arrField(0) = "dilig_cd"                     ' Field명(0)
			arrField(1) = "dilig_nm"                     ' Field명(1)
				    
			arrHeader(0) = "근태코드"                 ' Header명(0)
			arrHeader(1) = "근태코드명"                 ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
	     Frm1.txtdilig_cd.focus
		 Exit Function
	Else
		 Call SetCondArea(arrRet,iWhere)
	End If 
 
End Function

'======================================================================================================
' Name : SetCondArea()           
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCondArea(Byval arrRet, Byval iWhere) 
	With Frm1
		Select Case iWhere
			Case "1"
			    .txtdilig_cd.value = arrRet(0)
			    .txtdilig_nm.value = arrRet(1)
			    .txtdilig_cd.focus
		End Select
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
'   Event Name : vspdData_Change 
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    Select Case Col                                                                 
         Case C_DILIG_CD                                                         '근태코드 
                IntRetCD =  CommonQueryRs(" DILIG_CD "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
                If IntRetCD=True And Trim(frm1.vspdData.Text)<>""  Then
                    Call  DisplayMsgBox("800095","X","X","X")                         '☜ : 입력된자료가 있습니다.
                    frm1.vspdData.Text=""
                Else
	                frm1.vspdData.Col = C_DILIG_CD
                    frm1.vspdData.Text=Trim(frm1.vspdData.Text)
                End If
   End Select    

    If Frm1.vspdData.CellType =  Parent.SS_CELL_TYPE_FLOAT Then
	      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
			 Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		  End If
	End If
 
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick 
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
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
		Call GetSpreadColumnPos("A")
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
'   Event Name : txtDilig_cd_OnChange
'   Event Desc :
'========================================================================================================
Function txtDilig_cd_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtDilig_cd.value = "" Then
        frm1.txtDilig_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
        Else
            frm1.txtDilig_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function
'========================================================================================================
'   Event Name : txtDilig_nm_Onchange
'   Event Desc :
'========================================================================================================
Sub txtDilig_nm_Onchange()
    Dim IntRetCd

    If  frm1.txtDilig_cd.value = "" Then
		frm1.txtDilig_nm.value = ""        
    Else    
        IntRetCd =  CommonQueryRs(" DILIG_NM "," HCA010T "," DILIG_CD =  " & FilterVar(frm1.txtDilig_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			frm1.txtDilig_nm.value = ""
        Else
			frm1.txtDilig_nm.value = Trim(Replace(lgF0,Chr(11),""))   			
        End if 
    End if    
End Sub


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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>근태코드등록</font></td>
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
  <TD width=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
       <TR>
           <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
       </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
      <TABLE <%=LR_SPACE_TYPE_40%>>
       <TR>
        <TD CLASS="TD5" NOWRAP>근태코드</TD>
        <TD CLASS="TD6" NOWRAP><INPUT ID=txtDilig_cd NAME="txtDilig_cd" ALT="근태코드" TYPE="Text" SiZE="10" MAXLENGTH="2" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
                                                       <INPUT ID=txtDilig_nm NAME="txtDilig_nm" ALT="근태코드명" TYPE="Text" SiZE="20" MAXLENGTH="20" tag="14XXXU"></TD>
       
           <TD CLASS="TD5" NOWRAP>근태구분</TD>
        <TD CLASS="TD6" NOWRAP><SELECT NAME="CboDilig_type" ALT="근태구분" CLASS ="cbonormal" TAG="1XN"><OPTION VALUE=""></OPTION></SELECT></TD>
       </TR>
       <TR>
           <TD CLASS="TD5" NOWRAP></TD>
                 <TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" NAME="rbo_sort" ID="rbo_sort1" VALUE="1" CLASS="RADIO" TAG="11" CHECKED><LABEL FOR="rbo_sort1">코드순</LABEL>&nbsp;
                                       <INPUT TYPE="RADIO" NAME="rbo_sort" ID="rbo_sort2" VALUE="2" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_sort2">코드명순</LABEL>&nbsp;
                                      <INPUT TYPE="RADIO" NAME="rbo_sort" ID="rbo_sort3" VALUE="3" CLASS="RADIO" TAG="11"><LABEL FOR="rbo_sort3">순번순</LABEL>
           <TD CLASS="TD5" NOWRAP></TD>
        <TD CLASS="TD6" NOWRAP></TD>
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
      <TABLE <%=LR_SPACE_TYPE_20%> >
       <TR>
        <TD HEIGHT=100% WIDTH=100% >
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
    
   </TABLE>
  </TD>
 </TR>

 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
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
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

