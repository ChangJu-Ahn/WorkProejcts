<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 결산및장부관리
'*  3. Program ID           : h6152MA1_KO441
'*  4. Program Name         : 영세율첨부서류제출명세서
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/05/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Joo JiYeong
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :
'* 13. History              :
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
<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "a6152mb1_KO441.asp"                                      'Biz Logic ASP
Const C_SHEETMAXROWS    = 41	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_SEQ 
Dim C_DOCUMENT 
Dim C_DOCUMENT_NM
Dim C_ISSUE_CD 
Dim C_ISSUE_NM
Dim C_ISSUE_DT 
Dim C_SHIP_DT 
Dim C_DOCCUR
Dim C_DOCCUR_POP
Dim C_DOCCUR_NM 
Dim C_XCH_RATE
Dim C_PRESENT_AMT
Dim C_PRESENT_LOC_AMT
Dim C_REPORT_AMT
Dim C_REPORT_LOC_AMT 
Dim C_DESC 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 C_SEQ				= 1
	 C_DOCUMENT			= 2
	 C_DOCUMENT_NM		= 3
	 C_ISSUE_CD			= 4
	 C_ISSUE_NM			= 5
	 C_ISSUE_DT			= 6
	 C_SHIP_DT			= 7
	 C_DOCCUR			= 8
	 C_DOCCUR_POP		= 9
	 C_DOCCUR_NM		= 10
	 C_XCH_RATE			= 11
	 C_PRESENT_AMT		= 12
	 C_PRESENT_LOC_AMT	= 13
	 C_REPORT_AMT		= 14
	 C_REPORT_LOC_AMT	= 15
	 C_DESC				= 16
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
	
	frm1.txtFr_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtFr_dt.Month = strMonth 
	frm1.txtFr_dt.Day = "01"

	frm1.txtTo_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtTo_dt.Month = strMonth 
	frm1.txtTo_dt.Day = strDay

	frm1.txtPrt_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtPrt_dt.Month = strMonth 
	frm1.txtPrt_dt.Day = strDay
	
	frm1.txtcnt.value = 1	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "A","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Replace(Frm1.txtFr_dt.Text,"-","") & parent.gColSep                                           'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Replace(Frm1.txtTo_dt.Text,"-","") & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.cboBizArea.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR "," UD_MAJOR_CD = 'A03' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_DOCUMENT
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_DOCUMENT_NM

    Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR "," UD_MAJOR_CD = 'A04' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ISSUE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ISSUE_NM


    Call CommonQueryRs(" BIZ_AREA_CD,BIZ_AREA_NM "," B_BIZ_AREA "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.cboBizArea,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서 

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
			.Col = C_DOCUMENT
			intIndex = .value
			.col = C_DOCUMENT_NM
			.value = intindex

			.Col = C_ISSUE_CD
			intIndex = .value
			.col = C_ISSUE_NM
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
        .MaxCols = C_DESC + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0
	    ggoSpread.ClearSpreadData
		Call GetSpreadColumnPos("A") 

        ggoSpread.SSSetEdit     C_SEQ,			"순번", 4,,, 03
        ggoSpread.SSSetCombo    C_DOCUMENT,     "서류명",         05
        ggoSpread.SSSetCombo    C_DOCUMENT_NM,  "서류명", 15
        ggoSpread.SSSetCombo    C_ISSUE_CD,     "발급자",         05
        ggoSpread.SSSetCombo    C_ISSUE_NM,		"발급자", 15
        ggoSpread.SSSetDate     C_ISSUE_DT,     "발급일",   10,2, gDateFormat
        ggoSpread.SSSetDate     C_SHIP_DT,		"선적일",   10,2, gDateFormat
        
        ggoSpread.SSSetEdit     C_DOCCUR,       "통화", 8,,, 05
        ggoSpread.SSSetButton   C_DOCCUR_POP
        ggoSpread.SSSetEdit     C_DOCCUR_NM,    "통화명",   10,,, 30

        ggoSpread.SSSetFloat    C_XCH_RATE,     "환율", 10, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_PRESENT_AMT,  "당기제출(외화)", 14, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_PRESENT_LOC_AMT ,	"당기제출(원화)",	 14, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_REPORT_AMT,	"당기신고(외화)", 14, parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_REPORT_LOC_AMT ,	"당기신고(원화)",	 14, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit     C_DESC,     "비고",   27,,, 30

	    call ggoSpread.MakePairsColumn(C_DOCCUR, C_DOCCUR_POP)

	    Call ggoSpread.SSSetColHidden(C_DOCUMENT,C_DOCUMENT,True)
	    Call ggoSpread.SSSetColHidden(C_ISSUE_CD,C_ISSUE_CD,True)
	    Call ggoSpread.SSSetColHidden(C_DOCCUR_NM,C_DOCCUR_NM,True)

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
   ' ggoSpread.SpreadLock    C_DOCUMENT, -1, C_DOCUMENT
    ggoSpread.SpreadLock    C_DOCCUR_NM, -1, C_DOCCUR_NM
    ggoSpread.SpreadLock    C_SEQ, -1, C_SEQ
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
  '  ggoSpread.SSSetRequired		C_DOCUMENT, pvStartRow, pvEndRow
  '  ggoSpread.SSSetRequired		C_DOCUMENT_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired			C_SEQ, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_DOCCUR_NM, pvStartRow, pvEndRow
   ' ggoSpread.SSSetRequired		C_ISSUE_DT, pvStartRow, pvEndRow
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
			 C_SEQ				= iCurColumnPos(1)
			 C_DOCUMENT			= iCurColumnPos(2)
			 C_DOCUMENT_NM		= iCurColumnPos(3)
			 C_ISSUE_CD			= iCurColumnPos(4)
			 C_ISSUE_NM			= iCurColumnPos(5)
			 C_ISSUE_DT			= iCurColumnPos(6)
			 C_SHIP_DT			= iCurColumnPos(7)
			 C_DOCCUR			= iCurColumnPos(8)
			 C_DOCCUR_POP		= iCurColumnPos(9)
			 C_DOCCUR_NM		= iCurColumnPos(10)
			 C_XCH_RATE			= iCurColumnPos(11)
			 C_PRESENT_AMT		= iCurColumnPos(12)
			 C_PRESENT_LOC_AMT	= iCurColumnPos(13)
			 C_REPORT_AMT		= iCurColumnPos(14)
			 C_REPORT_LOC_AMT	= iCurColumnPos(15)
			 C_DESC				= iCurColumnPos(16)           
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
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")		
	
	Call ggoOper.FormatDate(frm1.txtFr_dt, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtTo_dt, Parent.gDateFormat, 2)

    Call InitSpreadSheet                                                            'Setup the Spread sheet
   
    Call InitVariables  
    Call SetDefaultVal                                                            'Initializes local global variables
    Call InitComboBox

    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 

    Call InitComboBox
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
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

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
    Dim lRow
	dim strSch,iDx
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
    
	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
		
           Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                
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
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

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
    Call Initdata()
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
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜FncInsertRow:화면 유형, Tab 유무 
End Function

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

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    Err.Clear                                                                        '☜: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey                 '☜: Next key tag
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

	Dim strRes_no

    DbSave = False                                                          
    
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C"  & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & frm1.cboBizArea.value & parent.gColSep
                                                  strVal = strVal & replace(frm1.txtFr_dt.text,"-","") & parent.gColSep
                                                  strVal = strVal & replace(frm1.txtTo_dt.text,"-","") & parent.gColSep
                    .vspdData.Col = C_SEQ		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DOCUMENT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SHIP_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DOCCUR	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_XCH_RATE  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PRESENT_AMT : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PRESENT_LOC_AMT : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REPORT_AMT  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REPORT_LOC_AMT  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DESC  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U"  & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & frm1.cboBizArea.value & parent.gColSep
                                                  strVal = strVal & replace(frm1.txtFr_dt.text,"-","") & parent.gColSep
                                                  strVal = strVal & replace(frm1.txtTo_dt.text,"-","") & parent.gColSep
                    .vspdData.Col = C_SEQ		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DOCUMENT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SHIP_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DOCCUR	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_XCH_RATE  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PRESENT_AMT : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PRESENT_LOC_AMT : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REPORT_AMT  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REPORT_LOC_AMT  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DESC  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D"  & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & frm1.cboBizArea.value & parent.gColSep
                                                  strDel = strDel & replace(frm1.txtFr_dt.text,"-","") & parent.gColSep
                                                  strDel = strDel & replace(frm1.txtTo_dt.text,"-","") & parent.gColSep
                    .vspdData.Col = C_SEQ		: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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

    Dim strVal
    dim intRow
	Dim intIndex 

    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	Frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function


'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1  
			arrParam(0) = "통화코드 팝업"	
			arrParam(1) = "B_Currency"
			arrParam(2) = ""
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "통화코드"

			arrField(0) = "Currency"
			arrField(1) = "Currency_desc"

			arrHeader(0) = "통화코드"	
			arrHeader(1) = "통화코드명"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case 1   
		        frm1.vspdData.Col = C_DOCCUR
		    	frm1.vspdData.action =0
        End Select
	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 1   
		    	.vspdData.Col = C_DOCCUR_NM
		    	.vspdData.text = arrRet(1)   
		        .vspdData.Col = C_DOCCUR
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.action =0

        End Select

	End With

End Function

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
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
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
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx, IntRetCD
    Dim strSch, strMajor
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
                   
     
    End Select    

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
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
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
        Select Case Col
            Case C_DOCUMENT_NM
                .Col = Col
                intIndex = .Value
				.Col = C_DOCUMENT
				.Value = intIndex
            Case C_ISSUE_NM
                .Col = Col
                intIndex = .Value
				.Col = C_ISSUE_CD
				.Value = intIndex
			End Select							
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_DOCCUR_POP    
	        frm1.vspdData.Col = C_DOCCUR
            Call OpenCode(frm1.vspdData.Text, 1, Row)
    End Select
    
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


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
	Call BtnPrint("N")
End Function	

	
Function BtnPrint(ByVal pvStrPrint) 	
   
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
		Call BtnDisabled(0)
        Exit Function
    End If

	Call BtnDisabled(1)	


   Dim ObjName
   Dim strUrl
   Dim strPrt_dt
   Dim dtDate
    
	dtDate = UNIDateAdd("m", 1, frm1.txtTo_Dt.text & "-01" , parent.gAPDateFormat)
	dtDate = UNIDateAdd("d", -1, dtDate, parent.gAPDateFormat)

    'strPrt_dt = "2005 년 1 기 ( 01월 01일 ~ 01월 31일 )"
    
    strPrt_dt = frm1.txtFr_dt.Year & " 년 " & frm1.txtCNT.value & " 기 ( "
    strPrt_dt = strPrt_dt & frm1.txtFr_dt.Month & "월 01 일 ~ " 
    strPrt_dt = strPrt_dt & frm1.txtTo_dt.Month & "월 " & right(cstr(dtDate),2) & "일 )" 
    
	strUrl =           "fr_dt|" & replace(frm1.txtFr_Dt.text,"-","")
	strUrl = strUrl & "|to_dt|" & replace(frm1.txtTo_Dt.text,"-","")
	strUrl = strUrl & "|biz_area_cd|" & frm1.cboBizArea.value
	strUrl = strUrl & "|cause|" & frm1.txtCause.value
	strUrl = strUrl & "|prt_dt|" & frm1.txtPrt_dt.Text
	strUrl = strUrl & "|prt_dt2|" & strPrt_dt

	OBjName = AskEBDocumentName("a6152oa1_KO441","ebr")    

	
	If pvStrPrint = "N" Then
		Call FncEBRPreview(ObjName, strUrl)
	Else
		Call FncEBRprint(EBAction, ObjName, strUrl)
	End If

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>영세율첨부서류제출명세서</font></td>
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
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
			     		 <TD CLASS=TD5 NOWRAP>거래기간</TD>       
				    	 <TD CLASS=TD6 ><OBJECT classid=<%=gCLSIDFPDT%> id=txtFr_dt name=txtFr_dt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="시작일"></OBJECT>&nbsp;~&nbsp;
				    				    <OBJECT classid=<%=gCLSIDFPDT%> id=txtTo_dt name=txtTo_dt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="종료일"></OBJECT></TD>
			             <TD CLASS=TD5 NOWRAP>사업장</TD>
						 <TD CLASS="TD6" NOWRAP><SELECT NAME="cboBizArea" ALT="사업장" CLASS ="cbonormal" TAG="12N"></SELECT></TD>
			           </TR>
					  </TABLE>
				     </FIELDSET>

					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
		               <TR>
			     		 <TD CLASS=TD5 NOWRAP>작성일</TD>       
				    	 <TD CLASS=TD6 ><OBJECT classid=<%=gCLSIDFPDT%> id=txtPrt_dt name=txtPrt_dt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11X1" ALT="신고일"></OBJECT></TD>
						 <TD CLASS=TD5 NOWRAP>회기</TD>
					     <TD CLASS=TD6><OBJECT classid=<%=gCLSIDFPDS%> id=txtCNT name=txtCNT CLASS=FPDS65 title=FPDOUBLESINGLE tag="11X9Z" ALT="회기"></OBJECT></TD>
			           </TR>
		               <TR>
			     		 <TD CLASS=TD5 NOWRAP>제출사유</TD>       
						 <TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtCause" SIZE=100 MAXLENGTH=180 tag="11" ALT="제출사유"> <!-- <IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenVatNoInfo(frm1.txtVatNo1.value,'VatNo1')"> --> </TD> 
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
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;</TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

