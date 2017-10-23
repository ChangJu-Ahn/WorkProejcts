<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>
<!--
======================================================================================================
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
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "B1Z04MB1_KO441.asp"                                      '비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

Const TAB1 = 1									
Const TAB2 = 2

Dim arrCollectVatType

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_SEQ
Dim C_TODO_DOC
Dim C_COMBO_YN
Dim C_UD_MAJOR_CD
Dim C_UD_MINOR_CD
Dim C_UD_MINOR_POP
Dim C_SAMPLE_DATA
Dim C_PROCESS_TYPE
Dim C_MES_USE_YN
Dim C_CDN_BIZ
Dim C_CDN_BMP
Dim C_CDN_PKG
Dim C_CDN_PRD
Dim C_CDN_TQC
Dim C_REMARK

Dim IsOpenPop          
Dim gSelframeFlg
Dim lgTabClickFlag  
Dim lgOpenFlag

Dim FromDateOfDB
Dim ToDateOfDB

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("m",-1,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("m", 0,"<%=GetSvrDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	
	gSelframeFlg = TAB1
	   	
	Set gActiveElement = document.activeElement
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)

	lgOpenFlag	= False
	lgTabClickFlag = False
	gSelframeFlg = TAB2
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetClickflag, ResetClickflag()  -----------------------------
'	Name : SetClickflag, ResetClickflag()
'	Description :  
'---------------------------------------------------------------------------------------------------------
Function SetClickflag()

	lgTabClickFlag = True	
	
End Function

Function ResetClickflag()

	lgTabClickFlag = False
	
End Function

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  

	C_SEQ						= 1
	C_TODO_DOC			= 2
	C_COMBO_YN			= 3
	C_UD_MAJOR_CD		= 4
	C_UD_MINOR_CD		= 5
	C_UD_MINOR_POP	= 6
	C_SAMPLE_DATA		= 7
	C_PROCESS_TYPE	= 8
	C_MES_USE_YN		= 9
	C_CDN_BIZ				= 10
	C_CDN_BMP				= 11
	C_CDN_PKG				= 12
	C_CDN_PRD				= 13
	C_CDN_TQC				= 14
	C_REMARK				= 15

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
	frm1.txtInsUser.value = parent.gUsrNm
	frm1.txtNoteDt.Text = ToDateOfDB
	frm1.txtVatType.value = "A"
	Call txtVatType_OnChange()
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

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
			.MaxCols = C_REMARK + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
			.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
			.ColHidden = True
			.MaxRows = 0
			Call GetSpreadColumnPos("A")  	

			ggoSpread.SSSetEdit   C_SEQ     		, "순번", 10,,, 3, 2
			ggoSpread.SSSetEdit   C_TODO_DOC    , "항목", 20,,, 50, 1
			ggoSpread.SSSetCheck	C_COMBO_YN		, "선택",	10,,,true
			ggoSpread.SSSetEdit   C_UD_MAJOR_CD	, "코드그룹", 10,,, 10, 2
			ggoSpread.SSSetEdit   C_UD_MINOR_CD	, "공통코드", 10,,, 10, 2
			ggoSpread.SSSetButton C_UD_MINOR_POP
			ggoSpread.SSSetEdit   C_SAMPLE_DATA	, "Sample Data", 20,,, 50, 1
			ggoSpread.SSSetEdit   C_PROCESS_TYPE, "Process Type", 20,,, 20, 1
			ggoSpread.SSSetCheck	C_MES_USE_YN	, "Mes Code운영",10,,,true
			ggoSpread.SSSetCheck	C_CDN_BIZ			, "영업등록",10,,,true
			ggoSpread.SSSetCheck	C_CDN_BMP			, "기술등록(Bump)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PKG			, "기술등록(Pkg)",10,,,true
			ggoSpread.SSSetCheck	C_CDN_PRD			, "품질등록",10,,,true
			ggoSpread.SSSetCheck	C_CDN_TQC			, "생산등록",10,,,true			
			ggoSpread.SSSetEdit  	C_REMARK      , "비고", 50,,, 50, 1
 
      call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
      call ggoSpread.SSSetColHidden(C_UD_MAJOR_CD,C_UD_MAJOR_CD,True)
		        
	   .ReDraw = true
	
     Call SetSpreadLock 
    
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

						C_SEQ						= iCurColumnPos(1)
						C_TODO_DOC			= iCurColumnPos(2)
						C_COMBO_YN			= iCurColumnPos(3)
						C_UD_MAJOR_CD		= iCurColumnPos(4)
						C_UD_MINOR_CD		= iCurColumnPos(5)
						C_UD_MINOR_POP	= iCurColumnPos(6)
						C_SAMPLE_DATA		= iCurColumnPos(7)
						C_PROCESS_TYPE	= iCurColumnPos(8)
						C_MES_USE_YN		= iCurColumnPos(9)
						C_CDN_BIZ				= iCurColumnPos(10)
						C_CDN_BMP				= iCurColumnPos(11)
						C_CDN_PKG				= iCurColumnPos(12)
						C_CDN_PRD				= iCurColumnPos(13)
						C_CDN_TQC				= iCurColumnPos(14)
						C_REMARK				= iCurColumnPos(15)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    	C_TODO_DOC, -1, C_TODO_DOC
        ggoSpread.SpreadLock    	C_COMBO_YN, -1, C_COMBO_YN
        ggoSpread.SpreadLock    	C_UD_MAJOR_CD, -1, C_UD_MAJOR_CD
        ggoSpread.SpreadLock    	C_PROCESS_TYPE, -1, C_PROCESS_TYPE
        ggoSpread.SpreadLock    	C_MES_USE_YN, -1, C_MES_USE_YN
        ggoSpread.SpreadLock    	C_CDN_BIZ, -1, C_CDN_BIZ
        ggoSpread.SpreadLock    	C_CDN_BMP, -1, C_CDN_BMP
        ggoSpread.SpreadLock    	C_CDN_PKG, -1, C_CDN_PKG
        ggoSpread.SpreadLock    	C_CDN_PRD, -1, C_CDN_PRD
        ggoSpread.SpreadLock    	C_CDN_TQC, -1, C_CDN_TQC
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
       
         ggoSpread.SSSetProtected		C_TODO_DOC, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_COMBO_YN, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_UD_MAJOR_CD, pvStartRow, pvEndRow
         
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
  Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
  Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
  Call InitSpreadSheet                                                            'Setup the Spread sheet
  Call InitVariables                                                              'Initializes local global variables
        
  Call SetDefaultVal
  Call InitComboBox
  Call SetToolbar("1100100100101111")										        '버튼 툴바 제어 
  frm1.txtItemCd.focus
End Sub

Sub InitComboBox()
    On Error Resume Next
    Err.Clear
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs("     UD_MINOR_CD,UD_MINOR_NM "," b_user_defined_minor "," UD_MAJOR_CD = " & FilterVar("zz006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboPrefix, lgF0, lgF1, Chr(11))

	
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
    Dim strFrDept, strToDept
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
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

    Call InitVariables                                                        '⊙: Initializes local global variables

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
    Dim intRetCd

    FncDelete = False												
    
	'-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If

	'-----------------------
    'Delete function call area
    '-----------------------%>
    IntRetCd = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		           
	If IntRetCd = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then
		Exit Function
	End If													
    
    FncDelete = True                                                    
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim strReturn_value, strSQL
    Dim HFlag,MFlag,Rowcnt
    Dim strVdate
    Dim strWhere
    Dim strDay_time
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
	 '------ Precheck area ------ 
	If lgBlnFlgChgValue = False Then								 'Check if there is retrived data 
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")					 '⊙: No data changed!! 
	    Exit Function
	End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

    FncSave = True                                            
    
		Call DisableToolBar(parent.TBC_SAVE)
		If DbSave = False Then                                    '☜: Save db data     Processing is OK
			Call RestoreToolBar()
      Exit Function
    End If
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                         
    
	'-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")	           
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

				ggoOper.SetReqAttr	frm1.rdoPhantomType1, "N"
				ggoOper.SetReqAttr	frm1.rdoPhantomType2, "N"
				ggoOper.SetReqAttr	frm1.rdoDP1, "N"
				ggoOper.SetReqAttr	frm1.rdoDP2, "N"
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
    Call InitVariables															'⊙: Initializes local global variables
    frm1.txtItemCd.focus
    Set gActiveElement = document.activeElement
      
    FncNew = True																'⊙: Processing is OK

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
            .Col = C_SEQ
            .Text = ""
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
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
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
	Call ggoSpread.ReOrderingSpreadData()
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
        strVal = strVal     & "&txtItemCd="       & frm1.txtItemCd.value
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
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
	
    DbSave = False                                                          
    
     If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                    strVal = strVal & "C" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.UpdateFlag                                      '☜: Update
                    strVal = strVal & "U" & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strVal = strVal & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strVal = strVal & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete
                    strDel = strDel & "D" & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_SEQ,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_SAMPLE_DATA,lRow,"X","X") & parent.gColSep
										strDel = strDel & GetSpreadText(frm1.vspdData,C_REMARK,lRow,"X","X") & parent.gColSep
                    strDel = strDel & lRow & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

           End Select
       Next
       .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal
		.txtFlgMode.value = lgIntFlgMode

	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                               
    
    DbDelete = False													
    
    LayerShowHide(1)						
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd2.value)		
	
	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbDelete = True                                                     

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    
    Dim iCnt
    For iCnt=1 To frm1.vspdData.MaxRows
	    Call SetSpreadColorAfterQuery(C_COMBO_YN, iCnt)
    Next

	lgBlnFlgChgValue = False						

	If CommonQueryRs(" CDN_BIZ ", " B_CDN_REQ_HDR_KO441 ", " ITEM_CD=" & FilterVar(frm1.txtItemCd.value,"''","S") & " AND (CDN_BIZ='Y' or CDN_BMP='Y' OR CDN_PKG='Y' or CDN_PRD='Y' or CDN_TQC='Y')", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		If InStr(lgF0,"Y") > 0 Then
        ggoSpread.SpreadLock    	-1, -1
				ggoOper.SetReqAttr	frm1.txtCBMdescription, "Q"
				ggoOper.SetReqAttr	frm1.txtItemNm1, "Q"
				ggoOper.SetReqAttr	frm1.txtUnit, "Q"
				ggoOper.SetReqAttr	frm1.cboItemAcct, "Q"
				ggoOper.SetReqAttr	frm1.txtValidDt, "Q"
				ggoOper.SetReqAttr	frm1.txtItemSpec, "Q"
				ggoOper.SetReqAttr	frm1.txtVatType, "Q"
				ggoOper.SetReqAttr	frm1.txtVatRate, "Q"
				ggoOper.SetReqAttr	frm1.rdoPhantomType1, "Q"
				ggoOper.SetReqAttr	frm1.rdoPhantomType2, "Q"
				ggoOper.SetReqAttr	frm1.txtNoteDt, "Q"
				ggoOper.SetReqAttr	frm1.rdoDP1, "Q"
				ggoOper.SetReqAttr	frm1.rdoDP2, "Q"
				Call SetToolbar("111000000001111")			
				EXIT Function 
		End If
	End If
    
	Call SetToolbar("111110110001111")			
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
	Call InitVariables
	Call FncNew()
End Function

'==========================================================================================================
'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & "  "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtUnit.focus
	
End Function
'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(byval arrRet)
	frm1.txtUnit.Value		= arrRet(0)		
	lgBlnFlgChgValue		= True
End Function
'===========================================================================
' Function Name : OpenBillHdr
' Function Desc : OpenBillHdr Reference Popup
'===========================================================================
Function OpenBillHdr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
	If frm1.txtVatType.readOnly = True Then
		IsOpenPop = False
		Exit Function
	End If

	arrParam(1) = "B_MINOR ,B_CONFIGURATION "	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType.value)	' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = "B_MINOR.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " " _
					& " AND B_MINOR.MINOR_CD=B_CONFIGURATION.MINOR_CD " _
					& " AND B_MINOR.MAJOR_CD=B_CONFIGURATION.MAJOR_CD "	_
					& " AND B_CONFIGURATION.SEQ_NO = 1 "					' Where Condition
	arrParam(5) = "VAT유형"					' TextBox 명칭 
		
	arrField(0) = "B_MINOR.MINOR_CD"			' Field명(0)
	arrField(1) = "B_MINOR.MINOR_NM"			' Field명(1)
	arrField(2) = "F5" & parent.gColSep & "B_CONFIGURATION.REFERENCE"				' Field명(2)
	    	    
	arrHeader(0) = "VAT유형"				' Header명(0)
	arrHeader(1) = "VAT유형명"				' Header명(1)
	arrHeader(2) = "VAT율"					' Header명(2)

	arrParam(0) = arrParam(5)					' 팝업 명칭 

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBillHdr(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtVatType.focus
	
End Function
'------------------------------------------  SetBillHdr()  -----------------------------------------------
'	Name : SetBillHdr()
'	Description : Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetBillHdr(Byval arrRet)
	frm1.txtVatType.value = arrRet(0)
	frm1.txtVatTypeNm.value = arrRet(1)
	frm1.txtVatRate.Text = arrRet(2)
	lgBlnFlgChgValue = true

End Function

'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 부가세타입 내용이 변경되었을때 부가세율 계산 
'==========================================================================================
Sub txtVatType_OnChange()

	Dim VatType, VatTypeNm, VatRate

	VatType = Trim(frm1.txtVatType.value)
	
	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtVatTypeNm.value = VatTypeNm
	frm1.txtVatRate.text = VatRate
	lgBlnFlgChgValue = true
End Sub
Sub cboItemAcct_OnChange()
	lgBlnFlgChgValue = true
End Sub
Sub rdoPhantomType1_OnChange()
	lgBlnFlgChgValue = true
End Sub
Sub rdoPhantomType2_OnChange()
	lgBlnFlgChgValue = true
End Sub
Sub rdoDP1_OnChange()
	lgBlnFlgChgValue = true
End Sub
Sub rdoDP2_OnChange()
	lgBlnFlgChgValue = true
End Sub
'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================
Sub InitCollectType()
	Dim i
	Dim iCodeArr, iNameArr, iRateArr
	
    On Error Resume Next

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & "  And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = REPLACE(iRateArr(i),".",parent.gComNumDec)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef====>ado이용추가 
' Function Desc : 
'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)  
		If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
	
End Sub

'==========================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_CDN_REQ_HDR_KO441"		 		' TABLE 명칭 
	arrParam(2) = frm1.txtItemCd.Value				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "품목코드"			
	
    arrField(0) = "ITEM_CD"							' Field명(0)
    arrField(1) = "ITEM_NM"							' Field명(1)
    
    arrHeader(0) = "품목코드"				' Header명(0)
    arrHeader(1) = "품목명"			' Header명(1)    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
	End If	

End Function
'==========================================================================================================
Function OpenMinor(pRow)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,pRow,"X","X") = "" Then
		Call DisplayMsgBox("971012", "X", "코드그룹", "X")
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사용자Minor코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_User_Defined_MINOR"		 		' TABLE 명칭 
	arrParam(2) = GetSpreadText(frm1.vspdData,C_UD_MINOR_CD,pRow,"X","X")				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "UD_MAJOR_CD=" & FilterVar(GetSpreadText(frm1.vspdData,C_UD_MAJOR_CD,pRow,"X","X"),"''","S")									' Where Condition
	arrParam(5) = "사용자Minor코드"			
	
    arrField(0) = "UD_MINOR_CD"							' Field명(0)
    arrField(1) = "UD_MINOR_NM"							' Field명(1)
    arrField(2) = "UD_REFERENCE"						' Field명(2)
    
    arrHeader(0) = "사용자Major코드"				' Header명(0)
    arrHeader(1) = "사용자Major코드명"			' Header명(1)
    arrHeader(2) = "Reference"							' Header명(2)
    


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_UD_MINOR_CD,pRow,arrRet(0))
		Call frm1.vspdData.SetText(C_SAMPLE_DATA,pRow,arrRet(2))
		call vspdData_Change(C_UD_MINOR_CD , pRow)
	End If	

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
		Select Case Col
			Case C_UD_MINOR_POP
				call OpenMinor(Row)
		End Select    
	End If
            
End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Call SetSpreadColorAfterQuery(Col,Row)
    
	ggoSpread.Source = frm1.vspdData
  ggoSpread.UpdateRow Row
  lgBlnFlgChgValue=True
End Function

Function SetSpreadColorAfterQuery(Col, Row)
    With frm1
    
       .vspdData.ReDraw = False

    Select Case Col
         Case  C_COMBO_YN
         	If GetSpreadText(frm1.vspdData,C_COMBO_YN,Row,"X","X")="1" Then
		        ggoSpread.SpreadUnLock    	C_UD_MAJOR_CD, Row, C_UD_MAJOR_CD, Row
						ggoSpread.SSSetRequired			C_UD_MAJOR_CD, Row, Row
		        ggoSpread.SpreadUnLock    	C_UD_MINOR_CD, Row, C_UD_MINOR_CD, Row
						ggoSpread.SSSetRequired			C_UD_MINOR_CD, Row, Row
		        ggoSpread.SpreadUnLock    	C_UD_MINOR_POP, Row, C_UD_MINOR_POP, Row
         	Else
         		Call frm1.vspdData.SetText(C_UD_MAJOR_CD,Row,"")
         		Call frm1.vspdData.SetText(C_UD_MINOR_CD,Row,"")
		        ggoSpread.SSSetProtected    	C_UD_MAJOR_CD, Row, Row
		        ggoSpread.SSSetProtected    	C_UD_MINOR_CD, Row, Row
		        ggoSpread.SSSetProtected    	C_UD_MINOR_POP, Row, Row
         	End If
    End Select    
       .vspdData.ReDraw = True
    
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
'   Event Desc : 
'========================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
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

Sub GetItemCode()
    On Error Resume Next
    Err.Clear
    
    Dim iSeqNo
    
    Call CommonQueryRs(" isnull(left(max(replace(item_cd,'" & frm1.cboPrefix.value & "','')),3),0) "," b_item "," item_cd like " & FilterVar(frm1.cboPrefix.value&"%", "'%'", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		
		iSeqNo = cdbl(replace(lgF0,chr(11),"")) + 1
		
		frm1.txtItemCd2.value = frm1.cboPrefix.value & right("000" & iSeqNo ,3)   
	
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidDt_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtNoteDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtNoteDt_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtNoteDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtNoteDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtNoteDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtNoteDt.Focus
    End If
End Sub

Sub txtCBMdescription_OnChange()
	If Trim(frm1.txtItemNm1.value) = "" Then
		frm1.txtItemNm1.value =  frm1.txtCBMdescription.value
	End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>CDN정보등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" onMouseOver="vbscript:SetClickflag" onMouseOut="vbscript:ResetClickflag" onFocus="vbscript:SetClickflag" onBlur="vbscript:ResetClickflag">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>CDN요청등록</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="12XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=50 tag="14"></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR height="*">
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=50% valign=top>
									<FIELDSET>			
										<LEGEND>기본정보</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>채번코드</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPrefix" ALT="채번코드" STYLE="Width: 250px;" tag="23" OnChange="vbscript:GetItemCode()"></SELECT>
														&nbsp;<INPUT TYPE=TEXT NAME="txtItemCd2" CLASS=protected READONLY=true TABINDEX="-1" SIZE=30 tag="14"></TD>													
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>MES Device</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBMdescription" SIZE=40 MAXLENGTH=50 tag="23" ALT="MES Device"></TD>
												</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" CLASS=required SIZE=40 MAXLENGTH=40 tag="22" ALT="품목명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>단위</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnit" CLASS=required STYLE="text-transform:uppercase" SIZE=5 MAXLENGTH=3 tag="22XXXU" ALT="단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenUnit()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목계정</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" CLASS=required STYLE="text-transform:uppercase; Width: 168px;" ALT="품목계정" tag="22"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>사용시작일</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtValidDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="견적일"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목규격</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=50 tag="21" ALT="품목규격"></TD>
											</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>VAT유형</TD>
													<TD CLASS=TD6 NOWRAP>
														<INPUT NAME="txtVatType" STYLE="text-transform:uppercase" TYPE="Text"  MAXLENGTH="5" SIZE=10  ALT="VAT유형" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr">
														<INPUT NAME="txtVatTypeNm" CLASS=protected READONLY=true TABINDEX="-1" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24">
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>VAT율</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLSPACING=0 CELLPADDING=0>
															<TR>
																<TD>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS100 name=txtVatRate CLASSID=CLSID:DD55D13D-EBF7-11D0-8810-0000C0E5948C tag=24X5Z VIEWASTEXT> </OBJECT>');</SCRIPT>
																	&nbsp;<LABEL><b>%</b></LABEL>
																</TD>																
															</TR>
														</TABLE>
													</TD>
												</TR>												
											<TR>
												<TD CLASS=TD5 NOWRAP>Phantom구분</TD>
												<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType1" Value="Y" CLASS="RADIO" tag="2X"><LABEL FOR="rdoPhantomType1">예</LABEL>
															<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType2" Value="N" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoPhantomType2">아니오</LABEL></TD>
											</TR>
											</TABLE>		
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET>	
										<LEGEND>계획정보</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>작성자</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInsUser" CLASS=protected READONLY=true TABINDEX="-1" SIZE=50 tag="24"></TD>													
												</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>작성일자</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtNoteDt name=txtNoteDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="작성일자"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>D/P</TD>
												<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" NAME="rdoDP" ID="rdoDP1" Value="Y" CLASS="RADIO" tag="2X"><LABEL FOR="rdoDP1">Development</LABEL>
															<INPUT TYPE="RADIO" NAME="rdoDP" ID="rdoDP2" Value="N" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoDP2">Production</LABEL></TD>
											</TR>
											</TABLE>
									</FIELDSET>
								</TD>
							</TR>	
						</TABLE>
						</div>
						<!--두번째 탭 -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							</TABLE>
						</DIV>
					</TD>	
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
