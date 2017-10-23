
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         : Single-Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->			<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "f5501mb1.asp"										'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID2 = "f5501mb2.asp"										'☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_GL_DT	
Dim C_SEQ	
Dim C_DR_CR_FG	
Dim C_DR_CR_FG_NM
Dim C_ITEM_AMT	
Dim C_ACCT_CD	
Dim C_ACCT_NM	
Dim C_ITEM_DESC	
Dim C_GL_NO		
Dim C_TEMP_GL_NO
Dim C_COL_END	



'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim IsOpenPop
Dim ptxtCardNo

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE   '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '⊙: Indicates that no value changed
    lgIntGrpCount = 0           '⊙: Initializes Group View Size
	lgStrPrevKey = ""
	
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'☆: 사용자 변수 초기화 
    lgSortKey = 1    
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


Sub initSpreadPosVariables()
		 C_GL_DT	     = 1
		 C_SEQ	         = 2
		 C_DR_CR_FG	     = 3
		 C_DR_CR_FG_NM   = 4
		 C_ITEM_AMT	     = 5
		 C_ACCT_CD	     = 6
		 C_ACCT_NM	     = 7
		 C_ITEM_DESC	 = 8
		 C_GL_NO		 = 9
		 C_TEMP_GL_NO    = 10
		 C_COL_END	     = 11
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtIssueDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.txtDueDt.Text   = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
	frm1.hOrgChangeId.value = Parent.gChangeOrgId	
	frm1.txtCardNoQry.focus 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para :	1. Currency
'			2. I(Input),Q(Query),P(Print),B(Bacth)
'			3. "*" is for Common module
'				"A" is for Accounting
'				"I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*","NOCOOKIE","MA")%>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"
			strTemp = ReadCookie("NOTE_NO")
			Call WriteCookie("NOTE_NO", "")

			If strTemp = "" then Exit Function

			frm1.txtCardNoQry.value = strTemp

			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("NOTE_NO", "")
				Exit Function 
			End If

			Call MainQuery()
		Case Else
			Exit Function
	End Select
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
       Case "Q"
                  lgKeyStream = Frm1.txtCardNoQry.Value  & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "D"
                  lgKeyStream = Frm1.htxtSchoolCd.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "N"
                  lgKeyStream = Frm1.txtCardNoQry.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "P"
                  lgKeyStream = Frm1.txtCardNoQry.Value & parent.gColSep       'You Must append one character(parent.gColSep)
       Case "R"
                  lgKeyStream = Frm1.txtCardNoQry.Value & parent.gColSep       'You Must append one character(parent.gColSep)
   End Select 
                   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1012", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = replace(lgF0, Chr(11), vbTab)
	ggoSpread.SetCombo lgF0, C_DR_CR_FG
	lgF1 = replace(lgF1, Chr(11), vbTab)
	ggoSpread.SetCombo lgF1, C_DR_CR_FG_NM
   
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow
		frm1.vspdData.Col = C_DR_CR_FG
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_DR_CR_FG_NM
		frm1.vspdData.value = intindex
	Next
	
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim sList
    
    Call initSpreadPosVariables()
    
    With frm1
    
		.vspdData.MaxCols = C_COL_END

		.vspdData.Col = .vspdData.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData.ColHidden = True
		
	
		.vspdData.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_SEQ, "순번", 8, , , 3
									'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Format        Row
		ggoSpread.SSSetDate C_GL_DT, "일자", 12,2 , parent.gDateFormat
									'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  ComboEditable   Row
		ggoSpread.SSSetCombo C_DR_CR_FG, "차대구분", 12
		ggoSpread.SSSetCombo C_DR_CR_FG_NM, "차대구분", 12
									'ColumnPosition     Header            Width   Grp            IntegeralPart       DeciPointpart                             Align   Sep    PZ   Min       Max 
		ggoSpread.SSSetFloat C_ITEM_AMT, "금액", 17, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit C_ACCT_CD, "계정코드", 15, , , 20
		ggoSpread.SSSetEdit C_ACCT_NM, "계정명", 25, , , 30
		ggoSpread.SSSetEdit C_ITEM_DESC, "적요", 35, , , 128
		ggoSpread.SSSetEdit C_GL_NO, "전표번호", 15, , , 18
		ggoSpread.SSSetEdit C_TEMP_GL_NO, "결의전표번호", 15, , , 18

        Call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
        Call ggoSpread.SSSetColHidden(C_DR_CR_FG,C_DR_CR_FG,True)
        Call ggoSpread.SSSetColHidden(C_DR_CR_FG_NM,C_DR_CR_FG_NM,True)
        Call ggoSpread.SSSetColHidden(C_ACCT_CD,C_ACCT_CD,True)
        
		Call SetSpreadLock                                              '바뀐부분 

    End With    

End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()    		
'		ggoSpread.SpreadLock 1 , -1
		.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
   With frm1

		.vspdData.ReDraw = False

		.vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
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
'========================================================================================================
'                        5.2 Common Group-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029							'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")		'⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")		'⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet                                                        'Setup the Spread sheet
	Call InitComboBox
    Call InitVariables							'⊙: Initializes local global variables   

    Call FncNew

    'SetGridFocus
    Call CookiePage("FORM_LOAD")

	' 권한관리 추가 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 

	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing

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
   			C_GL_DT	   = iCurColumnPos(1)
			C_SEQ	    =     iCurColumnPos(2)
			C_DR_CR_FG	 =    iCurColumnPos(3)
			C_DR_CR_FG_NM =  iCurColumnPos(4)
			C_ITEM_AMT	   =  iCurColumnPos(5)
			C_ACCT_CD	   =  iCurColumnPos(6)
			C_ACCT_NM	   =  iCurColumnPos(7)
			C_ITEM_DESC	 = iCurColumnPos(8)
			C_GL_NO		 = iCurColumnPos(9)
			C_TEMP_GL_NO  =  iCurColumnPos(10)
			C_COL_END	  =   iCurColumnPos(11)
            
    End Select    
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()

	Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG    
    Err.Clear                                                               '☜: Protect system from crashing

	'-----------------------
    'Check previous data area
    '----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
    'Erase contents area
    '----------------------- 
'    Call ggoOper.ClearField(Document, "2")			'⊙: Clear Contents  Field
    Call InitVariables								'⊙: Initializes local global variables
	frm1.vspdData.MaxRows = 0
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then		'⊙: This function check indispensable field
       Exit Function
    End If
    
    Call ggoOper.LockField(Document, "N")		'⊙: This function lock the suitable field

  '-----------------------
    'Query function call area
    '----------------------- 
      If DbQuery("Q") = False Then                                                       '☜: Query db data
       Exit Function
    End If
       
    If Err.number = 0 Then
       FncQuery = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
	Dim IntRetCD 
    
    FncNew = False      '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
	    IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x") '☜ 바뀐부분 
	     If IntRetCD = vbNo Then
	         Exit Function
	     End If
	End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")	'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")  '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")   '⊙: Lock  Suitable  Field
    

    Call SetDefaultVal
    Call InitVariables						'⊙: Initializes local global variables
	frm1.vspdData.MaxRows = 0

    Call SetToolbar("1110100000000011")										'⊙: 버튼 툴바 제어 

    If Err.number = 0 Then
       FncNew = True                                                              '☜: Processing is OK
    End If   
    
	frm1.txtCardNoQry.focus 
	Set gActiveElement = document.activeElement

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
	 Dim IntRetCD 
    
    FncDelete = False														'⊙: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        intRetCD = DisplayMsgBox("900002","x","x","x")  '☜ 바뀐부분 
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")  '☜ 바뀐부분 
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete															'☜: Delete db data
    
    If Err.number = 0 Then
       FncDelete = True                                                           '☜: Processing is OK
    End If   
    
    Set gActiveElement = document.ActiveElement       
   
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG    
    Err.Clear                                                               '☜: Protect system from crashing
    
	'-----------------------
	'Precheck area
	'-----------------------
	
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","x","x","x")  '☜ 바뀐부분 
		Exit Function
	End If
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
    
  '-----------------------
    'Save function call area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area(Multi)
       Exit Function
    End If
    
    
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If (frm1.txtIssueDt.Text <> "") And (frm1.txtDueDt.Text <> "") Then	
		If CompareDateByFormat(frm1.txtIssueDt.Text, frm1.txtDueDt.Text, frm1.txtIssueDt.Alt, frm1.txtDueDt.Alt, _
					"970025", frm1.txtIssueDt.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtIssueDt.focus											
			Exit Function
		End if	
	End If
	
	If CommonQueryRs(" A.DEPT_NM ","B_ACCT_DEPT A","A.ORG_CHANGE_ID = " & FilterVar(frm1.hOrgChangeId.value, "''", "S") & _
						" AND a.dept_cd = " & FilterVar(Trim(frm1.txtDeptCd.value)	, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("800062","X","X","X")            '☜ : No data is found. 
		Exit Function
	End If

	If CommonQueryRs(" A.BP_NM ","B_BIZ_PARTNER A","A.BP_CD = " & FilterVar(frm1.txtBpCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("970000","X",frm1.txtBpCd.alt,"X")            '☜ : No data is found. 
		Exit Function
	End If

	If CommonQueryRs(" A.CARD_CO_NM ","B_CARD_CO A","A.RCPT_CARD_FG = " & FilterVar("Y", "''", "S") & "  AND A.CARD_CO_CD = " & FilterVar(frm1.txtCardCoCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("970000","X",frm1.txtCardCoCd.alt,"X")            '☜ : No data is found. 
		Exit Function
	End If

	If UniCdbl(frm1.txtCardAmt.text) <= 0 Then
		Call DisplayMsgBox("972001","X",frm1.txtCardAmt.alt,"0")            '☜ : No data is found. 
		Exit Function
	End If
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
    If Err.number = 0 Then
       FncSave = True                                                           '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel()
    
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()

End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()  

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False	                                                          '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                              '☜: Processing is NG
      
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
    	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables													         '⊙: Initializes local global variables

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    
    If DbQuery("P") = False Then                                                 '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncPrev = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                              '☜: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
       Call DisplayMsgBox("900002","x","x","x")
       Exit Function
    End If
	
    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
       If IntRetCD = vbNo Then
          Exit Function
       End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
    Call InitVariables														     '⊙: Initializes local global variables

    If DbQuery("N") = False Then                                                 '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncNext = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncExport(parent.C_MULTI)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncExcel = True                                                            '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	'------ Developer Coding part (Start )   -------------------------------------------------------------- 
	Call Parent.FncFind(parent.C_MULTI, True)
	'------ Developer Coding part (End   )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncFind = True                                                             '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			          '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    If Err.number = 0 Then
       FncExit = True                                                             '☜: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    DbQuery = False                                                               '☜: Processing is NG

    Call DisableToolBar(parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call MakeKeyStream(pDirect)

    strVal = BIZ_PGM_ID & "?txtMode="        & parent.UID_M0001                          '☜: Query
    strVal = strVal     & "&txtKeyStream="   & lgKeyStream                        '☜: Query Key
    strVal = strVal     & "&txtPrevNext="    & pDirect                            '☜: Direction
    strVal = strVal     & "&lgStrPrevKey="   & lgStrPrevKey                       '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="     & Frm1.vspdData.MaxRows              '☜: Max fetched data

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                            '☜:  Run biz logic
	
    If Err.number = 0 Then
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

	Call LayerShowHide(1)
	
	ptxtCardNo = frm1.txtCardNoQry.value 
	
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode	

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
				
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    If Err.number = 0 Then
       DbSave = True                                                   '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCardNoQry=" & Trim(frm1.txtCardNoQry.value)		'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    If Err.number = 0 Then
       DbDelete = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 	
    Frm1.vspdData.focus
	Call SetToolbar("1111100011000111")                                           '☆: Developer must customize	
    Call InitData()
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")	
    Set gActiveElement = document.ActiveElement   	
	
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()  
	On Error Resume Next															'☜: If process fails
    Err.Clear																				'☜: Clear error status    
    
    Call InitVariables   

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
'	Frm1.txtCardNoQry.Value = pCardNo
    call MainQuery()
'	Call SetToolbar("1111111111111111")                                       '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Set gActiveElement = document.ActiveElement
    
    

End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

    Call FncNew()

End Sub


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenNoteInfo()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("f5501ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f5502ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
  
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID, Array(window.Parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		Exit Function
	Else
		frm1.txtCardNoQry.value  = arrRet(0)
		frm1.txtCardNoQry.focus 
	End If	

End Function
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function



'==================================================================================
'	Name : OpenPopUp()
'	Description : 공통팝업 정의 
'==================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
'			
		Case 5		' 카드사 
			arrParam(0) = "카드사 팝업"					' 팝업 명칭 
			arrParam(1) = "B_CARD_CO"			 			' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""										' Name Cindition
			arrParam(4) = "RCPT_CARD_FG = " & FilterVar("Y", "''", "S") & "   "			' Where Condition
			arrParam(5) = "카드사코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "CARD_CO_CD"						' Field명(0)
			arrField(1) = "CARD_CO_NM"						' Field명(1)
    
			arrHeader(0) = "카드사코드"				' Header명(0)
			arrHeader(1) = "카드사명"					' Header명(1)

	End Select
  
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function OpenPopuptempGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'결의전표번호 
			arrParam(1) = ""			'Reference번호 
		End If
	End With						'Reference번호 
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_Gl_No
			arrParam(0) = Trim(.Text)	'전표번호 
			arrParam(1) = ""				'Reference번호 
		End If
	End With						'Reference번호 

	IsOpenPop = True
   
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'부서코드 
	arrParam(1) = frm1.txtIssueDt.Text			'날짜(Default:현재일)
	arrParam(2) = "1"							'부서권한(lgUsrIntCd)
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCD.focus
		Exit Function
	End If
	
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	frm1.txtIssueDt.text = arrRet(3)		
	Call txtDeptCD_OnChange()
	frm1.txtDeptCD.focus
	
	lgBlnFlgChgValue = True
End Function

 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere

			Case 4		' 거래처 
				.txtBpCd.value = arrRet(0)
				.txtBpNM.value = arrRet(1)
				.txtBpCd.focus
				lgBlnFlgChgValue = True
			Case 5		' 발행은행 
				.txtCardCoCd.value = arrRet(0)
				.txtCardCoNm.value = arrRet(1)
				.txtCardCoCd.focus
				lgBlnFlgChgValue = True
		End Select

	End With
End Function
'------------------------------------------  EscPopUp()  --------------------------------------------------
'	Name : EscPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere

			Case 2		' 부서 
				.txtDeptCD.focus				
			Case 3		' 코스트센타 
				.txtCostCD.focus
			Case 4		' 거래처 
				.txtBpCd.focus
			Case 5		' 발행은행 
				.txtCardCoCd.focus
		End Select

	End With
End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case parent.C_ZipCodePopUp
				.Col = Col - 1
				.Row = Row
				Call OpenZipCode(.Text,Row)
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col
         Case  parent.C_StudyOnOffnM
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = parent.C_StudyOnOffCd
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtIssueDt.Focus 
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt_Change()

 Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtIssueDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtIssueDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
	
    lgBlnFlgChgValue = True

End Sub
'=======================================================================================================
'   Event Name :txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus         
    End If
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCardAmt_Change()
    lgBlnFlgChgValue = True    
End Sub

Sub txtSttlAmt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : Onchange
'========================================================================================================
'Sub txtCardNo_OnChange()
'	lgBlnFlgChgValue = True
'End Sub

Sub txtDeptCD_OnChange()

  Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	If Trim(frm1.txtDeptCd.value) = "" and Trim(frm1.txtIssueDt.Text = "") Then		Exit Sub

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtIssueDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
		'----------------------------------------------------------------------------------------

     lgBlnFlgChgValue = True
End Sub

Sub txtBpCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtCardCoCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtCardDesc_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"	'Split 상태코드 
 
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
    
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
    
    

End Sub

Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    
  
End Sub
  
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if  frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
    End if
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData()
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopuptempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
									<TD CLASS=TD5 NOWRAP>수취구매카드번호</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtCardNoQry" NAME="txtCardNoQry" SIZE=30 MAXLENGTH=30 tag="12XXXU"ALT="구매카드번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteQry" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:call OpenNoteInfo()"></TD>
								<TR>		
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
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON"  ONCLICK="vbscript:Call OpenPopupDept(frm1.txtDeptCD.Value, 2)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="부서"></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="22XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 4)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="24X" ALT="거래처"> </TD>
								<TD CLASS=TD5 NOWRAP>카드사</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoCd" NAME="txtCardCoCd" SIZE=10 MAXLENGTH=10   tag="22XXXU" ALT="카드사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCardCoCd.Value, 5)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardCoNm" NAME="txtCardCoNm" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="은행"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발행일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="발행일" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>결제일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="결제일" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>카드금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtCardAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="카드금액" tag="22X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>결제금액</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtSttlAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="결제금액" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT CLASS="clstxt" TYPE=TEXT ID="txtCardDesc" NAME="txtCardDesc" SIZE=70 MAXLENGTH=128  tag="2XX" ALT="비고"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtCardNo" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtInternalCd" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd" tag="2">
<INPUT TYPE=HIDDEN NAME="htxtCostCd" tag="2">
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="2">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

