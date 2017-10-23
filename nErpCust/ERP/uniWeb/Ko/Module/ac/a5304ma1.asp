
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 결산장부관리 
*  3. Program ID           : a5304ma1
*  4. Program Name         : 계정대체등록 
*  5. Program Desc         : 계정대체등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/08/08
*  8. Modified date(Last)  : 2002/08/08
*  9. Modifier (First)     : 
* 10. Modifier (Last)      :
* 11. Comment              : 
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID  = "a5304mb1.asp"                                      'Biz Logic ASP 
Const BIZ_PGM_ID2 = "a5304mb2.asp"                                      'Biz Logic ASP 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_CODE
Dim C_CODEPopup
Dim C_CODENM
Dim C_FROMACCTCD
Dim C_FROMACCTCDPopup
Dim C_FROMACCTNM
Dim C_TOACCTCD
Dim C_TOACCTCDPopup
Dim C_TOACCTNM


'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim IsOpenPop          
Dim BaseDate,LastDate


'========================================================================================================
Sub initSpreadPosVariables()
	C_CODE				=  1
	C_CODEPopup			=  2
	C_CODENM			=  3
	C_FROMACCTCD		=  4
	C_FROMACCTCDPopup	=  5
	C_FROMACCTNM		=  6
	C_TOACCTCD			=  7
	C_TOACCTCDPopup		=  8
	C_TOACCTNM			=  9
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()

	frm1.txtYear.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)

End Sub
	
'========================================================================================================

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)

End Sub


	
'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_TOACCTCDPopup  :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
			.Col = C_TOACCTNM  :  .Value = intindex
		Next
	End With
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspddata
	ggoSpread.SpreadInit "V20030102",,parent.gAllowDragDropSpread
	With frm1.vspdData
	
       .MaxCols   = C_TOACCTNM + 1                                                  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True

        ggoSpread.Source = frm1.vspdData

	   .ReDraw = false


        Call GetSpreadColumnPos("A")
       
		ggoSpread.SSSetEdit    C_CODE			,"코드"				,10    ,                   ,     ,1      ,2
		ggoSpread.SSSetButton C_CODEPopup		    
		ggoSpread.SSSetEdit    C_CODENM			,"코드명"			,20    ,                   ,     ,20     ,2
		ggoSpread.SSSetEdit    C_FROMACCTCD     ,"From 계정코드"	,15    ,2                  ,     ,10     ,2
		ggoSpread.SSSetButton  C_FROMACCTCDPopUp
		ggoSpread.SSSetEdit    C_FROMACCTNM     ,"계정명"			,20    ,                   ,     ,20     ,2
		ggoSpread.SSSetEdit    C_TOACCTCD       ,"To 계정코드"		,15    ,2                  ,     ,10     ,2
		ggoSpread.SSSetButton  C_TOACCTCDPopup
		ggoSpread.SSSetEdit    C_TOACCTNM       ,"계정명"			,20    ,                   ,     ,20     ,2

		Call ggoSpread.SSSetColHidden(C_CODE,C_CODE,True)
		Call ggoSpread.SSSetColHidden(C_CODEPopup,C_CODEPopup,True)
		Call ggoSpread.SSSetColHidden(C_CODENM,C_CODENM,True)

		call ggoSpread.MakePairsColumn(C_FROMACCTCD,C_FROMACCTCDPopUp)
		call ggoSpread.MakePairsColumn(C_TOACCTCD,C_TOACCTCDPopup)

	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False 

      ggoSpread.SpreadLock       C_FROMACCTNM   , -1         , C_FROMACCTNM  , -1 
      ggoSpread.SpreadLock       C_TOACCTNM     , -1         , C_TOACCTNM    , -1 

      ggoSpread.SSSetRequired    C_FROMACCTCD         , -1         ,-1
      ggoSpread.SSSetRequired    C_TOACCTCD         , -1         ,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(lRow)
    With frm1
    
    .vspdData.ReDraw = False

	  ggoSpread.SSSetRequired    C_FROMACCTCD   , lRow, lRow
      ggoSpread.SSSetRequired    C_TOACCTCD     , lRow, lRow

      ggoSpread.SSSetProtected       C_FROMACCTNM   , lRow, lRow
      ggoSpread.SSSetProtected       C_TOACCTNM     , lRow, lRow
      
    .vspdData.ReDraw = True
    
    End With
End Sub
Sub SetSpreadColorOk()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected    C_FROMACCTCD   , -1, -1
      ggoSpread.SSSetProtected    C_FROMACCTCDPopUp     , -1, -1
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_CODE				=  iCurColumnPos(1)
            C_CODEPopup			=  iCurColumnPos(2)
            C_CODENM			=  iCurColumnPos(3)
            C_FROMACCTCD		=  iCurColumnPos(4)
            C_FROMACCTCDPopup	=  iCurColumnPos(5)
            C_FROMACCTNM		=  iCurColumnPos(6)
            C_TOACCTCD			=  iCurColumnPos(7)
            C_TOACCTCDPopup		=  iCurColumnPos(8)
            C_TOACCTNM			=  iCurColumnPos(9)
    End Select
End Sub

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
	Call InitVariables
    Call SetDefaultVal
    frm1.btnauto.disabled	=	True

	frm1.txtYear.focus
	Call SetToolbar("110011010010111")
			
End Sub


'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
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
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery() = False Then                                                      '☜: Query db data
       Exit Function
    End If
	
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 

    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	

	With Frm1
        .vspdData.Col  = C_FROMACCTCD
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
        .vspdData.Col  = C_FROMACCTNM
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
	End With

    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    Dim iDx
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo
    Set gActiveElement = document.ActiveElement   
    FncCancel = False
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 

    FncInsertRow = False
    Err.Clear

	Dim varMaxRow,iCurRowPos
	Dim strDoc
	Dim varXrate

	Dim IntRetCD
	Dim imRow
	Dim imRow2


    if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
	else
			imRow = AskSpdSheetAddRowcount()

			If ImRow="" then
			Exit Function
			End If
	End If
		with frm1

			For imRow2=1 to imRow
			ggoSpread.InsertRow ,1
			.vspddata.ReDraw = True
			Call SetSpreadColor(.vspddata.ActiveRow)

			Next
    End With
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Dim varMaxRow
	Dim strDoc
	Dim varXrate

	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData
	if frm1.vspdData.MaxRows < 1 then Exit Function

	ggoSpread.DeleteRow

	lgBlnFlgChgValue = True
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()

	Dim strVal
	
    Err.Clear
    On Error Resume Next
	
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="       & parent.UID_M0001						         
        strVal = strVal     & "&txtYear=" &	.txtYear.text         '☜: Query Key
        strVal = strVal     & "&txtKeyStream=" & lgKeyStream         '☜: Query Key
        strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey        '☜: Next key tag
    End With
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim lRow
    Dim lGrpCnt
	Dim strVal, strDel

    On Error Resume Next
    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    Frm1.txtMode.value        = parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
		
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

               Case ggoSpread.InsertFlag                                      '☜: Create

														  strVal = strVal & "C"                       & parent.gColSep
														  strVal = strVal & lRow                      & parent.gColSep
                    .vspdData.Col = C_CODE				: strVal = strVal & Trim(.vspdData.Text)      & parent.gColSep
                    .vspdData.Col = C_FROMACCTCD		: strVal = strVal & Trim(.vspdData.Text)      & parent.gColSep   
                    .vspdData.Col = C_TOACCTCD			: strVal = strVal & Trim(.vspdData.Text)      & parent.gRowSep   

                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update

														  strVal = strVal & "U"                       & parent.gColSep
														  strVal = strVal & lRow                      & parent.gColSep
                    .vspdData.Col = C_FROMACCTCD		: strVal = strVal & Trim(.vspdData.Text)      & parent.gColSep
                    .vspdData.Col = C_FROMACCTCD		: strVal = strVal & Trim(.vspdData.Text)      & parent.gColSep   
                    .vspdData.Col = C_TOACCTCD			: strVal = strVal & Trim(.vspdData.Text)      & parent.gRowSep   
                   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                     strDel = strDel & "D"							& parent.gColSep
                                                     strDel = strDel & lRow							& parent.gColSep
                    .vspdData.Col = C_FROMACCTCD        : strDel = strDel & Trim(.vspdData.Text)		& parent.gColSep
                    .vspdData.Col = C_FROMACCTCD		: strDel = strDel & Trim(.vspdData.Text)      & parent.gColSep   
                    .vspdData.Col = C_TOACCTCD			: strDel = strDel & Trim(.vspdData.Text)      & parent.gRowSep   
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      = strDel & strVal

	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   

End Function

Sub ShowMSG(msg)

End Sub

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
    lgIntFlgMode = parent.OPMD_UMODE    
	Call SetToolbar("110011110011111")                                              '☆: Developer must customize
	Frm1.vspdData.Focus
	Call ggoOper.LockField(Document, "Q")
	Call SetSpreadColorOk()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
	Call SetToolbar("1111111111111111")                                              '☆: Developer must customize
    If DbQuery() = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
End Sub


'========================================================================================================
' Name : OpenZipCode()
' Desc : developer describe this line 
'========================================================================================================
Function OpenZipCode(ZipCode,Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "우편번호 팝업"                             ' Popup Name
	arrParam(1) = "ADDRESS"                                       ' Table Name
	arrParam(2) = ZipCode                                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "우편코드"

    arrField(0) = "ZipCd"                                         ' Field명(0)
    arrField(1) = "Address"                                       ' Field명(1)

    arrHeader(0) = "우편번호"	                              ' Header명(0)
    arrHeader(1) = "주소"                                     ' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SubSetZipCode(arrRet,Row)
	End If	

End Function

'========================================================================================================
'
'
'========================================================================================================
Sub SubSetZipCode(arrRet,Row)

	With frm1.vspdData 
          .Row  = Row
          .Col  = C_FROMACCTNM
          .Text = arrRet(0)
          .Col  = C_TOACCTCD
          .Text = arrRet(1)
	End With

End Sub


Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strTempBankCd
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0

			arrParam(0) = "Minor Code팝업"						' 팝업 명칭 
			arrParam(1) = "B_Minor"								' TABLE 명칭 
			arrParam(2) = ""								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "Major_cd=" & FilterVar("A1037", "''", "S") & " "									' Where Condition
			arrParam(5) = "Minor코드"

		    arrField(0) = "Minor_cd"								' Field명(0)
			arrField(1) = "Minor_Nm"								' Field명(1)

		    arrHeader(0) = "Minor Code"							' Header명(0)
			arrHeader(1) = "Minor Code명"						' Header명(1)

		Case 1, 2
			arrParam(0) = "계정코드팝업"						' 팝업 명칭 
			arrParam(1) = "A_ACCT"								' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "계정코드"

		    arrField(0) = "ACCT_CD"								' Field명(0)
			arrField(1) = "ACCT_NM"								' Field명(1)

		    arrHeader(0) = "계정코드"							' Header명(0)
			arrHeader(1) = "계정명"							' Header명(1)

		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, iWhere)
		'Call SetRefOpenAr(arrRet)
	End If	

End Function

 '**********************  SetReturnValue  ****************************************
'	기능: 기준 POP-UP에서 선택한 값을 Matching
'************************************************************************************** 

Function SetReturnVal(Byval arrRet, Byval iWhere)

	With frm1.vspdData
		Select Case iWhere
			Case 0	
				.Row = .ActiveRow
				.Col = C_Code
				.Text = arrRet(0)
				.Col = C_Codenm
				.Text = arrRet(1)
			Case 1
				.Row = .ActiveRow
				.Col = C_FROMACCTCD
				.Text = arrRet(0)
				.Col = C_FROMACCTNM
				.Text = arrRet(1)
				call vspdData_Change(.ActiveCol, .ActiveRow)
			Case 2
				.Row = .ActiveRow
				.Col = C_TOACCTCD
				.Text = arrRet(0)
				.Col = C_TOACCTNM
				.Text = arrRet(1)
				call vspdData_Change(.ActiveCol, .ActiveRow)
			Case Else
				Exit Function
		End Select
	End With

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

Dim strCode
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 Then
			Select Case Col
				Case C_CODEPopup
					.Col = C_CODE
					.Row = .ActiveRow
					strCode = .Text
					Call OpenPopup(strCode ,0)
				Case C_FROMACCTCDPopup
					.Col = C_FROMACCTCD
					.Row = .ActiveRow
					strCode = .Text
					Call OpenPopup(strCode,1)
				Case C_TOACCTCDPopup
					.Col = C_TOACCTCD
					.Row = .ActiveRow
					strCode = .Text
					Call OpenPopup(strCode,2)
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
         Case  C_TOACCTNM
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col	= C_TOACCTCDPopup
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)

	
	Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"	'Split 상태코드 
      

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
    	frm1.vspdData.Row = Row
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
   	Call SetSpreadLock

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
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub

Sub txtYear_DblClick(Button)
    If Button = 1 Then
       frm1.txtYear.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtYear.Focus       
    End If
End Sub

Sub txtYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

Function FncBtnauto()
	Dim IntRetCd

	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")

	If IntRetCD = vbNo Then
		Exit Function
	End if

	Call LayerShowHide(1)

	With frm1
		.txtMode.value		= parent.UID_M0002
		.hstryear.value		= .txtYear.text 	

    END With
    Call ExecMyBizASP(frm1, BIZ_PGM_ID2)

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정대체등록</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>연도</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5304ma1_fpYear_txtYear.js'></script></TD>
                                    <TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5304ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnauto" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnauto()" Flag=1>자동생성</BUTTON>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hstryear"   TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

