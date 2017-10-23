<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 인사마스타등록 
*  3. Program ID           : H9102ma1
*  4. Program Name         : H9102ma1
*  5. Program Desc         : 연말정산관리/연말정산/소득.세액공제등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/04
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "hb001mb1.asp"						           '☆: Biz Logic ASP Name
Const TAB1 = 1
Const TAB2 = 2

Dim C_USE_DT
Dim C_DAILY_DED_AMT
Dim C_DAILY_TAX_RATE
Dim C_DAILY_TAX_DED_RATE

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

Dim IsOpenPop						                                    ' Popup

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_USE_DT			= 1
	C_DAILY_DED_AMT		= 2
	C_DAILY_TAX_RATE	= 3
	C_DAILY_TAX_DED_RATE = 4

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then 
		    Exit Function
		End if	
		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If

End Function

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtUSE_DT.Year = strYear 
	frm1.txtUSE_DT.Month = strMonth 

End Sub


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM
	strMaskYM = "9999-99"
	Call initSpreadPosVariables()  
	With frm1.vspdData

		Call AppendNumberPlace("6","2","0")
		Call AppendNumberPlace("7","10","0")
		
        ggoSpread.Source = frm1.vspdData
	
		ggoSpread.Spreadinit "V20060601",,parent.gForbidDragDropSpread    
	    .ReDraw = false    
        .MaxCols = C_DAILY_TAX_DED_RATE + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True	  
        .MaxRows = 0 
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData     

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetMask    C_USE_DT, "기준년월", 10, 2, strMaskYM
		
        ggoSpread.SSSetFloat	C_DAILY_DED_AMT,		"일일근로소득공제금액", 20,"7", ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_DAILY_TAX_RATE,		"산출세율(%)", 15, "6",			ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
		ggoSpread.SSSetFloat	C_DAILY_TAX_DED_RATE,	"세액공제율(%)", 15, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
        
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

			C_USE_DT			= iCurColumnPos(1)
			C_DAILY_DED_AMT		= iCurColumnPos(2)
			C_DAILY_TAX_RATE	= iCurColumnPos(3)
			C_DAILY_TAX_DED_RATE = iCurColumnPos(4)

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	Dim iMaxRows
	
    With frm1
    
    iMaxRows = .vspdData.MaxRows
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock    C_USE_DT, -1, C_USE_DT
	ggoSpread.SSSetRequired		C_DAILY_DED_AMT, 1, iMaxRows
	ggoSpread.SSSetRequired		C_DAILY_TAX_RATE, 1, iMaxRows
	ggoSpread.SSSetRequired		C_DAILY_TAX_DED_RATE, 1, iMaxRows
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
    
    ggoSpread.SSSetRequired		C_USE_DT, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DAILY_DED_AMT, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DAILY_TAX_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DAILY_TAX_DED_RATE, pvStartRow, pvEndRow
        
    .vspdData.ReDraw = True
    
    End With
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Replace(Frm1.txtUSE_DT.Text, "-", "") & parent.gColSep                                           'You Must append one character(parent.gColSep)
'    lgKeyStream = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
End Sub        


'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)

	Call ggoOper.FormatDate(frm1.txtUSE_DT, parent.gDateFormat,2)

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
	Call SetDefaultVal

	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
	Call SetToolbar("1100000000000111")												'⊙: Set ToolBar

	
	Call InitVariables

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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                           '⊙: Initializes local global variables

	Call MakeKeyStream("X")
    
    Call  DisableToolBar( Parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
       
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgIntFlgMode <>  Parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  Parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call  DisableToolBar( Parent.TBC_DELETE)
	If DBDelete=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '☜: Processing is OK
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
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
    Call Initdata()
End Function
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
    End Select    
            
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

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
	
    Set gActiveElement = document.ActiveElement   

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
    
    Call MakeKeyStream("X")
    
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
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
	Call Parent.FncExport( Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( Parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

    If LayerShowHide(1)=False Then
		Exit Function
    End If


    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

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
                                                    strVal = strVal & "C" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_USE_DT	  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_DED_AMT      : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_TAX_RATE	  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_TAX_DED_RATE	  : strVal = strVal & Trim(.vspdData.value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_USE_DT	  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_DED_AMT      : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_TAX_RATE	  : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_DAILY_TAX_DED_RATE	  : strVal = strVal & Trim(.vspdData.value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_USE_DT	  : strVal = strVal & Trim(.vspdData.value) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
		.txtKeyStream.value = lgKeyStream
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
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal

	lgIntFlgMode      =  Parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

	Call SetToolbar("1100111100100111")

    Call  ggoOper.LockField(Document, "Q")
    Call SetSpreadLock 
    
    Set gActiveElement = document.ActiveElement   
End Function

Function DbQueryFail()

	Call SetToolbar("1100111100100111")

End Function
		
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables	
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call SetToolbar("1100111100100111")												'⊙: Set ToolBar
	Call InitVariables()
	Call MainNew()	
End Function

'=======================================
'   Event Name :txtUSE_DT_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtUSE_DT_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtUSE_DT.Action = 7
        frm1.txtUSE_DT.focus
    End If
End Sub

Sub txtUSE_DT_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="AUTO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23><% ' 탭위치 %>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><img src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>급여관련사항</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR><% ' 탭위치 종료 %>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=10%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	        	<TD CLASS="TD5" NOWRAP>기준년월</TD>
			    	    		<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtUSE_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="기준년월" tag="11X1" id=txtUSE_DT></OBJECT>');</SCRIPT></TD>
			    	    		<TD CLASS="TD5" NOWRAP></TD>
			    	    		<TD CLASS="TD6"></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>


		
            <TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
   	            <TR>
   	                <TD WIDTH=100% VALIGN="TOP" HEIGHT="*">
                        <TABLE width=100%>
							<TR>
							    <TD VALIGN=TOP colspan="2">
							    </TD>    
							</TR>
					    </TABLE>
    
                    </TD>
                </TR>
            </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24"><INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>

</BODY>
</HTML>

