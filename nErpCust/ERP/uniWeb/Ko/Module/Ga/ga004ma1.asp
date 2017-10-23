
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        : 
*  3. Program ID           : GA004MA1
*  4. Program Name         : 경영손익작업진행조회 
*  5. Program Desc         : 경영손익작업진행조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/18
*  8. Modified date(Last)  : 2001/12/18
*  9. Modifier (First)     : Lee Kang Yeong
* 10. Modifier (Last)      : Lee Tae Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History        
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID = "GA004mb1.asp"												'Biz Logic ASP
Const CookieSplit = 1233

Dim C_JOB_GRP  
Dim C_JOB_CD  																'Spread Sheet의 Column별 상수 
Dim C_JOB_NM  													
Dim C_JOB_START  
Dim C_JOB_END  
Dim C_PROGRESS_FG  
Dim C_ERROR_NUM  
Dim C_ERROR_POP  

'Const C_SHEETMAXROWS    = 21													'한 화면에 보여지는 최대갯수*1.5%>
Const C_SHEETMAXROWS_D  = 30													'☆: Server에서 한번에 fetch할 최대 데이타 건수 

Dim IscookieSplit 

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

	Dim StartDate
	Dim EndDate
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
		
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "G", "COOKIE", "QA") %>
End Sub

'========================================================================================================
Function CookiePage(ByVal flgs)
	Dim strCookie, i

	Const CookieSplit = 4877						

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	WriteCookie CookieSplit , IsCookieSplit

End Function

'========================================================================================================
Sub MakeKeyStream(pRow)

	Dim strYYYYMM
	Dim strYear,strMonth,strDay 

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	strYYYYMM	= strYear & strMonth
	lgKeyStream	= strYYYYMM & Parent.gColSep                                           'You Must append one character(Parent.gColSep)

End Sub        

'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
End Sub

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_JOB_GRP			= 1	  
	C_JOB_CD  			= 2													'Spread Sheet의 Column별 상수 
	C_JOB_NM  			= 3										
	C_JOB_START         = 4
	C_JOB_END			= 5
	C_PROGRESS_FG		= 6
	C_ERROR_NUM			= 7
	C_ERROR_POP			= 8
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
	
       .MaxCols = C_ERROR_POP + 1                                                   ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True
       
       .Col = C_JOB_CD
       .ColHidden = True                                                            ' ☜:☜:
    
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021127", ,parent.gAllowDragDropSpread

		ggoSpread.ClearSpreadData
		
	   .ReDraw = false
	
		Call AppendNumberPlace("6","2","0")

		Call GetSpreadColumnPos("A")     
		     
       ggoSpread.SSSetEdit  C_JOB_GRP ,      "작업단계"              ,20,,, 30
       ggoSpread.SSSetEdit  C_JOB_CD ,      "경영손익작업코드"              ,2,,, 30,2
       ggoSpread.SSSetEdit  C_JOB_NM ,       "경영손익작업"              ,30,,, 40
       ggoSpread.SSSetEdit  C_JOB_START ,    "시작시간"           ,28,2,, 30
       ggoSpread.SSSetEdit  C_JOB_END ,      "종료시간"           ,28,2,, 30
       ggoSpread.SSSetEdit  C_PROGRESS_FG ,          "작업상태"              ,10,2,, 40,2
       ggoSpread.SSSetEdit  C_ERROR_NUM ,      "ERROR갯수"              ,13,1,, 30,2
       ggoSpread.SSSetButton  C_ERROR_POP
       
       Call ggoSpread.MakePairsColumn(C_ERROR_NUM,C_ERROR_POP)
       
       Call ggoSpread.SSSetColHidden(C_JOB_CD,C_JOB_CD,True)
   
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
      
      ggoSpread.SpreadUnLock C_ERROR_POP, -1, C_ERROR_POP,-1
End Sub

'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
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

			C_JOB_GRP		= iCurColumnPos(1)
			C_JOB_CD		= iCurColumnPos(2)
			C_JOB_NM		= iCurColumnPos(3)    
			C_JOB_START 	= iCurColumnPos(4)
			C_JOB_END		= iCurColumnPos(5)
			C_PROGRESS_FG	= iCurColumnPos(6)
			C_ERROR_NUM		= iCurColumnPos(7)
			C_ERROR_POP		= iCurColumnPos(8)    
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
End Sub

'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
    frm1.txtYYYYMM.focus

    Call InitComboBox
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

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
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
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    
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
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
Function FncInsertRow() 
	 Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    
    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	imRow = AskSpdSheetAddRowcount()
	
	If imRow = "" Then
		Exit function
	End If
		
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
       .vspdData.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    
    IF Err.number = 0 Then
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
    On Error Resume Next                                                  '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                   '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear																			'☜: Clear err status

	if LayerShowHide(1) = False then
	   Exit Function
	end if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)            '☜: Max fetched data at a time
    End With
		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)													  '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
Function DbSave() 

End Function

'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                            '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
Function DbQueryOk()													     
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")										    '버튼 툴바 제어 
	
End Function

'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    
    Call InitVariables															'⊙: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================================
Function DbDeleteOk()

End Function


'=======================================
Sub txtYyyymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
    End If
End Sub

'=======================================================================================================
Sub txtYyyymm_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 21, 22, 23     ' 학교 
		        .vspdData.Col = C_SCHOOL
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_SCHOOL_NM
		    	.vspdData.text = arrRet(1)   
		    Case 31, 32         ' 전공 
		        .vspdData.Col = C_MAJOR
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_MAJOR_NM
		    	.vspdData.text = arrRet(1)   
        End Select

	End With

End Function

'========================================================================================================
Sub txtEnd_dt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtEnd_dt.Action = 7
    End If
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData,Col,Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_ERROR_POP
            Call OpenCode("", C_ERROR_POP, Row)
    End Select    
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = false then    		                         
      			Call RestoreToolBar()
      			Exit sub
      		End if
    	End If
    End if
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

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

	If Row < 1 Then Exit Sub

	IscookieSplit = ""
	
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
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

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

Sub vspdData_LeaveCell(Col, Row, NewCol, NewRow, Cancel)
    
 '   frm1.vspdData.OperationMode = 3             

End Sub

'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim yyyymm, Job_cd, strYYYYMM
	Dim strYear,strMonth,strDay 
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth   
	yyyymm = FilterVar(strYYYYMM, "''", "S")

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_JOB_CD
	Job_cd = FilterVar(frm1.vspdData.Text, "''", "S")
	

	Select Case iWhere
	    Case C_ERROR_POP
	        arrParam(0) = "ERROR목록팝업"										' 팝업 명칭 
	    	arrParam(1) = "G_ERROR"													' TABLE 명칭 
	    	arrParam(2) = strCode                   								' Code Condition
	    	arrParam(3) = ""														' Name Cindition
	    	arrParam(4) = " YYYYMM = " & yyyymm & " and JOB_CD = " & Job_cd         ' Where Condition

	    	arrParam(5) = "ERROR목록" 											' TextBox 명칭 
			
			arrField(0) = "ED08" & parent.gColSep & "SEQ"	     			            <%' Field명(1)%>
			arrField(1) = "ED99" & parent.gColSep & "ERROR_CONTENTS"					<%' Field명(0)%>
    
	    	arrHeader(0) = "ERROR번호"
	    	arrHeader(1) = "ERROR설명"	   		    							' Header명(0)	    	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=615px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
'		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
'        ggoSpread.UpdateRow Row
	End If	

End Function


</SCRIPT>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>경영손익작업진행</font></td>
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
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
			            		<TD CLASS=TD5 NOWRAP>대상년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/ga004ma1_fpLoanDtFr_txtYyyymm.js'></script>&nbsp;</TD>
			            		<TD CLASS="TDT"></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/ga004ma1_vaSpread1_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

