<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211MA3
'*  4. Program Name         : 판매계획기간정보수정 
'*  5. Program Desc         : 판매계획기간정보수정 
'*  6. Comproxy List        : PS2G213.dll
'*  7. Modified date(First) : 2003/01/11
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "s2211mb3.asp"        '☆: Head Query 비지니스 로직 ASP명 

Const C_PopFrSpPeriod = 1
Const C_PopToSpPeriod = 2

Dim C_SpType ' 1           '☆: Spread Sheet의 Column별 상수 
Dim C_SpPeriod ' 2
Dim C_SpPeriodDesc ' 3
Dim C_FromDt ' 4
Dim C_ToDt ' 5
Dim C_SpYear ' 6
Dim C_SpQuarter ' 7
Dim C_SpMonth ' 8
Dim C_SpWeek ' 9
Dim C_SpCreateMethod ' 10

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgLngStartRow
Dim lgStrWhere

Dim IsOpenPop       'Popup

'========================================================================================================
Sub initSpreadPosVariables()  

	C_SpType = 1
	C_SpPeriod = 2
	C_SpPeriodDesc = 3
	C_FromDt = 4
	C_ToDt = 5
	C_SpYear = 6
	C_SpQuarter = 7
	C_SpMonth = 8
	C_SpWeek = 9
	C_SpCreateMethod = 10

End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0      
	lgSortKey = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboConSpType.focus
End Sub

'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
 Call initSpreadPosVariables()    
	
 With frm1.vspdData
 
	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021120",,parent.gAllowDragDropSpread    		
    
	.ReDraw = false

	.MaxCols = C_SpCreateMethod+1             '☜: 최대 Columns의 항상 1개 증가시킴 
	
    Call AppendNumberPlace("6","2","0")
    Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit C_SpType, "계획구분", 10,2,,1,2
	ggoSpread.SSSetEdit C_SpPeriod, "계획기간", 10,2,,8,2
	ggoSpread.SSSetEdit C_SpPeriodDesc, "계획기간설명", 30
	ggoSpread.SSSetEdit C_FromDt, "시작일", 12
	ggoSpread.SSSetEdit C_ToDt, "종료일", 12
	ggoSpread.SSSetEdit C_SpYear, "년", 10,2,,4
    ggoSpread.SSSetFloat C_SpQuarter,"분기" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","4"
    ggoSpread.SSSetFloat C_SpMonth,"월" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","12"
    ggoSpread.SSSetFloat C_SpWeek,"주" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","54"
	ggoSpread.SSSetEdit C_SpCreateMethod, "생성방법", 10,2,,2
	 
	Call ggoSpread.SSSetColHidden(C_SpType,C_SpType,True)
	Call ggoSpread.SSSetColHidden(C_SpCreateMethod,.MaxCols,True)
	Call SetSpreadLock
	.ReDraw = true
 End With
    
End Sub

'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_SpType, -1, C_SpPeriod
	ggoSpread.SpreadLock C_FromDt, -1, C_SpYear
End Sub

'========================================================================================================
Sub SetQuerySpreadColor()
	Dim iCnt

    With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired  C_SpPeriodDesc,lgLngStartRow, .vspdData.MaxRows
		ggoSpread.SSSetRequired  C_SpQuarter,	lgLngStartRow, .vspdData.MaxRows

     	For iCnt = lgLngStartRow To .vspdData.MaxRows
			.vspdData.Row = iCnt
			.vspdData.Col = C_SpCreateMethod
			Select Case .vspdData.Text
				Case "10" '월 
					ggoSpread.SSSetProtected C_SpMonth, iCnt, iCnt
					ggoSpread.SSSetProtected C_SpWeek, iCnt, iCnt
				Case "30" '순 
					ggoSpread.SSSetProtected C_SpMonth, iCnt, iCnt
					ggoSpread.SSSetProtected C_SpWeek, iCnt, iCnt
				Case "40" '주 
					ggoSpread.SSSetRequired  C_SpMonth, iCnt, iCnt
					ggoSpread.SSSetProtected C_SpWeek, iCnt, iCnt
				Case "50" '일 
					ggoSpread.SSSetProtected C_SpMonth, iCnt, iCnt
					ggoSpread.SSSetProtected C_SpWeek, iCnt, iCnt
			End Select
		Next

		.vspdData.ReDraw = True
    
    End With			

End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpType			= iCurColumnPos(1)
			C_SpPeriod			= iCurColumnPos(2)
			C_SpPeriodDesc		= iCurColumnPos(3)
			C_FromDt			= iCurColumnPos(4)
			C_ToDt				= iCurColumnPos(5)
			C_SpYear			= iCurColumnPos(6)
			C_SpQuarter			= iCurColumnPos(7)
			C_SpMonth			= iCurColumnPos(8)
			C_SpWeek			= iCurColumnPos(9)
			C_SpCreateMethod	= iCurColumnPos(10)
			

	End Select
End Sub	

'========================================================================================================
Sub InitComboBox()
	' 판매계획유형 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConSpType,lgF0,lgF1,parent.gColSep)
End Sub


<% '========================================================================================================
'	Description : 판매계획기간 Popop
'======================================================================================================== %>
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(4)
	Dim iCalledAspName
	
	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s2211pa3")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", "X", "s2211pa3", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	With frm1
		iArrParam(4) = .cboConSpType.value
		
		Select Case pvIntWhere
		Case C_PopFrSpPeriod
			iArrParam(0) = .txtConFrSpPeriod.value
			.txtConFrSpPeriod.focus
		Case C_PopToSpPeriod
			iArrParam(0) = .txtConToSpPeriod.value
			.txtConToSpPeriod.focus
		End Select
	End With
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtDisplayFlag=N", Array(window.parent,iArrParam), _
	 "dialogWidth=690px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
		
End Function

'======================================================================================================== 
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopFrSpPeriod
			.txtConFrSpPeriod.value = pvArrRet(0) 
			.txtConFrSpPeriodDesc.value = pvArrRet(1)   
		Case C_PopToSpPeriod
			.txtConToSpPeriod.value = pvArrRet(0)
			.txtConToSpPeriodDesc.value = pvArrRet(1)
		End Select
	End With

	SetConPopup = True
End Function

<%
'======================================================================================================== 
' Function Desc : form_load시 jump cookie처리 
'=======================================================================================================
%>
Sub CookiePage()
	On Error Resume Next

	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal

	With frm1
		iStrTemp = ReadCookie(CookieSplit)
			
		If Trim(Replace(iStrTemp, parent.gColSep, "")) = "" then Exit Sub
			
		iArrVal = Split(iStrTemp, Parent.gColSep)
			
		.cboConSpType.value = iArrVal(0)
		.txtConFrSpPeriod.value = iArrVal(1)
		.txtConFrSpPeriodDesc.value = iArrVal(2)
		If .cboConSpType.value <> "" And .txtConFrSpPeriod.value <> "" Then
			Call DbQuery
		End IF
		WriteCookie CookieSplit , ""
	End With
End Sub

'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call SetDefaultVal 
	Call InitVariables              '⊙: Initializes local global variables
	Call InitSpreadSheet
	Call InitCombobox

	'폴더/조회/입력 
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전 
	'/다음/복사/엑셀 
	'/인쇄/찾기 
	Call SetToolBar("1100100000011111")          '⊙: 버튼 툴바 제어 
	Call CookiePage()

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111") 
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	
End Sub

'========================================================================================================= 
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then
			Exit Sub
		End If 
		   
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
		Exit Sub
		End If
	End if    

End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================= 
Function FncQuery()                                            
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

<%    '-----------------------
    'Check previous data area
    '----------------------- %>
	'************ 싱글/멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>


<%  '-----------------------
    'Query function call area
    '----------------------- %>
    Call DbQuery                <%'☜: Query db data%>

    FncQuery = True                <%'⊙: Processing is OK%>
        
End Function

'========================================================================================================= 
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	<%  '-----------------------
	'Precheck area
	'-----------------------%>
	'************ 싱글/멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If lgBlnFlgChgValue = False Or ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If

    
<%  '-----------------------
    'Check content area
    '-----------------------%>
	ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then     <%'⊙: Check contents area%>
       Exit Function
    End If
    
    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '-----------------------%>
    Call DbSave                                                    <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================= 
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================================= 
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'========================================================================================================= 
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

'========================================================================================================= 
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================= 
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()

	' 조회된 자료가 있는 경우	
	If gActiveSpdSheet.MaxRows > 0 Then
		lgLngStartRow = 1
		Call SetQuerySpreadColor()
	End If
End Sub

'========================================================================================================= 
Function FncExit()
Dim IntRetCD
FncExit = False
'************ 싱글/멀티인 경우 **************
ggoSpread.Source = frm1.vspdData 
If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
	If IntRetCD = vbNo Then
		Exit Function
	End If
End If
FncExit = True
End Function

'========================================================================================================= 
 Function DbQuery() 

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
	If   LayerShowHide(1) = False Then
             Exit Function 
        End If
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
    
    Dim iStrVal
    
    With frm1
	    iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>
	    If lgIntFlgMode = Parent.OPMD_UMODE Then    
   			iStrVal = iStrVal & lgStrWhere
		Else
			lgStrWhere = "&txtWhere="
			lgStrWhere = lgStrWhere & Trim(.cboConSpType.value) & parent.gColSep
			lgStrWhere = lgStrWhere & Trim(.txtConFrSpPeriod.value) & parent.gColSep
			lgStrWhere = lgStrWhere & Trim(.txtConToSpPeriod.value) & parent.gColSep
			
			iStrVal = iStrVal & lgStrWhere
		End If		
		iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey

		lgLngStartRow = .vspdData.MaxRows + 1
    End With
    
	Call RunMyBizASP(MyBizASP, iStrVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True               <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
	lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'⊙: This function lock the suitable field%>
    Call SetToolBar("1100100000011111")          '⊙: 버튼 툴바 제어 

	If Trim(lgStrPrevKey) = "" Then
		lgStrWhere = ""
    End If

    Call SetQuerySpreadColor
    
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus  
	Else
		frm1.txtConFrSpPeriod.focus
	End If    

End Function

'========================================================================================================= 
 Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
 
	If   LayerShowHide(1) = False Then
        Exit Function 
    End If
 
    DbSave = False                                                          '⊙: Processing is NG
    
	With frm1
		.txtMode.value = Parent.UID_M0002
   
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 0
	strVal = ""
	    
	'-----------------------
	'Data manipulate area
	'-----------------------
	For lRow = 1 To .vspdData.MaxRows
	    
		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text
			'Case ggoSpread.InsertFlag       '☜: 신규 
			Case ggoSpread.UpdateFlag       '☜: 수정 
			strVal = strVal & lRow & Parent.gColSep'☜: U=Update
		End Select
		   
		Select Case .vspdData.Text

			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag  '☜: 수정, 신규 

			.vspdData.Col = C_SpType
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			.vspdData.Col = C_SpPeriod
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			.vspdData.Col = C_SpPeriodDesc
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			.vspdData.Col = C_SpQuarter
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			.vspdData.Col = C_SpMonth
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			.vspdData.Col = C_SpWeek
			strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
			              
			strVal = strVal & Parent.gRowSep
		               
		End Select
		lGrpCnt = lGrpCnt + 1
	              
	Next

	.txtMaxRows.value = lGrpCnt
	.txtSpread.value = strDel & strVal
	  
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
	 
	End With
	 
	DbSave = True                                                           '⊙: Processing is NG
	    
	End Function

'========================================================================================================= 
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>
	Call ggoOper.LockField(Document, "N")
	Call InitVariables
		frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function

'========================================================================================================= 
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>판매계획기간정보수정</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
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
			<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
			<TD CLASS="TD6"><SELECT Name="cboConSpType" ALT="판매계획유형" tag="12XXXU"></SELECT></TD>
		</TR>
		<TR>
			<TD CLASS="TD5" NOWRAP>계획기간</TD>
			<TD CLASS="TD6"><INPUT NAME="txtConFrSPPeriod" ALT="계획기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConFrSPPeriod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopFrSpPeriod)">&nbsp;<INPUT NAME="txtConFrSPPeriodDesc" TYPE="Text" SIZE=25 tag="14">&nbsp;~&nbsp;
							<INPUT NAME="txtConToSPPeriod" ALT="계획기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConToSPPeriod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopToSpPeriod)">&nbsp;<INPUT NAME="txtConToSPPeriodDesc" TYPE="Text" SIZE=25 tag="14"></TD>
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
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <script language =javascript src='./js/s2211ma3_OBJECT1_vspdData.js'></script>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
 <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
