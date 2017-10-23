
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit)
'*  3. Program ID           : B1f01ma1.asp
'*  4. Program Name         : B1f01ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/08
'*  7. Modified date(Last)  : 2002/06/21
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit								

Const BIZ_PGM_ID = "D2113MB1.ASP"												<%'비지니스 로직 ASP명 %>

Dim C_BP_CD
Dim C_BP_POP
Dim C_BP_NM
Dim C_RV_CD
Dim C_RV_NM
Dim C_DESC
 

Dim IsOpenPop          
Dim lgSortKey1
Dim lgSortKey2

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
    C_BP_CD		= 1
    C_BP_POP	= 2
    C_BP_NM		= 3
    C_RV_CD    = 4
    C_RV_NM    = 5
    C_DESC	   = 6

End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","BA") %>
    
End Sub

Sub InitSpreadSheet()
   Call initSpreadPosVariables()  

   With frm1.vspdData
      ggoSpread.Source = frm1.vspdData	
      'patch version
      ggoSpread.Spreadinit "V20090922",,parent.gAllowDragDropSpread    

      .ReDraw = false

      .MaxCols = C_DESC + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
      .Col = .MaxCols														'☆: 사용자 별 Hidden Column
      .ColHidden = True    

      .MaxRows = 0
      ggoSpread.ClearSpreadData

      Call GetSpreadColumnPos("A")  

      ggoSpread.SSSetEdit		C_BP_CD, "거래처",      15,,,15,2
      ggoSpread.SSSetButton	C_BP_POP
      ggoSpread.SSSetEdit		C_BP_NM, "거래처명",    20,,,20,2
      ggoSpread.SSSetCombo		C_RV_CD, "역발행업무",  15
      ggoSpread.SSSetCombo	   C_RV_NM, "역발행업무",  15
      ggoSpread.SSSetEdit		C_DESC,	"비고",        40,,,40,1

      Call ggoSpread.MakePairsColumn(C_BP_CD, C_BP_POP)

		Call ggoSpread.MakePairsColumn(C_RV_CD, C_RV_NM, "1")
      Call ggoSpread.SSSetColHidden(C_RV_CD, C_RV_CD, True)

      .ReDraw = true

		Call InitComboBox()
      Call SetSpreadLock 
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired	C_BP_CD,	-1, -1
    ggoSpread.SSSetProtected	C_BP_NM,	-1, -1
    
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_BP_CD,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_BP_NM,	pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor1(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected C_BP_CD, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_BP_CD	= iCurColumnPos(1)
            C_BP_POP	= iCurColumnPos(2)
            C_BP_NM  = iCurColumnPos(3)
            C_RV_CD	= iCurColumnPos(4)
            C_RV_NM	= iCurColumnPos(5)
            C_DESC   = iCurColumnPos(6)
    End Select
End Sub

'========================================================================================
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
   '자료유형(Data Type)
   Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("DT005", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    

   iCodeArr = vbTab & lgF0
   iNameArr = vbTab & lgF1

   ggoSpread.SetCombo Replace(iCodeArr, Chr(11), vbTab), C_RV_CD			'COLM_DATA_TYPE
   ggoSpread.SetCombo Replace(iNameArr, Chr(11), vbTab), C_RV_NM
End Sub

Function Open_User(Byval strCode, Byval iWhere)
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)
   Dim OriginCol, TempCd
   Dim IntRetCD

   If IsOpenPop = True Then Exit Function

   IsOpenPop = True

   arrParam(0) = "bp_cd, bp_nm"			<%' 팝업 명칭 %>
   arrParam(1) = "b_biz_partner"			<%' TABLE 명칭 %>
   arrParam(2) = strCode					<%' Code Condition%>
   arrParam(4) = ""						   <%' Name Cindition%>
   arrParam(5) = "거래처"				   <%' 조건필드의 라벨 명칭 %>

   arrField(0) = "bp_cd"				   <%' Field명(0)%>
   arrField(1) = "bp_nm"				   <%' Field명(1)%>

   arrHeader(0) = "거래처"					<%' Header명(0)%>
   arrHeader(1) = "거래처명"			   <%' Header명(1)%>

   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
                                   Array(arrParam, arrField, arrHeader), _
                                   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False


   If arrRet(0) = "" Then
      Exit Function
   Else
      Call SetUser(arrRet, iWhere)
   End If	
	
End Function

Function Open_Partner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "bp_cd, bp_nm"				<%' 팝업 명칭 %>
	arrParam(1) = "b_biz_partner"				<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBpCd.value		<%' Code Condition%>
	arrParam(4) = ""							   <%' Name Cindition%>
	arrParam(5) = "거래처"					   <%' 조건필드의 라벨 명칭%>
		
	arrField(0) = "bp_cd"					   <%' Field명(0)%>
	arrField(1) = "bp_nm"					   <%' Field명(1)%>

	arrHeader(0) = "거래처"				   <%' Header명(0)%>
	arrHeader(1) = "사용자명"				   <%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
	                                Array(arrParam, arrField, arrHeader), _
		                             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
	End If	
	
End Function


Function SetUser(Byval arrRet, Byval iWhere)
	With frm1 
			.vspdData.Col = C_BP_CD
			.vspdData.Text = arrRet(0)
			
			.vspdData.Col = C_BP_NM
			.vspdData.Text = arrRet(1)

			lgBlnFlgChgValue = True
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                           <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
          
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>   
 
    Call SetToolbar("1100111100101111")										<%'버튼 툴바 제어 %>
  
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	'----------  Coding part  -------------------------------------------------------------   
	' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData

		.Row = Row

		Select Case Col
			Case  C_RV_NM
				.Col = Col
				intIndex = .Value
				.Col = C_RV_CD
				.Value = intIndex
		End Select
	End With
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
End Sub


Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub


Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
	
End Sub


Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub


Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    


Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


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
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_BP_POP Then
		    .Row = Row
		    .Col = C_BP_CD

		    Call Open_User(.Text, 1)
		End If
    End With
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
    	If (lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "") Then <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End if
    
End Sub

Function FncQuery() 
	Dim IntRetCD 

	FncQuery = False                                                        

	Err.Clear                                                               <%'Protect system from crashing%>

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    If frm1.txtBpCd.value = "" then
	    frm1.txtBpNm.value = ""
    End If

	Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitVariables                                                      <%'Initializes local global variables%>

<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = False Then Exit Function		  					<%'Query db data%>
       
    FncQuery = True
            
End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
    

    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data

    FncSave = True

End Function

Function FncCopy() 

    frm1.vspdData.ReDraw = False

    if frm1.vspdData.maxrows < 1 then exit function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    'Key field clear
    frm1.vspdData.Col=C_BP_CD
    frm1.vspdData.Text=""

    frm1.vspdData.Col = C_BP_NM
    frm1.vspdData.Text=""

    frm1.vspdData.ReDraw = True

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow
	Dim iRow 

	'On Error Resume Next                                                          '☜: If process fails
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
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
		ggoSpread.InsertRow ,imRow
		ggoSpread.SSSetRequired		C_BP_CD,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		ggoSpread.SSSetProtected	C_BP_NM,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtBpCd=" & .hUserId.value 			'☆: 조회 조건 데이타  			 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)			'☆: 조회 조건 데이타
		strVal = strVal & "&txtBpNm=" & Trim(.txtBpNm.value) 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    End If        
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
       
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
   
	Call SetToolbar("1100111100101111")										<%'버튼 툴바 제어 %>

	Call SetSpreadColor1(-1, -1)
End Function

Function DbSave() 
	Dim lRow        
	Dim lGrpCnt  
	Dim strVal, strDel

	DbSave = False                                                          

	Call LayerShowHide(1)
	'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

    For lRow = 1 To .vspdData.MaxRows
		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text
			Case ggoSpread.InsertFlag								'☜: 신규 
				strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'☜: C=Create
				.vspdData.Col = C_BP_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_RV_CD : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_DESC : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
				lGrpCnt = lGrpCnt + 1
			Case ggoSpread.UpdateFlag								'☜: 수정 
				strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update
				.vspdData.Col = C_BP_CD : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_RV_CD	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C_DESC : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep		            
				lGrpCnt = lGrpCnt + 1
			Case ggoSpread.DeleteFlag								'☜: 삭제 
				strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update
				.vspdData.Col = C_BP_CD	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
				lGrpCnt = lGrpCnt + 1
		End Select
	Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>

	End With

	DbSave = True
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>역발행거래처관리</font></td>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">거래처</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10  MAXLENGTH=13 tag="11XXXU" ALT="ERP사용"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Open_Partner()">
										<INPUT TYPE=TEXT NAME="txtBpNm" tag="14X">
									</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1f01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

