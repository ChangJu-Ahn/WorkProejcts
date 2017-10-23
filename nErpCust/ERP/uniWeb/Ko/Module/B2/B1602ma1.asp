
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Province)
'*  3. Program ID           : B1601ma1.asp
'*  4. Program Name         : B1601ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/07
'*  7. Modified date(Last)  : 2002/12/13
'*  8. Modifier (First)     : Hwang Jeong-won,Lee Seok Gon
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

Const BIZ_PGM_ID = "B1602mb1.asp"


<!-- #Include file="../../inc/lgvariables.inc" -->
 
Dim C_Province 
Dim C_ProvinceNm
Dim C_Remarks
Dim C_Country

Dim IsOpenPop          
Dim lgStrPrevKey2

Sub InitSpreadPosVariables()
    C_Province      = 1
    C_ProvinceNm    = 2
    C_Remarks       = 3
    C_Country       = 4
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_Country + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetEdit C_Province, "지방코드", 20,,,5,2
    ggoSpread.SSSetEdit C_ProvinceNm, "지방", 48,,,50
    ggoSpread.SSSetEdit C_Remarks, "비고", 30,,,30
    ggoSpread.SSSetEdit C_Country, "국가코드", 18
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_Province, -1, C_Province
    ggoSpread.SpreadLock C_Country, -1, C_Country
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired C_Province, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Country, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Province      = iCurColumnPos(1)
            C_ProvinceNm    = iCurColumnPos(2)
            C_Remarks       = iCurColumnPos(3)
            C_Country       = iCurColumnPos(4)
    End Select    
End Sub

Function OpenCountry()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "국가 팝업"				<%' 팝업 명칭 %>
	arrParam(1) = "b_country"					<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtCountryCd.value		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "국가"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "country_cd"					<%' Field명(0)%>
    arrField(1) = "country_nm"					<%' Field명(1)%>
    
    arrHeader(0) = "국가코드"					<%' Header명(0)%>
    arrHeader(1) = "국가"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtCountryCd.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCountry(arrRet)
	End If	
	
End Function

Function OpenProvince()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtCountryCd.value) = "" Then
		Call DisplayMsgBox("970021", "X", "국가코드", "X")
		Exit Function
	Else
		IsOpenPop = True

		arrParam(0) = "지방 팝업"				<%' 팝업 명칭 %>
		arrParam(1) = "b_province"					<%' TABLE 명칭 %>
		arrParam(2) = frm1.txtProvince.value		<%' Code Condition%>
		arrParam(3) = ""							<%' Name Cindition%>
		arrParam(4) = "country_cd =  " & FilterVar(frm1.txtCountryCd.value, "''", "S") & ""	<%' Where Condition%>
		arrParam(5) = "지방"					<%' 조건필드의 라벨 명칭 %>
		
		arrField(0) = "province_cd"					<%' Field명(0)%>
		arrField(1) = "province_nm"					<%' Field명(1)%>
    
	    arrHeader(0) = "지방코드"					<%' Header명(0)%>
		arrHeader(1) = "지방"						<%' Header명(1)%>
    
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
		IsOpenPop = False
		
		frm1.txtProvince.focus 
	
		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetProvince(arrRet)
		End If	
	End If
	
End Function

Function SetCountry(Byval arrRet)
	With frm1
		.txtCountryCd.value = arrRet(0)
		.txtCountryNm.value = arrRet(1)
	End With
End Function

Function SetProvince(Byval arrRet)
	With frm1
		.txtProvince.value = arrRet(0)
		.txtProvinceNm.value = arrRet(1)
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    
    frm1.txtCountryCd.value = parent.gCountry
    frm1.txtCountryCd.focus
    Call FncQuery
    
End Sub

Sub txtCountryCd_onChange()
	frm1.txtProvince.value = ""
	frm1.txtProvinceNm.value = ""
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		If Col = C_RegionNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_Region
			.TypeComboBoxCurSel = index
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
    	If lgStrPrevKey <> "" And lgStrPrevKey2 <> "" Then                  <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
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

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    frm1.txtCountryNm.value = ""
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
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
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
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
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR   '⊙: Check contents area
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
            SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			.vspdData.Col = C_Province
			.vspdData.Text = ""
    
			.vspdData.ReDraw = True
		End If
	End With

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim iRow 
    
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
    
	If frm1.txtCountryCd.value = "" Then                                      'Check if there is retrived data
		Call DisplayMsgBox("970021", "X", "국가코드", "X")
        Exit Function
	End If

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'.vspdData.EditMode = True
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1   
        .vspdData.Row = iRow
		.vspdData.Col = C_Country
		.vspdData.Text = UCase(.txtCountryCd.value)
		Next
		.vspdData.ReDraw = True
    
		SetSpreadColor .vspdData.ActiveRow
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
    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
        strVal = strVal & "&txtCountry=" & Trim(.txtCountryCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProvince=" & Trim(.txtProvince.value)			'☆: 조회 조건 데이타 
       
	Call RunMyBizASP(MyBizASP, strVal)									    	'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>

	Dim iNameArr
    Dim iDx    
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
      
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
	
End Function

Function DbLookUp()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 
       
	call CommonQueryRs("country_nm "," b_country "," country_cd =  " & FilterVar(frm1.txtCountryCd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtCountrynm.value = Trim(Replace(lgF0,Chr(11),"")) 
    
    call CommonQueryRs("province_nm "," b_province "," province_cd =  " & FilterVar(frm1.txtProvince.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtProvincenm.value = Trim(Replace(lgF0,Chr(11),""))       
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
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
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
					strVal = strVal & "C" & parent.gColSep						'☜: C=Create
		        Case ggoSpread.UpdateFlag								'☜: 수정 
					strVal = strVal & "U" & parent.gColSep						'☜: U=Update
			End Select
			
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			'☜: 수정, 신규 
					
		            .vspdData.Col = C_Country		'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Province		'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ProvinceNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Remarks		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.DeleteFlag								'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep						

		            .vspdData.Col = C_Country		'4
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Province		'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>지방코드</font></td>
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
									<TD CLASS="TD5">국 가</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtCountryCd" SIZE=10 MAXLENGTH=2 tag="12XXXU"  ALT="국가코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCountry()">
										<INPUT TYPE=TEXT NAME="txtCountryNm" tag="14X">
									</TD>
									<TD CLASS="TD5">지 방</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtProvince" SIZE=10 MAXLENGTH=5 tag="11XXXU"  ALT="지방코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProvince()">
										<INPUT TYPE=TEXT NAME="txtProvinceNm" tag="14X">
									</TD>
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
									<script language =javascript src='./js/b1602ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1602mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hCountryCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProvinceCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

