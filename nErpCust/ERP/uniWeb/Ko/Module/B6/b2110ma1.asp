<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 기준 
'*  3. Program ID           : B2110MA1
'*  4. Program Name         : BIZ Unit 등록 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/22
'*  8. Modified date(Last)  : 2000/03/22
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "b2110mb1.asp"												'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================


Dim C_BizUnitCd    															'☆: Spread Sheet의 Column별 상수 
Dim C_BizUnitNm
Dim C_BizUnitEngNm
Dim IsOpenPop

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================= 
Sub initSpreadPosVariables()
	C_BizUnitCd = 1
	C_BizUnitNm = 2
	C_BizUnitEngNm = 3
End Sub


Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKeyIndex = ""

    lgSortKey = 1
    lgPageNo   = 0
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
'		For intRow = 1 To .MaxRows
'			.Row = intRow
'			.Col = C_TYPECd
'            intIndex = .value                            'Index , not value 
'			.col = C_TYPENm
'			.value = intindex
'		Next	
	End With
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  

	With frm1.vspdData
		.ReDraw = False 

		.MaxCols = C_BizUnitEngNm + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit C_BizUnitCd, "사업부코드", 30 ,,,10, 2 
		ggoSpread.SSSetEdit C_BizUnitNm, "사업부명", 44,,,30
        ggoSpread.SSSetEdit C_BizUnitEngNm, "사업부영문명", 44,,,30

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true

	End With
    Call SetSpreadLock 

End Sub


'========================================================================================

Sub SetSpreadLock()
    With frm1

		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_BizUnitCd,-1, C_BizUnitCd
		ggoSpread.SSSetRequired C_BizUnitNm, -1, C_BizUnitNm
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True

    End With
End Sub


'========================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)		'ByVal lRow
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired	C_BizUnitCd, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	C_BizUnitNm, pvStarRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BizUnitCd = iCurColumnPos(1)
			C_BizUnitNm = iCurColumnPos(2)
			C_BizUnitEngNm = iCurColumnPos(3)
	End Select
End Sub


 '------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode, Byval iWhere)'
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'IsOpenPop = True

	arrParam(0) = "사업부 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_UNIT"							' TABLE 명칭 
	arrParam(2) = strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "사업부"			

    arrField(0) = "BIZ_UNIT_CD"							' Field명(0)
    arrField(1) = "BIZ_UNIT_NM"							' Field명(1)

    arrHeader(0) = "사업부코드"					' Header명(0)
    arrHeader(1) = "사업부명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizUnitCd.focus
		Exit Function
	Else
		Call SetItemInfo(arrRet, iWhere)
	End If	

End Function


 '------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet, Byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtBizUnitCd.focus
			.txtBizUnitCd.value = arrRet(0)
			.txtBizUnitNm.value = arrRet(1)
			lgBlnFlgChgValue = True
		End if
	End With

End Function

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

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================= 
Sub Form_Load()

'	Call GetGlobalVar
'   Call ClassLoad                                                          '⊙: Load Common DLL
    Call LoadInfTB19029 
    
    Call ggoOper.LockField(Document, "N")                                                         '⊙: Load table , B_numeric_format
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolBar("1100110100101111")		
    frm1.txtBizUnitCd.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

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
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True

End Sub
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If
		
    End With

End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgPageNo <> "" Then                         
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
    
End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If Trim(frm1.txtBizUnitCd.value) = "" Then
		frm1.txtBizUnitNm.value = ""
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery
    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Err.Clear
End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD 

    FncDelete = False
    Err.Clear

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '☆:
        Exit Function
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    If IntRetCD = vbNo Then
       Exit Function
    End If

    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 

    FncSave = False    
    Err.Clear 

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
  
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave	

    FncSave = True
End Function


'========================================================================================
Function FncCopy()
    Dim  IntRetCD

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_BizUnitCd
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True

End Function


'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.EditUndo
End Function


'========================================================================================
Function FncInsertRow(Byval pvRowCnt) 
	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

		If ImRow="" then
			Exit Function
		End If
	End If

	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		lgBlnFlgChgValue = True									'Indicates that value changed
		'.vspdData.EditMode = True
        ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
    End With
    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow
		lgBlnFlgChgValue = True

    End With
End Function


'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function


'========================================================================================
Function FncPrev() 
    On Error Resume Next
End Function


'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function



'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
End Function



'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim StrNextKey
    Dim pB21019         'As New P21018ListIndReqSvr

    DbQuery = False
    Call LayerShowHide(1)
    Err.Clear

	Dim strVal

    With frm1

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode			=" & parent.UID_M0001
			strVal = strVal		& "&txtBizUnitCd	=" & Trim(.hBizUnitCd.value)	'☆: 조회 조건 데이타 
			strVal = strVal		& "&lgStrPrevKey	=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode			=" & parent.UID_M0001							'☜: 
			strVal = strVal		& "&txtBizUnitCd	=" & Trim(.txtBizUnitCd.value)		'☆: 조회 조건 데이타 
			strVal = strVal		& "&lgStrPrevKey	=" & lgStrPrevKey
		End If
			strVal = strVal		& "&lgPageNo		=" & lgPageNo
			strVal = strVal		& "&txtMaxRows		=" & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()

    lgIntFlgMode = parent.OPMD_UMODE
	If frm1.vspdData.MaxRows > 0 Then
		Call SetToolBar("1100111100111111")                                              '☆: Developer must customize
	Else
		Call SetToolBar("1100111100101111")                                              '☆: Developer must customize
	End If

    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================
Function DbSave() 
    Dim pB21011     'As New P21011ManageIndReqSvr
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt 
	Dim strVal
	
    DbSave = False
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1

    strVal = ""

    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData.MaxRows

        .vspdData.Row = lRow
        .vspdData.Col = 0

        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag											'☜: 신규 

													strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep  				'☜: C=Create, Row위치 정보 
                .vspdData.Col = C_BizUnitCd		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BizUnitNm		:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BizUnitEngNm	:   strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep              

'				lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag											'☜: 수정 

													strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep				'☜: U=Update, Row위치 정보 
                .vspdData.Col =  C_BizUnitCd	:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col =  C_BizUnitNm	:   strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C_BizUnitEngNm  :	strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep        

'				lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag											'☜: 삭제 

													strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep				'☜: D=Delete, Row위치 정보 
                .vspdData.Col =   C_BizUnitCd	:   strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep									

'				lGrpCnt = lGrpCnt + 1
        End Select

    Next

	.txtMode.value        = parent.UID_M0002
'	.txtUpdtUserId.value  = parent.gUsrID
'	.txtInsrtUserId.value = parent.gUsrID
'	.txtMaxRows.value     = lGrpCnt - 1	
	.txtSpread.value      = strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 

	End With

    DbSave = True
End Function


'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	'Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field    ggoSpread.ssdeleteflag 1
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call SetSpreadLock

	Call InitVariables

	FncQuery

End Function


'========================================================================================
Function DbDelete() 
End Function


'========================================================================================
' Function Name : txtBizUnitCd_OnChange
' Function Desc : 사업부명 Query
'===================================================================================
Sub txtBizUnitCd_OnChange()

	Dim strval
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    strval = frm1.txtBizUnitCd.value

    Call  CommonQueryRs( "BIZ_UNIT_NM" , "B_BIZ_UNIT  " , "BIZ_UNIT_CD =  " & FilterVar(strval , "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    frm1.txtBizUnitNm.value = Replace(lgF0,chr(11),"")	

End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					<TD WIDTH=*>&nbsp;</TD>
					</TD>			
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
									<TD CLASS="TD5" NOWRAP>사업부</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtBizUnitCd" MAXLENGTH="10" SIZE=15 ALT ="사업부코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItemInfo(txtBizUnitCd.value,0)">&nbsp;
														   <INPUT NAME="txtBizUnitNm" MAXLENGTH="30" SIZE=30 ALT ="사업부명"   tag="14X"></TD>
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
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		    <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hBizUnitCd" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

