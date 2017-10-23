<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        :
'*  3. Program ID           : ZM000MA1
'*  4. Program Name         : 멀티컴퍼니접속정보등록 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) :
'*  8. Modified date(Last)  : 2005/04/07
'*  9. Modifier (First)     : MJG
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit
'=======================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'=======================================================================================================================
Dim interface_Account

Const BIZ_PGM_ID = "KZM000MB1.asp"

Dim	C_CompanyCd			'법인코드 
Dim	C_CompanyPopup      '법인팝업 
Dim	C_CompanyNm         '법인명 
Dim C_McFlg				'법인구분 
Dim	C_Url               '법인URL
Dim	C_UseYn             '사용여부 

Dim lsBtnProtectedRow
Dim lgQuery
Dim lgCopyRow

Dim gblnWinEvent
'=======================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
End Sub

'=======================================================================================================================
Sub InitComboBox()
     ggoSpread.SetCombo "발주"  & vbtab & "수주"				        , C_McFlg
End Sub

'=======================================================================================================================
Sub initSpreadPosVariables()
	C_CompanyCd			= 1      '법인코드 
	C_CompanyPopup  	= 2      '법인팝업 
	C_CompanyNm     	= 3      '법인명 
	C_McFlg				= 4		 '법인구분 
	C_Url           	= 5      '법인URL
	C_UseYn         	= 6      '사용여부 
End Sub
'=======================================================================================================================
Sub SetDefaultVal()
    frm1.txtCompanyCd.focus
	Set gActiveElement = document.activeElement
	Call SetToolbar("1100110100001111")
	interface_Account = GetSetupMod(parent.gSetupMod, "a")
End Sub
'=======================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'=======================================================================================================================
 Sub InitSpreadSheet()

	Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20070302",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.ReDraw = false

        .MaxCols = C_UseYn + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
        .Col = .MaxCols									'☜: 공통콘트롤 사용 Hidden Column
        .ColHidden = True

        .MaxRows = 0


		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_CompanyCd			, "법인", 10,,,18,2
		ggoSpread.SSSetButton 	C_CompanyPopup
		ggoSpread.SSSetEdit 	C_CompanyNm     	, "법인명", 20,,,50,2
		ggoSpread.SSSetCombo	C_McFlg             , "법인구분코드", 15
		ggoSpread.SSSetEdit 	C_Url           	, "URL", 30,,,100,2
		ggoSpread.SSSetCheck	C_UseYn				, "사용여부", 15,,,true

		Call ggoSpread.MakePairsColumn(C_CompanyCd, C_CompanyPopup)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		'//Call ggoSpread.SSSetSplit2(2)
		Call SetSpreadLock

		.ReDraw = true
	End With

End Sub
'=======================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CompanyCd		= iCurColumnPos(1)
			C_CompanyPopup  = iCurColumnPos(2)
			C_CompanyNm     = iCurColumnPos(3)
			C_McFlg			= iCurColumnPos(4)
			C_Url           = iCurColumnPos(5)
			C_UseYn         = iCurColumnPos(6)
    End Select
End Sub


'============================= 2.2.4 SetSpreadLock() ====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
 Sub SetSpreadLock()
	With frm1.vspdData
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData

		ggoSpread.SpreadLock		C_CompanyCd					         , -1, -1
		ggoSpread.SpreadLock		C_CompanyPopup					     , -1, -1
		ggoSpread.SpreadLock		C_CompanyNm					         , -1, -1
		ggoSpread.SSSetRequired		C_McFlg				                 , -1, -1
		ggoSpread.SSSetRequired		C_Url				                 , -1, -1
		'ggoSpread.SpreadLock		C_UseYn				                 , -1, -1

		.ReDraw = True
	End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData
		.ReDraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetRequired		C_CompanyCd					         , pvStartRow, pvEndRow
		'ggoSpread.SpreadLock		C_CompanyPopup					 , pvStartRow, pvEndRow
		ggoSpread.SpreadLock		C_CompanyNm					         , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_McFlg				         , pvStartRow,pvEndRow
		ggoSpread.SSSetRequired		C_Url				     , pvStartRow, pvEndRow
		'ggoSpread.SpreadLock		C_UseYn				         , pvStartRow, pvEndRow




		.ReDraw = True
	End With
End Sub


'=======================================================================================================================
Function OpenCompany()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "법인"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.txtCompanyCd.Value)
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('C','S','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "법인"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "법인"
    arrHeader(1) = "법인명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtCompanyCd.Value= arrRet(0)
		frm1.txtCompanyNm.value= arrRet(1)
		frm1.txtCompanyCd.focus
	End If

End Function


Function OpenCompanyCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
	frm1.vspdData.Col=C_CompanyCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(0) = "법인"						<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>

	arrParam(2) = Trim(frm1.vspdData.Text)		<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtORGNm.Value)	<%' Name Cindition%>

	arrParam(4) = "BP_TYPE = 'CS' AND IN_OUT_FLAG = 'O'"
	arrParam(5) = "법인"							<%' TextBox 명칭 %>

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "법인"
    arrHeader(1) = "법인명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCompanyCd(arrRet)
	End If
End Function

'=======================================================================================================================
Function SetCompanyCd(byval arrRet)

	frm1.vspdData.Col = C_CompanyCd
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col  = C_CompanyNm
	frm1.vspdData.Text = arrret(1)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow frm1.vspdData.ActiveRow

End Function



'=======================================================================================================================
Sub Form_Load()
	call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call SetDefaultVal
	Call InitSpreadSheet
    Call InitVariables
    Call InitComboBox
End Sub
'=======================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

 	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0101111111")         '화면별 설정 
	End If

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		lgSpdHdrClicked = 0		'2003-03-01 Release 추가 
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)

	 	'------ Developer Coding part (End)
 	End If

End Sub
'=======================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)

End Sub
'=======================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub
'=======================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 그리드 열고정을 한다.
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub


'=======================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=======================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData()
	CALL DbQueryOk()
End Sub
'=======================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   Dim iColumnName

	If Row <= 0 Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End if
End Sub
'=======================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim strType

	If lgQuery = true then Exit Sub
	If lgCopyRow = true then Exit Sub

	If lsBtnProtectedRow >= Row Then Exit Sub

	frm1.vspdData.ReDraw = False
	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData

    If Row > 0 And Col = C_CompanyPopup Then
        .Col = Col
        .Row = Row
        Call OpenCompanyCd()

    End If

    End With

	frm1.vspdData.ReDraw = True

End Sub
'=======================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

 	Select Case Col
		Case C_CompanyCd		'// 법인 
			Call LookUpData(Col, Row, Trim(GetSpreadText(frm1.vspdData, C_CompanyCd, Row, "X", "X")))
	End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
'=======================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
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
    End if

End Sub
'=======================================================================================================================
Function FncQuery()
    Dim IntRetCD

	FncQuery = False

	Err.Clear

	ggoSpread.Source = frm1.vspdData

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")
	Call InitVariables

	If Not ChkField(Document, "1") Then
		Exit Function
	End If

	If DbQuery = False Then Exit Function

	FncQuery = True
    Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")
    Call SetDefaultVal
    Call InitVariables

    FncNew = True
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncDelete()

    Exit Function
    Err.Clear

    FncDelete = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If

    If DbDelete = False Then
       Exit Function
    End If

    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")

    FncDelete = True
    Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim index

	FncSave = False

	ggoSpread.Source = frm1.vspdData

	Err.Clear

	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001","X","X","X")
		Exit Function
	End If

	If interface_Account = "N" then

		for index=1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = index
			frm1.vspdData.Col = 0
			If frm1.vspdData.Text = ggoSpread.InsertFlag Or frm1.vspdData.Text = ggoSpread.UpdateFlag Or frm1.vspdData.Text = ggoSpread.DeleteFlag Then
				Call frm1.vspdData.SetText(M_TransType,	index, "*")
			End if
		Next
	End if

	If Not ChkField(Document, "2") Then
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If

	If DbSave = False Then Exit Function

	FncSave = True
    Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncCopy()
	If frm1.vspdData.Maxrows < 1 Then Exit Function

	lgCopyRow = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
'    Call frm1.vspdData.SetText(M_IvType,	frm1.vspdData.ActiveRow, "")
    frm1.vspdData.ReDraw = True
	lgCopyRow = False
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
    Dim imRow, iRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End if
    End If

	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			Call frm1.vspdData.SetText(C_UseYn,	iRow,	"1")
		Next
		.vspdData.ReDraw = True

	End With

	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim iDelRowCnt, i
    if frm1.vspdData.Maxrows < 1 then exit function

    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData
	lDelRows = ggoSpread.DeleteRow
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncExcel()
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncFind()
	Call parent.FncFind(parent.C_MULTI, False)
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	ggoSpread.Source = frm1.vspdData

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
	Set gActiveElement = document.ActiveElement
End Function
'=======================================================================================================================
Function DbQuery()

	Err.Clear

	DbQuery = False

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtCompanyCd=" & Trim(frm1.hdnCompanyCd.value)
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.hdnUseflg.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtCompanyCd=" & Trim(frm1.txtCompanyCd.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		if frm1.rdoUsageFlgAll.checked = True then
			strVal = strVal & "&rdoUsageFlg=" & ""
		elseif frm1.rdoUsageFlgYes.checked = True then
			strVal = strVal & "&rdoUsageFlg=" & "Y"
		else
			strVal = strVal & "&rdoUsageFlg=" & "N"
		end if

	End If

	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery = True

End Function
'=======================================================================================================================
Function DbQueryOk()

	Dim index

	lgIntFlgMode = parent.OPMD_UMODE

	Call ggoOper.LockField(Document, "Q")
	Call SetToolbar("1100111100111111")

	frm1.vspdData.ReDraw = False

	lgQuery = False

	frm1.vspdData.ReDraw = True
End Function
'=======================================================================================================================
Function DbSave()

    Err.Clear

    Dim lRow
    Dim lGrpCnt
	Dim strVal
	Dim PvArr
	Dim iColSep

    DbSave = False

	If LayerShowHide(1) = False Then
	     Exit Function
	End If

	With frm1
	 .txtMode.value = parent.UID_M0002

	 lGrpCnt = 0
	 strVal = ""
	 iColSep = parent.gColSep
	 ReDim PvArr(0)

	 For lRow = 1 To .vspdData.MaxRows

		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag, ggoSpread.DeleteFlag

				If  .vspdData.Text = ggoSpread.InsertFlag then
					strVal = "C" & iColSep    '☜: C=Create
				ElseIf  .vspdData.Text = ggoSpread.UpdateFlag then
					strVal = "U" & iColSep    '☜: U=Update
				Else
					strVal = "D" & iColSep    '☜: D=Delete
				End if

				strVal = strVal & lRow & iColSep

				.vspdData.Col = C_CompanyCd:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep

				.vspdData.Col = C_McFlg
				if Trim(.vspdData.Text) = "발주" then
				strVal = strVal &"M"& iColSep
				else
				strVal = strVal &"S"& iColSep
				end if

				.vspdData.Col = C_Url:		strVal = strVal & Trim(.vspdData.Text) & iColSep

				.vspdData.Col = C_UseYn
				if Trim(.vspdData.Text) = "1" then
				strVal = strVal &"Y"& iColSep
				else
				strVal = strVal &"N"& iColSep
				end if

				strVal = strVal & parent.gRowSep


				ReDim Preserve PvArr(lGrpCnt)
				PvArr(lGrpCnt) = strVal
				lGrpCnt = lGrpCnt + 1
	   End Select
	 Next

	 .txtMaxRows.value = lGrpCnt -1
	 .txtSpread.value = Join(PvArr, "")


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 

	End With

    DbSave = True                                                           '⊙: Processing is NG

End Function
'=======================================================================================================================
Function DbSaveOk()
    Call InitVariables
    frm1.vspdData.MaxRows = 0

    Call MainQuery()
End Function
'=======================================================================================================================
'--------------------------------------------------------------------------------------------------
' Name : LookUpData(ByVal Col, ByVal lRow, ByVal sCode)
' Desc :
'--------------------------------------------------------------------------------------------------
Function LookUpData(ByVal Col, ByVal lRow, ByVal sCode)
	LookUpData = False

	Dim IntRetCD
	Dim strSelect
	Dim strTbl
	Dim strWhere

   On Error Resume Next
    Err.Clear

	Select Case Col
		Case C_CompanyCd 		'// 법인 

			strSelect = " BP_NM "
			strTbl    = " B_BIZ_PARTNER "
			strWhere  = " BP_TYPE = 'CS' AND IN_OUT_FLAG = 'O' AND BP_CD = " & FilterVar(sCode, "''", "S")

			If CommonQueryRs(strSelect, strTbl, strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				Call DisplayMsgBox("970000", Parent.VB_INFORMATION,"법인","X")
				With frm1.vspdData
					Call .SetText(C_CompanyNm,  lRow, "")
				End With
				Exit Function
			End If

			lgF0 = Split(lgF0, Chr(11))

			With frm1.vspdData
				Call .SetText(C_CompanyNm,  lRow, lgF0(0))
                .Row = lRow
				.Col = C_CompanyCd
				.Action = Parent.SS_ACTION_ACTIVE_CELL
			End With
	End Select

	LookUpData = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--
'#########################################################################################################
'            6. Tag부 
'#########################################################################################################
-->
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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니 접속정보</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래법인</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="거래법인" NAME="txtCompanyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCompany() ">
														   <INPUT TYPE=TEXT ALT="거래법인" NAME="txtCompanyNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>사용여부</TD>
									<TD CLASS=TD6 NOWRAP><input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgAll" value="A" tag = "11" checked><label for="rdoUsageFlgAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgYes" value="Y" tag = "11"><label for="rdoUsageFlgYes">사용</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoUsageFlg" id="rdoUsageFlgNo" value="N" tag = "11"><label for="rdoUsageFlgNo">미사용</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<script language =javascript src='./js/kzm000ma1_OBJECT1_vspdData.js'></script>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>

    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<!--<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<td WIDTH="*" ALIGN="RIGHT"><a href="VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:WriteCookiePage()">구매요청조회</a></td>
					<td WIDTH="20"></td>
				</tr>
			</table>-->
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="<%=BIZ_PGM_ID%>" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCompanyCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUseflg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
