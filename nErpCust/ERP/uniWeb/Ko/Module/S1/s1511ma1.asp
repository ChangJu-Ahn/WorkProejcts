<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1511MA1
'*  4. Program Name         : 품목그룹별구성비등록 
'*  5. Program Desc         : 품목그룹별구성비등록 
'*  6. Comproxy List        : PS1G117.dll, PS1G118.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/03/26
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : choinkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/27 : Grid성능 적용, Kang Jun Gu
'=======================================================================================================
%>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s1511mb1.asp"												'☆: 비지니스 로직 ASP명 

Dim C_Item_Cd
Dim C_Item_Nm
Dim C_ItemSpec
Dim C_Percent
Dim C_ChgFlg

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
	C_Item_Cd             = 1
	C_Item_Nm             = 2
	C_ItemSpec            = 3
	C_Percent		      = 4
	C_ChgFlg			  = 5
End Sub

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtItemGroup.focus
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData

	    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	    .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
		.MaxCols = C_ChgFlg            '☜: 최대 Columns의 항상 1개 증가시킴 
		
	    Call GetSpreadColumnPos("A")
		.ReDraw = false
	    ggoSpread.SSSetEdit     C_Item_Cd,              "품목" ,40,,,18,2  
        ggoSpread.SSSetEdit     C_Item_Nm,              "품목명",50, 0 
        ggoSpread.SSSetEdit		C_ItemSpec,				"규격",			20  	
		ggoSpread.SSSetFloat    C_Percent,              "구성비(%)" ,28,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
  	    ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 1, 2    	    
		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)
'		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column

		.ReDraw = true 
	    SetSpreadLock "", 0, -1, ""
    End With
End Sub

'========================================================================================================
	Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
	    With frm1
			ggoSpread.Source = .vspdData

			.vspdData.ReDraw = False

			ggoSpread.SpreadLock C_Item_Nm, lRow, -1
			ggoSpread.SpreadLock C_ItemSpec, lRow, -1
			ggoSpread.SpreadUnLock C_Percent, lRow, -1

			.vspdData.ReDraw = True
		End With
	End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected   C_Item_Cd,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Nm,			 pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_ItemSpec,			 pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Percent,		     pvStartRow, pvEndRow

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
			C_Item_Cd             = iCurColumnPos(1)
			C_Item_Nm             = iCurColumnPos(2)
			C_ItemSpec             = iCurColumnPos(3)
			C_Percent		        = iCurColumnPos(4)
			C_ChgFlg				= iCurColumnPos(5)
    End Select    
End Sub

'========================================================================================================
Function OpenItemGroup(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "품목그룹"							<%' 팝업 명칭 %>
	arrParam(1) = "B_ITEM_GROUP"							<%' TABLE 명칭 %>
	arrParam(4) = "LEAF_FLG = " & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "품목그룹"							<%' TextBox 명칭 %>
	
	arrField(0) = "ITEM_GROUP_CD"							<%' Field명(0)%>
	arrField(1) = "ITEM_GROUP_NM"							<%' Field명(1)%>
    
	arrHeader(0) = "품목그룹"							<%' Header명(0)%>
	arrHeader(1) = "품목그룹명"							<%' Header명(1)%>

	arrParam(2) = Trim(frm1.txtItemGroup.Value)				<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtItemGroupNm.Value)			<%' Name Cindition%>
	
	arrParam(3) = ""	
	
	frm1.txtItemGroup.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemGroup(arrRet, iWhere)		                      
	End If	
	
End Function
	
'========================================================================================================
	Function OpenItemRef()
		Dim iCalledAspName
		Dim IntRetCD
		Dim lblnWinEvent
		
		Dim arrRet
		Dim strParam

		On Error Resume Next
		
		lblnWinEvent = False
		If lblnWinEvent = True Then Exit Function
		
		If Trim(frm1.txtItemGroup.Value) = "" Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If	
		
		strParam = ""
		strParam = strParam & Trim(frm1.txtItemGroup.value) & parent.gColSep
		strParam = strParam & Trim(frm1.txtItemGroupNm.Value)		

		iCalledAspName = AskPRAspName("S1511RA1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S1511RA1", "X")
			lblnWinEvent = False
			Exit Function
		End If
	
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False


		If arrRet(0, 0) = "" Then
			If Err.number <> 0 Then Err.Clear 
			Exit Function
		Else
			Call SetItemRef(arrRet)
		End If	
	End Function
	

'========================================================================================================
	Function SetItemGroup(Byval arrRet, Byval iWhere)
		With frm1		
			.txtItemGroup.value = arrRet(0) 
			.txtItemGroupNm.value = arrRet(1) 		
			.txtItemGroup.focus 	
		End With
	End Function
	
'========================================================================================================
	Function SetItemRef(arrRet)		
		Dim TempRow, I, j
		Dim intLoopCnt
		Dim intCnt
		Dim blnEqualFlg			
		Dim intCntRow		
		
		With frm1
			
			ggoSpread.Source = .vspdData			
			.vspdData.ReDraw = False	

			TempRow = .vspdData.MaxRows											<% '☜: 현재까지의 MaxRows %>
			intLoopCnt = Ubound(arrRet, 1)										<% '☜: Reference Popup에서 선택되어진 Row만큼 추가 %>
			intCntRow = 0	

			For intCnt = 1 to intLoopCnt	
				blnEqualFlg = False
										
				If blnEqualFlg = false then
					intCntRow = intCntRow + 1
					.vspdData.MaxRows = CLng(TempRow) + CLng(intCntRow)
					.vspdData.Row = CLng(TempRow) + CLng(intCntRow)

					.vspdData.Col = 0
					.vspdData.Text = ggoSpread.InsertFlag


					<% '품목' %>
					.vspdData.Col = C_Item_Cd
					.vspdData.text = arrRet(intCnt - 1, 0)
					<% '품목명' %>
					.vspdData.Col = C_Item_Nm										
					.vspdData.text = arrRet(intCnt - 1, 1)
					<% '규격' %>
					.vspdData.Col = C_ItemSpec
					.vspdData.text = arrRet(intCnt - 1, 2)
					<% '구성비' %>
					.vspdData.Col = C_Percent			
					.vspdData.text = arrRet(intCnt - 1, 3)
					SetSpreadColor CLng(TempRow) + CLng(intCntRow), CLng(TempRow) + CLng(intCntRow)
				End if
			Next
									
			.vspdData.ReDraw = True
			.vspdData.focus
		End With
		lgBlnFlgChgValue = True
		Call SumPercent
		Call SetToolbar("1110100100001111")
	End Function	

	
'========================================================================================================
	Sub SumPercent()
		Dim dblTotPercent
		Dim intCnt
		Dim intCntDeleteFlag
		
		dblTotPercent = 0
		intCntDeleteFlag = 0
				
		With frm1
			ggoSpread.Source = .vspdData
			.txtHItemGroupRate.value = 0	
			For intCnt=1 to .vspdData.MaxRows
				.vspdData.Col = 0
				.vspdData.Row = intCnt
				If .vspdData.text <> ggoSpread.DeleteFlag Then
					.vspdData.Col = C_Percent
					.vspdData.Row = intCnt
					If .vspdData.text <>"" Then
						dblTotPercent = dblTotPercent + UNICDbl(.vspdData.text)
					End If	
				ELSE
					intCntDeleteFlag = intCntDeleteFlag + 1	
				End If
			Next
			
			'모든 ROW를 삭제로 선택한 경우 
			IF intCntDeleteFlag = .vspdData.MaxRows THEN 
				.txtHItemGroupRate.value = 100				
			END IF 
			
			.txtPercent.text = dblTotPercent
		End With
	End Sub	

	
'========================================================================================================
'=	Event Name : ItemRateSort
'=  Event Desc : Column 별 정렬.
'=======================================================================================================
Sub ItemRateSort(ByVal SortCol, ByVal intKey)

    frm1.vspdData.BlockMode = True
    frm1.vspdData.Col = 0
    frm1.vspdData.Col2 = frm1.vspdData.MaxCols
    frm1.vspdData.Row = 1
    frm1.vspdData.Row2 = frm1.vspdData.MaxRows
    
    'Row기준 Sort
    frm1.vspdData.SortBy = 0
    
    'Sort기준 Column
    frm1.vspdData.SortKey(1) = SortCol
    
    '정렬방법 
    frm1.vspdData.SortKeyOrder(1) = intKey  '0: 정렬None 1 :오름차순  2: 내림차순 
    frm1.vspdData.Action = 25				'SS_ACTION_SORT : VB number
    
    frm1.vspdData.BlockMode = False
    
End Sub

'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)		'⊙: Format Contents  Field
	
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitSpreadSheet
	Call SetDefaultVal
    Call InitVariables     
    Call SetToolbar("1110100000001111")										'⊙: 버튼 툴바 제어 

End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")

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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
	Call SumPercent    
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							<% '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
				Call DbQuery()
			End If
		End If
	End With
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess Then  Exit Sub

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery()
    	End If
    End if	    

End Sub

'========================================================================================================
	Function FncQuery()
		Dim IntRetCD

		FncQuery = False													<% '⊙: Processing is NG %>

		Err.Clear															<% '☜: Protect system from crashing %>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			<% '⊙: "Will you destory previous data" %>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '⊙: Clear Contents  Field %>
		Call InitVariables													<% '⊙: Initializes local global variables %>

		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then							<% '⊙: This function check indispensable field %>
			Exit Function
		End If

		<% '------ Query function call area ------ %>
		Call DbQuery()														<% '☜: Query db data %>

		FncQuery = True														<% '⊙: Processing is OK %>
	End Function
	
'========================================================================================================
	Function FncNew()
		Dim IntRetCD 

		FncNew = False														<% '☜: Protect system from crashing %>

		<% '------ Check previous data area ------ %>
		If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		<% '------ Erase condition area ----- %>
		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "A")								<%'⊙: Clear Condition Field%>
		Call ggoOper.LockField(Document, "N")								<%'⊙: Lock  Suitable  Field%>
		Call SetDefaultVal
		Call SetToolbar("1110100000001111")									<% '⊙: 버튼 툴바 제어 %>
		Call InitVariables													<%'⊙: Initializes local global variables%>

		FncNew = True														<%'⊙: Processing is OK%>

	End Function
	
'========================================================================================================
	Function FncDelete()
		Dim IntRetCD

		FncDelete = False												<% '⊙: Processing is NG %>
		
		<% '------ Precheck area ------ %>
		If lgIntFlgMode <> parent.OPMD_UMODE Then								<% 'Check if there is retrived data %>
			Call DisplayMsgBox("900002","x","x","x")
			Exit Function
		End If

		IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")

		If IntRetCD = vbNo Then
			Exit Function
		End If

		<% '------ Delete function call area ------ %>
		Call DbDelete													<% '☜: Delete db data %>

		FncDelete = True												<% '⊙: Processing is OK %>
	End Function
	
'========================================================================================================
	Function FncSave()
		Dim IntRetCD
		
		FncSave = False																		<% '⊙: Processing is NG %>
		
		Err.Clear																			<% '☜: Protect system from crashing %>
		
		<% '------ Precheck area ------ %>
		If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then								<% 'Check if there is retrived data %>
		    IntRetCD = DisplayMsgBox("900001","x","x","x")					<% '⊙: No data changed!! %>
		    Exit Function
		End If
		
		<% '------ Check contents area ------ %>
		ggoSpread.Source = frm1.vspdData

		If Not chkField(Document, "2") Then		<% '⊙: Check contents area %>
			Exit Function
		End If

		If Not ggoSpread.SSDefaultCheck Then		<% '⊙: Check contents area %>
			Exit Function
		End If
		
		<% '------ Save function call area ------ %>
		Call DbSave																			<% '☜: Save db data %>
		
		FncSave = True																		<% '⊙: Processing is OK %>
	End Function

'========================================================================================================
	Function FncCopy()
		frm1.vspdData.ReDraw = False

		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		Call SumPercent()		
		frm1.vspdData.ReDraw = True
	End Function

'========================================================================================================
	Function FncCancel() 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo														<%'☜: Protect system from crashing%>
		Call SumPercent()		
	End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
  With frm1
		FncInsertRow = False                                                         '☜: Processing is NG

		If Not chkField(Document, "2") Then
		Exit Function
		End If

		If IsNumeric(Trim(pvRowCnt)) Then
		    imRow = CInt(pvRowCnt)
		Else
		    imRow = AskSpdSheetAddRowCount()
		    If imRow = "" Then
		        Exit Function
		    End If
		End If
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.Row= .vspdData.ActiveRow
		.vspdData.ReDraw = True
  End With
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    Set gActiveElement = document.ActiveElement   
   
   lgBlnFlgChgValue = True
End Function

'========================================================================================================
	Function FncDeleteRow()
		Dim lDelRows
		Dim iDelRowCnt, i
	
		With frm1.vspdData 
			If .MaxRows = 0 Then
				Exit Function
			End If

			.focus
			ggoSpread.Source = frm1.vspdData

			lDelRows = ggoSpread.DeleteRow
			Call SumPercent()		
			lgBlnFlgChgValue = True
		End With
	End Function

'========================================================================================================
	Function FncPrint()
	    ggoSpread.Source = frm1.vspdData
		Call parent.FncPrint()													<%'☜: Protect system from crashing%>
	End Function

'========================================================================================================
	Function FncExcel() 
		Call parent.FncExport(parent.C_SINGLEMULTI)
	End Function

'========================================================================================================
	Function FncFind() 
		Call parent.FncFind(parent.C_SINGLEMULTI, False)
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
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(-1, -1)
End Sub

'========================================================================================================
	Function FncExit()
		Dim IntRetCD

		FncExit = False

		If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			<%'⊙: "Will you destory previous data"%>

			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If

		FncExit = True
	End Function

'========================================================================================================
	Function DbQuery()
	
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		Dim strVal

		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemGroup=" & Trim(frm1.txtHItemGroup.value)		<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemGroup=" & Trim(frm1.txtItemGroup.value)		<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		End If
		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
	
		DbQuery = True														<%'⊙: Processing is NG%>
	End Function
	
'========================================================================================================
	Function DbSave() 	
		Dim IntRetCD		

'son (frm1.txtPercent.text) --> UNICDbl(frm1.txtPercent.text)바꾸어줌 

		If UNICDbl(frm1.txtPercent.text) <> 100 AND UNICDbl(frm1.txtHItemGroupRate.value) <> 100 Then
			IntRetCD = DisplayMsgBox("201020","x","x","x")
			Call SetToolbar("1110101100011111")			
			Exit Function			
		End If
	
		Dim lRow
		Dim lGrpCnt
		Dim strVal, strDel
		Dim intInsrtCnt

		DbSave = False														<% '⊙: Processing is OK %>
    
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		With frm1
			.txtMode.value = parent.UID_M0002
			.txtUpdtUserId.value = parent.gUsrID
			.txtInsrtUserId.value = parent.gUsrID

			lGrpCnt = 1

			strVal = ""
			strDel = ""
	
			For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = 0

				Select Case .vspdData.Text
					Case ggoSpread.InsertFlag								<% '☜: 신규 %>
						strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep	<% '☜: C=Create, Row위치 정보 %>
						
						.vspdData.Col = C_Item_Cd								<% '2 %>
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

						.vspdData.Col = C_Percent								<% '3 %>
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep

						lGrpCnt = lGrpCnt + 1
		
					Case ggoSpread.UpdateFlag								<% '☜: Update %>
						strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	<% '☜: U=Update, Row위치 정보 %>
						
						.vspdData.Col = C_Item_Cd								<% '2 %>
						strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
						
						.vspdData.Col = C_Percent								<% '3 %>
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gRowSep

						lGrpCnt = lGrpCnt + 1
	
					Case ggoSpread.DeleteFlag								<% '☜: 삭제 %>
						strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep	<% '☜: D=Update, Row위치 정보 %>
						
						.vspdData.Col = C_Item_Cd								<% '2 %>
						strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

						lGrpCnt = lGrpCnt + 1

				End Select
			Next

			.txtMaxRows.value = lGrpCnt-1	
			.txtSpread.value = strDel & strVal			
		 			
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)						<% '☜: 비지니스 ASP 를 가동 %>

		End With

		DbSave = True														<% '⊙: Processing is NG %>
	End Function

'========================================================================================================
	Function DbQueryOk()													<% '☆: 조회 성공후 실행로직 %>
		<% '------ Reset variables area ------ %>
		lgIntFlgMode = parent.OPMD_UMODE											<% '⊙: Indicates that current mode is Update mode %>
		lgBlnFlgChgValue = False

		Call SumPercent		
		Call ggoOper.LockField(Document, "Q")								<% '⊙: This function lock the suitable field %>
		Call SetToolbar("1110101100011111")									<% '⊙: 버튼 툴바 제어 %>	
		frm1.vspdData.focus
	End Function
	
'========================================================================================================
	Function DbSaveOk()														<%'☆: 저장 성공후 실행 로직 %>
		Call ggoOper.ClearField(Document, "2")
		Call InitVariables
		Call MainQuery()
	End Function

</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹별구성비</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>						    
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenItemRef">품목참조</A></TD>
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
								<TD CLASS="TD5" NOWRAP>품목그룹</TD>
								<TD CLASS="TD6"><INPUT NAME="txtItemGroup" ALT="품목그룹" TYPE="Text" MAXLENGTH="15" SiZE="15"  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemGroup 0">&nbsp;<INPUT NAME="txtItemGroupNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP>총구성비</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s1511ma1_fpDoubleSingle1_txtPercent.js'></script>&nbsp;%</TD>
										</TR>
									</TABLE>	
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
					<TD WIDTH=100% HEIGHT= 100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>							
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s1511ma1_vaSpread_vspdData.js'></script>
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
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHItemGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHItemGroupRate" TAG="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>

</BODY>
</HTML>		


