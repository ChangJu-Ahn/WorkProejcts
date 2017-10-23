<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Tracking Valid Check by Item Account)
'*  3. Program ID           : b1b07ma1.asp
'*  4. Program Name         : by Item Account Tracking No Control
'*  5. Program Desc         :
'*  6. Modified date(First) : 2006/06/23
'*  7. Modified date(Last)  : 2006/06/23
'*  8. Modifier (First)     : Lee, Seung Wook
'*  9. Modifier (Last)      : Lee, Seung Wook
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_QRY_ID = "B1B07MB1.asp"
Const BIZ_PGM_SAVE_ID = "B1B07MB2.asp"

Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_TrackingFlag

Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub initSpreadPosVariables()
	C_PlantCd = 1
	C_PlantNm = 2  
    C_ItemAcct  = 3
    C_ItemAcctNm  = 4
    C_TrackingFlag	= 5
End Sub

Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = "" 
    lgLngCurRows = 0
    lgSortKey = 1
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
		
			.Row = intRow
			.Col = C_PlantCd
			intIndex = .Value
			.Col = C_PlantNm
			.Value = intIndex
			
			.Row = intRow
			.col = C_ItemAcct
			intIndex = .value
			.Col = C_ItemAcctNm
			.value = intindex
			
		Next	
	End With
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
    ggoSpread.Source = frm1.vspdData
	
    'patch version
    ggoSpread.Spreadinit "V20060623",,parent.gAllowDragDropSpread    
	.ReDraw = false
	
    .MaxCols = C_TrackingFlag + 1
	.Col = .MaxCols
    .ColHidden = True
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")    
	
	ggoSpread.SSSetCombo C_PlantCd  ,		"공장"				,15
	ggoSpread.SSSetCombo C_PlantNm  ,		"공장명"			,35
	ggoSpread.SSSetCombo C_ItemAcct  ,		"품목계정"			,15
	ggoSpread.SSSetCombo C_ItemAcctNm  ,	"품목계정명"        ,35
	ggoSpread.SSSetCheck C_TrackingFlag,	"Tracking사용여부"	,15,2 ,"" , True
	
	Call ggoSpread.MakePairsColumn(C_PlantCd, C_PlantNm)
	Call ggoSpread.MakePairsColumn(C_ItemAcct, C_ItemAcctNm)
	Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
    
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock    C_PlantCd,  -1, C_PlantCd
    ggoSpread.SpreadLock    C_PlantNm,  -1, C_PlantNm
    ggoSpread.SpreadLock    C_ItemAcct,  -1, C_ItemAcct
    ggoSpread.SpreadLock    C_ItemAcctNm,  -1, C_ItemAcctNm
    ggoSpread.SpreadLock    C_TrackingFlag,  -1, C_TrackingFlag
    
    'ggoSpread.SSSetRequired	C_TrackingFlag,  -1, -1
    
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    Dim iRow
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired		C_PlantCd,  pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm,  pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemAcct,  pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemAcctNm,  pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_TrackingFlag,  pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_PlantCd		= iCurColumnPos(1)
            C_PlantNm		= iCurColumnPos(2)
            C_ItemAcct      = iCurColumnPos(3)
            C_ItemAcctNm    = iCurColumnPos(4)
            C_TrackingFlag	= iCurColumnPos(5)
            
    End Select    
End Sub


'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

	'List Minor code(Plant Cd)
	Call CommonQueryRs(" PLANT_CD,PLANT_NM "," B_PLANT ",,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_PlantCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_PlantNm
	
	'****************************
	'List Minor code(Item Acct)
	'****************************
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & Filtervar("P1001","''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ItemAcct
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ItemAcctNm
    
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업" 
	arrParam(1) = "B_PLANT"    
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""   
	arrParam(5) = "공장"   
 
	arrField(0) = "PLANT_CD" 
	arrField(1) = "PLANT_NM" 
 
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)  
		frm1.txtPlantNm.Value    = arrRet(1) 
		frm1.txtPlantCd.focus 
	End If 
 
End Function

'------------------------------------------  OpenItemAcct()  --------------------------------------------------
' Name : OpenItemAcct()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 
 IsOpenPop = True

 arrParam(0) = "품목계정 팝업"    
 arrParam(1) = "B_MINOR"      
 arrParam(2) = Trim(frm1.txtItemAcct.Value) 
 arrParam(3) = ""       
 arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""  
 arrParam(5) = "품목계정"   
 
 arrField(0) = "MINOR_CD"      
 arrField(1) = "MINOR_NM"      
 
 arrHeader(0) = "품목계정"     
 arrHeader(1) = "계정명"      
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtItemAcct.focus
	Exit Function
 Else
	frm1.txtItemAcct.Value		= arrRet(0)
	frm1.txtItemAcctNm.Value	= arrRet(1)
	frm1.txtItemAcct.focus
 End If 
End Function

Sub Form_Load()
	
    Call LoadInfTB19029
    
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    
    Call InitSpreadComboBox()
    
    Call SetToolBar("1100110100001111")

    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemAcct.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
    
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
		Select Case Col
			
			Case C_PlantCd
				.Col = Col
				intIndex = .Value
				.Col = C_PlantNm
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			
			Case C_PlantNm	
				.Col = Col
				intIndex = .Value
				.Col = C_PlantCd
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row

			Case  C_ItemAcct
				.Col = Col
				intIndex = Trim(.Value)
				.Col = C_ItemAcctNm
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			
			Case C_ItemAcctNm	
				.Col = Col
				intIndex = .Value
				.Col = C_ItemAcct
				.Value = intIndex
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
				
		End Select		
				
	End with			

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'	With frm1.vspdData 
	
'		ggoSpread.Source = frm1.vspdData
		
'		If Row > 0 And Col = C_TrackingFlag Then
'			.Col = Col
'			.Row = Row									
'			IF .Text = 1 Then
'				If lgIntFlgMode = Parent.OPMD_UMODE Then
'					.Col = 0
'					.Text = ggoSpread.UpdateFlag
'				ElseIf lgIntFlgMode = Parent.OPMD_CMODE Then
'					.Col = 0
'					.Text = ggoSpread.InsertFlag
'				End If
				
'				lgBlnFlgChgValue = True
				
'			Elseif .Text = 0 Then
'				If lgIntFlgMode = Parent.OPMD_UMODE Then
'					.Col = 0
'					.Text = ""
'					lgBlnFlgChgValue = False
'				ElseIf lgIntFlgMode = Parent.OPMD_CMODE Then
'					.Col = 0
'					.Text = ggoSpread.InsertFlag
					
'					lgBlnFlgChgValue = True
'				End If
'			End if  
							
'		End If	
'	End With
'End Sub

 
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
    Else
    	frm1.vspdData.Row = Row
		frm1.vspdData.Col = C_PlantCd		
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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      'Initializes local global variables
    															
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then										'This function check indispensable field
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
    End If
       
    FncQuery = True															
    
End Function

Function FncSave() 
    
    FncSave = False                                                         
    
    Err.Clear
    On Error Resume Next
    
    '-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
    Call InitData(1)
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    Err.Clear
    
    FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
    
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        .vspdData.Col = C_TrackingFlag
        .vspdData.Text = 1
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

Function FncDeleteRow() 
    With frm1.vspdData
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    iColumnLimit  =  5                 ' split 한계치  maxcol이 아님(5번째 칼럼이 split의 최고치)
                                       ' 5라는 값은 표준이 아닙니다.개발자가 업무에 맞게 수정요 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.SSSetSplit(ACol)    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

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
	Call InitData(1)
End Sub

Function DbQuery() 

    DbQuery = False
    
    Err.Clear

	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
       
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
								& "&txtPlantCd=" & .txtPlantCd.value _
								& "&txtItemAcct=" & .txtItemAcct.value _
								& "&txtMaxRows=" & .vspdData.MaxRows

    
		Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk(ByVal LngMaxRow)													'조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	lgIntFlgMode = Parent.OPMD_CMODE
    Call InitData(LngMaxRow)
	Call SetToolBar("110011110001111")
	
	
End Function

Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    On Error Resume Next

	With frm1
		.txtMode.value = Parent.UID_M0002

	'-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""

    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag
				
				strVal = strVal & "CREATE" & Parent.gColSep
				
				.vspdData.Col = C_PlantCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

                .vspdData.Col = C_ItemAcct
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_TrackingFlag
                If .vspdData.Text = "1" Then
					strVal = strVal & "Y" & Parent.gColSep
                Else
					strVal = strVal & "N" & Parent.gColSep
                End If
                
                strVal = strVal & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
                
                strVal = strVal & "UPDATE" & Parent.gColSep

                .vspdData.Col = C_PlantCd
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

                .vspdData.Col = C_ItemAcct
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_TrackingFlag
                If .vspdData.Text = "1" Then
					strVal = strVal & "Y" & Parent.gColSep
                Else
					strVal = strVal & "N" & Parent.gColSep
                End If
                
                strVal = strVal & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag

				strDel = strDel & "DELETE" & Parent.gColSep
				
				.vspdData.Col = C_PlantCd
				strDel = strDel & Trim(.vspdData.Text) & parent.gColSep

                .vspdData.Col = C_ItemAcct
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                
                .vspdData.Col = C_TrackingFlag
                If .vspdData.Text = "1" Then
					strDel = strDel & "Y" & Parent.gColSep
                Else
					strDel = strDel & "N" & Parent.gColSep
                End If
				
                strDel = strDel & lRow & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목계정별Tracking사용</font></td>
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
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=28 tag="14">
									</TD>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=6 MAXLENGTH=2  tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=27 MAXLENGTH=50 tag="14">
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

