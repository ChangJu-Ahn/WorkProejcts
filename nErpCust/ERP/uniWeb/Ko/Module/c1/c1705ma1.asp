
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 계정별 직과항목 선택 
'*  3. Program ID           : c1705ma1
'*  4. Program Name         : 계정별 직과항목 선택 
'*  5. Program Desc         : 계정별 직과항목 선택 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/19
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung / Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c1705mb1.asp"	                          'Biz Logic ASP


Dim C_AcctCd  
Dim C_AcctCdPopup  															
Dim C_AcctNm  
Dim C_CtrlCd  
Dim C_CtrlCdPopup  
Dim C_CtrlNm  


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	 C_AcctCd			= 1
	 C_AcctCdPopup		= 2															
	 C_AcctNm			= 3
	 C_CtrlCd			= 4
	 C_CtrlCdPopup		= 5
	 C_CtrlNm			= 6
End Sub


'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0 
    
    lgStrPrevKey = "" 
    lgLngCurRows = 0  
    
    lgSortKey = 1
End Sub


'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
    .MaxCols = C_CtrlNm+1
    .Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	
	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit C_AcctCd, "계정코드", 20,,,20,2
    ggoSpread.SSSetButton C_AcctCdPopup
    ggoSpread.SSSetEdit C_AcctNm, "계정명", 36
    ggoSpread.SSSetEdit C_CtrlCd, "관리항목코드", 20,,,3,2
    ggoSpread.SSSetButton C_CtrlCdPopup
    ggoSpread.SSSetEdit C_CtrlNm, "관리항목명", 37

	call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctCdPopup)
	call ggoSpread.MakePairsColumn(C_CtrlCd,C_CtrlCdPopup)
	
'	ggoSpread.SSSetSplit(C_AcctNm)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_AcctCd		, -1, C_AcctCd
    ggoSpread.SpreadLock C_AcctCdPopup	, -1, C_AcctCdPopup
    ggoSpread.SpreadLock C_AcctNm		, -1, C_AcctNm
    ggoSpread.SSSetRequired	C_CtrlCd	, -1, -1
    ggoSpread.SpreadLock C_CtrlNm		, -1, C_CtrlNm
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired		C_AcctCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected		C_AcctNm	,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired		C_CtrlCd	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected		C_CtrlNm	,pvStartRow	,pvEndRow
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
			C_AcctCd					= iCurColumnPos(1)
			C_AcctCdPopup				= iCurColumnPos(2)
			C_AcctNm					= iCurColumnPos(3)    
			C_CtrlCd					= iCurColumnPos(4)
			C_CtrlCdPopup				= iCurColumnPos(5)
			C_CtrlNm					= iCurColumnPos(6)
    End Select    
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenAcctCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정팝업"
	arrParam(1) = "A_ACCT a, A_ACCT_CTRL_ASSN b"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "a.TEMP_FG_3 IN (" & FilterVar("M2", "''", "S") & " ," & FilterVar("M3", "''", "S") & " ," & FilterVar("M4", "''", "S") & " ) AND a.ACCT_CD = b.ACCT_CD"
	arrParam(5) = "계정"			
	
    arrField(0) = "a.ACCT_CD"
    arrField(1) = "a.ACCT_NM"
    
    arrHeader(0) = "계정코드"
    arrHeader(1) = "계정명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetAcctCd(arrRet, iWhere)
	End If	

End Function

Function SetAcctCd(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		 frm1.txtAcctCd.focus
    		.txtAcctCd.value = arrRet(0)
    		.txtAcctNm.value = arrRet(1)
    	Else
    		.vspdData.Col = C_AcctCd
    		.vspdData.Text = arrRet(0)
    		.vspdData.Col = C_AcctNm
    		.vspdData.Text = arrRet(1)
            
    		Call vspdData_Change(.vspdData.Col, .vspdData.Row)
    	End If
	
	End With
	
End Function

Function OpenCtrlCd(Byval strCode, Byval AcctCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "관리항목팝업"
	arrParam(1) = "A_CTRL_ITEM a, A_ACCT_CTRL_ASSN b"
	arrParam(2) = strCode
	arrParam(3) = ""	
	arrParam(4) = "a.CTRL_CD = b.CTRL_CD and b.ACCT_CD =  " & FilterVar(AcctCode, "''", "S") & " "
	arrParam(5) = "관리항목"			
	
    arrField(0) = "a.CTRL_CD"
    arrField(1) = "a.CTRL_NM"
    
    arrHeader(0) = "관리항목코드"
    arrHeader(1) = "관리항목명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCtrlCd(arrRet)
	End If	

End Function

Function SetCtrlCd(Byval arrRet)
	
	With frm1
	
		.vspdData.Col = C_CtrlCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_CtrlNm
		.vspdData.Text = arrRet(1)
		            
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)	
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029 
    
    Call ggoOper.LockField(Document, "N")                 
                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables  
    
    Call SetDefaultVal

    Call SetToolbar("110011010010111")
    frm1.txtAcctCd.focus 
   	Set gActiveElement = document.activeElement			    
    
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
  
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub



Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp, AcctCode
  Dim intRetCD
  
	With frm1 
	
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 And Col = C_AcctCdPopUp Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_AcctCd
		    Call OpenAcctCd(.vspdData.Text, 1)
		        
		ElseIf Row > 0 And Col = C_CtrlCdPopup Then
		    .vspdData.Col = Col
		    .vspdData.Row = Row
		        
			.vspdData.Col = C_AcctCd
			If .vspdData.Text = "" Then
				intRetCD = DisplayMsgBox("110100","x","x","x") 
				Exit Sub
			End If

			AcctCode = Trim(.vspdData.Text)
			
			.vspdData.Col = C_CtrlCd
		    Call OpenCtrlCd(.vspdData.Text, AcctCode)

		End If
    Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_AcctNm Or NewCol <= C_AcctNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	 
    	If lgStrPrevKey <> "" Then 
      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    	
    Call InitVariables 
    
    if frm1.txtAcctCd.value = "" then
		frm1.txtAcctNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    IF DbQuery = False Then
		Exit Function
	END IF	
    
    FncQuery = True															
    
End Function


'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear  
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")  
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then  
       Exit Function
    End If
    
    IF DbSave = False Then
		Exit function
	END IF	
    
    FncSave = True                                                          
    
End Function


'========================================================================================================
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	  
    if frm1.vspdData.maxrows < 1 then exit function 
	   
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_AcctCd
    frm1.vspdData.Text = ""
    
    frm1.vspdData.Col = C_AcctNm
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function

'
'========================================================================================================
Function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  
End Function



'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
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

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function


'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
End Function


'========================================================================================================
Function FncPrev() 
    On Error Resume Next     
End Function


Function FncNext() 
    On Error Resume Next 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub



'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================


'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit function
	END IF
    
    Err.Clear 
    
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtAcctCd=" & .hAcctCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtAcctCd=" & .txtAcctCd.value
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    
End Function


'========================================================================================================
Function DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")

	Call SetToolbar("110011110011111")
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
	
End Function


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
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False                                                          
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
	    
	With frm1
		.txtMode.value = Parent.UID_M0002

    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag
				
				strVal = strVal & "C" & iColSep & lRow & iColSep 

                .vspdData.Col = C_AcctCd
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_CtrlCd
                strVal = strVal & Trim(.vspdData.Text) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag	
					
				strVal = strVal & "U" & iColSep & lRow & iColSep

                .vspdData.Col = C_AcctCd
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_CtrlCd
                strVal = strVal & Trim(.vspdData.Text) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag	

				strDel = strDel & "D" & iColSep & lRow & iColSep	

                .vspdData.Col = C_AcctCd
                strDel = strDel & Trim(.vspdData.Text) & iRowSep									
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()	
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	
    Call MainQuery()
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정별직과항목등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>계정</TD>
									<TD CLASS="TD656" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenAcctCd txtAcctCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=20 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1705ma1_I186048591_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hAcctCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

