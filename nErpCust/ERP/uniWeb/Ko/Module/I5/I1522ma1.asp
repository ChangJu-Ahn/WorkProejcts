<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 검사결과등록 
'*  3. Program ID           : I1522MA1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

Const BIZ_PGM_QRY_ID = "i1522mb1.asp"									
Const BIZ_PGM_ID     = "i1522mb2.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim StartDate

Dim C_ItemCd 									
Dim C_ItemNm
Dim C_EntryQty
Dim C_GoodQty
Dim C_BadQty
Dim C_InspQty
Dim C_EntryUnit
Dim C_BpCd
Dim C_BpNm
Dim C_SlCd
Dim C_SlNm
Dim C_DocumentDt
Dim C_InspFlg
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_Specification
Dim C_BasicUnit
Dim C_DocumentNo
Dim C_DocumentYear
DIm C_SeqNo
Dim C_SubSeqNo
DIm C_InspReqNo

'#########################################################################################################
'					2. Function부 
'######################################################################################################### 
'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE         
	lgBlnFlgChgValue = False     	           
	lgIntGrpCount = 0                          
	lgStrPrevKey  = ""                         
	lgStrPrevKey1 = ""                         
	lgStrPrevKey2 = ""                         
	lgStrPrevKey3 = ""                         
	lgLngCurRows = 0                           
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	
	frm1.txtDocumentFrDt.Text = UNIDateAdd("m", -1, StartDate, Parent.gDateFormat)
	frm1.txtDocumentToDt.Text = StartDate
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtPlantCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030425", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_InspReqNo + 1							
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit       C_ItemCd,        "품목",         15
		ggoSpread.SSSetEdit       C_ItemNm,        "품목명",       20
		ggoSpread.SSSetFloat      C_EntryQty,      "입고수량",     15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat      C_GoodQty,       "양품수량",     15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat      C_BadQty,        "불량수량",     15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
	    ggoSpread.SSSetFloat      C_InspQty,       "검사요청수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
	    ggoSpread.SSSetEdit       C_EntryUnit,     "입고단위",      8
	    ggoSpread.SSSetEdit       C_BpCd,          "공급처",       10
	    ggoSpread.SSSetEdit       C_BpNm,          "공급처명",     20
	    ggoSpread.SSSetEdit       C_SlCd,          "VMI창고",      10
	    ggoSpread.SSSetEdit       C_SlNm,          "VMI창고명",    20
	    ggoSpread.SSSetDate       C_DocumentDt,    "입고일자",     10, 2,Parent.gDateFormat  
	    ggoSpread.SSSetCheck      C_InspFlg,       "검사여부",      8,,,1
	    ggoSpread.SSSetEdit       C_TrackingNo,    "Tracking No.", 20
	    ggoSpread.SSSetEdit       C_LotNo,         "Lot No.",      10
	    ggoSpread.SSSetEdit       C_LotSubNo,      "순번",          8
	    ggoSpread.SSSetEdit       C_Specification, "규격",          8
	    ggoSpread.SSSetEdit       C_BasicUnit,     "재고단위",      8
	    ggoSpread.SSSetEdit       C_DocumentNo,    "입고번호",     16
	    ggoSpread.SSSetEdit       C_DocumentYear,  "년도",          8
	    ggoSpread.SSSetEdit       C_SeqNo,         "입고순번",      8
	    ggoSpread.SSSetEdit       C_SubSeqNo,      "상세순번",      8
	    ggoSpread.SSSetEdit       C_InspReqNo,     "검사요청번호", 18
		
		Call ggoSpread.SSSetColHidden(C_DocumentYear, C_DocumentYear, True)
		Call ggoSpread.SSSetColHidden(C_SubSeqNo,     C_SubSeqNo,     True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

	    Call SetSpreadLock

		.ReDraw = true
		
   		ggoSpread.SSSetSplit2(2)  
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadUnLock C_EntryQty, -1, C_EntryQty
	ggoSpread.SSSetRequired C_EntryQty, -1
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ItemCd        = 1
	C_ItemNm        = 2
	C_EntryQty      = 3
	C_GoodQty       = 4
	C_BadQty        = 5
	C_InspQty       = 6
	C_EntryUnit     = 7
	C_BpCd          = 8
	C_BpNm          = 9
	C_SlCd          = 10
	C_SlNm          = 11
	C_DocumentDt    = 12
	C_InspFlg       = 13
	C_TrackingNo    = 14
	C_LotNo         = 15
	C_LotSubNo      = 16
	C_Specification = 17
	C_BasicUnit     = 18
	C_DocumentNo    = 19
	C_DocumentYear  = 20
	C_SeqNo         = 21
	C_SubSeqNo      = 22
	C_InspReqNo     = 23
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCd        = iCurColumnPos(1)
		C_ItemNm        = iCurColumnPos(2)
		C_EntryQty      = iCurColumnPos(3)
		C_GoodQty       = iCurColumnPos(4)
		C_BadQty        = iCurColumnPos(5)
		C_InspQty       = iCurColumnPos(6)
		C_EntryUnit     = iCurColumnPos(7)
		C_BpCd          = iCurColumnPos(8)
		C_BpNm          = iCurColumnPos(9)
		C_SlCd          = iCurColumnPos(10)
		C_SlNm          = iCurColumnPos(11)
		C_DocumentDt    = iCurColumnPos(12)
		C_InspFlg       = iCurColumnPos(13)
		C_TrackingNo    = iCurColumnPos(14)
		C_LotNo         = iCurColumnPos(15)
		C_LotSubNo      = iCurColumnPos(16)
		C_Specification = iCurColumnPos(17)
		C_BasicUnit     = iCurColumnPos(18)
		C_DocumentNo    = iCurColumnPos(19)
		C_DocumentYear  = iCurColumnPos(20)
		C_SeqNo         = iCurColumnPos(21)
		C_SubSeqNo      = iCurColumnPos(22)
		C_InspReqNo     = iCurColumnPos(23)
 	End Select
End Sub

'========================================== 2.4.2 Open???()  =============================================
'------------------------------------------  OpenPlant()  ------------------------------------------------
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
		Call SetPlant(arrRet)
	End If	
End Function

'==========================================  2.4.3 Set???()  =============================================
'------------------------------------------  SetPlant()  --------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	frm1.txtPlantCd.focus	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                                        						
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
    Call ggoOper.LockField(Document, "N")                                   						

 	StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
   
    Call InitVariables                                                      						
    Call SetdefaultVal
    Call InitSpreadSheet                                                    					
	
	Call SetToolbar("11100000000011")								
End Sub


'=======================================================================================================
'   Event Name : txtDocumentFrDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentFrDt_Change()
'=======================================================================================================
Sub txtDocumentFrDt_Change()
	If lgIntFlgMode = Parent.OPMD_CMODE Then	
		lgBlnFlgChgValue = False
	Else
		lgBlnFlgChgValue = True	
	End if
End Sub

'=======================================================================================================
'   Event Name : txtDocumentFrDt_KeyPress()
'=======================================================================================================
Sub txtDocumentFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentToDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDtt_Change()
'=======================================================================================================
Sub txtDocumentToDt_Change()
	If lgIntFlgMode = Parent.OPMD_CMODE Then	
		lgBlnFlgChgValue = False
	Else
		lgBlnFlgChgValue = True	
	End if
End Sub

'=======================================================================================================
'   Event Name : txtDocumentToDt_KeyPress()
'=======================================================================================================
Sub txtDocumentToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'******************************  3.2.1 Object Tag 처리  *********************************************
'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	
		If (lgStrPrevKey <> "" and lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" and lgStrPrevKey3 <> "") Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End If
	End if  
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
    If lgIntFlgMode = Parent.OPMD_CMODE Then
 		Call SetPopupMenuItemInf("0000111111") 
 	Else
 	 	Call SetPopupMenuItemInf("0001111111")
 	End If
 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
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
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'#########################################################################################################
'					5. Interface부 
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim TempInspDt 
    FncQuery = False                                                      
    
    Err.Clear
    
    If GetSetupMod(Parent.gSetupMod, "q") <> "Y" then
       Call DisplayMsgBox("169967","X", "X", "X")
       Exit Function
	End if                                                              

    '-----------------------
    'Check previous data area
    '----------------------- 
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function						
     '-----------------------
    'Erase contents area
    '-----------------------
	ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData  
	Call InitVariables

    If ValidDateCheck(frm1.txtDocumentFrDt, frm1.txtDocumentToDt) = False Then Exit Function

    
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True									
    
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    Err.Clear                                                            
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015",Parent.VB_YES_NO,"X", "X")    	
		If IntRetCD = vbNo Then Exit Function
    End If
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                       
	ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData  
	Call ggoOper.LockField(Document, "N")                                        
    Call InitVariables                                                    
    Call SetDefaultVal    
    Call SetToolbar("11100000000011")
    
    FncNew = True                                                         

End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
	If lgIntFlgMode <> Parent.OPMD_UMODE Then                                     
	    Call DisplayMsgBox("900002", "X", "X", "X")                               
		Exit Function
	End If
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                       
    Err.Clear                                                             
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2")  Then Exit Function        
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X", "X", "X")                           
        Exit Function
    End If

	'------------------------------------------------------
    If Not ggoSpread.SSDefaultCheck Then Exit Function

    If frm1.vspdData.MaxRows < 1 then
       Call DisplayMsgBox("900002","X", "X", "X")  
	   exit function
	End if 
    '-----------------------
    'Save function call area
    '-----------------------
	If DBSave() = False Then Exit Function

    FncSave = True                                                         
    
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
    If frm1.vspdData.maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)								
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub  FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
End Sub 

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
	Dim strVal
    
    Err.Clear                                                             
    Call LayerShowHide(1)       
    
    DbQuery = False
    With frm1    
		strVal = BIZ_PGM_QRY_ID &	"?txtMode="			& Parent.UID_M0001				& _					
									"&txtPlantCd="      & Trim(.txtPlantCd.value)		& _				
									"&txtDocumentFrDt=" & Trim(.txtDocumentFrDt.Text)	& _
									"&txtDocumentToDt=" & Trim(.txtDocumentToDt.Text)	& _
									"&lgStrPrevKey="    & Trim(lgStrPrevKey)			& _
									"&lgStrPrevKey1="   & Trim(lgStrPrevKey1)			& _
									"&lgStrPrevKey2="   & Trim(lgStrPrevKey2)			& _
									"&lgStrPrevKey3="   & Trim(lgStrPrevKey3)			& _		
									"&txtMaxRows="      & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE							
    Call SetToolbar("11101001000111")
    Call ggoOper.LockField(Document, "Q")
    Call SetActiveCell(frm1.vspdData,C_EntryQty,1,"M","X","X")

End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim strVal
	Dim iGoodQty
	Dim PvArr
	
    Call LayerShowHide(1)

    Err.Clear		
	
    DbSave = False                                                         
	
	frm1.txtMode.value = Parent.UID_M0002
 
	With frm1.vspdData
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0
		ReDim PvArr(0)
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .MaxRows
            .Row = lRow
			.Col = 0
        
			Select Case .Text
				Case ggoSpread.UpdateFlag	
					
					.Row = lRow
					.Col = C_GoodQty
					iGoodQty = uniCdbl(.Value)
			
					.Col = C_EntryQty
					If iGoodQty = 0 And uniCdbl(.Value) > 0 Then
						Call DisplayMsgBox("162061","X", "X", "X")
						Call LayerShowHide(0)
						Exit Function
					ElseIf uniCdbl(.Value) = 0 And iGoodQty > 0 Then
						Call DisplayMsgBox("162062","X", "X", "X")
						Call LayerShowHide(0)
						Exit Function
					ElseIf uniCdbl(.Value) > iGoodQty then 
						Call DisplayMsgBox("162063","X", "X", "X")
						Call LayerShowHide(0)
						Exit Function
					Else
						.Col = C_DocumentNo
						strVal = Trim(.Text) & Parent.gColSep    
						.Col = C_DocumentYear
						strVal = strVal & Trim(.Text) & Parent.gColSep   
						.Col = C_SeqNo
						strVal = strVal & Trim(.Text) & Parent.gColSep     
						.Col = C_ItemCd
						strVal = strVal & Trim(.Text) & Parent.gColSep     
						.Col = C_EntryQty
						strVal = strVal & Trim(.Value) & Parent.gColSep    
						strVal = strVal & lRow & parent.gRowSep            
	
						ReDim Preserve PvArr(lGrpCnt)
						PvArr(lGrpCnt) = strVal
						lGrpCnt = lGrpCnt + 1
					End If
			End Select
		Next
	End With

	frm1.txtMaxRows.value = lGrpCnt
	frm1.txtSpread.value  = Join(PvArr,"")
	If lGrpCnt <= 0 then				
		Call DisplayMsgBox("800161","X", "X", "X")    '
		Call LayerShowHide(0)
		Exit function
	End if

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)						
	
    DbSave = True                                                          
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()								
	Call InitVariables
	ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData    
    Call FncQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>VMI 검사결과등록</font></td>
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
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>입고일자</TD>
								<TD CLASS="TD6" NOWRAP>
								    <script language =javascript src='./js/i1522ma1_fpDateTime1_txtDocumentFrDt.js'></script>
							        &nbsp;~&nbsp;
							        <script language =javascript src='./js/i1522ma1_fpDateTime2_txtDocumentToDt.js'></script>
								</TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=*>
					<TD WIDTH=100% VALIGN=TOP>
					<TABLE <%=LR_SPACE_TYPE_60%>>
					<TR>
						<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
							<script language =javascript src='./js/i1522ma1_OBJECT1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtpvCommandMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hDocumentYear" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

