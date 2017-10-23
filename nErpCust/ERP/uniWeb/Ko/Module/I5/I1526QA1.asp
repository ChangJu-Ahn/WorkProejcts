<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 수불현황 조회 
'*  3. Program ID           : I1526QA1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/28
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

Const BIZ_PGM_QRY_ID = "i1526QB1.asp"				

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
DIm lgStrPrevKey1
DIm lgStrPrevKey2
DIm lgStrPrevKey3
DIm SetComboList

Dim C_ItemCd 									
Dim C_ItemNm
Dim C_DocumentDt
Dim C_VMISlCd
Dim C_SlCd
Dim C_BpCd
Dim C_BpNm
Dim C_TrnsType
Dim C_Qty
Dim C_BasicUnit
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_ItemSpec
Dim C_ItemDocumentNo
Dim C_DocumentYear
Dim C_SeqNo
Dim C_SubSeqNo

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

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim StartDate
	
	StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtTrnsFrDt.Text = UNIDateAdd("m", -1, StartDate, Parent.gDateFormat)
	frm1.txtTrnsToDt.Text = StartDate

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
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "QA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD =" & FilterVar("I0006", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTrnsType,lgF0,lgF1,Chr(11))
	
	SetComboList = lgF0 & Chr(12) & lgF1
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ==================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_SubSeqNo + 1							
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_ItemCd,         "품목",         15
		ggoSpread.SSSetEdit  C_ItemNm,         "품명",         20
		ggoSpread.SSSetDate  C_DocumentDt,     "수불일자",     16, 2,Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_VMISlCd,        "VMI창고",      10
		ggoSpread.SSSetEdit  C_SlCd,           "입고창고",     10
		ggoSpread.SSSetEdit  C_BpCd,           "공급처",       15
		ggoSpread.SSSetEdit  C_BpNm,           "공급처명",     20
		ggoSpread.SSSetEdit  C_TrnsType,       "수불구분",     10
		ggoSpread.SSSetFloat C_Qty,            "수량",         13, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit  C_BasicUnit,      "단위",          8
		ggoSpread.SSSetEdit  C_TrackingNo,     "Tracking No.", 20
		ggoSpread.SSSetEdit  C_LotNo,          "LOT NO",       10
		ggoSpread.SSSetEdit  C_LotSubNo,       "순번",          8
		ggoSpread.SSSetEdit  C_ItemSpec,       "규격",         20
		
		ggoSpread.SSSetEdit  C_ItemDocumentNo, "수불번호",     15
		ggoSpread.SSSetEdit  C_DocumentYear,   "년도",          8
		ggoSpread.SSSetEdit  C_SeqNo,          "순번",          8
		ggoSpread.SSSetEdit  C_SubSeqNo,       "상세순번",      8
		
		Call ggoSpread.SSSetColHidden(C_DocumentYear, C_DocumentYear, True)
		Call ggoSpread.SSSetColHidden(C_SubSeqNo,     C_SubSeqNo, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = true
		
	    ggoSpread.SSSetSplit2(2)  
    End With
    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ItemCd         = 1
	C_ItemNm         = 2
	C_DocumentDt     = 3
	C_VMISlCd        = 4
	C_SlCd           = 5
	C_BpCd           = 6
	C_BpNm           = 7
	C_TrnsType       = 8
	C_Qty            = 9
	C_BasicUnit      = 10
	C_TrackingNo     = 11
	C_LotNo          = 12
	C_LotSubNo       = 13
	C_ItemSpec       = 14
	C_ItemDocumentNo = 15
	C_DocumentYear   = 16
	C_SeqNo          = 17
	C_SubSeqNo       = 18
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCd         = iCurColumnPos(1)
		C_ItemNm         = iCurColumnPos(2)
		C_DocumentDt     = iCurColumnPos(3)
		C_VMISlCd        = iCurColumnPos(4)
		C_SlCd           = iCurColumnPos(5)
		C_BpCd           = iCurColumnPos(6)
		C_BpNm           = iCurColumnPos(7)
		C_TrnsType       = iCurColumnPos(8)
		C_Qty            = iCurColumnPos(9)
		C_BasicUnit      = iCurColumnPos(10)
		C_TrackingNo     = iCurColumnPos(11)
		C_LotNo          = iCurColumnPos(12)
		C_LotSubNo       = iCurColumnPos(13)
		C_ItemSpec       = iCurColumnPos(14)
		C_ItemDocumentNo = iCurColumnPos(15)
		C_DocumentYear   = iCurColumnPos(16)
		C_SeqNo          = iCurColumnPos(17)
		C_SubSeqNo       = iCurColumnPos(18)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
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

'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")  
		frm1.txtPlantCd.focus
		Exit Function
	End If

	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "I_VMI_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSlCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")		
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSlCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
End Function

'------------------------------------------  OpenBp()  -------------------------------------------------
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처팝업"	
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""		
	arrParam(5) = "공급처"			
	
	arrField(0) = "BP_CD"	
	arrField(1) = "BP_NM"	
	
	arrHeader(0) = "공급처"		
	arrHeader(1) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBp(arrRet)
	End If	
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItemCd()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam1, arrParam2
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    
		frm1.txtPlantCd.focus
		Exit Function
	End If

	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	arrParam1 = Trim(frm1.txtPlantCd.Value)		
	arrParam2 = Trim(frm1.txtItemCd.Value)

	iCalledAspName = AskPRAspName("I1522PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1522PA1","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If
End Function

'==========================================  2.4.3 Set???()  =============================================
'------------------------------------------  SetPlant()  --------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	frm1.txtPlantCd.focus	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
Function SetSL(byRef arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)	
	frm1.txtSLCd.focus	
End Function

'------------------------------------------  SetBp()  --------------------------------------------------
Function SetBp(byRef arrRet)
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)	
	frm1.txtBpCd.focus	
End Function

'------------------------------------------  SetTrackingNo()  --------------------------------------------------
Function SetItemCd(byRef arrRet)
	frm1.txtItemCd.Value  = arrRet(0)		
	frm1.txtItemNm.Value  = arrRet(1)	
	frm1.txtBpCd.focus	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                                        						
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
    Call ggoOper.LockField(Document, "N")                                   						
    
    Call InitVariables                                                      						
    Call SetdefaultVal
    Call InitComboBox
    Call InitSpreadSheet                                                    						
	Call SetToolbar("11000000000011")          
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_DblClick(Button)
'=======================================================================================================
Sub txtTrnsFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtTrnsFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_Change()
'=======================================================================================================
Sub txtTrnsFrDt_Change()
		lgBlnFlgChgValue = False
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_KeyPress()
'=======================================================================================================
Sub txtTrnsFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsToDt_DblClick(Button)
'=======================================================================================================
Sub txtTrnsToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtTrnsToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDtt_Change()
'=======================================================================================================
Sub txtTrnsToDt_Change()
		lgBlnFlgChgValue = False
End Sub

'=======================================================================================================
'   Event Name : txtTrnsToDt_KeyPress()
'=======================================================================================================
Sub txtTrnsToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
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
 	
 	Call SetPopupMenuItemInf("0000111111") 
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

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim TempInspDt 
    FncQuery = False                                                     
    
    Err.Clear
    If Not chkField(Document, "1") Then Exit Function                                                            
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")	
		If IntRetCD = vbNo Then Exit Function
    End If

    If ValidDateCheck(frm1.txtTrnsFrDt, frm1.txtTrnsToDt) = False Then 
   		frm1.txtTrnsFrDt.focus 
		Exit Function
	End If

	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	Else
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
								
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus
			Exit function
		End IF
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	If Trim(frm1.txtSLCd.Value) <> "" Then
		If 	CommonQueryRs(" A.SL_NM "," I_VMI_STORAGE_LOCATION A, B_PLANT B ", " A.PLANT_CD = B.PLANT_CD AND " & _
																			   " A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
																			   " AND A.SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
			If 	CommonQueryRs(" SL_NM "," I_VMI_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
								
				Call DisplayMsgBox("162001","X","X","X")
				frm1.txtSLNm.Value = ""
				frm1.txtSLCd.focus
				Exit function
			Else
				Call DisplayMsgBox("169922","X","X","X")
				frm1.txtSLNm.Value = ""
				frm1.txtSLCd.focus
				Exit function
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtSLNm.Value = lgF0(0)
	Else
		frm1.txtSLNm.Value = ""
	End If
	
	If Trim(frm1.txtBpCd.Value) <> "" Then
		If 	CommonQueryRs(" BP_NM "," B_Biz_Partner ", " BP_CD = " & FilterVar(frm1.txtBpCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("229927","X","X","X")
			frm1.txtBpNm.Value = ""
			frm1.txtBpCd.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtBpNm.Value = lgF0(0)	
	Else
		frm1.txtBpNm.Value = ""
	End If	
	
	If Trim(frm1.txtItemCd.Value) <> "" Then
		If 	CommonQueryRs(" A.item_nm "," B_ITEM A, B_ITEM_BY_PLANT B ", " A.item_cd = B.item_cd AND B.material_type = " & FilterVar("30", "''", "S") & " " & _
	   																	 " AND B.plant_cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
																		 " AND B.item_cd  = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			frm1.txtItemNm.Value = ""
		Else
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)	
		End If
	Else
		frm1.txtItemNm.Value = ""
	End If

    '-----------------------
    'Erase contents area
    '-----------------------
	ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 
    Call InitVariables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
	Set gActiveElement = document.activeElement
    FncQuery = True										
    
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

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
	Dim strVal
    
	On error resume next
    Err.Clear                                                            
    
    Call LayerShowHide(1)

    DbQuery = False
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID &	"?txtPlantCd="		& Trim(.txtPlantCd.value)	& _
										"&txtTrnsFrDt="		& Trim(.txtTrnsFrDt.Text)	& _
										"&txtTrnsToDt="		& Trim(.txtTrnsToDt.Text)	& _
										"&txtSlCd="			& .hSlCd.value				& _
										"&txtBpCd="			& .hBpCd.value				& _
										"&txtItemCd="		& .hItemCd.value			& _
										"&cboTrnsType="		& .hTrnsType.value			& _
										"&lgStrPrevKey="    & Trim(lgStrPrevKey)		& _
										"&lgStrPrevKey1="   & Trim(lgStrPrevKey1)		& _
										"&lgStrPrevKey2="   & Trim(lgStrPrevKey2)		& _
										"&lgStrPrevKey3="   & Trim(lgStrPrevKey3)		& _
										"&txtMaxRows="      & .vspdData.MaxRows			& _
										"&SetComboList="    & SetComboList 
		Else    
			strVal = BIZ_PGM_QRY_ID &	"?txtPlantCd="      & Trim(.txtPlantCd.value)	& _
										"&txtTrnsFrDt="     & Trim(.txtTrnsFrDt.Text)	& _
										"&txtTrnsToDt="     & Trim(.txtTrnsToDt.Text)	& _
										"&txtSlCd="         & Trim(.txtSlCd.Value)		& _
										"&txtBpCd="         & Trim(.txtBpCd.Value)		& _
										"&txtItemCd="       & Trim(.txtItemCd.Value)	& _
										"&cboTrnsType="     & Trim(.cboTrnsType.value)	& _
										"&lgStrPrevKey="    & Trim(lgStrPrevKey)		& _
										"&lgStrPrevKey1="   & Trim(lgStrPrevKey1)		& _
										"&lgStrPrevKey2="   & Trim(lgStrPrevKey2)		& _
										"&lgStrPrevKey3="   & Trim(lgStrPrevKey3)		& _
										"&txtMaxRows="      & .vspdData.MaxRows			& _
										"&SetComboList="    & SetComboList
		End IF
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()								
    lgIntFlgMode = Parent.OPMD_UMODE
    Call SetToolbar("11000000000111")
    frm1.vspdData.focus
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>VMI 수불현황조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
								<TD CLASS="TD5" NOWRAP>수불기간</TD>
								<TD CLASS="TD6" NOWRAP>
								    <script language =javascript src='./js/i1526qa1_fpDateTime1_txtTrnsFrDt.js'></script>
							        &nbsp;~&nbsp;
							        <script language =javascript src='./js/i1526qa1_fpDateTime2_txtTrnsToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>VMI 창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=8 MAXLENGTH=7 tag="11xxxU" ALT="VMI 창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=20 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=8 MAXLENGTH=10 tag="11xxxU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>								
								<TD CLASS="TD5" NOWRAP>수불구분</TD>
								<TD CLASS="TD6" NOWRAP>
								<SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 100px" tag="11"><OPTION Value=""></OPTION></SELECT>
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
								<script language =javascript src='./js/i1526qa1_OBJECT1_vspdData.js'></script>
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
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hTrnsFrDt" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hTrnsToDt" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hSlCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hBpCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hTrnsType" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

