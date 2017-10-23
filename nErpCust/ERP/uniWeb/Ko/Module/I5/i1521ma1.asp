<%@ LANGUAGE="VBSCRIPT" %>
<!--********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 입고등록 
'*  3. Program ID           : I1521MA1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/26
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

Const BIZ_PGM_QRY_ID = "i1521mb1.asp"								
Const BIZ_PGM_ID     = "i1521mb2.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
DIm lgStrPrevKey1

Dim C_ItemCd 										
Dim C_ItemPopup
Dim C_ItemNm
Dim C_EntryQty
Dim C_EntryUnit
Dim C_EntryUnitPopup
Dim C_InspFlag
Dim C_TrackingNo
Dim C_TrackingNoPopup
Dim C_LotNo
Dim C_LotSubNo
Dim C_LotNoPopup
Dim C_Specification
Dim C_BasicUnit
Dim C_InspReqNo
Dim C_InspSts
Dim C_SeqNo
Dim C_SeqSubNo
Dim C_LotFlg
Dim C_TrackingFlg

Dim StartDate

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
	lgLngCurRows = 0                            
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	frm1.txtDocumentDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	Call ExtractDateFrom(StartDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
	frm1.txtDocumentYear.Year = strYear  
		
	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemDocumentNo.focus 
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
	ggoSpread.Spreadinit "V20050728", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_TrackingFlg + 1						
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
		
		ggoSpread.SSSetEdit       C_ItemCd,        "품목",         15, 0, -1, 18, 2	
		ggoSpread.SSSetButton 	  C_ItemPopup
		ggoSpread.MakePairsColumn C_ItemCd, C_ItemPopup

		ggoSpread.SSSetEdit       C_ItemNm,        "품명",         20
		ggoSpread.SSSetFloat      C_EntryQty,      "입고수량",     15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit       C_EntryUnit,     "입고단위",     10, 0, -1, 3, 2 
		ggoSpread.SSSetButton 	  C_EntryUnitPopup
		ggoSpread.MakePairsColumn C_EntryUnit, C_EntryUnitPopup

		ggoSpread.SSSetCheck      C_InspFlag,      "검사여부",     10,,,1
		ggoSpread.SSSetEdit       C_TrackingNo,    "Tracking No",  25, 0, -1, 25, 2
		ggoSpread.SSSetButton 	  C_TrackingNoPopup
		ggoSpread.MakePairsColumn C_TrackingNo, C_TrackingNoPopup

		ggoSpread.SSSetEdit       C_LotNo,         "LOT NO",       20, 0, -1, 25, 2
		ggoSpread.SSSetFloat      C_LotSubNo,      "순번",          8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetButton     C_LotNoPopup
		ggoSpread.MakePairsColumn C_LotNo, C_LotNoPopup

		ggoSpread.SSSetEdit       C_Specification, "규격",         20, 0, -1, 50	
		ggoSpread.SSSetEdit       C_BasicUnit,     "재고단위",     10, 0, -1, 3
		ggoSpread.SSSetEdit       C_InspReqNo,     "검사요청번호", 18
		ggoSpread.SSSetEdit       C_InspSts,       "검사상태",     10
		ggoSpread.SSSetEdit       C_SeqNo,         "일련번호",     10
		ggoSpread.SSSetEdit       C_SeqSubNo,      "상세일련번호", 10
		ggoSpread.MakePairsColumn C_SeqNo, C_SeqSubNo
		ggoSpread.SSSetEdit       C_LotFlg,      "LotFlg", 5
		ggoSpread.SSSetEdit       C_TrackingFlg,      "TrackingFlg", 5

		Call ggoSpread.SSSetColHidden(C_InspReqNo, C_InspReqNo, True)
		Call ggoSpread.SSSetColHidden(C_InspSts, C_InspSts, True)
		Call ggoSpread.SSSetColHidden(C_SeqNo, .MaxCols, True)
		
		ggoSpread.SpreadLock -1, -1
   	
   		ggoSpread.SSSetSplit2(3)  
	
		.ReDraw = true
    End With
    
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.SSSetRequired  C_ItemCd,          pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm,		    pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_EntryQty,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_EntryUnit,       pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Specification,   pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_BasicUnit,       pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_InspReqNo,       pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_InspSts,         pvStartRow, pvEndRow
		
	ggoSpread.SSSetProtected C_SeqNo,           pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SeqSubNo,        pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_LotFlg,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TrackingFlg,     pvStartRow, pvEndRow
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_ItemCd          = 1
	C_ItemPopup       = 2
	C_ItemNm          = 3
	C_EntryQty        = 4
	C_EntryUnit       = 5
	C_EntryUnitPopup  = 6
	C_InspFlag        = 7
	C_TrackingNo      = 8
	C_TrackingNoPopup = 9
	C_LotNo           = 10
	C_LotSubNo        = 11
	C_LotNoPopup      = 12
	C_Specification   = 13
	C_BasicUnit       = 14
	C_InspReqNo       = 15
	C_InspSts         = 16
	C_SeqNo           = 17
	C_SeqSubNo        = 18
	C_LotFlg		  = 19
	C_TrackingFlg	  = 20
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCd          = iCurColumnPos(1)
		C_ItemPopup       = iCurColumnPos(2)
		C_ItemNm          = iCurColumnPos(3)
		C_EntryQty        = iCurColumnPos(4)
		C_EntryUnit       = iCurColumnPos(5)
		C_EntryUnitPopup  = iCurColumnPos(6)
		C_InspFlag        = iCurColumnPos(7)
		C_TrackingNo      = iCurColumnPos(8)
		C_TrackingNoPopup = iCurColumnPos(9)
		C_LotNo           = iCurColumnPos(10)
		C_LotSubNo        = iCurColumnPos(11)
		C_LotNoPopup      = iCurColumnPos(12)
		C_Specification   = iCurColumnPos(13)
		C_BasicUnit       = iCurColumnPos(14)
		C_InspReqNo       = iCurColumnPos(15)
		C_InspSts         = iCurColumnPos(16)
		C_SeqNo           = iCurColumnPos(17)
		C_SeqSubNo        = iCurColumnPos(18)
		C_LotFlg		  = iCurColumnPos(19)
		C_TrackingFlg	  = iCurColumnPos(20)
		
 	End Select
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'========================================== 2.4.2 Open???()  =============================================
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

'------------------------------------------  OpenItemDocumentNo()  -------------------------------------------------
Function OpenItemDocumentNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(3)
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")   
	    frm1.txtPlantCd.Focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtItemDocumentNo.Value)		
	arrParam(1) = Trim(frm1.txtDocumentYear.Text) 		
	arrParam(2) = "VR"
	arrParam(3) = Trim(frm1.txtPlantCd.Value)

	iCalledAspName = AskPRAspName("I1521PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1521PA1","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemDocumentNo.focus
		Exit Function
	Else
		Call SetItemDocumentNo(arrRet)
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
	'------------------------------------------------------

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
Function OpenItem(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam1, arrParam2
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus
		Exit Function
	ElseIf Trim(frm1.txtPlantCd.Value) <> Parent.gPlant Then
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	arrParam1 = Trim(frm1.txtPlantCd.Value)		
	arrParam2 = strCode							

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
		Exit Function
	Else
		Call SetItem(arrRet)
	End If
End Function

'------------------------------------------  OpenTrackingNo()  -------------------------------------------------
Function OpenTrackingNo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "Tracking No"					
	arrParam(1) = "s_so_tracking"					
	arrParam(2) = strCode 		                    
	
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "Tracking No"			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking No"		
    arrHeader(1) = "품목"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If	
End Function

'------------------------------------------  OpenEntryUnit()  -------------------------------------------------
Function OpenEntryUnit(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "단위팝업"					
	arrParam(1) = "B_Unit_Of_Measure"				
	arrParam(2) = strCode 		                    
	
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "단위"			
	
    arrField(0) = "Unit"	
    arrField(1) = "Unit_nm"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEntryUnit(arrRet)
	End If	
End Function

'------------------------------------------  OpenLotNo()  -------------------------------------------------
Function OpenLotNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrParam7, arrParam8
	

	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus
		Exit Function
	End If


	If Trim(frm1.txtSLCd.Value) = "" then
		Call DisplayMsgBox("169902","X", "X", "X")    
		frm1.txtSLCd.focus
		Exit Function
	End If


	If Trim(frm1.txtBpCd.Value) = "" then
		Call DisplayMsgBox("17A002","X", "공급처", "X")   
		frm1.txtBpCd.focus
		Exit Function
	End If

  
	If 	CommonQueryRs(" A.SL_NM, B.PLANT_NM "," I_VMI_STORAGE_LOCATION A, B_PLANT B ", " A.PLANT_CD = B.PLANT_CD AND " & _
	    " A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
		If 	CommonQueryRs(" SL_NM "," I_VMI_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							
			Call DisplayMsgBox("162001","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Exit function
		ElseIf Trim(frm1.txtPlantCd.Value) <> Parent.gPlant Then
			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus
				Exit function
			End If
		Else
			Call DisplayMsgBox("169922","X","X","X")
			frm1.txtSLCd.focus
			Exit function	
		End If
	End If
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	frm1.txtSLNm.Value = lgF0(0)
	frm1.txtPlantNm.Value = lgF1(0)

	If 	CommonQueryRs(" BP_NM "," B_Biz_Partner ", " BP_CD = " & FilterVar(frm1.txtBpCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("229927","X","X","X")
		frm1.txtBpNm.Value = ""
		frm1.txtBpCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpNm.Value = lgF0(0)		
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	arrParam1 = Trim(frm1.txtPlantCd.Value)    
	arrParam2 = Trim(frm1.txtPlantNm.Value)    
	arrParam3 = Trim(frm1.txtSLCd.Value)       
	arrParam4 = Trim(frm1.txtSLNm.Value)       
	arrParam5 = Trim(frm1.txtBpCd.Value)       
	arrParam6 = Trim(frm1.txtBpNm.Value)       
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_ItemCd
		arrParam7 = Trim(.Text)                
		.Col = C_ItemNm
		arrParam8 = Trim(.Text) 
	End With
	
	iCalledAspName = AskPRAspName("I1523PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1523PA1","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrParam7, arrParam8), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetLotNo(arrRet)
	End If
End Function

'==========================================  2.4.3 Set???()  =============================================
'------------------------------------------  SetPlant()  --------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	frm1.txtPlantCd.focus	
End Function

'------------------------------------------  SetItemDocumentNo()  --------------------------------------------------
Function SetItemDocumentNo(byRef arrRet)
	frm1.txtItemDocumentNo.Value    = arrRet(0)		
	frm1.txtItemDocumentNo.focus	
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
Function SetItem(byRef arrRet)
	With frm1.vspdData
		.ReDraw = False
		Call .SetText(C_ItemCd,	.ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm,	.ActiveRow, arrRet(1))
		Call .SetText(C_Specification,	.ActiveRow, arrRet(2))
		Call .SetText(C_BasicUnit,	.ActiveRow, arrRet(3))
		Call .SetText(C_EntryUnit,	.ActiveRow, arrRet(3))
		Call .SetText(C_LotNo,		.ActiveRow, "")
		Call .SetText(C_LotSubNo,	.ActiveRow, "")
		Call .SetText(C_TrackingNo,	.ActiveRow, "")
		Call .SetText(C_LotFlg,		.ActiveRow, arrRet(7))
		Call .SetText(C_TrackingFlg,.ActiveRow, arrRet(12))
		
		If arrRet(7) = "Y" and arrRet(13) = "M" Then
			ggoSpread.SpreadUnLock  C_LotNo,       .ActiveRow, C_LotNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetRequired  C_LotSubNo,       .ActiveRow, .ActiveRow
		ElseIf arrRet(7) = "Y" and arrRet(13) = "A" Then
			Call DisplayMsgBox("169968","X", "X", "X") 
			ggoSpread.SpreadUnLock  C_LotNo,       .ActiveRow, C_LotNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetRequired  C_LotSubNo,       .ActiveRow, .ActiveRow
		Else
			ggoSpread.SSSetProtected  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotSubNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotNoPopup,       .ActiveRow, .ActiveRow
		End If
		
		If arrRet(12) = "Y" Then
			ggoSpread.SpreadUnLock  C_TrackingNo,       .ActiveRow, C_TrackingNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_TrackingNo,       .ActiveRow, .ActiveRow
		Else 
			ggoSpread.SSSetProtected  C_TrackingNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_TrackingNoPopup,       .ActiveRow, .ActiveRow
		End If
		
		.Col = C_InspFlag
		.Row = .ActiveRow
		If arrRet(11) = "Y" Then
			.Text = "1"
		Else
			.Text = "0"
		End If
		.ReDraw = True
	End With
End Function

'------------------------------------------  SetTrackingNo()  --------------------------------------------------
Function SetTrackingNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo,	.ActiveRow, arrRet(0))
	End With
End Function

'------------------------------------------  SetEntryUnit()  --------------------------------------------------
Function SetEntryUnit(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_EntryUnit,	.ActiveRow, arrRet(0))
	End With
End Function

'------------------------------------------  SetLotNo()  --------------------------------------------------
Function SetLotNo(byRef arrRet)
	With frm1.vspdData
		.ReDraw = False
		
		Call .SetText(C_ItemCd,	.ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm,	.ActiveRow, arrRet(1))
		Call .SetText(C_Specification,	.ActiveRow, arrRet(7))
		Call .SetText(C_BasicUnit,	.ActiveRow, arrRet(3))
		Call .SetText(C_EntryUnit,	.ActiveRow, arrRet(3))
		Call .SetText(C_TrackingNo,	.ActiveRow, arrRet(4))
		Call .SetText(C_LotNo,	.ActiveRow, arrRet(5))
		Call .SetText(C_LotSubNo,	.ActiveRow, arrRet(6))
		Call .SetText(C_LotFlg,	.ActiveRow, arrRet(8))
		Call .SetText(C_TrackingFlg,	.ActiveRow, arrRet(9))
		If arrRet(8) = "Y" Then
			ggoSpread.SpreadUnLock  C_LotNo,       .ActiveRow, C_LotNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetRequired  C_LotSubNo,       .ActiveRow, .ActiveRow
		Else 
			ggoSpread.SSSetProtected  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotSubNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotNoPopup,       .ActiveRow, .ActiveRow
		End If
		If arrRet(9) = "Y" Then
			ggoSpread.SpreadUnLock  C_TrackingNo,       .ActiveRow, C_TrackingNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_TrackingNo,       .ActiveRow, .ActiveRow
		Else 
			ggoSpread.SSSetProtected  C_TrackingNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_TrackingNoPopup,       .ActiveRow, .ActiveRow
		End If

		.Col = C_InspFlag
		.Row = .ActiveRow
		If arrRet(10) = "Y" Then
			.Text = "1"
		Else
			.Text = "0"
		End If

		.ReDraw = True
	End With
End Function

'#########################################################################################################
'												3. Event부 
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                                        					
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
    Call ggoOper.FormatDate(frm1.txtDocumentYear, Parent.gDateFormat, 3)                   
    Call ggoOper.LockField(Document, "N")                                   				
    
    StartDate = "<%=GetSvrDate%>"
    
    Call InitVariables                                                      				
    Call SetdefaultVal
    Call InitSpreadSheet                                                    				
	
	Call SetToolbar("11101101000011")									
End Sub

'=======================================================================================================
'   Event Name : txtDocumentYear_DblClick(Button)
'=======================================================================================================
Sub txtDocumentYear_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentYear.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentYear.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentYear_Change()
'=======================================================================================================
Sub txtDocumentYear_Change()
	If lgIntFlgMode = Parent.OPMD_CMODE Then	
		lgBlnFlgChgValue = False
	Else
		lgBlnFlgChgValue = True	
	End if
End Sub

'=======================================================================================================
'   Event Name : txtDocumentYear_KeyPress()
'=======================================================================================================
Sub txtDocumentYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDtt_Change()
'=======================================================================================================
Sub txtDocumentDt_Change()
	If lgIntFlgMode = Parent.OPMD_CMODE Then	
		lgBlnFlgChgValue = False
	Else
		lgBlnFlgChgValue = True	
	End if
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'******************************  3.2.1 Object Tag 처리  *********************************************
'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

With frm1.vspdData

	If Col = C_ItemCd Then
		
		'공장코드가 있는 지 체크 
		If Trim(frm1.txtPlantCd.Value) = "" then
			Call DisplayMsgBox("169901","X", "X", "X")   
			frm1.txtPlantCd.focus
			Exit Sub
		ElseIf Trim(frm1.txtPlantCd.Value) <> Parent.gPlant Then
		'공장 체크 함수 호출 
			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus
				Exit Sub
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)
		End If
		
		.Row = Row		
		.Col = Col
		If 	CommonQueryRs(" B.item_nm,A.recv_inspec_flg,B.spec,B.basic_unit,A.lot_flg,A.tracking_flg,C.lot_gen_mthd","B_ITEM_BY_PLANT A left outer join B_ITEM B ON A.item_cd = B.item_cd left outer join B_LOT_CONTROL C ON A.plant_cd = C.plant_cd and A.item_cd = C.item_cd", _
		    "A.material_type = " & FilterVar("30", "''", "S") & " AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.item_cd = " & FilterVar(.Text, "''", "S"), _
		    lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			    
			Call .SetText(C_ItemNm,	Row, "")
			Call .SetText(C_InspFlag,	Row, "0")
			Call .SetText(C_Specification,	Row, "")
			Call .SetText(C_BasicUnit,	Row, "")
			Call .SetText(C_EntryUnit,	Row, "")
			Call .SetText(C_LotNo,	Row, "")
			Call .SetText(C_LotSubNo,	Row, "")
			Call .SetText(C_TrackingNo,	Row, "")
			ggoSpread.SSSetProtected  C_TrackingNo,			.ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_TrackingNoPopup,	.ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotNo,				.ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotSubNo,			.ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotNoPopup,			.ActiveRow, .ActiveRow
			.focus
			Exit Sub
		End If

		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		lgF4 = Split(lgF4, Chr(11))
		lgF5 = Split(lgF5, Chr(11))
		lgF6 = Split(lgF6, Chr(11))

		.ReDraw = False

		Call .SetText(C_ItemNm,	Row, lgF0(0))
		Call .SetText(C_Specification,	Row, lgF2(0))
		Call .SetText(C_BasicUnit,	Row, lgF3(0))
		Call .SetText(C_EntryUnit,	Row, lgF3(0))
		Call .SetText(C_LotNo,	Row, "")
		Call .SetText(C_LotSubNo,	Row, "")
		Call .SetText(C_TrackingNo,	Row, "")
		Call .SetText(C_LotFlg,	Row, lgF4(0))
		Call .SetText(C_TrackingFlg,	Row, lgF5(0))
		
		.Col = C_InspFlag
		If Trim(lgF1(0)) = "Y" Then
			.Text = "1"
		Else
			.Text = "0"
		End if
		
		If lgF4(0) = "Y" AND lgF6(0) = "M" Then
			ggoSpread.SpreadUnLock  C_LotNo,       .ActiveRow, C_LotNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_LotNo,      .ActiveRow, .ActiveRow
			ggoSpread.SSSetRequired  C_LotSubNo,   .ActiveRow, .ActiveRow
		ElseIf lgF4(0) = "Y" AND lgF6(0) = "A" Then
			Call DisplayMsgBox("169968","X", "X", "X") 
			ggoSpread.SpreadUnLock  C_LotNo,       .ActiveRow, C_LotNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetRequired  C_LotSubNo,       .ActiveRow, .ActiveRow
		Else
			ggoSpread.SSSetProtected  C_LotNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotSubNo,    .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_LotNoPopup,  .ActiveRow, .ActiveRow
		End If
		
		If lgF5(0) = "Y" Then
			ggoSpread.SpreadUnLock  C_TrackingNo,       .ActiveRow, C_TrackingNoPopup, .ActiveRow
			ggoSpread.SSSetRequired  C_TrackingNo,      .ActiveRow, .ActiveRow
		Else 
			ggoSpread.SSSetProtected  C_TrackingNo,       .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected  C_TrackingNoPopup,  .ActiveRow, .ActiveRow
		End If
		
		.ReDraw = True
	End If
End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			.Row = Row

			Select case Col 
			Case C_ItemPopup
				.Col = C_ItemCd
				Call OpenItem(.Text)
			
			Case C_TrackingNoPopup
				.Col = C_TrackingNo
				Call OpenTrackingNo(.Text)
			
			Case C_EntryUnitPopup
				.Col = C_EntryUnit
				Call OpenEntryUnit(.Text)

			Case C_LotNoPopup
				Call OpenLotNo()
			End Select		
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData 
		If Row = NewRow Then Exit Sub
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_InspFlag Then
			.Col = Col
			.Row = Row									
			IF .Text = 1 Then
				.Col = 0
				.Text = ggoSpread.UpdateFlag
			Elseif .Text = 0 Then
				.Col = 0
				.Text = ""
			End if  
		End If	
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	
		If (lgStrPrevKey <> "" and lgStrPrevKey1 <> "") Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
 		Call SetPopupMenuItemInf("1001111111") 
 	Else
 	 	Call SetPopupMenuItemInf("0101111111")
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
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData.MaxRows = 0 Then Exit Sub
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

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim TempInspDt 
    FncQuery = False                                             
    
    Err.Clear                                                             

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function						
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
     '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					
    Call ggoOper.LockField(Document, "N")                                     
    Call InitVariables
    Call SetToolbar("11101101000011")

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
    On Error Resume Next             
    
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
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call InitVariables                                                      
    Call SetDefaultVal    
    Call ggoSpread.SSSetColHidden(C_InspReqNo, C_InspReqNo, True)
    Call ggoSpread.SSSetColHidden(C_InspSts, C_InspSts, True)
    
    Call SetToolbar("11101101000011")
    
    FncNew = True                                                         

End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                    
    
    Err.Clear                                                          
    On Error Resume Next  

	'-----------------------
    'Check content area
    '-----------------------
    If frm1.vspdData.MaxRows < 1 then
       Call DisplayMsgBox("900001","X", "X", "X")  
	   exit function
	End if 
	  
    If Not chkField(Document, "2")  Then Exit Function       
    If Not ggoSpread.SSDefaultCheck Then Exit Function

    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X", "X", "X")                           
        Exit Function
    End If


	If 	CommonQueryRs(" A.SL_NM, B.PLANT_NM "," I_VMI_STORAGE_LOCATION A, B_PLANT B ", " A.PLANT_CD = B.PLANT_CD AND " & _
	    " A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
		If 	CommonQueryRs(" SL_NM "," I_VMI_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							
			Call DisplayMsgBox("162001","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Set gActiveElement = document.activeElement
			Exit function
		Else
			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				Call DisplayMsgBox("169922","X","X","X")
				frm1.txtSLCd.focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If
	End If
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1, Chr(11))
	frm1.txtSLNm.Value = lgF0(0)
	frm1.txtPlantNm.Value = lgF1(0)

	If 	CommonQueryRs(" BP_NM "," B_Biz_Partner ", " BP_CD = " & FilterVar(frm1.txtBpCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("229927","X","X","X")
		frm1.txtBpNm.Value = ""
		frm1.txtBpCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpNm.Value = lgF0(0)	

    '-----------------------
    'Save function call area
    '-----------------------
	If DBSave() = False Then Exit Function

    FncSave = True                                                        
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncInsertRow
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow ="" Then Exit Function
	End if
	
	With frm1.vspdData
		.focus
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData
    	ggoSpread.InsertRow  .ActiveRow,  imRow
    	SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
    	
    	.Col = C_InspFlag
    	.Text = "0"
    	.ReDraw = True
    End With
    
    Set gActiveElement = document.activeElement
    If Err.number = 0 Then
		FncInsertRow = True
	End If
End Function

'========================================================================================
' Function Name : FncDeleteRow
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
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
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
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
    
    Err.Clear                                                          
    
    Call LayerShowHide(1)       
    
    DbQuery = False
    With frm1    
		strVal = BIZ_PGM_QRY_ID &	"?txtMode="				& Parent.UID_M0001					& _						
									"&txtPlantCd="			& Trim(.txtPlantCd.value)			& _				
									"&txtItemDocumentNo="	& Trim(.txtItemDocumentNo.value)	& _
									"&txtDocumentYear="		& Trim(.txtDocumentYear.Year)		& _
									"&lgStrPrevKey="		& Trim(lgStrPrevKey)				& _
									"&lgStrPrevKey1="		& Trim(lgStrPrevKey1)				& _
									"&txtMaxRows="			& .vspdData.MaxRows

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
    Call SetToolbar("11101011000111")
    Call ggoSpread.SSSetColHidden(C_InspReqNo, C_InspReqNo, False)
    Call ggoSpread.SSSetColHidden(C_InspSts, C_InspSts, False)
    Call ggoOper.LockField(Document, "Q")
    frm1.txtItemDocumentNo.focus					
End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim strVal
	Dim PvArr
	
    Call LayerShowHide(1)

    Err.Clear		
	
    DbSave = False                                                     
 
	frm1.txtMode.value = Parent.UID_M0002
	frm1.hDocumentYear.value = frm1.txtDocumentYear.Year

	With frm1.vspdData
    '-----------------------
    'Data manipulate area
    '-----------------------
		lGrpCnt = 0
		ReDim PvArr(0)
		
		For lRow = 1 To .MaxRows
    
			.Row = lRow
			.Col = 0
        
	        Select Case .Text
				Case ggoSpread.InsertFlag	
					ReDim Preserve PvArr(lGrpCnt)

					frm1.txtpvCommandMode.value = "C"
			
					.Col = C_ItemCd
					strVal = Trim(.Text) & Parent.gColSep	
					
					.Col = C_EntryQty
					if uniCdbl(.Value) = 0 then
						Call DisplayMsgBox("169918","X", "X", "X")
						Call LayerShowHide(0)
						Exit Function
					End if
					strVal = strVal & Trim(.Value) & Parent.gColSep 
					
					.Col = C_EntryUnit
					strVal = strVal & Trim(.Text) & Parent.gColSep 
					
					.Col = C_InspFlag
					If .Text <> "0" Then
						If GetSetupMod(Parent.gSetupMod, "q") = "Y" then
							strVal = strVal & "Y" & Parent.gColSep
						Else
							Call DisplayMsgBox("169966","X", "X", "X")
							Call LayerShowHide(0)
							Exit Function
						End if 
					Else
						strVal = strVal & "N" & Parent.gColSep 
					End if
											
					.Col = C_TrackingNo
					strVal = strVal & Trim(.Text) & Parent.gColSep 
					.Col = C_LotNo
					strVal = strVal & Trim(.Text) & Parent.gColSep 
					.Col = C_LotSubNo
					strVal = strVal & Trim(.Text) & Parent.gColSep 
					strVal = strVal & lRow & parent.gRowSep               

    				PvArr(lGrpCnt) = strVal
    				lGrpCnt = lGrpCnt + 1

			    Case ggoSpread.UpdateFlag	
					frm1.txtpvCommandMode.value = "U"
					
			    Case ggoSpread.DeleteFlag	
					ReDim Preserve PvArr(lGrpCnt)

					frm1.txtpvCommandMode.value = "D"
					
					.Col = C_SeqNo
					strVal = Trim(.Text) & Parent.gColSep	
					.Col = C_ItemCd
					strVal = strVal & Trim(.Text) & Parent.gColSep	
					strVal = strVal & lRow & parent.gRowSep              
					
					PvArr(lGrpCnt) = strVal
					lGrpCnt = lGrpCnt + 1
		    End Select
		Next
	End With
	
	frm1.txtMaxRows.value = lGrpCnt
	frm1.txtSpread.value  = Join(PvArr, "")
	
	If lGrpCnt <= 0 then				
		Call DisplayMsgBox("800161","X", "X", "X")    
		Call LayerShowHide(0)
		Exit Function
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>VMI 입고등록</font></td>
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
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemDocumentNo" SIZE=16 MAXLENGTH=16 tag="12XXXU" ALT="입고번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemDocumentNoPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemDocumentNo()"></TD>
								<TD CLASS="TD5" NOWRAP>년도</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/i1521ma1_fpDateTime1_txtDocumentYear.js'></script>
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
					<TD CLASS="TD5" NOWRAP>VMI 창고</TD>
					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=8 MAXLENGTH=7 tag="23X1XU" ALT="VMI 창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=20 tag="24"></TD>
					<TD CLASS="TD5" NOWRAP>공급처</TD>
					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=8 MAXLENGTH=10 tag="23X1XU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24"></TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP>입고일자</TD>
					<TD CLASS="TD6" NOWRAP>
							<script language =javascript src='./js/i1521ma1_I303495872_txtDocumentDt.js'></script></TD>
					<TD CLASS="TD5" NOWRAP>입고번호</TD>
					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemDocumentNo2" SIZE=16 MAXLENGTH=16 tag="25XXXU" ALT="입고번호"></TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP>비고</TD>
					<TD CLASS="TD6" NOWRAP  COLSPAN=3><INPUT TYPE=TEXT NAME="txtDocumentText" SIZE=60 MAXLENGTH=60 tag="25" ALT="비고"></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
						<script language =javascript src='./js/i1521ma1_OBJECT1_vspdData.js'></script>
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

