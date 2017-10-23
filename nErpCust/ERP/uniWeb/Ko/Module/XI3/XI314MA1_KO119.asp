<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : INTERFACE   
'*  2. Function Name        : 
'*  3. Program ID           : XI314MA1_KO119
'*  4. Program Name         : Tray/대차정보 수신
'*  5. Program Desc         : Tray/대차정보 수신
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("d", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "XI314MB1_KO119.asp"												'☆: Head Query 비지니스 로직 ASP명 

'PLANT_CD, PALLET_NO, TRAY_NO, ITEM_CD, SUB_LOT_NO, IF_SEQ, CREATE_TYPE
'☆: Spread Sheet의 Column별 상수 
Dim C_HoldFlag
Dim C_PalletNo
Dim C_TrayNo
Dim C_SecItemCD
Dim C_SecItemNM
Dim C_ItemCD
Dim C_ItemNM
Dim C_SLotNO
Dim C_IfSeq
Dim C_PLotNO
Dim C_PalQty
Dim C_TryQty
Dim C_ProdDT
Dim C_ProdNO
Dim C_CrtFlg
Dim C_Status
Dim C_SndDHM
Dim C_RcvDHM
Dim C_LabeNO
Dim	C_ErrDsc
Dim	C_SEC_INVOICE_NO

Dim IsOpenPop			'Popup

'==================================================================================================================
Sub initSpreadPosVariables()  
	C_HoldFlag	     =  1
	C_PalletNo	     =  2	
	C_TrayNo	     =  3
	C_SecItemCD	     =  4	
	C_SecItemNM	     =  5	
	C_ItemCD	     =  6	
	C_ItemNM	     =  7
	C_SLotNO	     =  8
	C_IfSeq		     =  9
	C_PLotNO	     = 10
	C_PalQty	     = 11
	C_TryQty	     = 12
	C_ProdDT	     = 13
	C_ProdNO	     = 14
	C_LabeNO	     = 15
	C_CrtFlg	     = 16
	C_Status	     = 17
	C_SndDHM	     = 18
	C_RcvDHM	     = 19
	C_ErrDsc	     = 20
	C_SEC_INVOICE_NO = 21
	
End Sub

'==================================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'==================================================================================================================
Sub SetDefaultVal()
	frm1.txtPlant.focus
	lgBlnFlgChgValue = False
	frm1.txtConSoFrDt.text = EndDate
	frm1.txtConSoToDt.text = EndDate
End Sub

'==================================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub


'==================================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021104",, parent.gAllowDragDropSpread
		.ReDraw = false

	    .MaxCols = C_SEC_INVOICE_NO + 1													'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols															'☜: 공통콘트롤 사용 Hidden Column
	    .ColHidden = True
	
	    .MaxRows = 0
	    
        Call GetSpreadColumnPos("A")
        
        ggoSpread.SSSetCheck	C_HoldFlag,	      "출하중지",		8,,,true		
		ggoSpread.SSSetEdit		C_PalletNo,	      "Pallet번호",		30,,,30,2
		ggoSpread.SSSetEdit		C_TrayNo,	      "Tray 번호",		15,,,17,2
		ggoSpread.SSSetEdit		C_SecItemCD,      "SEC 품목코드",	18,,,18,2
	    ggoSpread.SSSetEdit		C_SecItemNM,      "SEC 품목명",		25,,,40
		ggoSpread.SSSetEdit		C_ItemCD,	      "품목코드",		18,,,18,2
	    ggoSpread.SSSetEdit		C_ItemNM,	      "품목명",			25,,,40
	    ggoSpread.SSSetEdit		C_SLotNO,	      "Sub LOT 번호",	15,,,25
	    ggoSpread.SSSetEdit		C_IfSeq,	      "전송순번",		 8
	    ggoSpread.SSSetEdit		C_PLotNO,	      "LOT 번호",		15,,,25
	    ggoSpread.SSSetFloat	C_PalQty,	      "Pallet당 수량",	12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_TryQty,	      "Tray당 수량",		12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetDate		C_ProdDT,	      "제조일자",		10, 2,Parent.gDateFormat
	    ggoSpread.SSSetEdit		C_ProdNO,	      "제조오더번호",	15,,,18,2
	    ggoSpread.SSSetEdit		C_LabeNO,	      "통합라벨번호",	22,,,22,2
	    ggoSpread.SSSetEdit		C_CrtFlg,	      "생성구분",		 8
	    ggoSpread.SSSetEdit		C_Status,	      "상태",			 8
	    ggoSpread.SSSetEdit		C_SndDHM,	      "MES전송일시",	25,2
		ggoSpread.SSSetEdit		C_RcvDHM,	      "ERP수신일시",	25,2
	    ggoSpread.SSSetEdit		C_ErrDsc,	      "에러내역",		25,,,40
	    ggoSpread.SSSetEdit		C_SEC_INVOICE_NO, "Invoice번호",	15,,,16
	    
	    Call ggoSpread.SSSetSplit2(C_HoldFlag)

		.ReDraw = true
    End With
End Sub

'==================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1
		.vspdData.ReDraw = False
	    
		ggoSpread.SSSetProtected C_PalletNo,       pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrayNo,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SecItemCD,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SecItemNM,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCD,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNM,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SLotNO,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IfSeq,          pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PLotNO,         pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected C_PalQty,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TryQty,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProdDT,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProdNO,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LabeNO,         pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected C_CrtFlg,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Status,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SndDHM,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RcvDHM,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ErrDsc,         pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SEC_INVOICE_NO, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
    End With
End Sub


'==================================================================================================================
' Name : OpenSecItem()
' Description : OpenItem PopUp
'==================================================================================================================
Function OpenSecItem()

	Dim arrRet
	Dim arrParam(5), arrField(7), arrHeader(7)
	Dim iCalledAspName

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "SEC품목"
	arrParam(1) = "(SELECT FR1_CD PAR_ITEM_CD, FR2_CD PAR_ITEM_NM, FR3_CD PAR_SPEC, TO1_CD ITEM_CD FROM J_CODE_MAPPING  WHERE MAJOR_CD = 'J0010' AND MINOR_CD = '0000' ) A, B_ITEM B"
	arrParam(2) = FilterVar(Trim(frm1.txtSecItemCd.Value),"","SNM")

	arrParam(4) = " RTRIM(A.ITEM_CD) *= B.ITEM_CD "
	arrParam(4) = arrParam(4) & " AND B.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(EndDate), "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B.VALID_TO_DT   >=  " & FilterVar(UNIConvDate(EndDate), "''", "S") & " "

	arrParam(5) = "품목"

	arrField(0) = "A.PAR_Item_Cd"
	arrField(1) = "A.PAR_Item_NM"
	arrField(2) = "A.PAR_SPEC"
	arrField(3) = "B.ITEM_CD"
	arrField(4) = "B.ITEM_NM"
	arrField(5) = "B.SPEC"

	arrHeader(0) = "SEC 품목코드"
	arrHeader(1) = "SCE 품목명"
	arrHeader(2) = "SEC SPEC"
	arrHeader(3) = "내부 품목코드"
	arrHeader(4) = "내부 품목명"
	arrHeader(5) = "내부 SPEC"

	iCalledAspName = AskPRAspName("J2b02pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "J2b02pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField, arrHeader), "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtSecItemCd.Value   = arrRet(0)
		frm1.txtSecItemNm.Value   = arrRet(1)
		frm1.txtSecItemCd.focus
	End If

End Function

'==================================================================================================================
Function OpenMakOrdNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtMakOrdNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlant.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlant.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlant.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = Trim(frm1.txtMakOrdNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""

	arrRet = window.showModalDialog(iCalledAspName, _
									Array(Window.Parent, arrParam), _
									"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtMakOrdNo.value = arrRet(0)
		frm1.txtMakOrdNo.focus
	End If
End Function


'==================================================================================================================
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"					
	arrParam(1) = "B_PLANT"						
	arrParam(2) = Trim(frm1.txtPlant.value)		
	arrParam(3) = ""
	arrParam(4) = ""					
	arrParam(5) = "공장"					
		
	arrField(0) = "Plant_cd"						
	arrField(1) = "Plant_NM"					
	    
	arrHeader(0) = "공장"					
	arrHeader(1) = "공장명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlant.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlant.focus
	End If	
End Function

'==================================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	Dim Indx
	
    With frm1
		.vspdData.ReDraw = False

		With frm1.vspdData
			For Indx = 1 To .MaxRows
				.Row = Indx
				.Col = C_Status

				Select Case .Text
						Case "D"
							ggoSpread.SSSetProtected C_HoldFlag, Indx, Indx
							ggoSpread.SSSetProtected C_PalQty,   Indx, Indx
				End Select
		    Next
		End With

		ggoSpread.SSSetProtected C_PalletNo,       lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_TrayNo,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SecItemCD,      lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SecItemNM,      lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemCD,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemNM,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SLotNO,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_IfSeq,          lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PLotNO,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_TryQty,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ProdDT,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ProdNO,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LabeNO,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_CrtFlg,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_Status,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SndDHM,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_RcvDHM,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ErrDsc,         lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SEC_INVOICE_NO, lRow, .vspdData.MaxRows

		.vspdData.ReDraw = True
    End With
End Sub

'==================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_HoldFlag	     = iCurColumnPos(1)
			C_PalletNo	     = iCurColumnPos(2)
			C_TrayNo	     = iCurColumnPos(3)
			C_SecItemCD	     = iCurColumnPos(4)
			C_SecItemNM	     = iCurColumnPos(5)
			C_ItemCD	     = iCurColumnPos(6)
			C_ItemNM	     = iCurColumnPos(7)
			C_SLotNO	     = iCurColumnPos(8)
			C_IfSeq		     = iCurColumnPos(9)
			C_PLotNO	     = iCurColumnPos(10)
			C_PalQty	     = iCurColumnPos(11)
			C_TryQty	     = iCurColumnPos(12)
			C_ProdDT	     = iCurColumnPos(13)
			C_ProdNO	     = iCurColumnPos(14)
			C_LabeNO	     = iCurColumnPos(15)
			C_CrtFlg	     = iCurColumnPos(16)
			C_Status	     = iCurColumnPos(17)
			C_SndDHM	     = iCurColumnPos(18)
			C_RcvDHM	     = iCurColumnPos(19)
			C_ErrDsc	     = iCurColumnPos(20)
			C_SEC_INVOICE_NO = iCurColumnPos(21)

    End Select    
End Sub

'==================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	
	Call InitSpreadSheet
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어

    If parent.gPlant <> "" Then
		frm1.txtPlant.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSecItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlant.focus 
		Set gActiveElement = document.activeElement
	End If

End Sub

'==================================================================================================================
Sub txtConSoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConSoFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoFrDt.Focus
	End If
End Sub

'==================================================================================================================
Sub txtConSoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConSoToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoToDt.Focus
	End If
End Sub

'==================================================================================================================
Sub txtConSoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==================================================================================================================
Sub txtConSoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
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
 	End If
End Sub

'==================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
    End If

    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If
End Sub

'==================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'==================================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_HoldFlag Or NewCol <= C_HoldFlag Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
	
	Dim iDx
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Select Case Col
		Case C_PalQty

		Case Else
			Exit Sub
	End Select
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)		

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub

'==================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown) 
	If Col = C_HoldFlag And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
'			frm1.vspdData.Col = 0
'			frm1.vspdData.Row = Row 
'			frm1.vspdData.text = "" 
'			lgBlnFlgChgValue = False					
	    End Select
	End If
End Sub

'==========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    

    	If lgStrPrevKey <> "" Then		    							
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery("R")
    	End If
    End If    
End Sub


'==================================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                      
    
    Err.Clear      

    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtConSoFrDt, frm1.txtConSoToDt) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			  
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
    Call DbQuery("R")

    FncQuery = True	
End Function

'========================================================================================
' Function Name : RecMes
' Function Desc : This function is data query and display
'========================================================================================
Function RecMes() 
    Dim IntRetCD 

    RecMes = False                                                      
    
    Err.Clear      

    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtConSoFrDt, frm1.txtConSoToDt) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			  
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    Call DbQuery("T")

    RecMes = True	
End Function



'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                       '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                          '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'==================================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'==================================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_MULTI, False)
End Function

'==================================================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'==================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'==================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()          
	Call ggoSpread.ReOrderingSpreadData()	
	Call SetQuerySpreadColor(1)
End Sub

'==================================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows, lDelRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
	Frm1.vspdData.Row = Frm1.vspdData.ActiveRow
	Frm1.vspdData.Col = C_SEC_INVOICE_NO
	If Trim(Frm1.vspdData.Text) <> "" Then
		Call DisplayMsgBox("XX1134","x","x","x")
		Exit Function	
	End If
	
    With Frm1.vspdData3
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
   	
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    lgBlnFlgChgValue = True 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

    Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
'     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
'     Frm1.vspdData.Col = C_ISSUE_FLAG    : iDx = Frm1.vspdData.value
'     Frm1.vspdData.Col = C_ISSUE_FLAG_NM : Frm1.vspdData.value = iDx
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==================================================================================================================
Function DbQuery(strMode) 
    Dim strVal
	Dim rdoFlag

    Err.Clear


    DbQuery = False

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	strVal = BIZ_PGM_ID
	strVal = strVal & "?txtMode="      & strMode
	strVal = strVal & "&txtPlant="     & Trim(frm1.txtPlant.value)
	strVal = strVal & "&txtConSoFrDt=" & Trim(frm1.txtConSoFrDt.text)
	strVal = strVal & "&txtConSoToDt=" & Trim(frm1.txtConSoToDt.text)
	strVal = strVal & "&txtSecItemCd=" & Trim(frm1.txtSecItemCd.value)
	strVal = strVal & "&txtMakOrdNo="  & Trim(frm1.txtMakOrdNo.value)
	strVal = strVal & "&txtPalletNo="  & Trim(frm1.txtPalletNo.value)
	strVal = strVal & "&txtTrayNo="    & Trim(frm1.txtTrayNo.value)
	If frm1.rdoFlagAll.checked = True Then
		strVal = strVal & "&rdoFlag="      & "A"
	ElseIf frm1.rdoFlagCom.checked = True Then
		strVal = strVal & "&rdoFlag="      & "C"
	Else
		strVal = strVal & "&rdoFlag="      & "N"
	End If
	strVal = strVal & "&txtInvoiceNo="    & Trim(frm1.txtInvoiceNo.value)

	
	strVal = strVal & "&txtUserId="    & Parent.gUsrID
	strVal = strVal & "&txtMaxRows="   & frm1.vspdData.MaxRows
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True																	

End Function

'==================================================================================================================
Function DbQueryOk()														
    lgIntFlgMode = Parent.OPMD_UMODE	 
							
    Call SetToolbar("11101011000111")					   
	Call SetQuerySpreadColor(1)
	lgBlnFlgChgValue = False
		
End Function


'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim Indx
    Dim strVal, strDel
    Dim ColSep, RowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
	
	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size


	On Error Resume Next
	Err.Clear

    DbSave = False                                     '⊙: Processing is NG
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value         = "U"				   '☜: 저장 상태 
		.txtFlgMode.value      = lgIntFlgMode          '☜: 신규입력/수정 상태 
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	'한번에 설정한 버퍼의 크기 설정 
    iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    iFormLimitByte       = parent.C_FORM_LIMIT_BYTE     '102399byte
    
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			'버퍼의 초기화 
    ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	ColSep = parent.gColSep : RowSep = parent.gRowSep 
	iTmpCUBufferCount = -1  : iTmpDBufferCount = -1
	strCUTotalvalLen  = 0   : strDTotalvalLen  = 0

	With frm1.vspdData
		For Indx = 1 To .MaxRows
			.Row = Indx
			.Col = 0
			Select Case .Text
			    Case ggoSpread.UpdateFlag

			    							  strVal = ""
			    							  strVal = strVal & "U"									& ColSep	'⊙: U=Update
											  strVal = strVal & UCase(Trim(frm1.txtPlant.value))	& ColSep	'공장코드(PK)
					.Col = C_PalletNo		: strVal = strVal & UCase(Trim(.Text))					& ColSep	'Pallet 번호(PK)
					.Col = C_TrayNo			: strVal = strVal & UCase(Trim(.Text))					& ColSep	'Tray 번호(PK)
					.Col = C_ItemCD			: strVal = strVal & UCase(Trim(.Text))					& ColSep	'품목코드(PK)
					.Col = C_SLotNO			: strVal = strVal & UCase(Trim(.Text))					& ColSep	'Sub Lot번호(PK)
					.Col = C_IfSeq			: strVal = strVal & UNIConvNum(Trim(.Text), 0)			& ColSep	'전송순번(PK)
					.Col = C_CrtFlg    		: strVal = strVal & UCase(Trim(.Text))					& ColSep	'생성구분(PK)

					.Col = C_HoldFlag
										If UCase(Trim(.Text)) = "1" Then
											  strVal = strVal & "Y" & ColSep									'출하중지여부 : Yes
										Else
											  strVal = strVal & "N" & ColSep									'출하중지여부 : No
										End If
					
					.Col = C_PalQty			: strVal = strVal & UNIConvNum(Trim(.Text), 0)			& ColSep	'Pallet당 수량
											  strVal = strVal & Indx & RowSep
											  
				Case ggoSpread.DeleteFlag
											  strDel = ""
			    							  strDel = strDel & "D"									& ColSep	'⊙: U=Update
											  strDel = strDel & UCase(Trim(frm1.txtPlant.value))	& ColSep	'공장코드(PK)
					.Col = C_PalletNo		: strDel = strDel & UCase(Trim(.Text))					& ColSep	'Pallet 번호(PK)
					.Col = C_TrayNo			: strDel = strDel & UCase(Trim(.Text))					& ColSep	'Tray 번호(PK)
					.Col = C_ItemCD			: strDel = strDel & UCase(Trim(.Text))					& ColSep	'품목코드(PK)
					.Col = C_SLotNO			: strDel = strDel & UCase(Trim(.Text))					& ColSep	'Sub Lot번호(PK)
					.Col = C_IfSeq			: strDel = strDel & UNIConvNum(Trim(.Text), 0)			& ColSep	'전송순번(PK)
					.Col = C_CrtFlg    		: strDel = strDel & UCase(Trim(.Text))					& ColSep	'생성구분(PK)

					.Col = C_HoldFlag
										If UCase(Trim(.Text)) = "1" Then
											  strDel = strDel & "Y" & ColSep									'출하중지여부 : Yes
										Else
											  strDel = strDel & "N" & ColSep									'출하중지여부 : No
										End If
					
					.Col = C_PalQty			: strDel = strDel & UNIConvNum(Trim(.Text), 0)			& ColSep	'Pallet당 수량
											  strDel = strDel & Indx & RowSep
											  
			End Select
			
			.Col = 0
			Select Case .Text
			    Case ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then
			            Set objTEXTAREA = document.createElement("TEXTAREA")
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
				Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			         iTmpDBufferCount = iTmpDBufferCount + 1
			         If iTmpDBufferCount > iTmpDBufferMaxCount Then             '초기설정 버퍼의 조정 Max값을 넘으면 버퍼 크기 증가
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    '나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then	 
       DbSave = True
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Call InitVariables()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	
	Call RemovedivTextArea() 
	
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
	                                              '☆: Developer must customize
    If DbQuery("R") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement   
    
End Sub

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim Indx
		
	For Indx = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12xxxU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
									<TD CLASS=TD5 NOWRAP>MES송신기간</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConSoFrDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="MES송신시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtConSoToDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="MES송신종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Pallet/대차번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPalletNo" SIZE="30" MAXLENGTH="25" ALT="Pallet/대차번호" TAG="11XXXU" ></TD>
									<TD CLASS=TD5 NOWRAP>Tray번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrayNo" SIZE="20" MAXLENGTH="18" ALT="Tray번호호" TAG="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMakOrdNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMakOrdNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMakOrdNo()"></TD>
									<TD CLASS=TD5 NOWRAP>SEC 품목코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSecItemCd" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="SEC 품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSecItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSecItem()">&nbsp;<INPUT NAME="txtSecItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전송여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagAll tag="12"><LABEL FOR=rdoFlagAll>전체</LABEL>
												   &nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagCom tag="12"><LABEL FOR=rdoFlagCom>완료</LABEL>
												   &nbsp;<INPUT TYPE=radio CLASS=Radio NAME=rdoFlag ID=rdoFlagNo  tag="12" CHECKED><LABEL FOR=rdoFlagNo>미완료(선행라벨)</LABEL></TD>
								<TD CLASS=TD5 NOWRAP>Invoice 번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=text NAME=txtInvoiceNo SIZE=15 MAXLENGTH=16 tag=11xxxU ALT="Invoice번호"><!--IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btn1 ALIGN=top TYPE=button ONCLICK="vbScript:Call OpenInvoiceNoHdr(1)"--></TD>
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
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD align=left>
						<BUTTON NAME="btnRec" CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()" disabled>MES정보수신</BUTTON>&nbsp;
					</TD>
					<TD WIDTH=10>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX = -1></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID=divTextArea></P>
<TEXTAREA CLASS=hidden NAME=txtSpread   tag=24 TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtCfmFlag" tag="14" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHCfmFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="24"TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoFrDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHSoToDt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtStatusFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHStatusFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtBOFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHBOFlag" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
