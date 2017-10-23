<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : INTERFACE  
'*  2. Function Name        : 
'*  3. Program ID           : XI224MA1_KO441
'*  4. Program Name         : Compose 현황
'*  5. Program Desc         : Compose 현황
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/03.27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Han cheol
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

Const BIZ_PGM_ID = "XI224MB1_KO441.asp"												'☆: Head Query 비지니스 로직 ASP명 

'☆: Spread Sheet의 Column별 상수 
Dim C_Select
Dim C_ProdNO
Dim C_ProSeq
Dim C_ItemCD
Dim C_ItemNM
Dim C_PLotNO
Dim C_ChkQty
Dim C_CrtFlg
Dim C_EReqNo
Dim C_SndDHM
Dim C_RcvDHM
Dim C_APPFlg
Dim C_ErrDsc

Dim IsOpenPop			'Popup

'==================================================================================================================
Sub initSpreadPosVariables()  
	C_Select =  1	
	C_ProdNO =  2	
	C_ProSeq =  3	
	C_ItemCD =  4	
	C_ItemNM =  5	
	C_PLotNO =  6	
	C_ChkQty =  7
	C_CrtFlg =  8	
	C_EReqNo =  9
	C_SndDHM = 10
	C_RcvDHM = 11
	C_APPFlg = 12
	C_ErrDsc = 13
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
	frm1.rdoCfmAll.checked = True
	frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
	
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True

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
		ggoSpread.Spreadinit "V20051105",, parent.gAllowDragDropSpread
		.ReDraw = false
	    .MaxCols = C_ErrDsc + 1													'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols															'☜: 공통콘트롤 사용 Hidden Column
	    .ColHidden = True

	    .MaxRows = 0
	    
        Call GetSpreadColumnPos("A")	    

		ggoSpread.SSSetCheck	C_Select,	"선택",				 5,,,true		
		ggoSpread.SSSetEdit		C_ProdNO,	"제조오더번호",		15
		ggoSpread.SSSetEdit		C_ProSeq,	"의뢰순번",			 8,2,,3
		ggoSpread.SSSetEdit		C_ItemCD,	"품목코드",			15,,,18,2
	    ggoSpread.SSSetEdit		C_ItemNM,	"품목명",			20,,,40
	    ggoSpread.SSSetEdit		C_PLotNO,	"LOT 번호",			18,,,25,2
		ggoSpread.SSSetFloat	C_ChkQty,	"검사의뢰수량",		12, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_CrtFlg,	"생성구분",			 8,2,,1,2
		ggoSpread.SSSetEdit		C_EReqNo,	"ERP검사의뢰번호",	15,,,18,2
	    ggoSpread.SSSetEdit		C_SndDHM,	"MES 전송일시",		20,,,30
	    ggoSpread.SSSetEdit		C_RcvDHM,	"ERP 수신일시",		20,,,30
	    ggoSpread.SSSetEdit		C_APPFlg,	"ERP반영여부",		10,2,,,2
	    ggoSpread.SSSetEdit		C_ErrDsc,	"에러내역",			25,,,40

		Call ggoSpread.SSSetColHidden(C_Select,  C_Select,  True)

		.ReDraw = true
    End With
End Sub

'==================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1
		.vspdData.ReDraw = False
	    
		ggoSpread.SSSetProtected C_Select, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_ProdNO, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCD, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PLotNO, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ChkQty, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CrtFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_EReqNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SndDHM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RcvDHM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_APPFlg, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ErrDsc, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
    End With
End Sub

'==================================================================================================================
'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()        
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtInspReqNo.ClassName) = UCase(Parent.UCN_PROTECTED)  Then
		Exit Function
	End If
	
	If Trim(frm1.txtPlant.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'공장정보가 필요합니다 
		frm1.txtPlant.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlant.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo.Value)
	Param4 = ""				'검사분류 
	Param5 = ""				'검사진행현황 
	
	iCalledAspName = AskPRAspName("q2512pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q2512pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, _
									Array(Window.Parent, Param1, Param2, Param3,  Param4, Param5), _
									"dialogWidth=820px; dialogHeight=500px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtInspReqNo.value = arrRet(0)
	End If
	
	frm1.txtInspReqNo.Focus	
	Set gActiveElement = document.activeElement
End Function

'==================================================================================================================
Function OpenItemCD()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목"
	arrParam(1) = "B_ITEM"
	arrParam(2) = Trim(frm1.txtItemCD.value)
	arrParam(4) = ""
	arrParam(5) = ""								

	arrField(0) = "ITEM_CD"									
	arrField(1) = "ITEM_NM"									
	    
	arrHeader(0) = "품목코드"								
	arrHeader(1) = "품 목 명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCD.value = arrRet(0)
		frm1.txtItemNM.value = arrRet(1)
		frm1.txtItemCD.focus
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
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_Select, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ProdNO, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ProSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemCD, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemNM, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_PLotNO, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ChkQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_CrtFlg, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_EReqNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SndDHM, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_RcvDHM, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_APPFlg, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ErrDsc, lRow, .vspdData.MaxRows
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
			
			C_Select = iCurColumnPos(1)
			C_ProdNO = iCurColumnPos(2)
			C_ProSeq = iCurColumnPos(3)
			C_ItemCD = iCurColumnPos(4)
			C_ItemNM = iCurColumnPos(5)
			C_PLotNO = iCurColumnPos(6)
			C_ChkQty = iCurColumnPos(7)
			C_CrtFlg = iCurColumnPos(8)
			C_EReqNo = iCurColumnPos(9)
			C_SndDHM = iCurColumnPos(10)
			C_RcvDHM = iCurColumnPos(11)
			C_APPFlg = iCurColumnPos(12)
			C_ErrDsc = iCurColumnPos(13)
    End Select    
End Sub

'==================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------

	frm1.btnSelect.style.display = "none"
	frm1.btnDisSelect.style.display = "none"

	Call InitVariables														'⊙: Initializes local global variables

	Call SetDefaultVal

	Call InitSpreadSheet

	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어

    If parent.gPlant <> "" Then
		frm1.txtPlant.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

End Sub

'==================================================================================================================
Function rdoCfmAll_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
End Function

Function rdoCfmYes_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmYes.value
End Function

Function rdoCfmNo_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmNo.value
End Function

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
    
    If frm1.vspdData.MaxRows = 0 Then           
       Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) ------------------------------------------------------------------------
    frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) ------------------------------------------------------------------------
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
    If Col <= C_Select Or NewCol <= C_Select Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
End Sub

'==================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown) 
	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0
			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row 
			frm1.vspdData.text = "" 
			lgBlnFlgChgValue = False					
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
Sub btnSelect_OnClick()
	Dim i

	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows

			frm1.vspdData.Col = C_APPFlg
			frm1.vspdData.Row = i
			
			If Ucase(Trim(frm1.vspdData.Text)) <> "Y" Then
				frm1.vspdData.Col = C_Select
				frm1.vspdData.Row = i
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_Select, i, 1)
			End If

		Next	
	End If	
End Sub

'==================================================================================================================
Sub btnDisSelect_OnClick()
	Dim i	
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows

			frm1.vspdData.Col = C_APPFlg
			frm1.vspdData.Row = i
			
			If Ucase(Trim(frm1.vspdData.Text)) <> "Y" Then
				frm1.vspdData.Col = C_Select
				frm1.vspdData.Row = i
				frm1.vspdData.value = 0

				Call vspdData_ButtonClicked(C_Select, i, 0)
			End If

		Next	
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

'========================================================================================
' Function Name : RegProd
' Function Desc : This function is update 
'========================================================================================
Function RegProd() 
    Dim IntRetCD 

    RegProd = False                                                      
    
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

    Call DbQuery("A")

    RegProd = True
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

'==================================================================================================================
Function DbDelete() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function DbDeleteOk()														
    On Error Resume Next                                                    
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

    If frm1.rdoCfmAll.checked then
		rdoFlag=""
	ElseIf frm1.rdoCfmYes.checked Then 
		rdoFlag ="Y"
	Else
		rdoFlag ="N"
	End If

	strVal = BIZ_PGM_ID & _
			"?txtMode="		 & strMode & _
			"&txtPlant="	 & Trim(frm1.txtPlant.value) & _
			"&txtConSoFrDt=" & Trim(frm1.txtConSoFrDt.text) & _
			"&txtConSoToDt=" & Trim(frm1.txtConSoToDt.text) & _
			"&txtItemCD="	 & Trim(frm1.txtItemCD.value) & _
			"&txtMakOrdNo="	 & Trim(frm1.txtMakOrdNo.value) & _
			"&txtInspReqNo=" & Trim(frm1.txtInspReqNo.value) & _
			"&txtCfmFlag="	 & rdoFlag & _
			"&txtUserId="	 & Parent.gUsrID & _
			"&txtMaxRows="	 & frm1.vspdData.MaxRows & _
			"&lgStrPrevKey=" & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True																	

End Function

'==================================================================================================================
Function DbQueryOk()														
	
    lgIntFlgMode = Parent.OPMD_UMODE	 
							
    Call SetToolbar("11000000000111")					   
	Call SetQuerySpreadColor(1)
	lgBlnFlgChgValue = False
	
	If frm1.vspdData.Maxrows > 0 then
		frm1.btnSelect.disabled = False
		frm1.btnDisSelect.disabled = False
	End If
		
End Function

'==================================================================================================================
Function DbSave() 
    Err.Clear																
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	
    DbSave = False                                                    

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		lGrpCnt = 0    
		strVal = ""
    
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

			If Trim(.vspdData.Text) <> "" Then
				'--- 수주번호
				.vspdData.Col = C_SoNo
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

				'--- 수주순번
				.vspdData.Col = C_SoSeq
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

				'--- 마감여부
				.vspdData.Col = C_CloseFlag

				If Trim(.vspdData.Text) = "Y" Then
					strVal = strVal & "N" & Parent.gColSep
				Else
					strVal = strVal & "Y" & Parent.gColSep
				End If

				strVal = strVal & lRow & Parent.gRowSep

			    lGrpCnt = lGrpCnt + 1 
			End If
		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)									'☜: 비지니스 ASP 를 가동 
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
End Function

'==================================================================================================================
Function DbSaveOk()

    Call InitVariables

	Call ggoOper.ClearField(Document, "2")										
    
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
										<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConSoFrDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="납기시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtConSoToDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="납기종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemCD" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="11XXXU" ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCD" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCD()">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMakOrdNo" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMakOrdNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMakOrdNo()"></TD>
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>ERP검사의뢰번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE="20" MAXLENGTH="18" ALT="검사의뢰번호" TAG="11XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnInspReqNoPopup ONCLICK=vbscript:OpenInspReqNo() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON"></TD>
									<TD CLASS=TD5 NOWRAP>ERP반영여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio class="RADIO" name="rdoCfmflag" id="rdoCfmAll" value="A" tag="11X" checked><label for="rdoCfmAll">전체</label>&nbsp;&nbsp;
										<input type=radio class="RADIO" name="rdoCfmflag" id="rdoCfmYes" value="Y" tag="11X"><label for="rdoCfmYes">반영</label>&nbsp;&nbsp;
										<input type=radio class="RADIO" name="rdoCfmflag" id="rdoCfmNo" value="N" tag="11X"><label for="rdoCfmNo">미반영</label>
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
						<BUTTON NAME="btnRec" style="WIDTH: 110px "CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()">MES정보수신</BUTTON>&nbsp;&nbsp;
						<BUTTON NAME="btnReg" style="WIDTH: 110px "CLASS="CLSMBTN" ONCLICK="vbscript:RegProd()">검사요청등록</BUTTON>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
						<BUTTON NAME="btnSelect" style="WIDTH: 100px "  CLASS="CLSMBTN">일괄선택</BUTTON>&nbsp;&nbsp;	
						<BUTTON NAME="btnDisSelect" style="WIDTH: 100px " CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</TD>
					<TD align=right>
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

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
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
