<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S3112RA20
'*  4. Program Name         : BOM참조(수주내역등록)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/18
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
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
<TITLE></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                              
Dim lgMark                                                  
Dim IscookieSplit 

Dim arrReturn					
Dim arrReturn2

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
'StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
'Const BIZ_PGM_ID        = "s3112rb20.asp"
Const C_MaxKey          = 30                                   '☆☆☆☆: Max key value
'Const gstPaytermsMajor = "B9004"           
Const BIZ_PGM_QRY_ID = "s3112rb20.asp"		
'Const UID_M0004 ="site"
'spread sheet var
Dim C_Chk				
Dim C_ChildItem
Dim C_ChildItemNm
Dim C_ChildItemSpec
Dim C_ReqQty		
Dim C_ChildItemUnit
Dim C_JiGubun
Dim C_JoGubun
Dim C_HSCd
Dim C_ChildItemAcct
Dim C_VATType
Dim C_VATRate

'Const gstrPayTermsMajor = "B9004"
'Const gstrIncoTermsMajor = "B9006"

Dim strReturn					
Dim arrParam	
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim gblnWinEvent		                
                               
 
                               
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'=============================================================================================================
Sub InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    
    lgSortKey        = 1
	
	Redim arrReturn(0,0)
	Redim arrReturn2(1)
	
	arrReturn2(0) = arrReturn
	
	Self.Returnvalue = arrReturn2
End Sub

'================================================================================================================
Sub initSpreadPosVariables()
	
	C_Chk					=1
	C_ChildItem			=2
	C_ChildItemNm		=3
	C_ChildItemSpec	=4
	C_ReqQty				=5
	C_ChildItemUnit		=6
	C_JiGubun				=7
	C_JoGubun			=8
	C_HSCd				=9
	C_ChildItemAcct			=10
	C_VATType			=11
	C_VATRate			=12

End Sub
'=============================================================================================================

Sub SetDefaultVal()	

	Dim iArrayCodeArr
	Dim iArrayStrArguments
	
	iArrayStrArguments = Split(arrParent(1), PopupParent.gRowSep)

	frm1.txtPlant.value = iArrayStrArguments(0)
	frm1.txtPlantNm.value = iArrayStrArguments(1)
	frm1.txtHSoldToParty.value = iArrayStrArguments(2)
	frm1.txtHCurrency.value = iArrayStrArguments(3)	

	Err.Clear

	frm1.txtSODt.text =enddate
	frm1.txtPoNoSeq.text=""
	
	frm1.btnSelectAll.disabled =true
	frm1.btnDeSelect.disabled =true
	 

End Sub
'==========================================  InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	On Error Resume Next
    Err.Clear
	'****************************
    '지급구분 
    'List Minor code(유무상구분)    
    '****************************	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0= "" & chr(11) & lgF0
    lgF1= "" & chr(11) & lgF1
    Call SetCombo2(frm1.cboGubun, lgF0, lgF1, Chr(11)) 
    
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
		<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=============================================================================================================
Sub InitSpreadSheet()	
'	Call SetZAdoSpreadSheet("S3112ra41","S","A","V20030918", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
'	Call SetSpreadLock           
	Call initSpreadPosVariables()	
	
	ggoSpread.Source =frm1.vspdData			
	ggoSpread.Spreadinit "V20021214",,PopupParent.gAllowDragDropSpread	

	with frm1
		.vspdData.ReDraw = False
		.vspdData.MaxCols = C_VATRate+ 1
		.vspdData.MaxRows = 0
			
		Call GetSpreadColumnPos("A")	 
			
		ggoSpread.SSSetCheck	C_Chk, "선택", 10, , ,true,-1
		ggoSpread.SSSetEdit		C_ChildItem, 		"자품목", 20
		ggoSpread.SSSetEdit		C_ChildItemNm,		"자품목명", 30
		ggoSpread.SSSetEdit		C_ChildItemSpec,	"규격", 30	

		ggoSpread.SSSetFloat	C_ReqQty,			"필요량", 15, PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,1,,"Z"    					
		
		ggoSpread.SSSetEdit 		C_ChildItemUnit,	"기준단위", 12
		ggoSpread.SSSetEdit 		C_JiGubun,	 	"지급구분", 12
		ggoSpread.SSSetEdit 		C_JoGubun,	 	"조달구분", 12
		ggoSpread.SSSetEdit		C_HSCd,			"HS코드", 12
		ggoSpread.SSSetEdit		C_ChildItemAcct,			"품목계정", 12
		ggoSpread.SSSetEdit		C_VATType,			"VAT유형", 12
		ggoSpread.SSSetFloat	C_VATRate,			"VAT율", 15, "8" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,1,,"Z"    
	
		'ggoSpread.SSSetSplit2(1)											'frozen 기능 추가 
	
		Call ggoSpread.SSSetColHidden(C_HSCd, C_HSCd, True)
		Call ggoSpread.SSSetColHidden(C_ChildItemAcct, C_ChildItemAcct, True)
		Call ggoSpread.SSSetColHidden(C_VATType, C_VATType, True)
		Call ggoSpread.SSSetColHidden(C_VATRate, C_VATRate, True)
			
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
		SetSpreadLock "", 0, -1, ""
		.vspdData.ReDraw = True
	end with
End Sub
'================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            	C_Chk				= iCurColumnPos(1)   
				C_ChildItem		= iCurColumnPos(2)   
				C_ChildItemNm	= iCurColumnPos(3)   
				C_ChildItemSpec	= iCurColumnPos(4)   
				C_ReqQty			= iCurColumnPos(5)
				C_ChildItemUnit				= iCurColumnPos(6)
				C_JiGubun			= iCurColumnPos(7)
				C_JoGubun		= iCurColumnPos(8)
				C_HSCd			= iCurColumnPos(9)
				C_ChildItemAcct		= iCurColumnPos(10)
				C_VATType		= iCurColumnPos(11)
				C_VATRate		= iCurColumnPos(12)
    End Select    
End Sub


'=============================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)

   Dim i	
   
	ggoSpread.Source = frm1.vspdData
			
	frm1.vspdData.ReDraw = False			
	ggoSpread.SpreadLockWithOddEvenRowColor()
	
	ggoSpread.SpreadUnLock C_Chk, lRow
	for i = 2 to frm1.vspdData.MaxCols-1
		ggoSpread.SpreadLock i, lRow
	next	
	frm1.vspdData.ReDraw = True
End Sub

'=============================================================================================================
Function OpenSortPopup()	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=============================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	If frm1.txtPlant.readOnly = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"   
	arrParam(1) = "B_PLANT"       
	arrParam(2) = Trim(frm1.txtPlant.value)  
	arrParam(4) = ""       
	arrParam(5) = "공장"    
 
	arrField(0) = "PLANT_CD"    
	arrField(1) = "PLANT_NM"    
	   
	arrHeader(0) = "공장"     
	arrHeader(1) = "공장명"    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If 
	
End Function

'=============================================================================================================

Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD	

	
	If chkPoNo =false then 		
		Exit Function 
	ELSE
		If len(frm1.txtPoNo.value) <>0 and len(frm1.txtPoNoSeq.text)<>0 then
			CALL  fncLookUp
			EXIT Function 
		End If			
	END IF
		

	If lgIsOpenPop = True Then 
		lgIsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlant.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlant.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	lgIsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlant.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItem.value)	' Item Code
	arrParam(2) = ""						' Combo Set Data:"12!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    arrField(2) = 3 							' Field명(2) : "item_spec"		
    arrField(3) = 4 							' Field명(3) : "basic_unit"	
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "B1B11PA4", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItem(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItem.focus
	
End Function

'=============================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam
	Dim iCalledAspName
	Dim IntRetCD
	
	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function		
		
	lgIsOpenPop = True

	Redim arrParam(2)

	arrParam(0) = frm1.txtPlant.value
	arrParam(1) = frm1.txtPlantNm.value

	iCalledAspName = AskPRAspName("s3112ra21")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "s3112ra21", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False	
	If strRet(0) = "" Then
		Exit Function
	Else	
		frm1.txtPoNo.value = strRet(0)			'발주번호 
		frm1.txtPoNoSeq.text = strRet(1)		'발주순번 
		frm1.txtItem.value = strRet(2)			'품목 
		frm1.txtItemNm.value = strRet(3)		'품목명 
		frm1.txtItemSpec.value = strRet(4)		'규격 
		frm1.txtAmt.value = strRet(5)				'수량 
		frm1.txtUnit.value = strRet(6)				'기준단위 
		frm1.txtPlant.focus
	End If	
		
End Function
'=============================================================================================================
Function SetPlant(Byval arrRet)
	With frm1
		.txtPlant.value = arrRet(0) 
		.txtPlantNm.value = arrRet(1) 
		.txtPoNo.focus  
	End With
End Function
'=============================================================================================================
Function SetItem(byval arrRet)
	With frm1	
		.txtItem.Value    = arrRet(0)				'품목 
		.txtItemNm.Value    = arrRet(1)		'품목명 
		.txtITemSpec.value = arrRet(2)		'스펙 
		.txtUnit.value = arrRet(3)					'기준단위	
		.txtItem.focus
	End With
End Function

'=============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                     

	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables
	Call InitSpreadSheet()	
	
End Sub
'=============================================================================================================
Sub txtSODt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSODt.Action = 7 
        Call SetFocusToDocument("P")
		frm1.txtSODt.Focus
    End If
End Sub
'=============================================================================================================
Sub txtSODt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtPlant_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtItem_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtAmt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtPoNo_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtPoNoSeq_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
		'If trim(frm1.txtPoNo.value)<>"" and trim(frm1.txtPoNoSeq.text)<>"" then
		'	Call fncLookUp()
		'Else
			Call FncQuery()        			
		'End If    		
    End if    
End Sub

'=============================================================================================================
Sub txtGubun_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'=============================================================================================================
Sub txtItem_onFocus()
	Dim strVal
	If chkPoNo = false then Exit Sub
	
	With frm1
		If trim(.txtPoNo.value)  <>""  and  trim(.txtPoNoSeq.text) <>"" then
		                                                        
			Call fncLookUp()	
			.txtPoNo.focus	
			Exit sub			
		End If
	End With
End Sub
'=============================================================================================================
Sub txtItem_onChange()
	Dim strVal
	If chkPoNo = false then Exit Sub
	
	With frm1
		If trim(.txtPoNo.value)  <>""  and  trim(.txtPoNoSeq.text) <>"" then                                                           
			Call fncLookUp()		
		Else
			Call fncItem()
		End If
	End With
End Sub
'=============================================================================================================
Sub txtPoNo_onChange()
	With frm1
		If trim(.txtPoNo.value)  <>""  and  trim(.txtPoNoSeq.text) <>"" then		
			call fncLookUp()
		End If
	End With
End Sub
'=============================================================================================================
Sub txtPoNoSeq_onChange()

	With frm1
		If trim(.txtPoNo.value)  <>""  and  trim(.txtPoNoSeq.text) <>"" then			
			Call fncLookUp()			
		End If
	End With
End Sub

'=============================================================================================================
'발주번호와 발주 순번이 입력되었을경우 자동으로 해당 품목정보를 셋팅해준다.
'=============================================================================================================
Function fncLookUp()
	Dim strVal
	
	with frm1
		If trim(.txtPoNo.value)<>"" and trim(.txtPoNoSeq.text) <>"" then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001	
		strVal = strVal & "&txtMFlg=" & "ITEM"'Trim(.txtMFlg.value)				
		strVal = strVal & "&txtItem=" & Trim(.txtItem.value)				
		strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
		strVal = strVal & "&txtSoDt=" & Trim(.txtSoDt.text)	'
		strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)	'
		strVal = strVal & "&txtPoNoSeq=" & Trim(.txtPoNoSeq.text)	'		
		strVal = strVal & "&txtCur="	& Trim(.txtHCurrency.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2				
			       
		Call RunMyBizASP(MyBizASP, strVal)			
	End If
	End with
End Function
'=============================================================================================================
Function FncLookUpOk()
	'frm1.txtItem.focus
End Function

function fncItem()
	Dim strVal
	
	with frm1

		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001	
		strVal = strVal & "&txtMFlg=" & "ITEM"'Trim(.txtMFlg.value)				
		strVal = strVal & "&txtItem=" & Trim(.txtItem.value)				
		strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2				
			       
		Call RunMyBizASP(MyBizASP, strVal)			
'	End If
	End with

End Function 
'=============================================================================================================
'Function vspdData_DblClick(ByVal Col, ByVal Row)
 '   If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then  
  '      Exit Function
   ' End If
        
	'If frm1.vspdData.MaxRows > 0 Then
	'	If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
	'		Call OKClick
	'	End If
	'End If
'End Function	
'=============================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
		If NewRow = .MaxRows Then
			If lgPageNo <> "" Then							
				DbQuery
			End If
		End If
	End With
End Sub
	
'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub

	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
		If lgPageNo <> "" Then									       
	       Call DbQuery
		End If
	End If        
End Sub
	
'=============================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
		frm1.vspdData.Col = C_Chk
        If frm1.vspdData.Text Then
			frm1.vspdData.Text =False			
		Else
			frm1.vspdData.Text = True
		End If        
		frm1.btnSelectAll.disabled = False
		frm1.btnDeselect.disabled = False
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================	
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000011111")
    gMouseClickStatus = "SPC"    
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then             'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
		  ggoSpread.SSSort Col				'Sort in Ascending
		  lgSortkey = 2
       Else
		  ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
		  lgSortkey = 1
	   End If
       Exit Sub
	'--- process check box 
	Else
		frm1.vspdData.Row = Row
		If frm1.vspdData.Col = C_Chk then						
			frm1.btnDeselect.disabled = false
			frm1.btnSelectAll.disabled = false   
		End If
    End If	    
End Sub
'=============================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow
	Dim iRowCnt
	
	with frm1.vspdData		
		Redim arrReturn2(4)
	
		'arrange array
		iRowCnt =0 :		.Col=C_Chk
		For intInsRow =0  To .MaxRows-1	
			.Row = intInsRow +1
			If  (.text)  Then 
				iRowCnt= iRowCnt +1
			End If
		Next
		Redim arrReturn(iRowCnt, frm1.vspdData.MaxCols - 1)		
			
			intInsRow = 0	
			For intRowCnt = 0 To .MaxRows - 1

				.Row = intRowCnt + 1
				.Col = C_Chk
				If .Text  Then	
						
					For intColCnt = 2  To .MaxCols - 1						
						.Col = intColCnt		
						arrReturn(intInsRow, intColCnt - 2) = .Text						
					Next
					intInsRow = intInsRow + 1
				End IF
			Next	
	End With

	arrReturn2(0) = arrReturn
	arrReturn2(1) = Trim(frm1.txtPlant.value)
	arrReturn2(2) = Trim(frm1.txtPlantNm.value	)
	arrReturn2(3) = Trim(frm1.txtPoNo.value)
	arrReturn2(4) = Trim(frm1.txtPoNoSeq.text)	
	
	Self.Returnvalue = arrReturn2
	Self.Close()
End Function

'=============================================================================================================
Function CancelClick()
	Self.Close()
End Function
'=============================================================================================================
Function SelectAll()
	Dim iRows
	Dim iRow	
	
	with frm1.vspdData	
		iRows = .maxRows		
		.Col = C_Chk		
		for iRow=1 to iRows
			.Row = iRow
			.text=True			
		next 		
	end with	
	
	frm1.btnSelectAll.disabled=true
	frm1.btnDeselect.disabled=False
End Function
'=============================================================================================================
Function DeSelect()
	Dim iRows
	Dim iRow
	
	with frm1.vspdData	
		iRows = .maxRows
		.Col = C_Chk		
		for iRow=1 to iRows
			.Row = iRow
			.text=False
		next 		
	end with	
	frm1.btnSelectAll.disabled=False
	frm1.btnDeselect.disabled=True
End Function

'=============================================================================================================
Function FncQuery() '                                                   

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
   End If					
	FncQuery = False								

	Err.Clear											
	If Not chkPoNo then
		Exit Function
	End if
	
	Call ggoOper.ClearField(Document, "2")	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	If frm1.txtAmt.text <= 0 then
		Call DisplayMsgBox("169918", "X", "", "X")		
		frm1.txtAmt.Focus
		Exit Function
	End If
		
	Call InitVariables									
	Call DbQuery()										

	FncQuery = True		
End Function

'=============================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
		
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001	
		
		strVal = strVal & "&txtItem=" & Trim(.txtItem.value)				
		strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
		strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)	'
		strVal = strVal & "&txtPoNoSeq=" & Trim(.txtPoNoSeq.text)	'
		strVal = strVal & "&txtSoDt=" & Trim(.txtSoDt.text) 
		strVal = strVal & "&txtUnit="	& Trim(.txtUnit.value)
		strVal = strVal & "&txtAmt="	& Trim(.txtAmt.text)
		strVal = strVal & "&txtGubun="	& Trim(.cboGubun.value)
		strVal = strVal & "&txtSoldToParty="	& Trim(.txtHSoldToParty.value)
		strVal = strVal & "&txtCur="	& Trim(.txtHCurrency.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2    		
	
        Call RunMyBizASP(MyBizASP, strVal)										
    End With    
    DbQuery = True
End Function

'=============================================================================================================
Function DbQueryOk()										

	If frm1.vspdData.MaxRows < 0 then
		frm1.btnDeselect.disabled = True
		frm1.btnSelectAll.disabled = True
	Else
		frm1.btnSelectAll.disabled = False
	End If
		
	frm1.vspdData.Focus
													
End Function
'=============================================================================================================
Function chkPoNo()
	chkPoNo=true
	With frm1
	'발주번호 또는 발주 순번이 입력되었을 경우 필수 입력사항이된다.	
		If Trim(.txtPoNo.value) <>"" and Trim(.txtPoNoSeq.text)=""  then
			Call DisplayMsgBox("17A002", "X", "발주순번", "X")			
			.txtPoNoSeq.focus	
			chkPoNo=false	
			Exit Function
		ElseIf Trim(.txtPoNo.value) ="" and Trim(.txtPoNoSeq.text)<>""  then
			Call DisplayMsgBox("17A002", "X", "발주번호", "X")
			.txtPoNo.focus			
			chkPoNo=false	
			Exit Function 					
		End If
	End With
End Function 
'=============================================================================================================

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>	
							<TD CLASS=TD5>공장</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 tag="12XXXU" class=required  STYLE="text-transform:uppercase" ALT="공장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" OnClick="vbscript:OpenPlant">&nbsp;
								<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" class = protected readonly = true></TD>								
							</TD>
							<TD CLASS=TD5>발주번호</TD>
							<TD CLASS=TD6>		
								<TABLE >
									<TR>
										<TD WIDTH=30%>						
										<INPUT TYPE=TEXT NAME="txtPoNo" SIZE=10 MAXLENGTH=18 tag="11XXXU"   ALT="발주번호" onChange="vbscript:fncLookUp"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrder" align=top TYPE="BUTTON" OnClick="vbscript:OpenPoNo">&nbsp;
										</TD><TD><script language =javascript src='./js/s3112ra20_fpDoubleSingle1_txtPoNoSeq.js'></script>
										</TD>
									</TR>
								</TABLE>										<!--<INPUT TYPE=TEXT NAME="txtPoNoSeq" SIZE=5 MAXLENGTH=3 tag="11Xxu"   ALT="발주순번" onChange="vbscript:fncLookUp">-->
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5>모품목</TD>
							<TD CLASS=TD6 colspan=3>
								<INPUT TYPE=TEXT NAME="txtItem" SIZE=18 MAXLENGTH=18 tag="12XXXU" class=required ALT="모품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenItem">												
								<INPUT TYPE=TEXT NAME="txtITemNm" SIZE=25 MAXLENGTH=50 TAG="14" ALT="모품목명" class = protected readonly = true>
								<INPUT TYPE=TEXT NAME="txtITemSpec" SIZE=25 MAXLENGTH=50 TAG="14" ALT="규격" class = protected readonly = true>								
							</TD>
						</TR>	
						<TR>	
							<TD CLASS=TD5>기준일</TD>
							<TD CLASS=TD6>									
								<script language =javascript src='./js/s3112ra20_fpDateTime1_txtSODt.js'></script>				
														
							</TD>
							<TD CLASS=TD5>수량</TD>
							<TD CLASS=TD6>
								<TABLE WIDTH=100%>
									<TR>
										<TD WIDTH=50%>
											<script language =javascript src='./js/s3112ra20_fpDoubleSingle1_txtAmt.js'></script>
										</TD>
										<TD WIDTH=50%>&nbsp;<INPUT 
											TYPE=TEXT NAME="txtUnit" SIZE=3 MAXLENGTH=2 TAG="14XXXU" ALT="기준단위" class = protected readonly = true>
										</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>	
						<TR>	
							<TD CLASS=TD5>지급구분</TD>
							<TD CLASS=TD6 >
								<select name="cboGubun" class="cboSmall" tag="11" alt="지급구분"></select>															
							</TD>
							<TD CLASS=TD5>&nbsp;</TD>
							<TD CLASS=TD6 ></td>
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
							<script language =javascript src='./js/s3112ra20_I297345298_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>	
					<TR>					
					  <TD>&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						<!--<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>-->
					  </td>
					  <td align=left valign=top>
					  	<BUTTON NAME="btnSelectAll" CLASS="CLSSBTN" alt="일괄선택" onClick="SelectAll()">일괄선택</BUTTON>&nbsp;
						<BUTTON NAME="btnDeselect" CLASS="CLSSBTN" alt="일괄취소" onClick ="DeSelect()">일괄취소</BUTTON>
					</td>									
					<TD ALIGN=RIGHT> 
						<IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100%  HEIGHT=<%=BizSize%> SCROLLING=No noresize  FRAMEBORDER=0  framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
	
	
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<!--<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPoNoSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSoDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHAmt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHUnit" TAG="24">
<INPUT TYPE=HIDDEN NAME="cboHGubun" TAG="24">-->
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
