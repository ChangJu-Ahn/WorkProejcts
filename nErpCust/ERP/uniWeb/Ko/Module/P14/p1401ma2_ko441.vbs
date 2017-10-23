Const BIZ_PGM_QRY_ID	= "p1401mb10_ko441.asp"		

Const C_Sep  = "/"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM ="PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"
Const C_SUBCON  = "SUBCON"

Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/Assembly.gif"
Const C_IMG_SUBCON = "../../../CShared/image/subcon.gif"

Const tvwChild = 4	

Dim lgBlnFlgConChg				'☜: Condition 변경 Flag

Dim IsOpenPop
Dim lgBlnBizLoadMenu
Dim lgSelNode

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False				'⊙: Indicates that no value changed
    lgIntGrpCount = 0					'⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
	lgSelNode = ""
End Sub

'========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=================================================================================================== 
Sub SetDefaultVal()
	frm1.rdoSrchType1.checked = True
	frm1.txtBaseDt.Text = StartDate
	frm1.txtBomNo.value = "1"
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"					' 팝업 명칭 
	arrParam(1) = "B_PLANT"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
   	arrField(0) = "PLANT_CD"						' Field명(0)
   	arrField(1) = "PLANT_NM"						' Field명(1)
    
   	arrHeader(0) = "공장"						' Header명(0)
   	arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenIremCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"	
    arrField(2) = 3								' Field명(1) : "ITEM_ACCT"
    arrField(3) = 8								' Field명(1) : "PHANTOM_FLG"	
    arrField(4) = 5								' Field명(1) : "PROCUR_TYPE"
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
   '---------------------------------------------
	 ' Validation Check Area
	 '--------------------------------------------- 
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	  '---------------------------------------------
	 ' Parameter Setting
	 '--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM팝업"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"							' TABLE 명칭 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
        
    arrHeader(0) = "BOM Type"					' Header명(0)
    arrHeader(1) = "BOM 특성"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBomNo.focus
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup에서 return된 값 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)		
End Function

'==========================================================================================
'   Function Name :LookUpHdr
'   Function Desc :선택한 품목의 Header Data를 읽는다.
'==========================================================================================
Sub LookUpHdr(ByVal txtItemCd,ByVal txtBomNo)
	Dim strVal
	
	Call ggoOper.ClearField(Document, "2")
	Call ggoOper.LockField(Document, "Q") 
	
	LayerShowHide(1)
				    
	'------------------------------
	' Server Logic Call
	'------------------------------
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0002				'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'☜: LookUP 조건 데이타 
	strVal = strVal & "&txtItemCd=" & txtItemCd							'☜: LookUP 조건 데이타 
	strVal = strVal & "&txtBomNo=" & txtBomNo							'☜: LookUP 조건 데이타    

	If Trim(frm1.txtSrchType.value) = "2" Then
		strVal = strVal & "&rdoSrchType=" & 2							'☜: LookUP 조건 데이타    
	ELSE
		strVal = strVal & "&rdoSrchType=" & 4							'☜: LookUP 조건 데이타    
	END IF
			
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
End Sub

Sub LookUpHdrOk()
End Sub

'==========================================================================================
'   Function Name :LookUpDtl
'   Function Desc :선택한 품목의 Item Acct를 읽는다.
'==========================================================================================

Function LookUpDtl(ByVal txtPrntItemCd,ByVal txtPrntBomNo,ByVal intChildItemSeq,ByVal intLevel,ByVal txtChildBomNo,ByVal txtChildItemCd)
	Dim strVal
	
	Call ggoOper.ClearField(Document, "2")
	Call ggoOper.LockField(Document, "Q")
	
	LayerShowHide(1)
		
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0003			'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'☜: LookUP 조건 데이타 
	strVal = strVal & "&intChildItemSeq=" & intChildItemSeq				'☜: LookUP 조건 데이타 
	strVal = strVal & "&intLevel=" & intLevel							'☜: LookUP 조건 데이타 

	If Trim(frm1.txtSrchType.value) = "2" Then
		strVal = strVal & "&txtPrntItemCd=" & txtPrntItemCd				'☜: LookUP 조건 데이타 
		strVal = strVal & "&txtPrntBomNo=" & txtPrntBomNo				'☜: LookUP 조건 데이타    
		strVal = strVal & "&txtChildBomNo=" & txtChildBomNo				'☜: LookUP 조건 데이타 
		strVal = strVal & "&txtChildItemCd=" & txtChildItemCd
		strVal = strval & "&rdoSrchType=" & 2
	Else
		strVal = strVal & "&txtPrntItemCd=" & txtChildItemCd			'☜: LookUP 조건 데이타 
		strVal = strVal & "&txtPrntBomNo=" & txtChildBomNo				'☜: LookUP 조건 데이타    
		strVal = strVal & "&txtChildBomNo=" & txtPrntBomNo				'☜: LookUP 조건 데이타 
		strVal = strVal & "&txtChildItemCd=" & txtPrntItemCd
		strVal = strval & "&rdoSrchType=" & 4
	End If	
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

End Function

Sub LookUpDtlOk()
End Sub

'========================================================================================
' Function Name : InitTreeImage
' Function Desc : 이미지 초기화 
'========================================================================================
Function InitTreeImage()
	Dim NodX, lHwnd
	
	With frm1

		.uniTree1.SetAddImageCount = 4
		.uniTree1.Indentation = "200"									' 줄 간격 
		.uniTree1.AddImage C_IMG_PROD, C_PROD, 0						'⊙: TreeView에 보일 이미지 지정 
		.uniTree1.AddImage C_IMG_MATL, C_MATL, 0
		.uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0				'⊙: TreeView에 보일 이미지 지정 
		.uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
		.uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0

		.uniTree1.OLEDragMode = 0										'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
		.uniTree1.OLEDropMode = 0
	
	End With

End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtItemCd_OnKeyPress()
	frm1.txtItemNm.value = ""
End Sub

'==========================================================================================
'   Event Name : rdoSrchType1_OnClick
'   Event Desc : 정전개 선택시 BOM No Field 필수 입력 
'==========================================================================================
Sub rdoSrchType1_OnClick()
	Call ggoOper.SetReqAttr(frm1.txtBomNo,"N")
End Sub

'==========================================================================================
'   Event Name : rdoSrchType2_OnClick
'   Event Desc : 역전개 선택시 BOM No Field Default
'==========================================================================================
Sub rdoSrchType2_OnClick()
	Call ggoOper.SetReqAttr(frm1.txtBomNo,"N")
End Sub

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click시 Look Up Call
'==========================================================================================
Sub uniTree1_NodeClick(ByVal Node)
    Dim NodX
    
	Dim iPos1
	Dim iPos2
	Dim iPos3
	
	Dim txtPrntBomNo
	Dim txtPrntItemCd
	Dim intChildItemSeq
	Dim intLevel
	Dim txtChildBomNo
	Dim txtChildItemCd
	Dim prntNode

	Err.Clear                                                               '☜: Protect system from crashing
			
	With frm1
	
    Set NodX = .uniTree1.SelectedItem
    
    If Not NodX Is Nothing Then				' 선택된 폴더가 있으면 

		'-------------------------------------
		'If Same Node Clicked, Exit
		'---------------------------------------
			
		If NodX.Key = lgSelNode Then
			Set NodX = Nothing
			Exit Sub
		Else
			lgSelNode = NodX.Key
		End If
		
		'-------------------------------------
		'Hidden Value Init
		'---------------------------------------
		
		Set PrntNode = NodX.Parent
		
		If PrntNode is Nothing Then				' Root일 경우 
			iPos1 = InStr(1,NodX.Key, "|^|^|")										'Parent Bom No
			txtPrntItemCd  = Trim(Mid(NodX.Key,1,iPos1-1))

			'--------------------------------------
			'Child Bom NO Setting
			'--------------------------------------				
			txtPrntBomNo   = Trim(Right(NodX.Key,Len(NodX.Key)-(iPos1+4)))
				
			Call LookUpHdr(txtPrntItemCd ,txtPrntBomNo) 
		Else
			'--------------------------------------
			'Child Item Seq Setting
			'--------------------------------------				
			iPos1 = InStr(1,NodX.Key, "|^|^|")    
			
			intChildItemSeq = Mid(NodX.Key,1,iPos1-1)

			iPos2 = InStr(PrntNode.Text, "    (")   
			txtPrntItemCd = Trim(Left(PrntNode.Text,iPos2-1))
				
			iPos2 = InStr(NodX.Text, "    (")   
			txtChildItemCd = Trim(Left(NodX.Text,iPos2-1))

			iPos2 = InStr(iPos1+5,NodX.Key,"|^|^|")
			txtChildBomNo = Trim(Mid(NodX.Key,iPos1+5,iPos2-(iPos1+5)))

			iPos2 = InStr(iPos1+5,NodX.Key,"|^|^|")
			txtPrntBomNo = Trim(Mid(NodX.Key,iPos1+5,iPos2-(iPos1+5)))
							
			'--------------------------------------
			'Level Setting
			'--------------------------------------				
			iPos3 = InStr(iPos2+5,NodX.Key,"|^|^|")								'Level
			intLevel = Mid(NodX.Key,iPos2+5,iPos3-(iPos2+5))

			'--------------------------------------
			'Prnt Bom NO Setting
			'--------------------------------------				
			iPos1 = InStr(1,PrntNode.Key, "|^|^|")       
			iPos2 = InStr(iPos1+5,PrntNode.Key,"|^|^|")								'Child Item Seq			
				
			If iPos2 <> 0 Then
				txtPrntBomNo = Trim(Mid(PrntNode.Key,iPos1+5,iPos2-(iPos1+5)))
			Else 
				txtPrntBomNo = Trim(Right(PrntNode.Key,Len(PrntNode.Key)-(iPos1+4)))
			End If
	
		    Call LookUpDtl(txtPrntItemCd,txtPrntBomNo,intChildItemSeq,intLevel,txtChildBomNo,txtChildItemCd)
		End IF
	End If
    
    Set NodX = Nothing
    Set PrntNode = Nothing
    
    End With

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	'-----------------------
    'Erase contents area
    '----------------------- 

	frm1.uniTree1.Nodes.Clear		
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
												'⊙: Tree View Content	 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
            
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	'-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then   
		Exit Function           
    End If     																		'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim PrntKey
    Dim strVal
    
    Err.Clear															'☜: Protect system from crashing
    
    DbQuery = False														'⊙: Processing is NG
   
	LayerShowHide(1)
		
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
    
    frm1.txtUpdtUserId.value= parent.gUsrID
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'☜: 조회 조건 데이타 
    strVal = strVal & "&txtBaseDt=" & Trim(frm1.txtBaseDt.Text)	
    
    If frm1.rdoSrchType1.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value 
		frm1.txtSrchType.value = 2
    ElseIf frm1.rdoSrchType2.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value 
		frm1.txtSrchType.value = 4
    End If       
    
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBomNo.value)
    
    Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True														'⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()													'☆: 조회 성공후 실행로직 
	
    Dim NodX
    
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		
		Set NodX = frm1.uniTree1
		NodX.SetFocus
		Set NodX = Nothing
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = parent.OPMD_UMODE											'⊙: Indicates that current mode is Update mode
    
    Call SetToolbar("11000000000111")
    
End Function
