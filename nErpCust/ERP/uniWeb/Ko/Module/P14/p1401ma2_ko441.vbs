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

Dim lgBlnFlgConChg				'��: Condition ���� Flag

Dim IsOpenPop
Dim lgBlnBizLoadMenu
Dim lgSelNode

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False				'��: Indicates that no value changed
    lgIntGrpCount = 0					'��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
	lgSelNode = ""
End Sub

'========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
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

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
   	arrField(0) = "PLANT_CD"						' Field��(0)
   	arrField(1) = "PLANT_NM"						' Field��(1)
    
   	arrHeader(0) = "����"						' Header��(0)
   	arrHeader(1) = "�����"						' Header��(1)
    
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
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
    arrField(2) = 3								' Field��(1) : "ITEM_ACCT"
    arrField(3) = 8								' Field��(1) : "PHANTOM_FLG"	
    arrField(4) = 5								' Field��(1) : "PROCUR_TYPE"
    
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
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "ǰ��", "X")
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

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
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
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)		
End Function

'==========================================================================================
'   Function Name :LookUpHdr
'   Function Desc :������ ǰ���� Header Data�� �д´�.
'==========================================================================================
Sub LookUpHdr(ByVal txtItemCd,ByVal txtBomNo)
	Dim strVal
	
	Call ggoOper.ClearField(Document, "2")
	Call ggoOper.LockField(Document, "Q") 
	
	LayerShowHide(1)
				    
	'------------------------------
	' Server Logic Call
	'------------------------------
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0002				'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & txtItemCd							'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtBomNo=" & txtBomNo							'��: LookUP ���� ����Ÿ    

	If Trim(frm1.txtSrchType.value) = "2" Then
		strVal = strVal & "&rdoSrchType=" & 2							'��: LookUP ���� ����Ÿ    
	ELSE
		strVal = strVal & "&rdoSrchType=" & 4							'��: LookUP ���� ����Ÿ    
	END IF
			
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	
End Sub

Sub LookUpHdrOk()
End Sub

'==========================================================================================
'   Function Name :LookUpDtl
'   Function Desc :������ ǰ���� Item Acct�� �д´�.
'==========================================================================================

Function LookUpDtl(ByVal txtPrntItemCd,ByVal txtPrntBomNo,ByVal intChildItemSeq,ByVal intLevel,ByVal txtChildBomNo,ByVal txtChildItemCd)
	Dim strVal
	
	Call ggoOper.ClearField(Document, "2")
	Call ggoOper.LockField(Document, "Q")
	
	LayerShowHide(1)
		
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0003			'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&intChildItemSeq=" & intChildItemSeq				'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&intLevel=" & intLevel							'��: LookUP ���� ����Ÿ 

	If Trim(frm1.txtSrchType.value) = "2" Then
		strVal = strVal & "&txtPrntItemCd=" & txtPrntItemCd				'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&txtPrntBomNo=" & txtPrntBomNo				'��: LookUP ���� ����Ÿ    
		strVal = strVal & "&txtChildBomNo=" & txtChildBomNo				'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & txtChildItemCd
		strVal = strval & "&rdoSrchType=" & 2
	Else
		strVal = strVal & "&txtPrntItemCd=" & txtChildItemCd			'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&txtPrntBomNo=" & txtChildBomNo				'��: LookUP ���� ����Ÿ    
		strVal = strVal & "&txtChildBomNo=" & txtPrntBomNo				'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&txtChildItemCd=" & txtPrntItemCd
		strVal = strval & "&rdoSrchType=" & 4
	End If	
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Function

Sub LookUpDtlOk()
End Sub

'========================================================================================
' Function Name : InitTreeImage
' Function Desc : �̹��� �ʱ�ȭ 
'========================================================================================
Function InitTreeImage()
	Dim NodX, lHwnd
	
	With frm1

		.uniTree1.SetAddImageCount = 4
		.uniTree1.Indentation = "200"									' �� ���� 
		.uniTree1.AddImage C_IMG_PROD, C_PROD, 0						'��: TreeView�� ���� �̹��� ���� 
		.uniTree1.AddImage C_IMG_MATL, C_MATL, 0
		.uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0				'��: TreeView�� ���� �̹��� ���� 
		.uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
		.uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0

		.uniTree1.OLEDragMode = 0										'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
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
'   Event Desc : ������ ���ý� BOM No Field �ʼ� �Է� 
'==========================================================================================
Sub rdoSrchType1_OnClick()
	Call ggoOper.SetReqAttr(frm1.txtBomNo,"N")
End Sub

'==========================================================================================
'   Event Name : rdoSrchType2_OnClick
'   Event Desc : ������ ���ý� BOM No Field Default
'==========================================================================================
Sub rdoSrchType2_OnClick()
	Call ggoOper.SetReqAttr(frm1.txtBomNo,"N")
End Sub

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click�� Look Up Call
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

	Err.Clear                                                               '��: Protect system from crashing
			
	With frm1
	
    Set NodX = .uniTree1.SelectedItem
    
    If Not NodX Is Nothing Then				' ���õ� ������ ������ 

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
		
		If PrntNode is Nothing Then				' Root�� ��� 
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
        
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

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
												'��: Tree View Content	 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
            
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	'-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then   
		Exit Function           
    End If     																		'��: Query db data
       
    FncQuery = True																'��: Processing is OK
        
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
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
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
    
    Err.Clear															'��: Protect system from crashing
    
    DbQuery = False														'��: Processing is NG
   
	LayerShowHide(1)
		
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
    
    frm1.txtUpdtUserId.value= parent.gUsrID
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBaseDt=" & Trim(frm1.txtBaseDt.Text)	
    
    If frm1.rdoSrchType1.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value 
		frm1.txtSrchType.value = 2
    ElseIf frm1.rdoSrchType2.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value 
		frm1.txtSrchType.value = 4
    End If       
    
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBomNo.value)
    
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True														'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()													'��: ��ȸ ������ ������� 
	
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
    lgIntFlgMode = parent.OPMD_UMODE											'��: Indicates that current mode is Update mode
    
    Call SetToolbar("11000000000111")
    
End Function
