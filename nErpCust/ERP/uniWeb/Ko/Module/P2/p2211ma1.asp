<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2211ma1.asp
'*  4. Program Name         : MPS���������ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'**********************************************************************************************-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p2211mb1.asp"

Dim C_ItemCd 
Dim C_ItemNm 
Dim C_Spec
Dim C_TrackingNo 
Dim C_StartDt 
Dim C_PlanQty 
Dim C_Unit 
Dim C_ItemGroupCd
Dim C_ItemGroupNm

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,parent.gDateFormat)

'==========================================================================================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'==========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

Dim IsOpenPop						' Popup

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCd        =  1
	C_ItemNm        =  2
	C_Spec			=  3
	C_TrackingNo    =  4
	C_StartDt		=  5
	C_PlanQty		=  6
	C_Unit			=  7
	C_ItemGroupCd	=  8
	C_ItemGroupNm	=  9
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""    
    lgLngCurRows = 0
    lgSortKey    = 1

End Sub

'==========================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromPlanDt.text	= StartDate
	frm1.txtToPlanDt.text	= LastDate
End Sub

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
    With frm1
    
    ggoSpread.Source = .vspdData
    ggoSpread.Spreadinit "V20021123",,parent.gAllowDragDropSpread    
    
    .vspdData.Redraw = False
    
    .vspdData.MaxCols = C_ItemGroupNm + 1
    .vspdData.MaxRows = 0
    
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit		C_ItemCd, 		"ǰ��"			, 18
    ggoSpread.SSSetEdit 	C_ItemNm,       "ǰ���"		, 25
    ggoSpread.SSSetEdit 	C_Spec,			"�԰�"			, 25
    ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No."	, 25
    ggoSpread.SSSetDate 	C_StartDt,		"��ȹ��"		, 14, 2, gDateFormat        
    ggoSpread.SSSetFloat	C_PlanQty,		"��ȹ����"		, 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    ggoSpread.SSSetEdit 	C_Unit, 		"����"			, 10
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"ǰ��׷�",		15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"ǰ��׷��",	30
    
    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)
    
    ggoSpread.SSSetSplit2(1)
    .vspdData.Redraw = True
    
    End With
    
    Call SetSpreadLock()

End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec          = iCurColumnPos(3)    
			C_TrackingNo	= iCurColumnPos(4)    
			C_StartDt		= iCurColumnPos(5)
			C_PlanQty		= iCurColumnPos(6)
			C_Unit			= iCurColumnPos(7)
			C_ItemGroupCd	= iCurColumnPos(8)
			C_ItemGroupNm	= iCurColumnPos(9)
			
    End Select    

End Sub

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
End Function


'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = strCode
	arrParam(2) = "12!MO"							' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""
	
    arrField(0) = 1
    arrField(1) = 2
    
    iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet, iPos)
	End If	

End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	Dim iCalledAspName, IntRetCD

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
'	arrParam(3) = UniConvYYYYMMDDToDate(gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
End Function
'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "ǰ��׷�"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function
'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(ByRef arrRet, Byval Pos)
	With frm1
		If Pos = 0 Then
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)
			.txtItemCd.focus
			Set gActiveElement = document.activeElement
		Else
			.vspdData.Col = C_ItemCd
			.vspdData.Text = arrRet(0)	
		End If	
	End With
End Function

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement	
End Function
'=========================================================================================================
Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function
'==========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables
	
	Call SetToolBar("11000000000011")

    If Parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================================
'   Event Name : txtFromPlanDt_DblClick
'   Event Desc :
'=========================================================================================================
Sub txtFromPlanDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromPlanDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromPlanDt.Focus
	End if
End Sub

'========================================================================================================
'   Event Name : txtToPlanDt_DblClick
'   Event Desc :
'========================================================================================================
Sub txtToPlanDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPlanDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToPlanDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtFromPlanDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromPlanDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToPlanDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToPlanDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
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
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       Exit Sub
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)   ': Query ��ư�� disable ��Ŵ.
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()

    Dim IntRetCD 

    FncQuery = False
    Err.Clear

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtFromPlanDt, frm1.txtToPlanDt)  = False Then		
		Exit Function
	End If

    If DbQuery = False Then
		Exit Function
	End If
	
    FncQuery = True
   
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
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
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4		
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & frm1.hTrackingNo.value 
		strVal = strVal & "&txtFromPlanDt=" & Trim(frm1.hFromPlanDt.value)
		strVal = strVal & "&txtToPlanDt=" & Trim(frm1.hToPlanDt.value)	
		strVal = strVal & "&rdoProcType=" & frm1.hProcType.value
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & frm1.txtTrackingNo.value	
		strVal = strVal & "&txtFromPlanDt=" & Trim(frm1.txtFromPlanDt.Text)
		strVal = strVal & "&txtToPlanDt=" & Trim(frm1.txtToPlanDt.Text)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

    Call SetToolBar("11000000000111")
    
    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
    frm1.vspdData.focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MPS���������ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>��ȹ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p2211ma1_OBJECT1_txtFromPlanDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2211ma1_OBJECT2_txtToPlanDt.js'></script>
									</TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>																						
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>									
								</TR>
								<TR>								
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
			 						<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
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
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p2211ma1_I777864009_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProcType" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromPlanDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToPlanDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

