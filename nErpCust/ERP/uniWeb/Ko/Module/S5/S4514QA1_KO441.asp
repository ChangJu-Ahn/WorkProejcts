<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
                                            
Dim lgIsOpenPop                                                                                 
Dim IscookieSplit 
Dim lgSaveRow                                           
Dim StartDate
Dim EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

Const BIZ_PGM_ID 		= "s4514qb1_ko441.asp"                     
Const C_MaxKey          = 27					             

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
<%'=========================================================================================================%>
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
	lgIntFlgMode = Parent.OPMD_CMODE
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	'frm1.txtFrDt.Text	= StartDate
	'frm1.txtToDt.Text	= EndDate
	' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
	' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub
'======================================================================================
' Function Name : InitComboBox()
' Function Desc : Initialize ComboBox
'========================================================================================
Sub InitComboBox()
End Sub
'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA")%>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
<%'========================================================================================%>
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S4514QA1_KO441","S","A","V20080513", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End If
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���				
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function
'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"					
	arrParam(1) = "B_PLANT"      			
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		
	arrParam(4) = ""								
	arrParam(5) = "����"						
	
    arrField(0) = "PLANT_CD"					
    arrField(1) = "PLANT_NM"					
    
    arrHeader(0) = "����"					
    arrHeader(1) = "�����"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

 '------------------------------------------  OpenSLCode()  -------------------------------------------------
' Name : OpenSLCode()
' Description : Storage Location Display PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenSLCode()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)
 
 If Trim(frm1.txtPlantCd.Value) = "" then 
  Call DisplayMsgBox("169901","X", "X", "X")    '���������� �ʿ��մϴ� 
  frm1.txtPlantCd.focus
  Exit Function
 End if

 '-----------------------
 'Check Plant CODE  '�����ڵ尡 �ִ� �� üũ 
 '-----------------------
 If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("125000","X","X","X")
  frm1.txtPlantNm.value = ""
  frm1.txtPlantCd.focus
  Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtPlantNm.value = lgF0(0)

 If lgIsOpenPop = True Then Exit Function

 lgIsOpenPop = True

 arrParam(0) = "â����ȸ�˾�"   
 arrParam(1) = "B_STORAGE_LOCATION"  
 arrParam(2) = Trim(frm1.txtSL_Cd.value)  
 arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")  ' Where Condition    
 arrParam(5) = "â��"    ' TextBox ��Ī 
 
 arrField(0) = "SL_CD"     
 arrField(1) = "SL_NM"     
    
 arrHeader(0) = "â��"   
 arrHeader(1) = "â���"    
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 lgIsOpenPop = False

 If arrRet(0) = "" Then
	frm1.txtSL_Cd.focus 
	Exit Function
 Else
	frm1.txtSL_Cd.value = arrRet(0) 
	frm1.txtSL_Nm.value	= arrRet(1)   
	frm1.txtSL_Cd.focus  
 End If 
 
End Function


'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderBy("A")
End Sub

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


'------------------------------------  OpenGroupPopup()  ----------------------------------------------
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'====================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						

	If Kubun = 1 Then								
	ElseIf Kubun = 0 Then							
	ElseIf Kubun = 2 then
	End IF
		
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
    ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")										
    Call InitComboBox()
    Call CookiePage(0)
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

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

'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================%>
Sub vspdData_Click(ByVal Col, ByVal Row)

	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	Call SetPopupMenuItemInf("00000000001")
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
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
	End If    		

End Sub
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End If
    
End Sub

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
    FncQuery = False                                            
    Err.Clear                                                   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

'	with frm1
'		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
'			Call DisplayMsgBox("17a003", "X","�����", "X")	
'			Exit Function
'		End if   
'	End with
	
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

	Set gActiveElement = document.activeElement
    FncQuery = True													
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                            
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
	Set gActiveElement = document.activeElement
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim strClsFlg
	
    DbQuery = False
    
    Err.Clear                                                       
	If LayerShowHide(1) = False Then Exit Function
    
    With frm1

		If lgIntFlgMode = Parent.OPMD_UMODE Then	
			strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.hdnPlantCd.value)
			strVal = strVal	& "&rdoQty="	  & Trim(.hdnQty.value)		
			strVal = strVal	& "&txtItemCd="	  & Trim(.hdnItemCd.value)		
			strVal = strVal	& "&txtSL_Cd=" & Trim(.hdnSL_Cd.value)
			strVal = strVal & "&lgPageNo="    & lgPageNo         
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Else
			strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
  		if frm1.rdoQty1.Checked=True Then
  			strVal = strVal & "&rdoQty="
  		Elseif frm1.rdoQty2.Checked=True Then
  			strVal = strVal & "&rdoQty=Y" 
  		Else
	  		strVal = strVal & "&rdoQty=N" 
  		End If
  		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)    	
			strVal = strVal	& "&txtSL_Cd=" & Trim(.txtSL_Cd.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo         
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		End If

        Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								
    
    
			'strVal = strVal	& "&txtFrDt="	  & Trim(.hdnFrDt.value)
			'strVal = strVal	& "&txtToDt="	  & Trim(.hdnToDt.value)
  		'strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)
  		'strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function

'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setretflg()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function Setretflg()
    Setretflg = False
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iretflg
    Dim iPlsFlg
	
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = GetKeyPos("A",1)
    Err.Clear

	Call CommonQueryRs(" ret_flg ", " m_pur_ord_hdr ", " po_no = " & FilterVar(Trim(frm1.vspdData.Text), " " , "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


    iretflg = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
		Err.Clear 
		Exit Function
	End If

    if Trim(lgF0) <> "" then
        frm1.hdnretflg.value = UCase(Trim(iretflg(0)))   
        if UCase(Trim(iretflg(0))) = "Y" then
            Setretflg = False
            Exit Function 
        end if
    end if 
    Setretflg = True
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������Ȳ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemCd" SIZE=18 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemNm" SIZE=34 tag="14"></TD>
								</TR>					   
								<TR>						   
									<TD CLASS="TD5" NOWRAP>â��</TD>
									<TD CLASS="TD6" NOWRAP >
										<input NAME="txtSL_Cd" TYPE="Text" CLASS=required STYLE="Text-Transform: uppercase" MAXLENGTH="7" tag="11XXXU" ALT = "â��" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSLCode()">&nbsp;<input NAME="txtSL_Nm" TYPE="Text" CLASS=protected readonly=true TABINDEX="-1" MAXLENGTH="40" SIZE=25 tag="14N"></TD>    
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoQty" id="rdoQty1" value="" tag = "11" checked>
											<label for="rdoQty1">��ü</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoQty" id="rdoQty2" value="Y" tag = "11">
											<label for="rdoQty2">��������>0</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoQty" id="rdoQty3" value="N" tag = "11">
											<label for="rdoQty3">��������<0</label>&nbsp;									
<!--
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="�����"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtToDt CLASSID=<%=gCLSIDFPDT%> ALT="�����" tag="11X1"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>-->
							         </TD>
	                            </TR>	
	                            <TR>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSL_Cd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnQty" tag="24">


<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
