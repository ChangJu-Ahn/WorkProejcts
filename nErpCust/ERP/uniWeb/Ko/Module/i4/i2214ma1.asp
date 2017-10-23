<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : ��ǰ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'############################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'===========================================================================================================-->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                             

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop                                            

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "i2214mb1.asp"
Const BIZ_PGM_JUMP_ID   = "i2213ma1"				  	      
Dim lsPoNo                                                 
Dim IsOpenPop
'--------------- ������ coding part(��������,End)-------------------------------------------------------------
Dim	C_ItemCode  
Dim	C_ItemName  
Dim	C_ItemSpec
Dim	C_BasicUnit
Dim	C_ReqDate    
Dim	C_RemainQty

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
'   Call GetAdoFiledInf("i2214ma1","S","A")                        '��: spread sheet �ʵ����� query   -----
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no   
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
'#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBaseDt.Text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemAccnt.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_RemainQty + 1         '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
'		.Col = .MaxCols               '��: ������Ʈ�� ��� Hidden Column
'		.ColHidden = True
		.MaxRows = 0
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_ItemCode,   "ǰ��",     18
		ggoSpread.SSSetEdit  C_ItemName,   "ǰ���",   25
		ggoSpread.SSSetEdit  C_ItemSpec,   "�԰�",     20
		ggoSpread.SSSetEdit  C_BasicUnit,  "����",      9
        ggoSpread.SSSetDate  C_ReqDate,    "��ǰ��",   10, 2,             Parent.gDateFormat
		ggoSpread.SSSetFloat C_RemainQty,  "�������", 25, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

		'ggoSpread.MakePairsColumn()
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true
		
	    Call SetSpreadLock 
		ggoSpread.SSSetSplit2(2)

    End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode   = 1
	C_ItemName   = 2
	C_ItemSpec   = 3
	C_BasicUnit  = 4
	C_ReqDate    = 5
	C_RemainQty  = 6
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode   = iCurColumnPos(1)
		C_ItemName   = iCurColumnPos(2)
		C_ItemSpec   = iCurColumnPos(3)
		C_BasicUnit  = iCurColumnPos(4)
		C_ReqDate    = iCurColumnPos(5)
		C_RemainQty  = iCurColumnPos(6)
	End Select

End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
'------------------------------------------ OpenPlant()  --------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "����"		
	arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

'------------------------------------------  OpenItemAccnt()  --------------------------------------------------
'	Name : OpenItemAccnt()
'	Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAccnt()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemAccnt.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ����� �˾�"			' �˾� ��Ī 
	arrParam(1) = "B_MINOR"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemAccnt.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""			' Where Condition
	arrParam(5) = "ǰ�����"			
	
	arrField(0) = "MINOR_CD"					' Field��(0)
	arrField(1) = "MINOR_NM"					' Field��(1)
	
	arrHeader(0) = "ǰ�����"				' Header��(0)
	arrHeader(1) = "ǰ�������"				' Header��(1)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemAcct(arrRet)
		
	End If	
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroup.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroup.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  " 			
	arrParam(5) = "ǰ��׷�"			
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemGroup(arrRet)
	End If	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : OpenPlant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	frm1.txtPlantCd.focus	
End Function

'------------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : OpenItemGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroup.Value    = arrRet(0)		
	frm1.txtItemGroupNm.Value    = arrRet(1)
	frm1.txtItemGroup.focus		
End Function

'------------------------------------------  SetItemAcct()  --------------------------------------------------
'	Name : SetItemAcct()
'	Description : ItemAcct Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byval arrRet)
	frm1.txtItemAccnt.Value	    =arrRet(0)
	frm1.txtItemAccntNm.Value	=arrRet(1)
	frm1.txtItemAccnt.focus
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

	Const CookieSplit = 4877						        'Cookie Split String : CookiePage Function Use

	If Kubun = 1 Then								        'Jump�� ȭ���� �̵��� ��� 

		'Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		With frm1.vspdData	    
		  If .MaxRows = 0 Then
		      Call DisplayMsgBox("169903","X", "X", "X")   'ǰ���ڷᰡ �ʿ��մϴ� 
			  Exit Function
           else
			   .Row = .ActiveRow
			   .Col = C_ItemCode
			    lsPoNo = Trim(.Text )
		   End if	  
    	 End With
		 if lsPoNo = "" then
    		Call DisplayMsgBox("169903","X", "X", "X")     'ǰ���ڷᰡ �ʿ��մϴ� 
    		Exit Function
    	End If

		WriteCookie "PoNo", lsPoNo						   'Jump�� ȭ���� �̵��Ҷ� �ʿ��� Cookie �������� 
		WriteCookie "BaseDt", frm1.txtBaseDt.Text
		WriteCookie "PlantCd", frm1.txtPlantCd.Value

		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then								   'Jump�� ȭ���� �̵��� ������� 
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, gRowSep)
		If arrVal(0) = "" Then Exit Function
		Dim iniSep

'--------------- ������ coding part(�������,Start)---------------------------------------------------
		'�ڵ���ȸ�Ǵ� ���ǰ��� �˻����Ǻ� Name�� Match 
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("ǰ��")
				frm1.txtItemCd.value =  arrVal(iniSep + 1)						
			End Select
		Next
'--------------- ������ coding part(�������,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 														'��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")    

	Call InitVariables						
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")		

	Call CookiePage(0)
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'==========================================================================================
'   Event Name : txtBaseDt
'   Event Desc :
'==========================================================================================
Sub txtBaseDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBaseDt.Action = 7
		Call SetFocusToDocument("M")        
        frm1.txtBaseDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyPress(KeyAscii)
'   Event Desc : 
'=======================================================================================================
Function txtBaseDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
			ggoSpread.SSSort Col					'Sort in Ascending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortKey = 1
        End If  
        Exit Sub  
    End If
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
	lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	'------ Developer Coding part (Start)
	'------ Developer Coding part (End)
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
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	
		If lgStrPrevKey <> "" Then							
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
                Call RestoreToolBar()
				Exit Sub
			End if
		End if
	End if
End Sub


'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
    On Error Resume Next
    
    FncQuery = False                                                        '��: Processing is NG
 
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

	'-----------------------
	'Check Plant CODE		
	'-----------------------
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If

	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	
	'-----------------------
	'Check txtItemAccnt CODE	    'ǰ������ڵ尡 �ִ� �� üũ 
	'-----------------------
	If 	CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & Trim(FilterVar(frm1.txtItemAccnt.Value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
		Call DisplayMsgBox("169952",vbOKOnly, "x", "x")   'ǰ����� �ڵ尡 �ʿ��մϴ�.
		frm1.txtItemAccntNm.value = ""
		frm1.txtItemAccnt.focus 
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
    frm1.txtItemAccntNm.value = lgF0(0)
		
	'-----------------------
	'Check txtItemGroup CODE	    'ǰ��׷��ڵ尡 �ִ� �� üũ 
	'-----------------------
	frm1.txtItemGroupNm.value = ""
	If frm1.txtItemGroup.value <> "" Then
		If 	CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP ", " DEL_FLG = " & FilterVar("N", "''", "S") & "  AND ITEM_GROUP_CD= " & FilterVar(frm1.txtItemGroup.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then

			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemGroupNm.value = lgF0(0)
		End If
	End If
		

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then															'��: Query db data
		Exit Function
	End if

    FncQuery = True		
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
    Call parent.FncExport(Parent.C_MULTI)									'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                        '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()	
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim iStr

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
	
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtBaseDate="    & Trim(.txtBaseDt.Text)		
		strVal = strVal & "&txtItemAccnt="   & Trim(.txtItemAccnt.value)
		strVal = strVal & "&txtItemGroup="   & Trim(.txtItemGroup.value)	
		
'--------------- ������ coding part(�������,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                     
        Call RunMyBizASP(MyBizASP, strVal)										
        
    End With
    
    DbQuery = True


End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
     Call SetToolbar("11000000000111")

 With frm1
    
       if Trim(.txtItemGroup.value) = "" Then       	       
          .txtItemGroupNm.value = ""
       End if
       
       if Trim(.txtItemAccnt.value) = "" Then
          .txtItemAccntNm.value = ""
       End if
 End With
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ��ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
					<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>����</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
							<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
							<TD CLASS="TD5" NOWRAP>������</TD>
							<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i2214ma1_fpDateTime2_txtBaseDt.js'></script> </TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>ǰ�����</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ�����" NAME="txtItemAccnt" SIZE=6 MAXLENGTH=2 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAccnt()">
							<INPUT TYPE=TEXT NAME="txtItemAccntNm" SIZE=25 tag="14"></TD>
							<TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="ǰ��׷�" NAME="txtItemGroup" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">
							<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 tag="14"></TD>
						</TR>						
				</TABLE>
				</RIELDSET>
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
							<script language =javascript src='./js/i2214ma1_OBJECT1_vspdData.js'></script>
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
	
	<TR HEIGHT="20">
		<TD WIDTH=100%>
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:CookiePage(1)">Stock Req. List</a></TD>
				<TD WIDTH=50>&nbsp;</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
