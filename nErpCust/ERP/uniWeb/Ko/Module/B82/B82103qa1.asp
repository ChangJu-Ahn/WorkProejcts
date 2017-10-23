<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         :  
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit												'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID	   = "B82103qb1.asp"                     '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID1 = "B82101ma1"                         '��: Cookie���� ����� ��� 
Const BIZ_PGM_JUMP_ID2 = "B82102ma1"

Dim C_ReqNo         '�Ƿڹ�ȣ 
Dim C_ReqId         '�Ƿ��� 
Dim C_ReqIdNm       '�Ƿ��� 
Dim C_ReqDt         '�Ƿ����� 
Dim C_Status        '���� 
Dim C_itemKind      'ǰ�񱸺� 
Dim C_ItemKindNm    'ǰ�񱸺и� 
Dim C_ItemCd        'ǰ���ڵ� 
Dim C_ItemNm        'ǰ��� 
Dim C_Spec          '�԰� 
Dim C_AcctR         '�������� 
Dim C_AcctT         '������� 
Dim C_AcctP         '���Ű��� 
Dim C_AcctQ         'ǰ������ 
Dim C_TransDt       '�̰����� 
Dim C_DocNo         '�����ȣ 
Dim C_Remark        '��� 

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim StartDate, EndDate

StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)

'--------------- ������ coding part(�������,End)------------------------------------------------------------- 
                 
'==========================================  InitComboBox()  ======================================
'	Name : InitComboBox()
'	Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1001' ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboItemAcct , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False   
End Sub 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtDtFr.Text	= StartDate
	frm1.txtDtTo.Text	= EndDate
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877				
	Dim strTemp, arrVal

	If Kubun = 1 Then
		
		frm1.vspddata.row = frm1.vspddata.activerow
		frm1.vspddata.col = C_ReqNo
		WriteCookie CookieSplit , frm1.vspddata.value & parent.gRowSep 

	ElseIf Kubun = 0 Then

		WriteCookie CookieSplit , ""
		
	End If

End Function


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_Remark + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("7", "13", "2")
		
		ggoSpread.SSSetEdit  C_ReqNo,	    "�Ƿڹ�ȣ",	  15
		ggoSpread.SSSetEdit  C_ReqId,       "�Ƿ���",	  10
		ggoSpread.SSSetEdit  C_ReqIdNm,     "�Ƿ���",	  10
		ggoSpread.SSSetDate  C_ReqDt,	    "�Ƿ�����",   10, 2, Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_Status,  	"Status",         10
		ggoSpread.SSSetEdit  C_ItemKind,  	"ǰ�񱸺�",   10
		ggoSpread.SSSetEdit  C_ItemKindNm,  "ǰ�񱸺�",   10
		ggoSpread.SSSetEdit  C_ItemCd,  	"ǰ���ڵ�",   15
		ggoSpread.SSSetEdit  C_ItemNm,	    "ǰ���",     20
		ggoSpread.SSSetEdit  C_Spec,	    "�԰�",	      20
		ggoSpread.SSSetEdit  C_AcctR,   	"��������",   10
		ggoSpread.SSSetEdit  C_AcctT,    	"�������",   10
		ggoSpread.SSSetEdit  C_AcctP,	    "���Ű���",   10
		ggoSpread.SSSetEdit  C_AcctQ,		"ǰ������",   10
		ggoSpread.SSSetDate  C_TransDt,     "�̰�����",   10, 2, Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_DocNo,	    "�����ȣ",	  15
 		ggoSpread.SSSetEdit  C_Remark,	    "���",	      50
 		
 		Call ggoSpread.SSSetColHidden(C_ReqId,    C_ReqId, True)
 		Call ggoSpread.SSSetColHidden(C_ItemKind, C_ItemKind, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    ggoSpread.SSSetSplit2(2)  
		
		.ReDraw = true
		
    End With
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

	C_ReqNo      = 1    '�Ƿڹ�ȣ 
	C_ReqId      = 2    '�Ƿ��� 
    C_ReqIdNm    = 3    '�Ƿ��� 
    C_ReqDt      = 4    '�Ƿ����� 
    C_Status     = 5    '����    
    C_itemKind   = 6    'ǰ�񱸺� 
    C_itemKindNm = 7    'ǰ�񱸺и� 
    C_ItemCd     = 8    'ǰ���ڵ� 
    C_ItemNm     = 9    'ǰ��� 
    C_Spec       = 10   '�԰� 
    C_AcctR      = 11   '�������� 
    C_AcctT      = 12   '������� 
    C_AcctP      = 13   '���Ű��� 
    C_AcctQ      = 14   'ǰ������ 
    C_TransDt    = 15   '�̰����� 
    C_DocNo      = 16   '�����ȣ 
    C_Remark     = 17   '��� 
    
End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
		C_ReqNo		= iCurColumnPos(1)
	    C_ReqId     = iCurColumnPos(2)
	    C_ReqIdNm   = iCurColumnPos(3)
		C_ReqDt		= iCurColumnPos(4)
		C_Status	= iCurColumnPos(5)
		C_ItemKind	= iCurColumnPos(6)
		C_ItemKindNm= iCurColumnPos(7)
		C_ItemCd	= iCurColumnPos(8)
		C_ItemNm	= iCurColumnPos(9)
		C_Spec		= iCurColumnPos(10)
		C_AcctR		= iCurColumnPos(11)
		C_AcctT		= iCurColumnPos(12)
		C_AcctP		= iCurColumnPos(13)
		C_AcctQ		= iCurColumnPos(14)
		C_TransDt	= iCurColumnPos(15)
		C_DocNo     = iCurColumnPos(16)
		C_Remark	= iCurColumnPos(17)
		
 	End Select
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	
	Call InitVariables                                                      '��: Initializes local global variables                 														
	Call SetDefaultVal	
	Call InitComboBox()
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	
	frm1.txtItem_Kind.focus()
'--------------- ������ coding part(�������,Start)----------------------------------------------------

'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
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
 	End If
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDtFr.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtDtTo.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function


'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery()

    Dim IntRetCD
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")		    
		If IntRetCD = vbNo Then Exit Function
    End If
        
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function						'��: This function check indispensable field
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtTo) = False Then
   		frm1.txtDtFr.focus 
		Set gActiveElement = document.activeElement
		Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

	If Valid_Check() = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If
								                                            '��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then	Exit Function

    FncQuery = True															'��: Processing is OK
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
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
	Dim strVal
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
    
        If .rdoStatus1.checked = True Then
           .txtrdoStatus.value = "1" 
        ElseIf .rdoStatus2.checked = True Then
           .txtrdoStatus.value = "2" 
        ElseIf .rdoStatus3.checked = True Then
           .txtrdoStatus.value = "3" 
        End If
        
		'--------------- ������ coding part(�������,Start)----------------------------------------------------------------------------
		strVal = BIZ_PGM_ID & "?txtDtFr="		& Trim(.txtDtFr.Text) & _
							  "&txtDtTo="		& Trim(.txtDtTo.Text) & _
							  "&txtrdoStatus="	& Trim(.txtrdoStatus.value) & _
							  "&cboItemAcct="	& Trim(.cboItemAcct.value) & _
							  "&txtItem_Kind="	& Trim(.txtItem_Kind.value) & _
							  "&txtreq_user="		& Trim(.txtreq_user.value) & _
							  "&txtItemSpec="	& Trim(.txtItemSpec.value) & _
							  "&txtMaxRows="	& .vspdData.MaxRows & _
							  "&lgStrPrevKey="	& lgStrPrevKey                      '��: Next key tag
							  
		Call RunMyBizASP(MyBizASP, strVal)
		
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'��: ��ư ���� ���� 
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim arrParam(7), arrField(8), arrHeader(8)
        Dim sItemAcct , sItemKind, sItemLvl1, sItemLvl2, sItemLvl3

        If IsOpenPop = True  Then  
           Exit Function
        End If   

        IsOpenPop = True
        Select Case arPopUp
               Case 1 'ǰ�񱸺� 
                                   
                    arrParam(0) = frm1.txtItem_Kind.Alt
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtItem_Kind.value)
                    arrParam(4) = "MAJOR_CD = 'Y1001'"
                    arrParam(5) = frm1.txtItem_Kind.Alt

 
 
                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"
    
                    arrHeader(0) = frm1.txtItem_Kind.Alt
                    arrHeader(1) = frm1.txtItem_Kind_nm.Alt
                    frm1.txtItem_Kind.focus()
               Case 2 '�Ƿ��� 
                                   
                    arrParam(0) = frm1.txtreq_user.Alt
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtreq_user.value)
                    arrParam(4) = "MAJOR_CD = 'Y1006' "
                    arrParam(5) = frm1.txtreq_user.Alt

                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"
    
                    arrHeader(0) = frm1.txtreq_user.Alt
                    arrHeader(1) = frm1.txtreq_user_Nm.Alt                                 
					 frm1.txtreq_user.focus()
               
               Case Else
                    Exit Function
      End Select
        
      arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

      IsOpenPop = False
                
      If arrRet(0) = "" Then
         Exit Function
      Else
         Call SetConPopup(arrRet,arPopUp)
      End If        
        
End Function

'======================================================================================================
Function SetConPopup(Byval arrRet,ByVal arPopUp)

     SetConPopup = False

     Select Case arPopUp
            Case 1 'ǰ�񱸺� 
                 frm1.txtItem_Kind.value   = arrRet(0) 
                 frm1.txtItem_Kind_nm.value = arrRet(1)   
            Case 2 '�Ƿ��� 
                 frm1.txtreq_user.value      = arrRet(0) 
                 frm1.txtreq_user_Nm.value    = arrRet(1) 
            Case 3 'ǰ���ڵ� 
                 frm1.txtItemCd.value     = arrRet(0) 
                 frm1.txtItemNm.value     = arrRet(1)            
     End Select

     SetConPopup = True

End Function

'========================================================================================
' Function Name : Valid_Check
'========================================================================================
Function Valid_Check()

	Valid_Check = False
	
	With frm1
		'-----------------------
		'ǰ�񱸺� Check
		'-----------------------		
		If Trim(.txtItem_Kind.Value) <> "" Then
		   If CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = 'Y1001' AND MINOR_CD = " & FilterVar(.txtItem_Kind.Value,"","S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		      .txtItem_Kind_nm.Value = ""			  
			   Call DisplayMsgBox("970000","X","ǰ�񱸺�","X")
			  .txtItem_Kind.focus 
			  Exit function
			else
			  lgF0 = Split(lgF0, Chr(11))
		    .txtItem_Kind_nm.Value = lgF0(0)
			End If
			
		End If
		
		'-----------------------
		'�Ƿ��� Check
		'-----------------------
		If Trim(.txtreq_user.Value) <> "" Then
			If CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = 'Y1006' AND MINOR_CD = " & FilterVar(.txtreq_user.Value,"","S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			  .txtreq_user_Nm.Value = ""			  
			  Call DisplayMsgBox("970000","X","�Ƿ���","X")
			  .txtreq_user.focus 
			  Exit function
			else
			lgF0 = Split(lgF0, Chr(11))
		    .txtreq_user_nm.Value = lgF0(0)   
			End If
			
		End If
		
	End With
	
	Valid_Check = True

End Function

'========================================================================================
' Function Name : CookiePage
'========================================================================================
Function CookiePage()

	On Error Resume Next

	Const CookieSplit = 4877						
	
	If frm1.vspdData.ActiveRow > 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_ReqNo
		WriteCookie CookieSplit , frm1.vspdData.Text
	Else
		WriteCookie CookieSplit , ""
	End If

End Function

'========================================================================================
' Function Name : JumpChgCheck
'========================================================================================
Function JumpChgCheck(ByVal arJump)

    Call CookiePage()
    
	Select Case arJump
	       Case 1
		        PgmJump(BIZ_PGM_JUMP_ID1)
	       Case 2
		        PgmJump(BIZ_PGM_JUMP_ID2)
	End Select

End Function

'========================================================================================
' Function Name : txtreq_user_OnChange
' Function Desc : 
'========================================================================================
Function txtreq_user_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtreq_user.value = "" Then
        frm1.txtreq_user_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1006' and minor_cd="&filterVar(frm1.txtreq_user.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtreq_user_nm.value=""
        Else
            frm1.txtreq_user_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


'========================================================================================
' Function Name : txtitem_acct_cd_OnChange
' Function Desc : 
'========================================================================================
Function cboItemAcct_OnChange()
    Dim iDx
    Dim IntRetCd
 

    //call txtItem_kind_OnChange()
End Function



'========================================================================================
' Function Name : txtItem_kind_OnChange
'========================================================================================
Function txtItem_kind_OnChange()
    Dim iDx
    Dim IntRetCd
 
	
    If frm1.txtItem_kind.value = "" Then
        frm1.txtItem_kind_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm ","  B_MINOR A, B_CIS_CONFIG B "," major_cd='Y1001' AND A.MINOR_CD = B.ITEM_KIND AND minor_cd="&filterVar(frm1.txtItem_kind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         //IntRetCd =  CommonQueryRs(" minor_nm ","  B_MINOR A, B_CIS_CONFIG B "," major_cd='Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT like "&filtervar(frm1.cboItemAcct.value&"%","''","S")&" and minor_cd="&filterVar(frm1.txtItem_kind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_kind_nm.value=""
        Else
            frm1.txtItem_kind_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></TD>
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
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								  <TD CLASS=TD5 NOWRAP>�Ƿ�����</TD>
								  <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82103qa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
								                       <script language =javascript src='./js/b82103qa1_fpDateTime6_txtDtTo.js'></script>
								</TD>
								   <TD CLASS=TD5 NOWRAP>Status</TD>
								   <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus1" Value="1" CLASS="RADIO" tag="1X"><LABEL FOR="rdoStatus1">��ü</LABEL>
								                        <INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus2" Value="2" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoStatus2">������</LABEL>
								                        <INPUT TYPE="RADIO" NAME="rdoStatus" ID="rdoStatus3" Value="3" CLASS="RADIO" tag="1X"><LABEL FOR="rdoStatus3">�Ϸ�</LABEL></TD>
								</TR>
								<TR>
								   <TD CLASS=TD5 NOWRAP>ǰ�����</TD>
								   <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="11" ALT="ǰ�����"><OPTION VALUE=""></OPTION></SELECT></TD>
								   <TD CLASS=TD5 NOWRAP>ǰ�񱸺�</TD>
								   <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem_Kind" ALT="ǰ�񱸺�" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')">
								                        <INPUT NAME="txtItem_Kind_nm" ALT="ǰ�񱸺и�" TYPE="Text" SiZE=25   tag="14XXXU"></TD>
								</TR>
								<TR>
								  <TD CLASS=TD5 NOWRAP>�Ƿ���</TD>
								  <TD CLASS=TD6 NOWRAP><INPUT NAME="txtreq_user" ALT="�Ƿ���" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('2')">
								                       <INPUT NAME="txtreq_user_Nm" ALT="�Ƿ��ڸ�" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
								  <TD CLASS=TD5 NOWRAP>�԰�</TD>
								  <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemSpec" ALT="�԰�" TYPE="Text" SiZE=40   tag="11XXXU"></TD>
								</TR>  	                
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<script language =javascript src='./js/b82103qa1_I277439552_vspdData.js'></script>
							</TD>	
						</TR>	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=12>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>						
						<TD ALIGN =right>
							 <A HREF = "VBSCRIPT:JumpChgCheck(1)" ONCLICK="VBSCRIPT:CookiePage" >ǰ��ű��Ƿڵ��</A>&nbsp;
							 <A HREF = "VBSCRIPT:JumpChgCheck(2)" ONCLICK="VBSCRIPT:CookiePage" >ǰ��ű��Ƿڽ���</A>
						</TD>
					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtrdoStatus" TAG="24" TABINDEX="-1"></INPUT>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
</BODY>
</HTML>