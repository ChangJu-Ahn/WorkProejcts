<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ra1
'*  4. Program Name         : ������ǥ��ȣ PopUp
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/12/09
'*  8. Modified date(Last)  : 2003/10/31
'*  9. Modifier (First)     : Ahn Hye Jin
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"  SRC = "Acctctrl.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'**********************************************************************************************************
Const BIZ_PGM_ID        = "a5101rb1.asp"
Const C_SHEETMAXROWS    = 17                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey			= 1

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgAuthorityFlag
Dim lgIsOpenPop                                             '��: Popup status     
Dim lgMark
Dim IsOpenPop                                                  '��: ��ũ                                 
Dim lsPoNo 
Dim arrReturn
Dim arrParent
Dim arrParam					

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)


top.document.title = "������ǥ�˾�"

'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0)
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgAuthorityFlag = arrParam(4)                          '���Ѱ��� �߰� 
    
	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd	= arrParam(5)
		lgInternalCd	= arrParam(6)
		lgSubInternalCd	= arrParam(7)
		lgAuthUsrID		= arrParam(8)    
	End If
	    
	Self.Returnvalue = arrReturn
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	Dim EndDate

	If "" & Trim(arrParam(0)) <> "" Then
		frm1.txtfrtempgldt.Text	= arrParam(0)
		frm1.txttotempgldt.Text	= arrParam(1)
	Else
	    EndDate = UniConvDateAToB("<%=GetSvrDate%>" ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)

		frm1.txtfrtempgldt.Text	= EndDate
		frm1.txttotempgldt.Text	= EndDate
	End If 
	
	frm1.txtDrLocAmtFr.Text	= ""
	frm1.txtDrLocAmtTo.Text	= ""
End Sub

'========================================================================================================= 
Function OpenPopUp(Byval iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'���Ѱ��� �߰�   							  
	
	If IsOpenPop = True Then Exit Function

	Select Case iwhere	
		Case 2
			arrParam(0) = "��ǥ�Է°���˾�"
			arrParam(1) = "B_MINOR"
			arrParam(2) = UCase(Trim(frm1.txtInputType.Value))
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1001", "''", "S") & " "
			arrParam(5) = "��ǥ�Է°���ڵ�"

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"
    
			arrHeader(0) = "��ǥ�Է°���ڵ�"
			arrHeader(1) = "��ǥ�Է°�θ�"
		Case 3
			arrParam(0) = "����� �˾�"						' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 							' TABLE ��Ī 
			arrParam(2) = UCase(Trim(frm1.txtBizArea.Value))								' Code Condition
			arrParam(3) = ""									' Name Cindition
			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			arrParam(5) = "������ڵ�"							' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"							' Field��(0)
			arrField(1) = "BIZ_AREA_NM"							' Field��(1)
    
			arrHeader(0) = "������ڵ�"				' Header��(0)
			arrHeader(1) = "������"				' Header��(1)
		Case 4
			arrParam(0) = "�ۼ��� �˾�"						' �˾� ��Ī 
			arrParam(1) = "Z_USR_MAST_REC" 							' TABLE ��Ī 
			arrParam(2) = UCase(Trim(frm1.txtUsrNm.Value))								' Code Condition
			arrParam(3) = ""									' Name Cindition
			' ���Ѱ��� �߰� 
			If lgAuthUsrID <>  "" Then
				arrParam(4) = " USR_ID=" & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If
			arrParam(5) = "�ۼ����ڵ�"							' �����ʵ��� �� ��Ī 

			arrField(0) = "USR_ID"							' Field��(0)
			arrField(1) = "USR_NM"							' Field��(1)
    
			arrHeader(0) = "�ۼ����ڵ�"				' Header��(0)
			arrHeader(1) = "�ۼ��ڸ�"				' Header��(1)				
	End Select
	
	IsOpenPop = True
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet,iwhere)
	End If	
End Function

'========================================================================================================= 
Function SetPopUp(ByRef arrRet, Byval iwhere)
	With frm1
		Select Case iwhere
			Case 1
				.txtDeptCd.focus
				.txtDeptCd.value = UCase(Trim(arrRet(0)))
				.txtDeptNm.value = arrRet(1)
			Case 2
				.txtInputType.focus
				.txtInputType.value = UCase(Trim(arrRet(0)))
				.txtInputTypeNm.value = arrRet(1)							
			Case 3		' Biz area
				.txtBizArea.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value		= arrRet(1)
				
				.txtBizArea.focus
			Case 4		' Usr Id
				.txtUsrId.value		= UCase(Trim(arrRet(0)))
				.txtUsrNm.value		= arrRet(1)
				
				.txtUsrId.focus									
		End Select
	End With
End Function

'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1  
				.txtDeptCd.focus
			Case 2 
				.txtInputType.focus
		End Select    
	End With
End Function

'========================================================================================================= 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)
	Dim arrStrRet

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	arrStrRet =  AutorityMakeSql("DEPT",PopupParent.gChangeOrgId, "","","","")	'���Ѱ��� �߰�   							  

	arrParam(0) = frm1.txtfrtempgldt.text
   	arrParam(1) = frm1.txttotempgldt.Text
	'arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' �������� ���� Condition  

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(PopupParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'========================================================================================================= 
Function SetDept(Byref arrRet)
	frm1.txtDeptCd.focus
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)		
End Function

'========================================================================================================= 
Sub InitComboBox()	
	Err.clear
	
'-- eWare Inf Begin 
	If	Trim(PopupParent.gEware) = "" Then
		Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
	Else
		Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("AI001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
	End If
'-- eWare Inf End 		
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()		
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
End Sub

'========================================================================================================	
Function OKClick()
	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(1)
		
		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		arrReturn(0)	  = frm1.vspdData.Text
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================================================================================
Function MousePointer(pstr1)
    Select Case UCase(pstr1)
        Case "PON"
	  	  window.document.search.style.cursor = "wait"
        Case "POFF"
	  	  window.document.search.style.cursor = ""
    End Select
End Function

'==========================================================================================================
Sub InitSpreadSheet()    
	frm1.vspdData.OperationMode = 3
	Call SetZAdoSpreadSheet("A5101RA1", "S", "A", "V20051010", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock()          
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1    
		.vspdData.ReDraw = False
         ggoSpread.SpreadLockWithOddEvenRowColor()		 
		.vspdData.ReDraw = True
    End With
End Sub

'===========================================================================
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

'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal		
	Call InitSpreadSheet()
	Call InitComboBox()

End Sub

'==========================================================================================
Sub txtfrtempgldt_DblClick(Button)
	If Button = 1 Then
		frm1.txtfrtempgldt.Action = 7
        Call SetFocusToDocument("P")
        frm1.txtfrtempgldt.focus    
	End If
End Sub

'==========================================================================================
Sub txttotempgldt_DblClick(Button)
	If Button = 1 Then
		frm1.txttotempgldt.Action = 7
        Call SetFocusToDocument("P")
        frm1.txttotempgldt.focus    		
	End If
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
   End If
End Sub

'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
End Sub

'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'==========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'==========================================================================================
Sub txtfrtempgldt_Keypress(KeyAscii)
    On Error Resume Next      

    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call fncQuery()
    End If
End Sub

'==========================================================================================
Sub txttotempgldt_Keypress(KeyAscii)
    On Error Resume Next
    
    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call fncQuery()
    End If
End Sub

'==========================================================================================
Sub txtFrTempGlNo_OnKeyPress()		
	If window.event.keycode = 39 Then	'Single quotation mark �ԷºҰ� 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
Sub txtFrTempGlNo_onpaste()	
	Dim iStrTempGlNo

	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)
End Sub

'==========================================================================================
Sub txtToTempGlNo_OnKeyPress()	
	If window.event.keycode = 39 Then	'Single quotation mark �ԷºҰ� 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
Sub txtToTempGlNo_onpaste()	
	Dim iStrTempGlNo 	

	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)		
End Sub

'==========================================================================================
Sub txtDrLocAmtFr_Keypress(KeyAscii)
    On Error Resume Next

    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call fncQuery()
    End If
End Sub

'==========================================================================================
Sub txtDrLocAmtTo_Keypress(KeyAscii)
    On Error Resume Next
    
    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call fncQuery()
    End If
End Sub

'==========================================================================================
Function CompareTempGlNoByDB(ByVal FromNo , ByVal ToNo)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareTempGlNoByDB = False

    If FromNo.value <> "" And ToNo.value <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromNo.value), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToNo.value), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromNo.value), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToNo.value), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareTempGlNoByDB = True
End Function

'==========================================================================================
Function CompareTempGlAmtByDB(ByVal FromAmt , ByVal ToAmt)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareTempGlAmtByDB = False

    If FromAmt.text <> "" And ToAmt.text <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromAmt.text), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToAmt.text), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromAmt.text), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToAmt.text), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareTempGlAmtByDB = True
End Function

'========================================================================================================= 
Function FncQuery() 
	Dim IntRetCD
	
    FncQuery = False                                                        '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing

    If Not chkField(Document, "1") Then										'��: This function check indispensable field
		Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFrTempGlDt.text,frm1.txtToTempGlDt.text,frm1.txtFrTempGlDt.Alt,frm1.txtToTempGlDt.Alt, _
                        "970025",frm1.txtFrTempGlDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then		
		Exit Function
    End If
    
    If CompareTempGlNoByDB(frm1.txtfrtempglNo,frm1.txttotempglNo) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtfrtempglNo.Alt, frm1.txttotempglNo.Alt)
        frm1.txtfrtempglNo.focus
		Exit Function
	End If		

    If CompareTempGlAmtByDB(frm1.txtDrLocAmtFr,frm1.txtDrLocAmtTo) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtDrLocAmtFr.Alt, frm1.txtDrLocAmtTo.Alt)
        frm1.txtDrLocAmtFr.focus
		Exit Function
	End If		

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData()
    
    Call InitVariables 														'��: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then													'��: Query db data
    	Exit Function
    End If															

    FncQuery = True		
End Function

'========================================================================================
Function FncPrint() 
    Call PopupParent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call PopupParent.FncExport(PopupParent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
    Call PopupParent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False

    Err.Clear																	'��: Protect system from crashing

	Call LayerShowHide(1)

    With frm1
		strVal = BIZ_PGM_ID & "?txtfrtempgldt=" & Trim(.txtfrtempgldt.Text)
		strVal = strVal & "&txttotempgldt=" & Trim(.txttotempgldt.Text)
		strVal = strVal & "&txtfrtempglno=" & Trim(.txtfrtempglNo.value)
		strVal = strVal & "&txttotempglno=" & Trim(.txttotempglNo.value)
		strVal = strVal & "&txtdeptcd=" & Trim(.txtdeptcd.value)
		strVal = strVal & "&txtrefno=" & UCase(Trim(.txtRefNo.value))
		strVal = strVal & "&txtdesc=" & Trim(.txtDesc.value)		
'		strVal = strVal & "&txtConfFg=" & Trim(.cboConfFg.value)
		strVal = strVal & "&txtInputType=" & Trim(.txtInputType.value)
		strVal = strVal & "&txtDrLocAmtFr=" & .txtDrLocAmtFr.text
		strVal = strVal & "&txtDrLocAmtTo=" & .txtDrLocAmtTo.text
		strVal = strVal & "&txtBizArea=" & .txtBizArea.value
		strVal = strVal & "&txtUsrId=" & .txtUsrId.value

        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")         
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)       

		'-- eWare Inf Begin 
		If Trim(PopupParent.gEware)  = "" Then
			strVal = strVal & "&txtConfFg=" & Trim(.cboConfFg.value)
		Else
			strVal = strVal & "&txtConfFg=" & Trim(.cboAppFg.value)
		End If		
		'-- eWare Inf End

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

        Call RunMyBizASP(MyBizASP, strVal)										
    End With

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()															'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = True														'Indicates that no value changed
'    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field

	If frm1.vspdData.MaxRows > 0  Then
		frm1.vspdData.focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtfrtempgldt CLASSID=<%=gCLSIDFPDT%> ALT="��������" tag="12"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
											 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txttotempgldt CLASSID=<%=gCLSIDFPDT%> ALT="��������" tag="12"></OBJECT>');</SCRIPT></TD>
						<TD CLASS=TD5 NOWRAP>���ǹ�ȣ</TD>				
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtfrtempglNo" SIZE=10 MAXLENGTH=18 tag="1XXXXU" ALT="���۰��ǹ�ȣ">&nbsp;~&nbsp;
											 <INPUT TYPE="Text" NAME="txttotempglNo" SIZE=10 MAXLENGTH=18 tag="1XXXXU" ALT="������ǹ�ȣ"></TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
											 <INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=18 tag ="14XXXU"></TD>
<!--'-- eWare Inf Begin -->
<%	If Trim(gEware)  = "" Then	%>	  								
						<TD CLASS=TD5 NOWRAP>���λ���</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboConfFg" tag="21" STYLE="WIDTH:82px:" ALT="���λ���"><OPTION VALUE="" selected></OPTION></SELECT></TD>
<%	Else %>
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAppFg" tag="21" STYLE="WIDTH:82px:" ALT="�������"><OPTION VALUE="" selected></OPTION></SELECT></TD>
<%	End If %>								
<!-- '--eWare Inf End -->
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>������ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="������ȣ" MAXLENGTH="30" SIZE="20" tag="11XXXU" ></TD></TD>				
						<TD CLASS=TD5 NOWRAP>���</TD>
						<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="25" tag="11XXXX" ></TD>
					</TR>							
					<TR>				
						<TD CLASS=TD5 NOWRAP>��ǥ�Է°��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtInputType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="��ǥ�Է°���ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('2')"> <INPUT TYPE="Text" NAME="txtInputTypeNm" SIZE=18 tag="14X" ALT="��ǥ�Է°�θ�"></TD>								
						<TD CLASS=TD5 NOWRAP>��ǥ�ݾ�</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDrLocAmtFr" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="������ǥ�ݾ�" id=OBJECT1></OBJECT>');</SCRIPT> ~ 
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDrLocAmtTo" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="������ǥ�ݾ�" id=OBJECT2></OBJECT>');</SCRIPT></TD>				
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�����</TD>										
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea"   ALT="�����"   Size="10" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp('3')">
											 <INPUT NAME="txtBizAreaNm" ALT="������" Size="18" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
						<TD CLASS=TD5 NOWRAP>�ۼ���</TD>										
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtUsrId"   ALT="�ۼ���"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp('4')">
											 <INPUT NAME="txtUsrNm" ALT="�ۼ��ڸ�" Size="18" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
					</TR>									 
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>

	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"></OBJECT>');</SCRIPT>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
