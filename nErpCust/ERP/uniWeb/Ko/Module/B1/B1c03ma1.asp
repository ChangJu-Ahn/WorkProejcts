<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Message���)
'*  3. Program ID           : B1c03ma1.asp
'*  4. Program Name         : B1c03ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/14
'*  7. Modified date(Last)  : 2002/09/10
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	<%'��: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B1c03mb1.asp"												<%'�����Ͻ� ���� ASP�� %>
 
Dim C_Lang          
Dim C_Msg           
Dim C_MsgType       
Dim C_MsgTypeNm     
Dim C_MsgLevel      
Dim C_MsgLevelNm    
Dim C_MsgText       

Dim lgStrPrevLang

<!-- #Include file="../../inc/lgvariables.inc" -->
Sub InitSpreadPosVariables()
    C_Lang = 1
    C_Msg = 2
    C_MsgType = 3
    C_MsgTypeNm = 4
    C_MsgLevel = 5
    C_MsgLevelNm = 6
    C_MsgText = 7
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevLang = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()

    Call initSpreadPosVariables()  
    
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021204",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_MsgText + 1						'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
    .Col = .MaxCols									'��: ������Ʈ�� ��� Hidden Column
    .ColHidden = True
    
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    Call GetSpreadColumnPos("A") 

    ggoSpread.SSSetCombo C_Lang, "���", 10 '1
    ggoSpread.SSSetEdit C_Msg, "�޼��� �ڵ�", 15, , ,6 '2
	ggoSpread.SSSetCombo C_MsgType, "", 10 '3
	ggoSpread.SSSetCombo C_MsgTypeNm, "�޼��� ����", 17 '4
	ggoSpread.SSSetCombo C_MsgLevel, "", 10 '5
	ggoSpread.SSSetCombo C_MsgLevelNm, "�޼��� Level", 15 '6
    ggoSpread.SSSetEdit C_MsgText, "����", 58, , , 1024 '7

    Call ggoSpread.SSSetColHidden(C_MsgType,C_MsgType,True)
    Call ggoSpread.SSSetColHidden(C_MsgLevel,C_MsgLevel,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_Lang, -1, C_Lang
		ggoSpread.SpreadLock C_Msg, -1, C_Msg		
		ggoSpread.SpreadLock C_MsgTypeNm, -1, C_MsgTypeNm
		ggoSpread.SSSetRequired	C_MsgLevelNm, -1, -1
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
		.vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_Lang, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Msg, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_MsgTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_MsgLevelNm, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Lang          = iCurColumnPos(1)
            C_Msg           = iCurColumnPos(2)
            C_MsgType       = iCurColumnPos(3)
            C_MsgTypeNm     = iCurColumnPos(4)
            C_MsgLevel      = iCurColumnPos(5)
            C_MsgLevelNm    = iCurColumnPos(6)
            C_MsgText       = iCurColumnPos(7)
    End Select    
End Sub

Sub InitComboBox()
    Dim strCboData
    Dim strCboData2
    
    
    Call CommonQueryRs(" RTrim(LANG_CD),LANG_NM ", " B_LANGUAGE ", " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	        
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
	ggoSpread.SetCombo strCboData, C_Lang
	'ggoSpread.SetCombo strCboData2, C_CaptionCd
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0007", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.SetCombo strCboData,  C_MsgType
	ggoSpread.SetCombo strCboData2, C_MsgTypeNm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0008", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.SetCombo strCboData,  C_MsgLevel
	ggoSpread.SetCombo strCboData2, C_MsgLevelNm
End Sub

Sub InitComboBox2()
    Call CommonQueryRs(" RTrim(LANG_CD),LANG_NM ", " B_LANGUAGE ", " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboLang, lgF0, lgF1, Chr(11))	
	        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0007", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboMsgType, lgF0, lgF1, Chr(11))
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0008", "''", "S") & "  ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboMsgLevel, lgF0, lgF1, Chr(11))
	
End Sub

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call InitComboBox
    Call InitComboBox2
    Call SetToolbar("1100110100101111")										<%'��ư ���� ���� %>
    
    frm1.cboLang.value = gLang
    frm1.cboLang.focus
    
End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
    
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	
	
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		If Col = C_MsgTypeNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_MsgType
			.TypeComboBoxCurSel = index
		End If
		If Col = C_MsgLevelNm And Row > 0 Then
			.Row = Row
			.Col = Col
			index = .TypeComboBoxCurSel
			
			.Col = C_MsgLevel
			.TypeComboBoxCurSel = index
		End If
	End With
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If CheckRunningBizProcess = True Then					'��: ��ȸ���̸� ���� ��ȸ ���ϵ��� üũ 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'��: ������ üũ %>
    	If lgStrPrevKey <> "" And lgStrPrevLang <> "" Then      <%'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� %>
			Call DisableToolBar(parent.TBC_QUERY)					'�� : Query ��ư�� disable ��Ŵ.
			If DBQuery = False Then 
			   Call RestoreToolBar()
			   Exit Sub 
			End If 
    	End If

    End if
End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If

    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = False Then
       Exit Function
    End If
   
    
    FncQuery = True															

End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then  'Not chkField(Document, "2") OR  '��: Check contents area
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

    With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
			
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
			'Key field clear
			.Col = C_Msg
			.Text = ""
			
			.ReDraw = True
		End If
    End With
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim iRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
    
	With frm1
	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
		    .vspdData.Col = C_Lang
		    .vspdData.text = ""
		Next

		.vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtLang=" & .hLang.value 
		strVal = strVal & "&txtCode=" & .hMsg.value 
		strVal = strVal & "&txtType=" & .hMsgType.value 
		strVal = strVal & "&txtLevel=" & .hMsgLevel.value 
		strVal = strVal & "&txtText=" & .hMsgText.value 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevLang=" & lgStrPrevLang
    Else

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtLang=" & Trim(.cboLang.value)
		strVal = strVal & "&txtCode=" & Trim(.txtMsg.value)
		strVal = strVal & "&txtType=" & Trim(.cboMsgType.value)
		strVal = strVal & "&txtLevel=" & Trim(.cboMsgLevel.value)
		strVal = strVal & "&txtText=" & Trim(.txtMsgText.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevLang=" & lgStrPrevLang				
    End If
   
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

	Call SetToolbar("1100111100111111")										<%'��ư ���� ���� %>
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
<%  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data ���� ��Ģ 
    ' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ 

    For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep	'��: C=Create, Row��ġ ���� 
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'��: U=Update, Row��ġ ���� 
			End Select			

		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag			'��: �ű�, ���� 

		            .vspdData.Col = C_Lang	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Msg	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_MsgType	'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_MsgLevel	'5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_MsgText	'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag								'��: ���� 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'��: U=Update

		            .vspdData.Col = C_Lang	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Msg	'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
  
		            .vspdData.Col = C_MsgType	'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

  		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'��: �����Ͻ� ASP �� ���� %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Message</font></td>
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
									<TD CLASS="TD5">���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboLang" tag="11X" ALT="���" STYLE="WIDTH: 160px;"><OPTION value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5">�޼��� �ڵ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtMsg" SIZE=16 MAXLENGTH=6 tag="11XXXU" ALT="�޼��� �ڵ�"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�޼��� ����</TD>
									<TD CLASS="TD6"><SELECT NAME="cboMsgType" tag="11X"  CLASS=cboNormal ALT="�޼��� ����"><OPTION value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5">�޼��� Level</TD>
									<TD CLASS="TD6"><SELECT NAME="cboMsgLevel" tag="11X" CLASS=cboNormal" ALT="�޼��� Level"><OPTION value=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">����</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtMsgText" SIZE=30 MAXLENGTH=1024 tag="11" ALT="����"></TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b1c03ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGTH_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1c03mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hLang" tag="24">
<INPUT TYPE=HIDDEN NAME="hMsg" tag="24">
<INPUT TYPE=HIDDEN NAME="hMsgType" tag="24">
<INPUT TYPE=HIDDEN NAME="hMsgLevel" tag="24">
<INPUT TYPE=HIDDEN NAME="hMsgText" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

