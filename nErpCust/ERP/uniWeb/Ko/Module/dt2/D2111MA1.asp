
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit)
'*  3. Program ID           : B1f01ma1.asp
'*  4. Program Name         : B1f01ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/08
'*  7. Modified date(Last)  : 2002/06/21
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "D2111mb1.asp"												<%'�����Ͻ� ���� ASP�� %>

Dim C_user_id
Dim C_user_id_pop
Dim C_user_name
Dim C_smartbill_id
Dim C_smartbill_pw
 

Dim IsOpenPop          
Dim lgSortKey1
Dim lgSortKey2

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
    C_user_id			= 1
    C_user_id_pop		= 2
    C_user_name		= 3
    C_smartbill_id   = 4
    C_smartbill_pw	= 5
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","BA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20090922",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_smartbill_pw + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetEdit		C_user_id,		"ERP �����", 15,,,15,2
	ggoSpread.SSSetButton	C_user_id_pop
	ggoSpread.SSSetEdit		C_user_name,	"ERP ����ڸ�", 20,,,20,2
	ggoSpread.SSSetEdit		C_smartbill_id,	"SmartBill ID", 20,,,20,1
	ggoSpread.SSSetEdit		C_smartbill_pw,	"SmartBill PW", 20,,,20,1
	
	call ggoSpread.MakePairsColumn(C_user_id,C_user_id_pop)

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SSSetRequired		C_smartbill_id,	-1, -1
    ggoSpread.SSSetRequired		C_smartbill_pw,	-1, -1
    ggoSpread.SSSetRequired		C_user_id,			-1, -1
    'ggoSpread.SSSetRequired		C_user_name,		-1, -1
    
    
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_smartbill_id,	pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_smartbill_pw,	pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_user_id,			pvStartRow, pvEndRow
    'ggoSpread.SSSetProtected	C_user_id_pop,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_user_name,		pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadColor1(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected C_user_id, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_user_id_pop, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_user_name, pvStartRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_user_id			= iCurColumnPos(1)
            C_user_id_pop		= iCurColumnPos(2)
            C_user_name       = iCurColumnPos(3)
            C_smartbill_id		= iCurColumnPos(4)
            C_smartbill_pw    = iCurColumnPos(5)
    End Select    
End Sub

Function Open_User(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "usr_id, usr_nm"				<%' �˾� ��Ī %>
	arrParam(1) = "z_usr_mast_rec"				<%' TABLE ��Ī %>
	arrParam(2) = strCode						<%' Code Condition%>
	arrParam(4) = ""							<%' Name Cindition%>
	arrParam(5) = "�����"						<%' �����ʵ��� �� ��Ī %>

	arrField(0) = "usr_id"						<%' Field��(0)%>
	arrField(1) = "usr_nm"						<%' Field��(1)%>

	arrHeader(0) = "�����"						<%' Header��(0)%>
	arrHeader(1) = "����ڸ�"					<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUser(arrRet, iWhere)
	End If	

End Function

Function Open_User1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "usr_id, usr_nm"				<%' �˾� ��Ī %>
	arrParam(1) = "z_usr_mast_rec"				<%' TABLE ��Ī %>
	arrParam(2) = frm1.txtuserId.value			<%' Code Condition%>
	arrParam(4) = ""							<%' Name Cindition%>
	arrParam(5) = "�����"						<%' �����ʵ��� �� ��Ī %>

	arrField(0) = "usr_id"						<%' Field��(0)%>
	arrField(1) = "usr_nm"						<%' Field��(1)%>

	arrHeader(0) = "�����"						<%' Header��(0)%>
	arrHeader(1) = "����ڸ�"					<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtuserId.value = arrRet(0)
		frm1.txtuserNm.value = arrRet(1)
	End If	
	
End Function


Function SetUser(Byval arrRet, Byval iWhere)
	With frm1 
		.vspdData.Col = C_user_id
		.vspdData.Text = arrRet(0)

		.vspdData.Col = C_user_name
		.vspdData.Text = arrRet(1)

		lgBlnFlgChgValue = True
	End With
End Function

Sub Form_Load()
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                           <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
          
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>   
 
    Call SetToolbar("1100110100001111")										<%'��ư ���� ���� %>
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
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_user_id_pop Then
		    .Row = Row
		    .Col = C_user_id

		    Call Open_User(.Text, 1)
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'��: ������ üũ %>
    	If (lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "") Then <%'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� %>
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
    If frm1.txtUserId.value = "" then
        frm1.txtUserNm.value = ""
    End If
    
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
    If DbQuery = False Then Exit Function		  					<%'Query db data%>
       
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
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '��: Check contents area
       Exit Function
    End If
    

    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 

    frm1.vspdData.ReDraw = False

    if frm1.vspdData.maxrows < 1 then exit function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    'Key field clear
    frm1.vspdData.Col=C_user_id
    frm1.vspdData.Text=""

    frm1.vspdData.Col = C_user_name
    frm1.vspdData.Text=""

    frm1.vspdData.ReDraw = True

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow 

    'On Error Resume Next                                                          '��: If process fails
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
				ggoSpread.SSSetRequired		C_smartbill_id,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
				ggoSpread.SSSetRequired		C_smartbill_pw,	.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
				ggoSpread.SSSetRequired		C_user_id,			.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
				ggoSpread.SSSetProtected	C_user_name,		.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		.vspdData.ReDraw = True
    End With

    
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
		strVal = strVal & "&txtuserId=" & .hUserId.value 			'��: ��ȸ ���� ����Ÿ  			 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtUserId=" & Trim(.txtuserId.value)			'��: ��ȸ ���� ����Ÿ
		strVal = strVal & "&txtUserNm=" & Trim(.txtuserNm.value) 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

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

	call SetSpreadColor1(-1, -1)
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

		        Case ggoSpread.InsertFlag								'��: �ű� 
					
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'��: C=Create
										
		            .vspdData.Col = C_user_id		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_id	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_pw	'9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag								'��: ���� 
		
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'��: U=Update

		            .vspdData.Col = C_user_id		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_id	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_smartbill_pw	'9
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep		            
		         

		            lGrpCnt = lGrpCnt + 1
		            
		            
		        Case ggoSpread.DeleteFlag								'��: ���� 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'��: U=Update

		            .vspdData.Col = C_user_id		'3
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
	.txtMaxRows.value = lGrpCnt - 1	
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

Function CheckDelOk()
    Dim lRow,lRow2
    Dim lsDelUnit,lsDelToUnit,lsUpUnit,lsUpToUnit
    
	With frm1
         For lRow = 1 To .vspdData.MaxRows
             .vspdData.Row = lRow
             .vspdData.Col = 0
             if .vspdData.Text = ggoSpread.DeleteFlag then								'��: ���� 
                 .vspdData.Col = C_user_id		'3
                 lsDelUnit =  Trim(.vspdData.Text)
                 .vspdData.Col = C_ToUnit		'6
                 lsDelToUnit = Trim(.vspdData.Text) 
                 For lRow2 = 1 to  .vspdData.MaxRows
                     .vspdData.Row = lRow2
                     .vspdData.Col = 0
                     if .vspdData.Text = ggoSpread.UpdateFlag then						'��: ���� 
                         .vspdData.Col = C_user_id		'3
                         lsUpUnit =  Trim(.vspdData.Text)
                         .vspdData.Col = C_ToUnit		'6
                         lsUpToUnit = Trim(.vspdData.Text) 
                         if lsDelUnit = lsUpToUnit and lsDelToUnit = lsUpUnit then
                             CheckDelOk = False
                             exit function
                         end if
                     end if
                 next
              end if
        next
	End with
	CheckDelOk = true
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ڰ���</font></td>
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
									<TD CLASS="TD5">ERP�����</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtuserId" SIZE=10  MAXLENGTH=13 tag="11XXXU" ALT="ERP���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Open_User1()">
										<INPUT TYPE=TEXT NAME="txtuserNm" tag="14X">
									</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
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
									<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1f01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

