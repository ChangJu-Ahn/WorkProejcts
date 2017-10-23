
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Zip code)
'*  3. Program ID           : B1g01ma1.asp
'*  4. Program Name         : B1g01ma1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/14
'*  7. Modified date(Last)  : 2002/12/13
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

Option Explicit																	<%'��: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B1g01mb1.asp"												<%'�����Ͻ� ���� ASP�� %>
 
Dim C_ZipCd
Dim C_SerNo
Dim C_Address
Dim C_Ref1
Dim C_Ref2
Dim C_Ref3
Dim C_Country


<!-- #Include file="../../inc/lgvariables.inc" -->


Dim IsOpenPop         
Dim lgStrPrevSer, lgStrPrevAdd,lgStrNo 

Sub InitSpreadPosVariables()
    C_ZipCd     = 1
    C_SerNo     = 2
    C_Address   = 3
    C_Ref1      = 4
    C_Ref2      = 5
    C_Ref3      = 6
    C_Country   = 7
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevSer = ""
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevAdd = ""                           'initializes Previous Key
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
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
	.ReDraw = false

    .MaxCols = C_Country + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetEdit C_ZipCd, "�����ȣ", 10,,,12,2    
    ggoSpread.SSSetEdit C_SerNo, "Serial No", 10,,,12
    ggoSpread.SSSetEdit C_Address, "�ּ�", 60,,,100
    ggoSpread.SSSetEdit C_Ref1, "����", 20,,,50
    ggoSpread.SSSetEdit C_Ref2, "ȣ", 10,,,50
    ggoSpread.SSSetEdit C_Ref3, "��Ÿ", 10,,,50
    ggoSpread.SSSetEdit C_Country, "�����ڵ�", 15
	
    Call ggoSpread.SSSetColHidden(C_SerNo,C_SerNo,True)

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ZipCd, -1, C_ZipCd
		ggoSpread.SSSetRequired C_Address, -1, -1
		ggoSpread.SpreadLock C_Country, -1, C_Country
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
		.vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_ZipCd, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Address, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Country, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_ZipCd     = iCurColumnPos(1)
            C_SerNo     = iCurColumnPos(2)
            C_Address   = iCurColumnPos(3)
            C_Ref1      = iCurColumnPos(4)
            C_Ref2      = iCurColumnPos(5)
            C_Ref3      = iCurColumnPos(6)
            C_Country   = iCurColumnPos(7)
    End Select    
End Sub
 
Function OpenCountry()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"				' �˾� ��Ī 
	arrParam(1) = "b_country"					' TABLE ��Ī 
	arrParam(2) = frm1.txtCountryCd.value		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' �����ʵ��� �� ��Ī 
	
    arrField(0) = "country_cd"					' Field��(0)
    arrField(1) = "country_nm"					' Field��(1)
    
    arrHeader(0) = "�����ڵ�"				' Header��(0)
    arrHeader(1) = "����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtCountryCd.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCountry(arrRet)
	End If	
	
End Function

Function SetCountry(Byval arrRet)
	With frm1
		.txtCountryCd.value = arrRet(0)
		.txtCountryNm.value = arrRet(1)
	End With
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
      
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100101111")										<%'��ư ���� ���� %>
    frm1.txtCountryCd.value = parent.gCountry
    frm1.txtCountryCd.focus
    Call fncQuery
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	 
    	If lgStrPrevKey <> "" Then 
      	DbQuery
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
    frm1.txtCountryNm.value = ""
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
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
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
    If Not ggoSpread.SSDefaultCheck Then   'Not chkField(Document, "2") OR   '��: Check contents area
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
			.Col = C_ZipCd
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
		.vspdData.Row = iRow
		.vspdData.Col = C_Country
		.vspdData.Text = UCase(.txtCountryCd.value)
        Next    		
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
		
		
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtCountry="   & Trim(.hCountryCd.value) 				
		strVal = strVal & "&txtZipCd="     & Trim(.hZipCd.value) 					
		strVal = strVal & "&txtAddress="   & Trim(.hAddress.value)		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrNo="      & lgStrNo
		strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
		
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtCountry="   & Trim(.txtCountryCd.value)				
		strVal = strVal & "&txtZipCd="     & Trim(.txtZipCd.value)				
		strVal = strVal & "&txtAddress="   & Trim(.txtAddress.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrNo="      & lgStrNo								
		strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
	
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
	Call SetToolbar("1100111100111111")										'��: ��ư ���� ���� 
	
End Function

Function DbLookUp()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6 
       
	call CommonQueryRs("country_nm "," b_country "," country_cd =  " & FilterVar(frm1.txtCountryCd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    frm1.txtCountrynm.value = Trim(Replace(lgF0,Chr(11),"")) 

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
		        Case ggoSpread.InsertFlag									'��: �ű� 
					strVal = strVal & "C" & parent.gColSep	                		'��: C=Create
		        Case ggoSpread.UpdateFlag									'��: ���� 
					strVal = strVal & "U" & parent.gColSep		                 	'��: U=Update
			End Select
			
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					.vspdData.Col = C_Country		'7
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_SerNo			'1
					strVal = strVal & "0" & parent.gColSep
					        
					.vspdData.Col = C_ZipCd			'2
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					        
					.vspdData.Col = C_Address		'3
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref1			'4
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref2			'5
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref3			'6
					strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
					        
					lGrpCnt = lGrpCnt + 1
					
				Case ggoSpread.UpdateFlag
					.vspdData.Col = C_Country		'7
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_SerNo			'1
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					        
					.vspdData.Col = C_ZipCd			'2
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					        
					.vspdData.Col = C_Address		'3
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref1			'4
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref2			'5
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_Ref3			'6
					strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
					        
					lGrpCnt = lGrpCnt + 1
					
				Case  ggoSpread.DeleteFlag
					strDel = strDel & "D" & parent.gColSep                   		'��: U=Update

					.vspdData.Col = C_Country		'7
					strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_SerNo			'1
					strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_ZipCd			'2
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
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ȣ</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5">�� ��</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtCountryCd" SIZE=10 MAXLENGTH=2 tag="12XXXU"  ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCountry()">
										<INPUT TYPE=TEXT NAME="txtCountryNm" tag="14X">
									</TD>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�����ȣ</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtZipCd" SIZE=12  MAXLENGTH=12 tag="11XXXU">
									</TD>
									<TD CLASS="TD5">�ּ�</TD>
									<TD CLASS="TD6">
										<INPUT TYPE=TEXT NAME="txtAddress" SIZE=30 MAXLENGTH=100 tag="11">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%">
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b1g01ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1g01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hCountryCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSerNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hZipCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hAddress" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

