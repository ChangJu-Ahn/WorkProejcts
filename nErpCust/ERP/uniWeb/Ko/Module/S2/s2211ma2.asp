<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2211MA2
'*  4. Program Name         : �ǸŰ�ȹ�������ø����� 
'*  5. Program Desc         : �ǸŰ�ȹ�������ø����� 
'*  6. Comproxy List        : PS2G212.dll
'*  7. Modified date(First) : 2003/1/7
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Heeyoung Lee
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   1.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          1.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.Inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          1.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          1.3 Client Side Script
======================================================================================================== -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                             '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       2.1  Data Declaration Part
'########################################################################################################
'========================================================================================================
'=                       2.1.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "s2211mb2.asp"            '��: Head Query �����Ͻ� ���� ASP�� 

'========================================================================================================
'=                       2.1.2 Constant variables defined
'========================================================================================================
Const C_PopUnit		= 0


'========================================================================================================
'=                       2.1.3 Common variables 
'========================================================================================================= %>
<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       2.1.4.1 Variables For spreadsheet
'========================================================================================================
'��: Spread Sheet�� Column

Dim C_Unit
Dim C_UnitPopup
Dim C_Decimals
Dim C_RoundingUnit
Dim C_RoundingPolicy
Dim C_RoundingPolicy_NM

'========================================================================================================
'=                       2.1.4.2 User-defind Variables
'========================================================================================================
Dim lgIntSplitCol
Dim	lgIntSplitCol2
Dim lgBlnOpenPop

Dim lgLngStartRow		' Start row to be queryed

'########################################################################################################
'#                      3. Method Declaration Part
'########################################################################################################
'========================================================================================================
'========================================================================================================
'                       3.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_Unit				= 1
	C_UnitPopup			= 2
	C_Decimals			= 3
	C_RoundingUnit		= 4
	C_RoundingPolicy	= 5
	C_RoundingPolicy_NM	= 6
	
End Sub

'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgLngCurRows = 0  
    
    lgBlnOpenPop = False
    
End Sub

'=========================================================================================================
Sub SetDefaultVal()

	frm1.txtConUnit.focus
End Sub

<%
'==========================================================================================================
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
%>
Sub SetRowDefaultVal(ByVal pvRowCnt)
	Dim iIntRow
	
	With frm1.vspdData
		For iIntRow = 0 To pvRowCnt - 1
			.Row = .ActiveRow + iIntRow
				
			'�ø����� Default�� �����κ�. 
			.Col = C_RoundingPolicy		:	.Value = 1

			.Col = C_RoundingPolicy_NM  :	.Value = 1
			
			.Col = C_RoundingUnit		:	.Value = 0.1
		Next
	End With
End Sub

' Copy row
Sub SetRowCopyDefaultVal(ByVal pvRowCnt)

	With frm1.vspdData
	
		.Row = pvRowCnt
	
		.Col = C_Unit		:	.Text = ""
		
		' set the focus
		.Action = 0
	End With

End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables() 
	
		With frm1.vspdData		
			
		   	'��:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData
			ggoSpread.ClearSpreadData
			
			'patch version
		    ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_RoundingPolicy_NM + 1		'��: �ִ� Columns�� �׻� 1�� ������Ŵ 

            Call AppendNumberPlace("6","1","0")
		    Call GetSpreadColumnPos("A")
		    
		    ggoSpread.SSSetEdit C_Unit, "����",17,2,,,2	'3
            ggoSpread.SSSetButton C_UnitPopup       		'4
            ggoSpread.SSSetFloat C_Decimals,"�Ҽ����ڸ���" ,16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","4"
            ggoSpread.SSSetEdit C_RoundingUnit, "�ø�ó������", 18, 1	'10
            ggoSpread.SSSetCombo C_RoundingPolicy, "�ø�����", 20		'11
            ggoSpread.SSSetCombo C_RoundingPolicy_NM, "�ø�����", 20, 2		'11

			Call ggoSpread.MakePairsColumn(C_Unit,C_UnitPopup)

            Call ggoSpread.SSSetColHidden(C_RoundingPolicy,C_RoundingPolicy,True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
            
   		    Call SetSpreadLock()

			.ReDraw = True
		End With
    
    
End Sub

'==========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_RoundingUnit, -1, C_RoundingUnit
	ggoSpread.SpreadLock C_RoundingPolicy_NM, -1, C_RoundingPolicy_NM
End Sub


'==========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	ggoSpread.SSSetRequired		C_Unit		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_Decimals , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_RoundingUnit	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_RoundingPolicy_NM, pvStartRow, pvEndRow

End Sub

'==========================================================================================================
' Afetr query
Sub SetQuerySpreadColor(ByVal pvStartRow)

	With frm1.vspdData
		ggoSpread.SSSetProtected	C_Unit		, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected	C_UnitPopup	, pvStartRow, .MaxRows
		ggoSpread.SSSetRequired		C_Decimals	, pvStartRow, .MaxRows
	End With
End Sub

'==========================================================================================================
' Desc : This method set focus to position of error
'      : This method is called in MB area
'==========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
   	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_Unit				= iCurColumnPos(1)
			C_UnitPopup			= iCurColumnPos(2)
			C_Decimals			= iCurColumnPos(3)
			C_RoundingUnit		= iCurColumnPos(4) 
			C_RoundingPolicy	= iCurColumnPos(5)
			C_RoundingPolicy_NM = iCurColumnPos(6)
			
	End Select
	    
End Sub


'==========================================================================================================
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()
	Dim strCboData    ''lgF0
	Dim strCboData2    ''lgF1
	''FLAG(�ø�/�ݿø�)

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("B0004", "''", "S") & " ", _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData,  C_RoundingPolicy
	ggoSpread.SetCombo strCboData2, C_RoundingPolicy_NM

End Sub


'==========================================================================================================
' Desc : Reset ComboBox
'==========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			.col = C_RoundingPolicy     :	intIndex = .value
			.col = C_RoundingPolicy_NM  :	.value = intindex
		Next	
	End With
End Sub


'==========================================================================================================
'==========================================================================================================
'                       3.2 Common Group-2
'==========================================================================================================
'==========================================================================================================

'==========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029             '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("11001111001111")          '��: ��ư ���� ���� 

	Call InitSpreadSheet
	call InitSpreadComboBox()
	Call SetDefaultVal    
	Call InitVariables              '��: Initializes local global variables

End Sub

'==========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                  <%'��: Processing is NG%>
    
   ' Err.Clear             
                                                      <%'��: Protect system from crashing%>
    '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         <%'��: This function check indispensable field%>
       Exit Function
    End If

	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")          <%'��: Clear Contents  Field%>
    Call ggoSpread.ClearSpreadData()
    Call InitVariables               <%'��: Initializes local global variables%>

	'-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery                <%'��: Query db data%>

    FncQuery = True                <%'��: Processing is OK%>
        
End Function

'========================================================================================================
Function FncNew() 
End Function

'========================================================================================================
Function FncDelete() 
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'��: Processing is NG%>
    
    Err.Clear         
                                                    <%'��: Protect system from crashing%>

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		'Call MsgBox("No data changed!!", vbInformation)
		Exit Function
	End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then     <%'��: Check contents area%>
       Exit Function
    End If
    If ggoSpread.SSDefaultCheck = False Then     <%'��: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll  DbSave                                                   <%'��: Save db data%>
    
    FncSave = True                                                          <%'��: Processing is OK%>
    
End Function

'========================================================================================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	FncCopy = False
	
	With frm1.vspdData
		.ReDraw = False
		.focus
			 
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow

		Call SetRowCopyDefaultVal(.ActiveRow)
		.ReDraw = True
	End With

	lgBlnFlgChgValue = True

	If Err.number = 0 Then  FncCopy = True				                                '��: Processing is OK
	
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
Function FncCancel() 

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False                                                             '��: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    Call InitData()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncCancel = True                                                           '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 

	Dim IntRetCD
    Dim iIntInsertedRows
    On Error Resume Next                                                          '��: If process fails
    Err.Clear   

    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		iIntInsertedRows = CInt(pvRowCnt)
	Else
		iIntInsertedRows = AskSpdSheetAddRowcount()

		If iIntInsertedRows="" then Exit Function
	End If
   
   With frm1.vspdData
	
		.focus
		.ReDraw = False

		ggoSpread.Source = .vspdData

		ggoSpread.InsertRow,iIntInsertedRows
		
		' �����Էµ� Row�� Default �� ���� 
		Call SetRowDefaultVal(iIntInsertedRows)
		
		Call SetSpreadColor(.ActiveRow,.ActiveRow + iIntInsertedRows - 1)
		
		.ReDraw = True
    End With


    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
Function FncDeleteRow() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	With frm1.vspdData  
		.focus
		ggoSpread.Source = frm1.vspdData 
	
		Call ggoSpread.DeleteRow
		
		lgBlnFlgChgValue = True
	End With
	
    
End Function

'========================================================================================================
Function FncPrint() 
 Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
 On Error Resume Next                                                             '��: If process fails
    Err.Clear                                                                     '��: Clear error status
End Function

'========================================================================================================
Function FncFind()
	On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

End Function

'========================================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit, iColumnLimit2
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    iColumnLimit  = lgIntSplitCol         ' split �Ѱ�ġ  maxcol�� �ƴ�(6��° Į���� split�� �ְ�ġ)
                                       ' 6�̶�� ���� ǥ���� �ƴմϴ�.�����ڰ� ������ �°� ������ 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If gMouseClickStatus = "SPCRP" Then
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow

		If ACol > iColumnLimit Then
			Frm1.vspdData.Col = iColumnLimit : Frm1.vspdData.Row = 0
			iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
			Exit Function
		End If   
    
		Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE    
    
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
    
		Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL '0
		Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    End If

End Function
'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                       3.3 Common Group-3
'                           ���� : Fnc�Լ����� ȣ��Ǵ� ���� Function 
'========================================================================================================
'========================================================================================================

'========================================================================================================
Function DbQuery() 
	Err.Clear                                                               <%'��: Protect system from crashing%>
	    
	DbQuery = False                                                         <%'��: Processing is NG%>
	   
	If  LayerShowHide(1) = False Then
		Exit Function 
	End If
	    
	Dim iStrVal

	With Frm1
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'��: �����Ͻ� ó�� ASP�� ���� %>

	    If lgIntFlgMode = Parent.OPMD_UMODE Then    
			' Scroll
			iStrVal = iStrVal & "&txtWhere=" & .hConUnit.value			
		Else
			' Initial query
			iStrVal = iStrVal & "&txtWhere=" & .txtConUnit.value 			
		End If 
'		iStrVal = iStrVal & "&lgPageNo=" & lgPageNo
		iStrVal = iStrVal & "&txtSheetLastRow=" & frm1.vspdData.MaxRows
		
		lgLngStartRow = frm1.vspdData.MaxRows + 1
	End With

	Call RunMyBizASP(MyBizASP, iStrVal)            <%'��: �����Ͻ� ASP �� ���� %>
	DbQuery = True   
                          <%'��: Processing is NG%>
End Function

'========================================================================================================
Function DbQueryOk()              <%'��: ��ȸ ������ ������� %>
    '-----------------------
    'Reset variables area
    '-----------------------
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then		
		lgIntFlgMode = Parent.OPMD_UMODE            <%'��: Indicates that current mode is Update mode%>		
	End If
    
	frm1.vspdData.focus
	Call SetQuerySpreadColor(lgLngStartRow)
	Call InitData()
	Call ggoOper.LockField(Document, "Q")
	
    Set gActiveElement = document.ActiveElement  
End Function

'========================================================================================================
Function DbSave() 

	Err.Clear                <%'��: Protect system from crashing%>

	Dim iStrIns, iStrUpd, iStrDel, iStrKey
	Dim iLngRow
		 
	DbSave = False                                                          '��: Processing is NG
			    
	On Error Resume Next                                                   '��: Protect system from crashing
		   
	If LayerShowHide(1) = False Then
			Exit Function 
	End If

  '-----------------------
  'Data manipulate area
  '-----------------------
  iStrIns = ""
  iStrUpd = ""
  iStrDel = ""
   
  'Data manipulate area
  '-----------------------
  	With frm1.vspdData
	
		For iLngRow = 1 To .MaxRows
 
			.Row = iLngRow
			.Col = 0
		 if .Text <> "" Then
			Select Case .Text
			
					Case ggoSpread.InsertFlag     
					
						iStrIns = iStrIns & iLngRow & Parent.gColSep      '��: C=Create, Row��ġ ���� 
						.Col = C_Unit		' Unit
						iStrIns = iStrIns & Trim(.Text) & Parent.gColSep
						
						.Col = C_Decimals	' Decimals 
						iStrIns = iStrIns & Trim(.Text) & parent.gColSep
						
						.Col = C_RoundingUnit' Exchange rate Operator
						iStrIns = iStrIns & Trim(.Text) & parent.gColSep  'Round Unit  
						
						.Col = C_RoundingPolicy		' Item unit
						iStrIns = iStrIns & Trim(.Text) & Parent.gRowSep
						

					Case ggoSpread.UpdateFlag       '��: ���� 
					
						istrUpd = istrUpd & iLngRow & Parent.gColSep      '��: C=Create, Row��ġ ���� 
						.Col = C_Unit		' Unit
						istrUpd = istrUpd & Trim(.Text) & Parent.gColSep
						
						.Col = C_Decimals	' Decimals 
						istrUpd = istrUpd & Trim(.Text) & parent.gColSep
						
						.Col = C_RoundingUnit' Exchange rate Operator
						istrUpd = istrUpd & Trim(.Text) & parent.gColSep  'Round Unit  
						
						.Col = C_RoundingPolicy		' Item unit
						istrUpd = istrUpd & Trim(.Text) & Parent.gRowSep


					Case ggoSpread.DeleteFlag       '��: ���� 
						iStrDel = iStrDel & iLngRow & Parent.gColSep      '��: C=Create, Row��ġ ���� 
						.Col = C_Unit		' Unit
						iStrDel = iStrDel & Trim(.Text) & Parent.gRowSep
			end select
		 end if
		Next
	End With
	
	With frm1
	  .txtMode.value = Parent.UID_M0002
	  .txtSpreadIns.value = iStrIns
	  .txtSpreadUpd.value = istrUpd
	  .txtSpreadDel.value = iStrDel
	End With

 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         '��: �����Ͻ� ASP �� ���� 
 
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================================
Function DbSaveOk()               <%'��: ���� ������ ���� ���� %>

	Call InitVariables
	frm1.vspdData.MaxRows = 0
    Call MainQuery()
    
End Function

'========================================================================================================
Function DbDelete() 
    On Error Resume Next                                            <%'��: Protect system from crashing%>
End Function

'========================================================================================================
Function DbDeleteOk()              <%'��: ���� ������ ���� ���� %>
    On Error Resume Next                                            <%'��: Protect system from crashing%>
End Function

'========================================================================================================
'========================================================================================================
'                       3.4 User-defined Method 
'========================================================================================================
'========================================================================================================

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	Select Case pvIntWhere
	Case C_PopUnit
			iArrParam(1) = "dbo.B_UNIT_OF_MEASURE "			<%' TABLE ��Ī %>
			iArrParam(2) = frm1.txtConUnit.value						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = " DIMENSION <> " & FilterVar("TM", "''", "S") & " "			<%' Where Condition%>
			iArrParam(5) = "����"						<%' TextBox ��Ī %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "UNIT"
			iArrField(1) = "ED30" & Parent.gColSep & "UNIT_NM"
			    
			iArrHeader(0) = "����"
			iArrHeader(1) = "������"
			
			frm1.txtConUnit.focus
			
	End Select
	
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================================================================================
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	
	
	OpenSpreadPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	Select Case pvLngCol
	
		Case C_UnitPopup
			iArrParam(1) = "dbo.B_UNIT_OF_MEASURE "			<%' TABLE ��Ī %>
			iArrParam(2) = pvStrData						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = " DIMENSION <> " & FilterVar("TM", "''", "S") & " "			<%' Where Condition%>
			iArrParam(5) = "����"						<%' TextBox ��Ī %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "UNIT"
			iArrField(1) = "ED30" & Parent.gColSep & "UNIT_NM"
			    
			iArrHeader(0) = "����"
			iArrHeader(1) = "������"
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
		Call vspdData_Change(pvLngCol , pvLngRow)
	End If	



End Function
' Item Popup

'=======================================3.4.2 POP-UP (Set) ===============================================
' Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'===========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopUnit
			.txtConUnit.value = pvArrRet(0) 
			
		End Select
	End With

	SetConPopup = True
End Function


'========================================================================================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)

	SetSpreadPopup = False
	With frm1.vspdData
		.Row = pvLngRow

		Select Case pvLngCol
		
			Case C_UnitPopup
				.Col = C_Unit			: .Text = pvArrRet(0)				
		End Select
		
	End With

	SetSpreadPopup = True

End Function

'========================================================================================================
'   Event Desc : Update the Row Status
'===========================================================================================================
Sub SetRowStatus(intRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow intRow

 lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Desc : Update the Row Status
'===========================================================================================================
Sub SetRoundingUnit(ByVal pvLngRow,ByVal pvIntDecimals)
Dim i, j, x
Dim intIndex
Dim lRound, lRoundP
Dim strCellText
Dim Col
    ggoSpread.Source = frm1.vspdData

    lgBlnFlgChgValue = True
    
    With frm1.vspdData
    
    	   Col = C_Decimals  ''���� �ȵ� 
		  .Col = Col
		  .Row = pvLngRow
		  j = pvIntDecimals
		  
   		  lRound = 0.1
		  lRoundP = 1
		  		      
		    If j > 0 Then
		        For i = 1 To j
		            lRound = lRound * 0.1
		        Next
		        
		        .Col = C_RoundingUnit
		        .Row = pvLngRow
		        .value = lRound
		        
		    ElseIf j = 0 Then
		        .Col = C_RoundingUnit
		        .Row = pvLngRow
		        .value = lRound
		        
		    Else
		        For i = 1 To (j * -1)
		            lRoundP = lRoundP * 10
		        Next
		        
		        lRoundP = lRoundP / 10
		        .Col = C_RoundingUnit
		        .Row = pvLngRow
		        .value = lRoundP
		        
		    End If    		   
		  
		'***   frm1 �ȿ� ���� �ϴµ�...vspdData�̾ȿ� ���� ������ ������...0416
		
    End with
       
End Sub

'========================================================================================================
'========================================================================================================
'                       3.5 Spread Popup Method
'========================================================================================================
'========================================================================================================
'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	 '------ Developer Coding part (Start ) --------------------------------------------------------------
	If gMouseClickStatus = "SPCRP" Then	SetQuerySpreadColor(1)

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'========================================================================================================
'========================================================================================================
'                       3.6 Spread OCX Tag Event
'========================================================================================================
'========================================================================================================
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

 <% '----------  Coding part  -------------------------------------------------------------%>   
 'dim C_UnitPopup
 
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
			
	If Row > 0 Then
		Select Case Col
		
			CASE C_UnitPopup
				.Col = C_Unit
				.Row = Row
				call OpenSpreadPopup(col, Row, .Text) 
						
		End Select

		Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
	End If
		
	End With

End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("1101111111") 
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
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


'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim iStrData, iStrOldSpPeriod
	Dim iDblOldAmt, iDblQty, lDblAmt
	
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		CALL SetRowStatus(Row)

		.Col = Col	: iStrData = .Text
		
		If iStrData = "" Then Exit Sub
	
	end with
	
    If Col = C_Decimals  Then    
		call SetRoundingUnit(Row,frm1.vspdData.value)
    end if

End Sub
'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : 
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

'========================================================================================================
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

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



'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And frm1.hConUnit.value <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg	
	Dim ii
	Dim iChkAcctForVat

	 '---------- Coding part -------------------------------------------------------------
	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 

	With frm1
		.vspddata.Row = Row
		
		Select Case Col
			Case C_RoundingPolicy_NM 

				.vspddata.Col = Col				
				intIndex = .vspddata.Value
				.vspddata.Col = C_RoundingPolicy 
				.vspddata.Value = intIndex
		End Select
		
	End With

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<!--'======================================================================================================
'            6. Tag�� 
' ���: Tag�κ� ���� 
'======================================================================================================= -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ǸŰ�ȹ�������ø�����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD656"><INPUT NAME="txtConUnit" ALT="����" TYPE="Text" MAXLENGTH=3 SiZE=20 tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopUnit)"></TD>
									
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>    
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s2211ma2_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>    
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" src="../../blank.htm"  HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA class=hidden name=txtSpreadIns tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadUpd tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadDel tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hConUnit" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtSheetMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


