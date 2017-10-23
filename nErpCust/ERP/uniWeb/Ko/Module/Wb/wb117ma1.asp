
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : WB117MA1
'*  4. Program Name         : WB117MA1.asp
'*  5. Program Desc         : 작업진행조회 및 마감 
'*  6. Modified date(First) : 2005/02/14
'*  7. Modified date(Last)  : 2005/02/14
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "wb117ma1"
Const BIZ_PGM_ID = "wb117mb1.asp"											 '☆: 비지니스 로직 ASP명 

Dim C_PGM_ID
Dim C_GROUP_NM
Dim C_MNU_NM
Dim C_STATUS_FLG
Dim C_CONFIRM_FLG
Dim C_UPDT_USER_ID
Dim C_UPDT_DT

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_PGM_ID		= 1	
	C_GROUP_NM		= 2
	C_MNU_NM		= 3
	C_STATUS_FLG	= 4
	C_CONFIRM_FLG	= 5
	C_UPDT_USER_ID	= 6
	C_UPDT_DT		= 7

End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_UPDT_DT + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    'Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_PGM_ID,		"PGM ID", 7,,,10,1
	ggoSpread.SSSetEdit		C_GROUP_NM,		"작업단계 그룹", 20,,,50,1
	ggoSpread.SSSetEdit		C_MNU_NM,		"작업단계", 30,,,100,1
    ggoSpread.SSSetEdit		C_STATUS_FLG,	"진행상태", 10,,,10,2
    ggoSpread.SSSetCheck	C_CONFIRM_FLG,	"마감여부", 10,,,True
	ggoSpread.SSSetEdit		C_UPDT_USER_ID,	"최종작업자", 15,,,20,1
    ggoSpread.SSSetEdit		C_UPDT_DT,		"최종작업시간", 20,,,20,2
    	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_PGM_ID,C_PGM_ID,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With

End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()

End Sub


Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_PGM_ID, -1, C_STATUS_FLG
	ggoSpread.SpreadLock C_UPDT_USER_ID, -1, C_UPDT_DT
	
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_PGM_ID			= iCurColumnPos(1)
            C_GROUP_NM			= iCurColumnPos(2)
            C_MNU_NM			= iCurColumnPos(3)
            C_STATUS_FLG		= iCurColumnPos(4)
            C_CONFIRM_FLG		= iCurColumnPos(5)
            C_UPDT_USER_ID		= iCurColumnPos(6)
            C_UPDT_DT			= iCurColumnPos(7)
    End Select    
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

End Sub

'============================================  조회조건 함수  ====================================
Sub BtnAllChk()
	Dim iRow, iMaxRows 
	ggoSpread.Source = frm1.vspdData
    
	With frm1.vspdData
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_CONFIRM_FLG
			If .value = "0" Then 
				.value = "1"
				lgBlnFlgChgValue = True
				ggoSpread.UpdateRow iRow
			End If
		Next
		
	End With
End Sub

Sub BtnChkCancel()
	Dim iRow, iMaxRows
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_CONFIRM_FLG
			If .value = "1" Then 
				.value = "0"
				lgBlnFlgChgValue = True
				ggoSpread.UpdateRow iRow
			End If
		Next
		
	End With
End Sub

Sub BtnVerify()
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

    Call fncQuery()
    
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	With frm1.vspdData
	
    .Row = Row
    .Col = Col

    If .CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(.text) < UNICDbl(.TypeFloatMin) Then
         .text = .TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
	End With
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

'============================================  툴바지원 함수  ====================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100000000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
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
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, dblSum
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If    

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
 
End Function

Function FncCancel() 
                                                 '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function InitGrid2()    
	Dim i, iRow, iCol, iMaxRows, ret, sField, sFrom , sWhere, arrMinorNm, arrMinorCd, arrSeqNo, arrRef, arrW1_W2, sOldW1, sOldW2
	Dim soldMinorNm, soldMinorCd, iSpanRowW1, iSpanRowW2, iSpanCntW1, iSpanCntW2
	
	If frm1.vspdData2.MaxRows > 0 Then Exit Function
	
	sField	= "	A.MINOR_NM, B.MINOR_CD, B.SEQ_NO, B.REFERENCE"
	sFrom	= " B_MINOR A " & vbCrLf
	sFrom	= sFrom	& " 	INNER JOIN B_CONFIGURATION B WITH (NOLOCK) ON A.MAJOR_CD=B.MAJOR_CD AND A.MINOR_CD=B.MINOR_CD "
	sWhere	= " A.MAJOR_CD='W2003' "

	Call CommonQueryRs(sField, sFrom, sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0
		arrMinorNm	= Split(lgF0 , Chr(11))
		arrMinorCd	= Split(lgF1 , Chr(11))
		arrSeqNo	= Split(lgF2 , Chr(11))
		arrRef		= Split(lgF3 , Chr(11))
		
		iMaxRows = UBound(arrMinorNm)
		
		For i = 0 to iMaxRows -1
			
			ggoSpread.InsertRow , 1
			.Row = iRow + 1 : iRow = iRow + 1 : .Col = C_SEQ_NO : .Value = iRow
			
			.Col = C_MINOR_NM	: .Value = arrMinorNm(i)
			.Col = C_MINOR_CD	: .Value = arrMinorCd(i)
			
			arrW1_W2 = Split(arrMinorNm(i), "|")	' 코드명을 | 로 분리한다.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 비교 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 합침 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 합침 
			End If		
			
			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
				
			For iCol = 1 To 6	' 컬럼 갯수 
				.Col = C_W2 + iCol 
				Select Case iCol
					Case 1, 4
						.Value = ReadCombo1(arrRef(i)) : i = i + 1
					Case 2, 5
						.Value = ReadCombo2(arrRef(i)) : i = i + 1
					Case 6
						.Value = arrRef(i) ' For 문에서 i 값이 증가한다.
					Case 3
						.Value = "~"
				End Select

				
			Next

			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
		
		Call SetSpreadLock2
		
	End With

End Function

' -- 그리드 span
Function SetGridSpan()
	Dim soldMinorCd, iSpanCntW1, iSpanCntW2, iSpanRowW1, iSpanRowW2, iRow, i, sMinorNm, arrW1_W2, sOldW1, sOldW2, ret, iMaxRows
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0

		
		iMaxRows = .MaxRows
		
		For i = 1 to iMaxRows 
			
			.Row = i : .Col = C_MINOR_NM : sMinorNm = .Text
			
			arrW1_W2 = Split(sMinorNm, "|")	' 코드명을 | 로 분리한다.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 비교 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 합침 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 합침 
			End If	

			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
			
			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
	End With
End Function

Function FncDeleteRow() 

End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 

End Function
'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        strVal = strVal     & "&txtCurrGrid="        & lgCurrGrid      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.vspdData.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		'Call SetGridSpan
		
		Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>

	End If
	'frm1.vspdData.focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
   strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1.vspdData
		' ----- 1번째 그리드 
		ggoSpread.Source = frm1.vspdData
		lMaxRows = .MaxRows : lMaxCols = .MaxCols
				
		For lRow = 1 To lMaxRows
		    
		   .Row = lRow : .Col = 0
		   
		   ' I/U/D 플래그 처리 
		   Select Case .Text
		       Case  ggoSpread.UpdateFlag                                      '☜: Update                                                  
		                                           strVal = strVal & "U"  &  Parent.gColSep                                                 
		            lGrpCnt = lGrpCnt + 1                                                 
		  End Select
		 .Col = 0
		  ' 모든 그리드 데이타 보냄     
		  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
				For lCol = 1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Value) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
 
		Next
	End With

	Frm1.txtSpread.value      =  strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		.MaxRows = 0
		ggoSpread.ClearSpreadData
	End With
	
    Call MainQuery()
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/wb117ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
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
									<script language =javascript src='./js/wb117ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>>
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
					<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnAllChk()"   Flag=1>일괄마감</BUTTON>&nbsp;
					<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnChkCancel()"   Flag=1>일괄취소</BUTTON>&nbsp;
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

