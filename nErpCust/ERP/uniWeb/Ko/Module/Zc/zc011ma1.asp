<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : 
'*  3. Program ID           : zc011ma1.asp
'*  4. Program Name         : 화면이동정보등록 
'*  5. Program Desc         : 화면이동정보등록 
'*  6. Modified date(First) : 2005/02/24
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Shin Hyun Ho
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "ZC011mb1.asp"												'<%'비지니스 로직 ASP명 %>

<% 
	Dim StrCo
	StrCo = Request.Cookies("unierp")("gLang")
 %>
 
 Dim StrLang
    StrLang = "<%=Strco%>" 


<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_MOVE_ITEM_CD
Dim C_MOVE_ITEM_POP
Dim C_MOVE_ITEM_NM
Dim C_MINOR_NM
Dim C_TYPE_VALUE
Dim C_MOVE_DIRECTION
Dim C_MOVE_DIRECTION_NM
Dim C_NEXT_PGM_ID
Dim C_NEXT_PGM_ID_POP
Dim C_NEXT_PGM_ID_NM
Dim C_MOVE_PGM_DESC
Dim C_TYPE_VALUE2
Dim C_MOVE_DIRECTION2
Dim C_NEXT_PGM_ID2


Dim IsOpenPop          
Dim lsConcd

Sub initSpreadPosVariables()  
	C_MOVE_ITEM_CD		=	1
	C_MOVE_ITEM_POP		=	2
	C_MOVE_ITEM_NM		=	3
	C_MINOR_NM			=	4
	C_TYPE_VALUE		=	5
	C_MOVE_DIRECTION	=	6
	C_MOVE_DIRECTION_NM	=	7
	C_NEXT_PGM_ID		=	8
	C_NEXT_PGM_ID_POP	=	9
	C_NEXT_PGM_ID_NM	=	10
	C_MOVE_PGM_DESC		=	11
	C_TYPE_VALUE2		=	12
	C_MOVE_DIRECTION2	=	13
	C_NEXT_PGM_ID2		=	14
End Sub

Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0

    lgSortKey = 1
    lgPageNo = 0
End Sub


Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim pYesNo
	
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	.ReDraw = false
	
	.ReDraw = True
	End With
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
    ggoSpread.Source = frm1.vspdData

    'patch version
    ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
    .MaxCols = C_NEXT_PGM_ID2 + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	.Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
    .ColHidden = True
	
    .MaxRows = 0
    ggoSpread.ClearSpreadData

    Call GetSpreadColumnPos("A")    
   
    ggoSpread.SSSetEdit		C_MOVE_ITEM_CD		, "이동선택항목"	,15,  ,, 15,2
    ggoSpread.SSSetButton	C_MOVE_ITEM_POP		
    ggoSpread.SSSetEdit		C_MOVE_ITEM_NM		, "이동선택항목명"	,20,  ,, 30    
    ggoSpread.SSSetEdit		C_MINOR_NM			, "유형명"			,20,  ,, 30    
    ggoSpread.SSSetEdit		C_TYPE_VALUE		, "Value"			,10
    ggoSpread.SSSetCombo	C_MOVE_DIRECTION	, "방향"			,10
    ggoSpread.SSSetCombo	C_MOVE_DIRECTION_NM	, "방향"			,20
    ggoSpread.SSSetEdit		C_NEXT_PGM_ID		, "NEXT_PGM_ID"		,15,  ,, 15    
    ggoSpread.SSSetButton	C_NEXT_PGM_ID_POP
    ggoSpread.SSSetEdit		C_NEXT_PGM_ID_NM	, "NEXT_PGM_ID명"	,30,  ,, 30       
    ggoSpread.SSSetEdit		C_MOVE_PGM_DESC		, "비고"			,40
    ggoSpread.SSSetEdit		C_TYPE_VALUE2		, ""			,10
    ggoSpread.SSSetEdit		C_MOVE_DIRECTION2	, ""			,10
    ggoSpread.SSSetEdit		C_NEXT_PGM_ID2		, ""			,15


    Call ggoSpread.SSSetColHidden(C_MOVE_DIRECTION,C_MOVE_DIRECTION,True)
    Call ggoSpread.SSSetColHidden(C_MOVE_DIRECTION2,C_MOVE_DIRECTION2,True)
    Call ggoSpread.SSSetColHidden(C_TYPE_VALUE2,C_TYPE_VALUE2,True)
    Call ggoSpread.SSSetColHidden(C_NEXT_PGM_ID2,C_NEXT_PGM_ID2,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock		C_MOVE_ITEM_CD,			-1,	C_MINOR_NM
    ggoSpread.SSSetRequired		C_MOVE_DIRECTION_NM,	-1,	-1
    ggoSpread.SSSetRequired		C_NEXT_PGM_ID,			-1, -1
    ggoSpread.SpreadLock		C_NEXT_PGM_ID_NM,		-1,	C_NEXT_PGM_ID_NM
	ggoSpread.SSSetProtected	.vspdData.MaxCols,		-1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    Dim iRow
    
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired		C_MOVE_ITEM_CD,		pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_MOVE_DIRECTION_NM,pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_NEXT_PGM_ID,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_MOVE_ITEM_NM,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_MINOR_NM,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_NEXT_PGM_ID_NM,	pvStartRow, pvEndRow

    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_MOVE_ITEM_CD		=	iCurColumnPos(1)
			C_MOVE_ITEM_POP		=	iCurColumnPos(2)
			C_MOVE_ITEM_NM		=	iCurColumnPos(3)
			C_MINOR_NM			=	iCurColumnPos(4)
			C_TYPE_VALUE		=	iCurColumnPos(5)
			C_MOVE_DIRECTION	=	iCurColumnPos(6)
			C_MOVE_DIRECTION_NM	=	iCurColumnPos(7)
			C_NEXT_PGM_ID		=	iCurColumnPos(8)
			C_NEXT_PGM_ID_POP	=	iCurColumnPos(9)
			C_NEXT_PGM_ID_NM	=	iCurColumnPos(10)
			C_MOVE_PGM_DESC		=	iCurColumnPos(11)
			C_TYPE_VALUE2		=	iCurColumnPos(12)
			C_MOVE_DIRECTION2	=	iCurColumnPos(13)
			C_NEXT_PGM_ID2		=	iCurColumnPos(14)
                        
    End Select    
End Sub

Sub InitSpreadComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

	'앞,뒤 
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD='Z0040'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_MOVE_DIRECTION			'C_MOVE_DIRECTION
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_MOVE_DIRECTION_NM
End Sub

Function OpenPopup(Byval iWhere)

	Dim arrRet
	Dim arrParam(7), arrField(8), arrHeader(8)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	   
	   Case 0
			
			frm1.vspdData.Col = C_MOVE_ITEM_CD
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
	        arrParam(0) = "이동선택항목"
			arrParam(1) = "Z_MOVE_ITEM A,B_MAJOR B"
			arrParam(2) = Trim(frm1.vspdData.value)
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD *= B.MAJOR_CD"
			arrParam(5) = "이동선택항목"

			arrField(0) = "ED10" & Chr(11) & "MOVE_ITEM_CD"
			arrField(1) = "ED30" & Chr(11) & "MOVE_ITEM_NM"
			arrField(2) = "ED05" & Chr(11) & "A.MAJOR_CD"
			arrField(3) = "ED07" & Chr(11) & "MAJOR_NM"
			
			
			arrHeader(0) = "이동선택항목"
			arrHeader(1) = "이동선택항목명"
			arrHeader(2) = "종합코드"
			arrHeader(3) = "종합코드명"

         
       Case 1
			
			frm1.vspdData.Col = C_NEXT_PGM_ID
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
	        arrParam(0) = "NEXT_PGM_ID"
			arrParam(1) = "             (SELECT A.MNU_ID,A.MNU_TYPE,A.MNU_NM,B.UPPER_MNU_ID,(CASE WHEN B.USE_YN = 1 THEN 'TRUE' ELSE 'FALSE' END) AS USE_YN  "
			arrParam(1) = arrParam(1) & "  FROM Z_LANG_CO_MAST_MNU A,Z_CO_MAST_MNU      B "
			arrParam(1) = arrParam(1) & " WHERE A.MNU_ID = B.MNU_ID AND A.LANG_CD = 'KO' AND A.MNU_TYPE = 'P'  "
			arrParam(1) = arrParam(1) & "   AND UPPER_MNU_ID IN (SELECT MNU_ID "
			arrParam(1) = arrParam(1) &                           "FROM Z_CO_MAST_MNU )) A "
			arrParam(2) = Trim(frm1.vspdData.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "NEXT_PGM_ID"

			arrField(0) = "ED08" & Chr(11) & "A.MNU_ID"
			arrField(1) = "ED18" & Chr(11) & "A.MNU_NM"
			arrField(2) = "ED07" & Chr(11) & "A.UPPER_MNU_ID"
			arrField(3) = "ED07" & Chr(11) & "A.USE_YN"
			
			arrHeader(0) = "NEXT_PGM_ID"
			arrHeader(1) = "NEXT_PGM_ID명"
			arrHeader(2) = "상위메뉴"
			arrHeader(3) = "사용여부"

        
       Case 2 
			
			arrParam(0) = "이동선택항목"
			arrParam(1) = "Z_MOVE_PGM_INF A ,Z_MOVE_ITEM B"
			arrParam(2) = Trim(frm1.txtMoveItem.value)
			arrParam(3) = ""
			arrParam(4) = "A.MOVE_ITEM_CD = B.MOVE_ITEM_CD"
			arrParam(5) = "이동선택항목"

			arrField(0) = "ED10" & Chr(11) & "A.MOVE_ITEM_CD"
			arrField(1) = "ED30" & Chr(11) & "B.MOVE_ITEM_NM"
			arrField(2) = "ED05" & Chr(11) & "MOVE_DIRECTION"
			arrField(3) = "ED07" & Chr(11) & "NEXT_PGM_ID"
			
			
			arrHeader(0) = "이동선택항목"
			arrHeader(1) = "이동선택항목명"
			arrHeader(2) = "구분"
			arrHeader(3) = "다음항목"
			
       			
      End Select
	
	  
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SubSetPopup(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SubSetPopup()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetPopup(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    
		    CASE 0            
				
				.vspdData.Col = C_MOVE_ITEM_CD
				.vspdData.Text = Trim(arrRet(0))
				.vspdData.Col = C_MOVE_ITEM_NM
				.vspdData.Text = Trim(arrRet(1))
		        .vspdData.Col = C_MINOR_NM
				.vspdData.Text = Trim(arrRet(3))
				
				
		    
		    CASE 1            
				.vspdData.Col = C_NEXT_PGM_ID
				.vspdData.Text = Trim(arrRet(0))
				.vspdData.Col = C_NEXT_PGM_ID_NM
				.vspdData.Text = Trim(arrRet(1))
		        
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.activeRow

            CASE 2
                .txtMoveItem.Value		= Trim(arrRet(0))
		        .txtMoveItemNM.Value	= Trim(arrRet(1))  
				
				'lgBlnFlgChgValue		= True 
				
            
        End Select
	End With
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)


    If Row <= 0 Then
        Exit Sub
    End If

    With frm1
        Select Case Col
            Case C_MOVE_ITEM_POP '테이블ID
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(0)
            Case C_NEXT_PGM_ID_POP '컬럼ID
                .vspdData.Col = Col-1
                .vspdData.Row = Row
                Call OpenPopUp(1)
           
        End Select

'		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			
			Case  C_MOVE_DIRECTION_NM
				.Col = Col
				intIndex = .Value
				.Col = C_MOVE_DIRECTION
				.Value = intIndex
			
		End Select
	End With
End Sub

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
 
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call InitSpreadComboBox
    
    Call SetToolBar("1100110100101111")										<%'버튼 툴바 제어 %>
    Call CookiePage(0)
    
    
End Sub

Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	Call DbQuery
End Function

Sub vspdData_Change(ByVal Col , ByVal Row )
    
    Dim StrChk
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	Select Case Col
        Case C_MOVE_ITEM_CD
        	 
        	 StrChk =  CommonQueryRs("A.MOVE_ITEM_CD, A.MOVE_ITEM_NM,A.MAJOR_CD,B.MAJOR_NM", "Z_MOVE_ITEM A,B_MAJOR B", "A.MAJOR_CD *= B.MAJOR_CD AND A.MOVE_ITEM_CD = " & FilterVar(UCASE(Trim(frm1.vspdData.value)),"''","S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
			 If StrChk = FALSE Then
				Frm1.vspdData.Col = C_MOVE_ITEM_NM
				Frm1.vspdData.value = ""
				Frm1.vspdData.Col = C_MINOR_NM
				Frm1.vspdData.value = ""
			 Else
				Frm1.vspdData.Col = C_MOVE_ITEM_NM
				Frm1.vspdData.value = Replace(lgF1,Chr(11),"")
				Frm1.vspdData.Col = C_MINOR_NM
				Frm1.vspdData.value = Replace(lgF3,Chr(11),"")
			 End If
        Case C_NEXT_PGM_ID  
    End Select    
	
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
    Else
	
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

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
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
    If DbQuery	= False Then														<%'Query db data%>
       Exit Function
    End If
       
    FncQuery = True															
    
End Function

Function FncSave() 
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
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
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

 
 
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = false
		
			ggoSpread.Source = frm1.vspdData	
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
		
			.ReDraw = true
		End If
    End with

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
    
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

Function FncDeleteRow() 
	Dim lDeIRows
	Dim iDeIRowCnt, i
	Dim IntRetCD 

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 

		lDeIRows = ggoSpread.DeleteRow
    End With
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    iColumnLimit  =  5                 ' split 한계치  maxcol이 아님(5번째 칼럼이 split의 최고치)
                                       ' 5라는 값은 표준이 아닙니다.개발자가 업무에 맞게 수정요 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.SSSetSplit(ACol)    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

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

Function DbQuery() 
	Dim strVal

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtMoveItem=" & Trim(.HtxtMoveItem.value)	'조회 조건 데이타 
			strVal = strVal & "&txtLANG=" & StrLang	'국가 코드 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtMoveItem=" & Trim(.txtMoveItem.value)	'조회 조건 데이타 
			strVal = strVal & "&txtLANG=" & StrLang	'국가 코드 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 
    End With

    DbQuery = True
End Function

Function MakeKeyStream()

	lgKeyStream = UCase(Trim(frm1.txtMoveItem.value))  & Parent.gColSep

End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE

    Call InitData()

	Call SetToolBar("110011110011111")										<%'버튼 툴바 제어 %>

End Function

Function DbSave() 
	Dim IRow
	Dim lGrpCnt
	Dim strVal, strDel

    DbSave = False
    On Error Resume Next
    Err.Clear 

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For IRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = IRow
		    .vspdData.Col = 0


		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag															'☜: 신규 
					strVal = strVal & "C"  & Parent.gColSep & IRow & Parent.gColSep					'☜: C=Create
		            .vspdData.Col = C_MOVE_ITEM_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TYPE_VALUE
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_DIRECTION
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_NEXT_PGM_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_PGM_DESC
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
				Case ggoSpread.UpdateFlag															'☜: 수정 
					strVal = strVal & "U"  & Parent.gColSep & IRow & Parent.gColSep					'☜: U=Update
		            .vspdData.Col = C_MOVE_ITEM_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TYPE_VALUE
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_DIRECTION
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_NEXT_PGM_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_PGM_DESC
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TYPE_VALUE2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_DIRECTION2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_NEXT_PGM_ID2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				    lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag																'☜: 삭제 
					strDel = strDel & "D" & Parent.gColSep & IRow & Parent.gColSep
		            .vspdData.Col = C_MOVE_ITEM_CD
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TYPE_VALUE
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MOVE_DIRECTION2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_NEXT_PGM_ID2
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		Next
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)																'☜: 비지니스 ASP 를 가동 
	End With

    DbSave = True																						'⊙: Processing is NG
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>화면이동정보등록</font></td>
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
									<TD CLASS="TD5">이동선택항목</TD>
									<TD CLASS="TD656">
										<INPUT TYPE=TEXT NAME="txtMoveItem" SIZE=20 MAXLENGTH=15 tag="11XXXU" ALT="이동선택항목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Openpopup(2)">
										<INPUT TYPE=TEXT NAME="txtMoveItemNm" tag="14X">
									</TD>
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
									<script language =javascript src='./js/zc011ma1_I255496561_vspdData.js'></script>
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
    <tr HEIGHT="20">
      <td WIDTH="100%">
      <table <%=LR_SPACE_TYPE_30%>>
        <tr>
          
        </tr>
      </table>
      </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1a01mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24">
</TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtMoveItem" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

