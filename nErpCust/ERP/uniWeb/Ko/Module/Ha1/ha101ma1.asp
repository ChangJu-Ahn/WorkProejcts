<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 퇴직정산(근속가산 및 누진율 등록)
*  3. Program ID           : h1a01ma1.asp
*  4. Program Name         : h1a01ma1.asp
*  5. Program Desc         :
*  6. Modified date(First) : 2001/06/19
*  7. Modified date(Last)  : 2003/06/16
*  8. Modifier (First)     : Hwang Jeong-won
*  9. Modifier (Last)      : Lee SiNa
* 10. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "ha101mb1.asp"												<%'비지니스 로직 ASP명 %>
Const CookieSplit = 1233
Const TAB1 = 1
Const TAB2 = 2
Const C_SHEETMAXROWS = 30														 <%'한 화면에 보여지는 최대갯수*1.5%>
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
Dim lsConcd
Dim gSelframeFlg			   ' 현재 TAB의 위치를 나타내는 Flag

Dim C_PAY_GRD1
Dim C_POPUP
Dim C_PAY_GRD1_NM
Dim C_ADD_RATE

Dim C_DUTY_MM
Dim C_ACCUM_RATE
'========================================================================================================
'	Name : initSpreadPosVariables()
'	Description : 변수 초기화 
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
    
		 C_PAY_GRD1		= 1															<%'Spread Sheet의 Column별 상수 %>
		 C_POPUP		= 2
		 C_PAY_GRD1_NM	= 3
		 C_ADD_RATE		= 4		 

    ElseIf pvSpdNo = "B" Then
    
		 C_DUTY_MM		= 1															<%'Spread Sheet의 Column별 상수 %>
		 C_ACCUM_RATE	= 2
    End If

End Sub
'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = ""                      'initializes Previous Key Index
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
  Call AppendNumberPlace("6","3","0")

  If pvSpdNo = "" OR pvSpdNo = "A" Then

		Call initSpreadPosVariables("A")	

		With frm1.vspdData
           ggoSpread.Source = frm1.vspdData	
           ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread    
           .ReDraw = false
           .MaxCols = C_ADD_RATE + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
           .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
           .ColHidden = True
           .MaxRows = 0

		   Call GetSpreadColumnPos("A") 

		   ggoSpread.SSSetEdit		C_PAY_GRD1, "급호", 10,,, 2
		   ggoSpread.SSSetButton	C_POPUP
		   ggoSpread.SSSetEdit		C_PAY_GRD1_NM, "급호명", 30
		   ggoSpread.SSSetFloat		C_ADD_RATE,"가산율" ,20,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","999"
		   
		   Call ggoSpread.MakePairsColumn(C_PAY_GRD1,  C_POPUP)

		   .ReDraw = true

		Call SetSpreadLock("A") 
    
		End With
    
    End if
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then		

		Call initSpreadPosVariables("B")	
		With frm1.vspdData1
 
		    ggoSpread.Source = frm1.vspdData1	
		    ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

		    .ReDraw = false
		    .MaxCols = C_ACCUM_RATE + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
		    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
		    .ColHidden = True
		    .MaxRows = 0		
		    		   
		    Call GetSpreadColumnPos("B")
   
		    ggoSpread.SSSetFloat C_DUTY_MM,"근속개월수" , 20,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","999"
		    ggoSpread.SSSetFloat C_ACCUM_RATE,"누진근속개월수" ,20,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","999"
			
			.ReDraw = true
	
		Call SetSpreadLock("B") 
    
		End With
	End if
    
End Sub

'======================================================================================================
'	기능: ClickTab1()
'	설명: Tab Click시 필요한 기능을 수행한다.
'         Header Tab처리 부분 (Header Tab이 있는 경우만 사용)
'=======================================================================================================
Function ClickTab1()
	Dim IntRetCD

	If gSelframeFlg = TAB1 Then Exit Function
	
	 ggoSpread.Source = frm1.vspdData1
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	Call changeTabs(TAB1)                                               <%'첫번째 Tab%>
	Call InitSpreadSheet("")

	ggoSpread.Source = frm1.vspdData
	gSelframeFlg = TAB1

End Function

'======================================================================================================
'	기능: ClickTab2()
'	설명: Tab Click시 필요한 기능을 수행한다.
'======================================================================================================
Function ClickTab2()
	Dim IntRetCD

	If gSelframeFlg = TAB2 Then Exit Function
	
	 ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call changeTabs(TAB2)
	Call InitSpreadSheet("")
    
	ggoSpread.Source = frm1.vspdData1
	gSelframeFlg = TAB2

End Function

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    If pvSpdNo = "A" Then

        ggoSpread.Source = Frm1.vspdData

        With frm1.vspdData
        	.ReDraw = False

			ggoSpread.SpreadLock	C_PAY_GRD1		, -1, C_PAY_GRD1
			ggoSpread.SpreadLock	C_POPUP			, -1, C_POPUP
			ggoSpread.SpreadLock	C_PAY_GRD1_NM	, -1, C_PAY_GRD1_NM
			ggoSpread.SSSetRequired	C_ADD_RATE		, -1, -1  		
			ggoSpread.SSSetProtected  .MaxCols   , -1, -1

        	.ReDraw = True
        End With
        
    ElseIf pvSpdNo = "B" Then
        ggoSpread.Source = Frm1.vspdData1

        With frm1.vspdData1
        	.ReDraw = False

			ggoSpread.SpreadLock	C_DUTY_MM	, -1, C_DUTY_MM
			ggoSpread.SSSetRequired	C_ACCUM_RATE, -1, -1  
			ggoSpread.SSSetProtected  .MaxCols   , -1, -1      	

        	.ReDraw = True
        End With
    End If
                
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	If gSelframeFlg = TAB1 Then
		With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired		C_PAY_GRD1		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PAY_GRD1_NM	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ADD_RATE		, pvStartRow, pvEndRow

		.vspdData.ReDraw = True
    
		End With
    Else    
		With frm1
    
		.vspdData1.ReDraw = False

		ggoSpread.SSSetRequired		C_DUTY_MM		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ACCUM_RATE	, pvStartRow, pvEndRow

		.vspdData1.ReDraw = True
    
		End With
	End If

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_PAY_GRD1		= iCurColumnPos(1)
			C_POPUP			= iCurColumnPos(2)
			C_PAY_GRD1_NM	= iCurColumnPos(3)
			C_ADD_RATE		= iCurColumnPos(4)      
    
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)			
			
			C_DUTY_MM		= iCurColumnPos(1)
			C_ACCUM_RATE	= iCurColumnPos(2)		
   
    End Select    
End Sub

<%'======================================================================================================
'	Name : OpenMajor()
'	Description : Major PopUp
'=======================================================================================================%>
Function OpenMajor()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "급호 팝업"			<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"				 		<%' TABLE 명칭 %>
	arrParam(0) = frm1.vspdData.Text			<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = "major_cd = " & FilterVar("H0001", "''", "S") & ""							<%' Where Condition%>
	arrParam(5) = "급호"			
	
    arrField(0) = "minor_cd"					<%' Field명(0)%>
    arrField(1) = "minor_nm"				<%' Field명(1)%>
    
    arrHeader(0) = "급호"						<%' Header명(0)%>
    arrHeader(1) = "급호명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.vspdData.Col = C_PAY_GRD1
		frm1.vspdData.action =0
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetMajor(Byval arrRet)
	With frm1
		.vspdData.Col = C_PAY_GRD1_NM
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_PAY_GRD1
		.vspdData.Text = arrRet(0)
		.vspdData.action =0
	End With
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화 
'=======================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
            
    gSelframeFlg = TAB1
    Call InitSpreadSheet("")                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
	Call changeTabs(TAB1)
    Call SetToolbar("1100110100101111")										<%'버튼 툴바 제어 %>
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 2   ' Tab의 갯수를 적어 주세요    
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
       
    Dim IntRetCD

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col  

	Select Case Col
	    Case C_PAY_GRD1
           	IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           	If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                Call DisplayMsgbox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다.
           		Frm1.vspdData.Col = C_PAY_GRD1_NM
           		frm1.vspdData.Text=""
            Else
    	    	frm1.vspdData.Col = C_PAY_GRD1_NM
            	frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
           	End If
    End Select   

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

    Dim IntRetCD
       
   	Frm1.vspdData1.Row = Row
   	Frm1.vspdData1.Col = Col

   	If Frm1.vspdData1.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData1.text) < CDbl(Frm1.vspdData1.TypeFloatMin) Then
         Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
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
		If lgStrPrevKey <> "" Then
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
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    '☜: 재쿼리 체크 %>
    	If lgStrPrevKey <> "" Then                  '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
      		Call DisableToolBar(parent.TBC_QUERY)
      		If DBQuery = False Then
      			Call RestoreToolBar ()
      			Exit Sub
      		End If
    	End If

    End if
    
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Private Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_POPUP Then
		    .Row = Row
		    .Col = C_PAY_GRD1

		    Call OpenMajor()        
    End If
    
    End With
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    		If IntRetCD = vbNo Then
      		Exit Function
    		End If
		End If
	Else
		ggoSpread.Source = frm1.vspdData1
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    		If IntRetCD = vbNo Then
      		Exit Function
    		End If
		End If
	End If
    
    Call ggoOper.ClearField(Document, "2")								
    Call InitVariables                                                      'Initializes local global variables
    															
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If

    If DbQuery = False Then

		Exit Function
	End If															'☜: Query db data
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'======================================================================================================
Function FncSave() 
    DIM iRow
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgbox("900001", "X", "X", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	Else
		ggoSpread.Source = frm1.vspdData1
		If ggoSpread.SSCheckChange = False Then
			Call DisplayMsgbox("900001", "X", "X", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	End If
    
    If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData
		If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
			Exit Function
		End If
	Else
		ggoSpread.Source = frm1.vspdData1
		If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '⊙: Check contents area
			Exit Function
		End If
	End If
    
    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0
           Select Case .Text
               Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
   	                .col = C_PAY_GRD1_NM
   	                if Trim(.text) = "" then    
	                    Call DisplayMsgbox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다십시요.
  	                    .Col = C_PAY_GRD1
  	                    Set gActiveElement = document.activeElement
                        Exit Function
                    End if 
			End Select
        Next
    End With
    
    If DbSave = False Then
		Exit Function
	End If			                                                <%'☜: Save db data%>
    
    FncSave = True                                                          
    
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================%>
Function FncCopy()
	If gSelframeFlg = TAB1 Then
        lgCurrentSpd = "M"

        If Frm1.vspdData.MaxRows < 1 Then
           Exit Function
        End If
	   
        With frm1.vspdData
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData	
				ggoSpread.CopyRow
                SetSpreadColor .ActiveRow, .ActiveRow
                
               	.Col = C_PAY_GRD1
				.Text = ""
				
				.Col = C_PAY_GRD1_NM
				.Text = ""                                 
				
				.ReDraw = True
    		    .Focus
			End If
		End With
	Else
	    lgCurrentSpd = "S"

        If Frm1.vspdData1.MaxRows < 1 Then
           Exit Function
        End If

        With frm1.vspdData1
			If .ActiveRow > 0 Then
				.focus
				.ReDraw = False
		
				ggoSpread.Source = frm1.vspdData1	
				ggoSpread.CopyRow
                SetSpreadColor .ActiveRow, .ActiveRow
                
				.Col = C_DUTY_MM
				.Text = ""
    
				.ReDraw = True
    		    .Focus
			End If
		End With
	End If

    Set gActiveElement = document.ActiveElement   
	
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel()
	If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	Else
		ggoSpread.Source = frm1.vspdData1	
		ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	End If
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

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
	
	If gSelframeFlg = TAB1 Then
	   lgCurrentSpd = "M"

		With frm1
	    .vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
        End With

    Else
	   lgCurrentSpd = "S"

		With frm1
	    .vspdData1.ReDraw = False
		.vspdData1.focus
		ggoSpread.Source = .vspdData1
        ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
        SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow - 1
		.vspdData1.ReDraw = True
        End With
	End If

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	   
    Set gActiveElement = document.ActiveElement   
	
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

	If gSelframeFlg = TAB1 Then	
		With frm1.vspdData 
    		.focus
    		ggoSpread.Source = frm1.vspdData 
    		lDelRows = ggoSpread.DeleteRow
		End With
	Else
		With frm1.vspdData1
    		.focus
    		ggoSpread.Source = frm1.vspdData1
    		lDelRows = ggoSpread.DeleteRow
		End With
	End If    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 '☜: 화면 유형 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      '☜:화면 유형, Tab 유무 
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

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    Select Case gActiveSpdSheet.id
		Case "vaSpread"
			Call InitSpreadSheet("A")            
		Case "vaSpread1"
			Call InitSpreadSheet("B")      		            
	End Select 

	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
	If gSelframeFlg = TAB1 Then
		ggoSpread.Source = frm1.vspdData	
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	Else
		ggoSpread.Source = frm1.vspdData1
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	End If
    FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================%>
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

	
	Dim strVal
    
    With frm1
    
    If gSelframeFlg = TAB1 Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						<%'현재 검색조건으로 Query%>
		strVal = strVal & "&txtFlag=" & "1"		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						<%'현재 검색조건으로 Query%>
		strVal = strVal & "&txtFlag=" & "2"		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
        
    End With
    
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()													<%'조회 성공후 실행로직 %>
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")									<%'This function lock the suitable field%>
	Call SetToolbar("110011110011111")										<%'버튼 툴바 제어 %>
	
End Function

'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbSave()     
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

	With frm1
		.txtMode.value = parent.UID_M0002
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 
	If gSelframeFlg = TAB1 Then
		.txtFlag.value = "1"
		For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag									    <%'☜: 신규 %>
				
				strVal = strVal & "C" & parent.gColSep					  			<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_PAY_GRD1	'1
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ADD_RATE	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
                
                strVal = strVal & "U" & parent.gColSep								<%'☜: C=Create, Row위치 정보 %>

                .vspdData.Col = C_PAY_GRD1	'1
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_ADD_RATE	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>

				strDel = strDel & "D" & parent.gColSep								<%'☜: D=Update, Row위치 정보 %>
				
                .vspdData.Col = C_PAY_GRD1	'1
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep									
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
		Next
		
	Else
		.txtFlag.value = "2"
		For lRow = 1 To .vspdData1.MaxRows
    
        .vspdData1.Row = lRow
        .vspdData1.Col = 0
        
        Select Case .vspdData1.Text

            Case ggoSpread.InsertFlag									    <%'☜: 신규 %>
				
				strVal = strVal & "C" & parent.gColSep					  			<%'☜: C=Create, Row위치 정보 %>

                .vspdData1.Col = C_DUTY_MM	'1
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ACCUM_RATE	'2
                strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
                
                strVal = strVal & "U" & parent.gColSep								<%'☜: C=Create, Row위치 정보 %>

                .vspdData1.Col = C_DUTY_MM	'1
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ACCUM_RATE	'2
                strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag										<%'☜: 삭제 %>

				strDel = strDel & "D" & parent.gColSep								<%'☜: D=Update, Row위치 정보 %>
				
                .vspdData1.Col = C_DUTY_MM	'1
                strDel = strDel & Trim(.vspdData1.Text) & parent.gRowSep									
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
		Next
	End If
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================%>
Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	If gSelframeFlg = TAB1 Then
		frm1.vspdData.MaxRows = 0
	Else
		frm1.vspdData1.MaxRows = 0
	End If
	Call MainQuery()    
End Function
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
    
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
    
End Sub

'======================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'======================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101011111")
    
    gMouseClickStatus = "SP1C"   

    Set gActiveSpdSheet = frm1.vspdData1
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If    
    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData1.MaxRows = 0 then
		exit sub
	end if
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub    

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
    
End Sub   

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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>직급별가산율</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>근속누진율</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/ha101ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/ha101ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxrows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlag" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
