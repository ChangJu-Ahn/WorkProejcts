<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h9101ma1
*  4. Program Name         : h9101ma1
*  5. Program Desc         : 연말정산관리/연말정산/현장기술인력소득공제등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/01
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h9103mb1.asp"                                      'Biz Logic ASP 
Const C_TECH_STRT = 1															<%'Spread Sheet의 Column별 상수 %>
Const C_TECH_END = 2
Const C_DEDUCT_RATE = 3
Const C_DEDUCT_AMT = 4

Const C_SHEETMAXROWS =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

<%
StartDate = DateSerial(Year(Date),Month(Date),1)
StartDate = Year(StartDate) & @@@gComDateType & Right("0" & Month(StartDate),2) & @@@gComDateType & Right("0" & Day(StartDate),2) ' Start date
EndDate   = Year(Date)      & @@@gComDateType & Right("0" & Month(Date),2)      & @@@gComDateType & Right("0" & Day(Date),2)      ' End date
%>

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = @@@OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    	lgBlnFlgChgValue = False
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(@@@gCurrency, "I", "*") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	With frm1.vspdData
	
       .MaxCols = C_DEDUCT_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

       .MaxRows = 0
        @@@ggoSpread.Source = frm1.vspdData

	   .ReDraw = false
	
       Call @@@AppendNumberPlace("6","2","0")
       Call @@@AppendNumberPlace("7","5","0")
       @@@ggoSpread.Spreadinit
       @@@ggoSpread.SSSetFloat    C_TECH_STRT,    "근속시작개월" ,     30,"7",@@@ggStrIntegeralPart, @@@ggStrDeciPointPart,@@@gComNum1000,@@@gComNumDec
       @@@ggoSpread.SSSetFloat    C_TECH_END,     "근속종료개월" ,     30,"7",@@@ggStrIntegeralPart, @@@ggStrDeciPointPart,@@@gComNum1000,@@@gComNumDec
       @@@ggoSpread.SSSetFloat    C_DEDUCT_RATE , "공제율(%)" ,        28, "6",@@@ggStrIntegeralPart, @@@ggStrDeciPointPart,@@@gComNum1000,@@@gComNumDec,,,"Z"
       @@@ggoSpread.SSSetFloat    C_DEDUCT_AMT,   "공제액" ,           28,@@@ggAmtOfMoneyNo,@@@ggStrIntegeralPart, @@@ggStrDeciPointPart,@@@gComNum1000,@@@gComNumDec,,,"Z"
	 
	   .ReDraw = true

       Call SetSpreadLock 
    
    End With
    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
      @@@ggoSpread.SpreadLock      C_TECH_STRT , -1, C_TECH_STRT
      @@@ggoSpread.SSSetRequired	C_TECH_END, -1, -1
      @@@ggoSpread.SSSetRequired	C_DEDUCT_RATE, -1, -1
      @@@ggoSpread.SSSetRequired	C_DEDUCT_AMT, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(lRow)
    With frm1
    
    .vspdData.ReDraw = False
      @@@ggoSpread.SSSetRequired    C_TECH_STRT , lRow, lRow
      @@@ggoSpread.SSSetRequired    C_TECH_END , lRow, lRow
      @@@ggoSpread.SSSetRequired    C_DEDUCT_RATE , lRow, lRow
      @@@ggoSpread.SSSetRequired    C_DEDUCT_AMT , lRow, lRow
      
    .vspdData.ReDraw = True
    
    End With
End Sub
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,@@@gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> @@@UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
    
    Call @@@ggoOper.FormatField(Document, "1",@@@ggStrIntegeralPart, @@@ggStrDeciPointPart,@@@gDateFormat,@@@gComNum1000,@@@gComNumDec)
	Call @@@ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetDefaultVal
    Call Parent.MASetToolBar("1100111100101111")										        '버튼 툴바 제어 
    
	Call CookiePage (0)
	
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    @@@ggoSpread.Source = Frm1.vspdData
    If @@@ggoSpread.SSCheckChange = True Then
		IntRetCD = @@@DisplayMsgBox("900013", @@@VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call @@@ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call SetDefaultVal

    If DbQuery = False Then  
		Exit Function
	End If
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
   	Dim lRow
   	Dim intTechStrt
   	Dim intTechEnd
   	
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    @@@ggoSpread.Source = frm1.vspdData
    If @@@ggoSpread.SSCheckChange = False Then
        IntRetCD = @@@DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    @@@ggoSpread.Source = frm1.vspdData
    If Not @@@ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case @@@ggoSpread.InsertFlag, @@@ggoSpread.UpdateFlag

   	                .vspdData.Col = C_TECH_STRT
                    intTechStrt = .vspdData.value
                    
   	                .vspdData.Col = C_TECH_END
                    intTechEnd = .vspdData.value
                    
                    If Cdbl(intTechStrt) > Cdbl(intTechEnd) then
	                    Call @@@DisplayMsgBox("970024","X","근속시작개월","근속종료개월")	'근속시작개월은 근속종료개월보다 작아야합니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_TECH_END
                        .vspdData.value = ""
  	                    .vspdData.Col = C_TECH_STRT
                        .vspdData.value = ""
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    Else
                    End if 
            End Select
        Next
	End With

    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    @@@ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            @@@ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
           .Col  = C_TECH_STRT
           .Row  = .ActiveRow
           .Text = ""
    End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    @@@ggoSpread.Source = Frm1.vspdData	
    @@@ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        @@@ggoSpread.Source = .vspdData
        @@@ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow

       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	@@@ggoSpread.Source = frm1.vspdData 
    	lDelRows = @@@ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(@@@C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(@@@C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    @@@ggoSpread.Source = frm1.vspdData	
    If @@@ggoSpread.SSCheckChange = True Then
		IntRetCD = @@@DisplayMsgBox("900016", @@@VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & @@@UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = @@@OPMD_UMODE Then
    Else
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case @@@ggoSpread.InsertFlag                                      '☜: Update추가 
                                                       strVal = strVal & "C" & @@@gColSep 'array(0)
                                                       strVal = strVal & lRow & @@@gColSep
                    .vspdData.Col = C_TECH_STRT	: strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_TECH_END  : strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_DEDUCT_RATE	        : strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_DEDUCT_AMT     : strVal = strVal & Trim(.vspdData.Value) & @@@gRowSep
                    lGrpCnt = lGrpCnt + 1 
               
               Case @@@ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & @@@gColSep
                                                      strVal = strVal & lRow & @@@gColSep
                    .vspdData.Col = C_TECH_STRT	     : strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_TECH_END	     : strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_DEDUCT_RATE	 : strVal = strVal & Trim(.vspdData.Value) & @@@gColSep
                    .vspdData.Col = C_DEDUCT_AMT     : strVal = strVal & Trim(.vspdData.Value) & @@@gRowSep
                    lGrpCnt = lGrpCnt + 1
               
               Case @@@ggoSpread.DeleteFlag                                      '☜: Delete

                                                      strDel = strDel & "D" & @@@gColSep
                                                      strDel = strDel & lRow & @@@gColSep
                    .vspdData.Col = C_TECH_STRT    : strDel = strDel & Trim(.vspdData.Value) & @@@gRowSep	'삭제시 key만								
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = @@@UID_M0002
       .txtUpdtUserId.value  = @@@gUsrID
       .txtInsrtUserId.value = @@@gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> @@@OPMD_UMODE Then                                      'Check if there is retrived data
        Call @@@DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = @@@DisplayMsgBox("900003", @@@VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = @@@OPMD_UMODE    
    Call @@@ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call Parent.MASetToolBar("1100111100111111")									
	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call @@@ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
	call DBQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case C_DILIG_POP
	        arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HCA010T"				 		' TABLE 명칭 
	        arrParam(2) = ""		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = ""							' Where Condition
	        arrParam(5) = "근태코드"			    ' TextBox 명칭 
	
            arrField(0) = "dilig_cd"					' Field명(0)
            arrField(1) = "dilig_nm"				    ' Field명(1)
    
            arrHeader(0) = "근태코드"				' Header명(0)
            arrHeader(1) = "근태코드명"			    ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	@@@ggoSpread.Source = frm1.vspdData
        @@@ggoSpread.UpdateRow Row
	End If	

End Function
'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_DILIG_POP
		        .vspdData.Col = C_DILIG_CD
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_DILIG_NM
		    	.vspdData.text = arrRet(1)
        End Select

	End With

End Function





'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
   

	End With

   	@@@ggoSpread.Source = frm1.vspdData
    @@@ggoSpread.UpdateRow Row

End Sub


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col

    End Select    
End Sub



'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
   	If Frm1.vspdData.CellType = @@@SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	@@@ggoSpread.Source = frm1.vspdData
    @@@ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
     @@@gMouseClickStatus = "SPC"   
'    If Row = 0 Then
'        @@@ggoSpread.Source = frm1.vspdData
'        If lgSortKey = 1 Then
'            @@@ggoSpread.SSSort
'            lgSortKey = 2
'        Else
'            @@@ggoSpread.SSSort ,lgSortKey
'            lgSortKey = 1
'        End If
'    Else
  '  	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'	lsConcd = frm1.vspdData.Text		
 '   End If
 '   
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And @@@gMouseClickStatus = "SPC" Then
       @@@gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
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

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG src="../../image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>현장기술인력소득공제</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG src="../../image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h9103ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
