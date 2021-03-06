<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 직급별인상율등록 
*  3. Program ID           : Ha201ma1
*  4. Program Name         : Ha201ma1
*  5. Program Desc         : 퇴직정산관리/직급별인상율등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID = "Ha201mb1.asp"                                      'Biz Logic ASP 

Const CookieSplit = 1233
Dim C_PAY_GRD1
Dim C_PAY_GRD1_POP
Dim C_PAY_GRD1_NM
Dim C_PULL_RATE
Dim C_PULL_AMT

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_PAY_GRD1		= 1															<%'Spread Sheet의 Column별 상수 %>
	 C_PAY_GRD1_POP = 2
	 C_PAY_GRD1_NM	= 3
	 C_PULL_RATE	= 4
	 C_PULL_AMT		= 5	  
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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

    Dim strYear,strMonth,strDay

    frm1.txtRetire_yy.Focus
    Call ggoOper.FormatDate(frm1.txtRetire_yy, parent.gDateFormat, 3)
	Call ExtractDateFrom(UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat),parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)    
    frm1.txtRetire_yy.Year	= strYear
    frm1.txtRetire_yy.Month = strMonth
	frm1.txtRetire_yy.Day	= strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    
    Dim strYear,strMonth,strDay

    lgKeyStream = frm1.txtRetire_yy.YEAR  & parent.gColSep       'You Must append       
	lgKeyStream = lgKeyStream & Frm1.txtpay_grd1.value & parent.gColSep
End Sub

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

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_PULL_AMT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0
	
	    Call AppendNumberPlace("6","3","2")
        Call  GetSpreadColumnPos("A")

        ggoSpread.SSSetEdit   C_PAY_GRD1,     "직급",    10,,, 2,2
        ggoSpread.SSSetButton C_PAY_GRD1_POP
        ggoSpread.SSSetEdit   C_PAY_GRD1_NM,  "직급명",  44,,, 20,2
        ggoSpread.SSSetFloat  C_PULL_RATE,    "인상율",  30, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat  C_PULL_AMT,     "인상액",  30, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        
        Call ggoSpread.MakePairsColumn(C_PAY_GRD1,  C_PAY_GRD1_POP)
	
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
    ggoSpread.SpreadLock C_PAY_GRD1		, -1, C_PAY_GRD1
    ggoSpread.SpreadLock C_PAY_GRD1_POP	, -1, C_PAY_GRD1_POP
    ggoSpread.SpreadLock C_PAY_GRD1_NM	, -1, C_PAY_GRD1_NM
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1          
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_PAY_GRD1		, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PAY_GRD1_NM	, pvStartRow, pvEndRow
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_PAY_GRD1		= iCurColumnPos(1)															<%'Spread Sheet의 Column별 상수 %>
			C_PAY_GRD1_POP	= iCurColumnPos(2)
			C_PAY_GRD1_NM	= iCurColumnPos(3)
			C_PULL_RATE		= iCurColumnPos(4)
			C_PULL_AMT		= iCurColumnPos(5)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call SetDefaultVal
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    
	Call CookiePage (0)                                                             '☜: Check Cookie
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

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    if txtPay_grd1_Onchange() then
		exit function
	end if 

	ggoSpread.ClearSpreadData	
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If                                                                '☜: Query db data
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
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
    Dim dblPull_rate
    Dim dblPull_amt
    Dim lRow

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgbox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
               Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_PAY_GRD1_NM               
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					Call DisplayMsgBox("970000", "X","직급코드","x")
						Exit Function
					end if					
                    .vspdData.Col = C_PULL_RATE
                    if  .vspdData.Text <> "" then
                        dblPull_rate = UNICDbl(.vspdData.Text)
                    else
                        .vspdData.Text = 0
                        dblPull_rate = 0
                    end if

                    .vspdData.Col = C_PULL_AMT
                    if  .vspdData.Text <> "" then
                        dblPull_amt = UNICDbl(.vspdData.Text)
                    else
                        .vspdData.Text = 0
                        dblPull_amt = 0
                    end if

                    if  dblPull_rate = 0 and dblPull_amt = 0 then
                        Call DisplayMsgbox("970027","x","인상율/인상액","x")
                        .vspdData.Col = C_PULL_RATE
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    end if

                    if  dblPull_rate >= 0 and dblPull_rate <= 100 then
                    else
                        Call DisplayMsgbox("970027","x","인상율","x")
                        .vspdData.Col = C_PULL_RATE
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    end if

            End Select
        Next
    End With

    Call MakeKeyStream("X")
    If DbSave = False Then
		Exit Function
	End If			                                                    '☜: Save db data
    
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
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			 SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
	       .Col  = C_PAY_GRD1
           .Row  = .ActiveRow
           .Text = ""

           .Col  = C_PAY_GRD1_NM
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
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow
	Dim iRow

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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
           .vspdData.Row  = iRow
           .vspdData.Col  = C_PULL_RATE
           .vspdData.value = 0

           .vspdData.Col  = C_PULL_AMT
           .vspdData.value = 0           
        Next
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
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
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

	 if LayerShowHide(1) = false then
	    Exit Function
	end if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
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
    
     if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & .txtretire_yy.Year & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PULL_RATE	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PULL_AMT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & .txtretire_yy.Year & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PULL_RATE	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PULL_AMT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & .txtretire_yy.Year & parent.gColSep
                    .vspdData.Col = C_PAY_GRD1	: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    If DbDelete = False Then
		Exit Function
	End If															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode = parent.OPMD_UMODE    
    
	Call SetToolbar("110011110011111")
    
    Call InitData()
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	
	frm1.vspdData.focus
    lgBlnFlgChgValue = False
	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData	
	ggoSpread.ClearSpreadData	
    
    Call InitVariables															'⊙: Initializes local global variables
	call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
'	Name : OpenMajor()
'	Description : Major PopUp
'========================================================================================================
Function OpenMajor(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "직급코드 팝업"			' 팝업 명칭 
	arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	If  iWhere = 0 then
	    arrParam(2) = frm1.txtpay_grd1.value	' Code Condition
	else
	    frm1.vspddata.Col = C_PAY_GRD1
	    arrParam(2) = frm1.vspddata.Text	    ' Code Condition
	end if
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""			' Where Condition
	arrParam(5) = "직급코드"			
	
    arrField(0) = "minor_cd"					' Field명(0)
    arrField(1) = "minor_nm"				    ' Field명(1)
    
    arrHeader(0) = "직급코드"		        ' Header명(0)
    arrHeader(1) = "직급코드명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    if  iWhere = 0 then
		    frm1.txtpay_grd1.focus
		else
		    frm1.vspddata.Col = C_PAY_GRD1
			frm1.vspddata.action =0
		end if	
		Exit Function
	Else
		Call SetMajor(iWhere,arrRet)
	End If	

End Function

'========================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetMajor(Byval iWhere, Byval arrRet)
	With frm1
	    if  iWhere = 0 then
		    .txtpay_grd1.value = arrRet(0)
		    .txtpay_grd1_Nm.value = arrRet(1)
		    .txtpay_grd1.focus
		else
		    .vspddata.Col = C_PAY_GRD1_NM
	        .vspddata.Text = arrRet(1)
		    .vspddata.Col = C_PAY_GRD1
	        .vspddata.Text = arrRet(0)
			.vspddata.action =0
		end if
	End With
End Function

'========================================================================================================
'   Event Name : txtpay_grd1_change
'   Event Desc :
'========================================================================================================
function txtpay_grd1_Onchange()
    Dim IntRetCd
    txtpay_grd1_Onchange = false   
    If frm1.txtpay_grd1.value = "" Then
		frm1.txtpay_grd1_nm.value = ""
    Else
        IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " AND minor_cd =  " & FilterVar(frm1.txtpay_grd1.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgbox("800008","X","X","X")
			frm1.txtpay_grd1_nm.value = ""
            frm1.txtpay_grd1.focus
            Set gActiveElement = document.ActiveElement    
			txtpay_grd1_Onchange = true            
        Else
			frm1.txtpay_grd1_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
            Case C_PAY_GRD1
            IntRetCd = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " AND minor_cd= " & FilterVar(Frm1.vspdData.Value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If  IntRetCd = false then
		    	Call DisplayMsgbox("800008","X","X","X")
                Frm1.vspdData.Col = C_PAY_GRD1_NM
		    	Frm1.vspdData.Value = ""
                Frm1.vspdData.Action = 0 ' go to 
            Else
                Frm1.vspdData.Col = C_PAY_GRD1_NM
		    	Frm1.vspdData.Value = Trim(Replace(lgF0,Chr(11),""))
            End if 
    End Select    
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row     
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub

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
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	if  Col = C_PAY_GRD1_POP then
        Call OpenMajor(1)
    End if

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

Sub txtRetire_yy_DblClick(Button)
	If Button = 1 Then
		frm1.txtRetire_yy.Action = 7
        Call SetFocusToDocument("M")
        Frm1.txtRetire_yy.Focus
	End If
End Sub


Sub txtRetire_yy_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>직급별인상율등록</font></td>
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
							    	<TD CLASS=TD5>퇴직추계연도</TD>
							    	<TD CLASS=TD6><script language =javascript src='./js/ha201ma1_fpDateTime2_txtRetire_yy.js'></script></TD>
							        <TD CLASS="TD5" NOWRAP>직급</TD>
							        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtPay_grd1" MAXLENGTH=2 SIZE=10 tag=11XXXU ALT="직급코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMajor(0)">
							                               <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtPay_grd1_nm" MAXLENGTH=30 SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X" ALT="직급코드명"></TD>
							    </TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/ha201ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 <%=BizSize%> FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

