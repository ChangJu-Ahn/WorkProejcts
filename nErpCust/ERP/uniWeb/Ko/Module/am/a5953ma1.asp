<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5953MA1
'*  4. Program Name         : a5953ma1
'*  5. Program Desc         : 월차결산 진행현황 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/09
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : park jai hong
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance
	
'========================================================================================================
Const BIZ_PGM_ID = "a5953mb1.asp"												'Biz Logic ASP
Const BIZ_PGM_JUMP_ID   = "a5955ma1"


'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 4877

Dim C_JOB_CD 
Dim C_JOB_NM 
Dim C_JOB_OPTION 
Dim C_JOB_START
Dim C_JOB_END 
Dim C_PROGRESS_FG 
Dim C_ERROR_NUM 
Dim C_ERROR_POP 


<%
Dim lsSvrDate
lsSvrDate = GetsvrDate
%>
'Dim IscookieSplit

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
Sub initSpreadPosVariables()         '1.2 변수에 Constants 값을 할당 
 C_JOB_CD = 1
 C_JOB_NM = 2																'Spread Sheet의 Column별 상수 
 C_JOB_OPTION = 3													
 C_JOB_START = 4
 C_JOB_END = 5
 C_PROGRESS_FG = 6
 C_ERROR_NUM = 7
 C_ERROR_POP = 8



	
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE												'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False													'⊙: Indicates that no value changed
	lgIntGrpCount     = 0														'⊙: Initializes Group View Size
    lgStrPrevKey      = ""														'⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""														'⊙: initializes Previous Key Index
    lgSortKey         = 1														'⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom("<%=lsSvrDate%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.txtYyyymm.Year	= strYear
	frm1.txtYyyymm.Month	= strMonth
	frm1.txtYyyymm.Day	= strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat,2)
	frm1.txtYyyymm.focus

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
'    Dim iRow,iCol
	Dim strYear,strMonth,strDay
	Dim TempFrDt
   '------ Developer Coding part (Start ) --------------------------------------------------------------       
   With Frm1                      	
		
	Select Case Kubun		
	Case 0
				
			If ReadCookie("JumpFlag")	<>""	Then .txtJumpFlag.Value		= ReadCookie("JumpFlag")
			
			If UCase(Trim(.txtJumpFlag.Value)) = "A5953MA1" Then
				If ReadCookie("FrYYYYMM")	<>"" Then TempFrDt				= ReadCookie("FrYYYYMM")				
				Call ExtractDateFrom(TempFrDt, Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)
				.txtYyyymm.Year	= strYear
				.txtYyyymm.Month	= strMonth
							
			
				If ReadCookie("Unt_Code")	<>"" Then .txtMajorCd.value 			= ReadCookie("Unt_Code") 
						
				If Trim(.txtYyyymm.Text) <> "" and Trim(.txtMajorCd.value) <> ""  Then
					Call MainQuery()
      			End If
      		End If
      		
      		WriteCookie "FrYYYYMM", ""
      		WriteCookie "Unt_Code" ,""
      		WriteCookie "JumpFlag", ""
      		
     Case 1			
   		
		    WriteCookie "FrYYYYMM" , UniConvYYYYMMDDToDate(Parent.gDateFormat,Trim(.txtYyyymm.Year),Right("0" & Trim(.txtYyyymm.Month),2),"01")
		    WriteCookie "Unt_Code" , .txtMajorCd.value
		    WriteCookie "JumpFlag" , "a5955ma1"
	End Select 		
    
	End With
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'	Desc : 화면이동 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 계속하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

'   Call CookiePage(strPgmId)
   Call PgmJump(strPgmId)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

  Dim strYYYYMM
  Dim strYear,strMonth,strDay 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

'	strYYYYMM = strYear & strMonth
'	lgKeyStream       = strYYYYMM & Parent.gColSep                                           'You Must append one character(Parent.gColSep)

    strYYYYMM = strYear & Parent.gServerDateType & strMonth    
    lgKeyStream = strYYYYMM & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtMajorCd.value & Parent.gColSep                      '사업장코드     

   '------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_ERROR_POP + 1                                                  ' ☜:☜: Add 1 to Maxcols
	    .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
        .ColHidden = True           
       
		ggoSpread.Source= frm1.vspdData
    ggoSpread.ClearSpreadData

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
       call GetSpreadColumnPos("A")
       
       Call AppendNumberPlace("6","2","0")
       ggoSpread.SSSetEdit  C_JOB_CD ,      "작업코드"            ,15,,, 20
       ggoSpread.SSSetEdit  C_JOB_NM ,      "작업명"              ,15,,, 20
       ggoSpread.SSSetEdit  C_JOB_OPTION ,       "작업구분"       ,15,,, 20
       ggoSpread.SSSetEdit  C_JOB_START ,    "시작시간"           ,20,2,, 30
       ggoSpread.SSSetEdit  C_JOB_END ,      "종료시간"           ,20,2,, 30
       ggoSpread.SSSetEdit  C_PROGRESS_FG ,          "작업상태"   ,10,,, 20
       ggoSpread.SSSetFloat C_ERROR_NUM,      "ERROR COUNT",   20,3,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       'ggoSpread.SSSetEdit  C_ERROR_NUM ,      "ERROR 수"         ,15,1,, 20
       ggoSpread.SSSetButton  C_ERROR_POP
       Call ggoSpread.SSSetColHidden(C_ERROR_NUM,C_ERROR_NUM,True,"S")
       Call ggoSpread.SSSetColHidden(C_ERROR_POP,C_ERROR_POP,True,"S")
	   call ggoSpread.MakePairsColumn(C_JOB_CD,C_JOB_NM)		   
	   call ggoSpread.MakePairsColumn(C_JOB_START,C_JOB_END)		   

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
         ggoSpread.SpreadLock      C_JOB_CD , -1, C_JOB_CD
         ggoSpread.SpreadLock      C_JOB_NM , -1, C_JOB_NM
	     ggoSpread.SpreadLock      C_JOB_OPTION , -1, C_JOB_OPTION
	     ggoSpread.SpreadLock      C_JOB_START , -1, C_JOB_START
	     ggoSpread.SpreadLock      C_JOB_END , -1, C_JOB_END
	     ggoSpread.SpreadLock      C_PROGRESS_FG , -1, C_PROGRESS_FG
	     ggoSpread.SpreadLock      C_ERROR_NUM , -1, C_ERROR_NUM
	     ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols
    .vspdData.ReDraw = True
    

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pStartRow,ByVal pEndRow)
    With frm1
    
    .vspdData.ReDraw = False
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

			C_JOB_CD         =  iCurColumnPos(1)
			C_JOB_NM         =  iCurColumnPos(2)
			C_JOB_OPTION  =  iCurColumnPos(3)
			C_JOB_START  =  iCurColumnPos(4)
			C_JOB_END        =  iCurColumnPos(5)
			C_PROGRESS_FG        =  iCurColumnPos(6)
			C_ERROR_NUM            =  iCurColumnPos(7)
			C_ERROR_POP     =  iCurColumnPos(8)
			
    End Select    
End Sub



'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100000000001111")
	Call CookiePage(0)
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
'    Call SetDefaultVal
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    If DbQuery = False Then
        Exit Function
    End If
       
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
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call MakeKeyStream("X")
    
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
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

	With Frm1.VspdData
           '.Col  = C_MAJORCD
           '.Row  = .ActiveRow
           '.Text = ""
    End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	
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
Function FncInsertRow() 

Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
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
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                  '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                   '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncPasteRepeatedSpreadData
' Function Desc : 
'========================================================================================
'5.1. [FncPasteRepeatedSpreadData]를 신규로 추가 
Function FncPasteRepeatedSpreadData()			
    ggoSpread.Source = gActiveSpdSheet   
    ggoSpread.CPasteRepeatedSpreadData 
End Function

'========================================================================================
' Function Name : FncSaveSpreadInf
' Description   : 
'========================================================================================
Sub FncSaveSpreadInf()                              '11. user의 변경(컬럼이동,정렬,타이틀변경등)을 저장, 복원하는 함수 
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadInf()
End Sub

'========================================================================================
' Function Name : FncResetSpreadInf
' Description   : 
'========================================================================================
Sub FncResetSpreadInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ResetSpreadInf()
	Call initSpreadPosVariables()      
    Call InitSpreadSheet           
	Call InitVariables
	Call InitComboBox
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
End Sub

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear																			'☜: Clear err status

	if LayerShowHide(1) = False then
	   Exit Function
	end if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)													  '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                            '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
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
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")										    '버튼 툴바 제어 
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
   Call DisableToolBar(Parent.TBC_QUERY)
   If DBQuery = false Then
	    Call RestoreToolBar()
   	    Exit Function
   	   End If

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
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		    	<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"		<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtMajorCd.value	    <%' Code Condition%>
	arrParam(3) = "" 		            		<%' Name Cindition%>
	arrParam(4) = ""					<%' Where Condition%>
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
    arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>
    
    arrHeader(0) = "사업장코드"				<%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMajorCd.focus
		Exit Function
	Else
		Call SetMajor(arrRet)
	End If	

End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetMajor(Byval arrRet)
	With frm1
		.txtMajorCd.focus
		.txtMajorCd.value = arrRet(0)
		.txtMajorName.value = arrRet(1)		
	End With
End Function



'=======================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtYyyymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtYyyymm.Focus
     End If
End Sub
'=======================================================================================================
'   Event Name : txtYyyymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtYyyymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub


'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 21, 22, 23     ' 학교 
		        .vspdData.Col = C_SCHOOL
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_SCHOOL_NM
		    	.vspdData.text = arrRet(1)   
		    Case 31, 32         ' 전공 
		        .vspdData.Col = C_MAJOR
		    	.vspdData.text = arrRet(0) 
		    	.vspdData.Col = C_MAJOR_NM
		    	.vspdData.text = arrRet(1)   
        End Select

	End With

End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
Sub txtEnd_dt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtEnd_dt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtEnd_dt.Focus
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub



'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
    		End If
    	End If
    End if
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")    

    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If

	If Row < 1 Then Exit Sub
	
	    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    

'	IscookieSplit = ""

'    frm1.vspdData.Col = C_EMP_NO
'    frm1.vspdData.Row = Row
'	IscookieSplit = frm1.vspdData.text

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
'3.1 SpreadSheet의 이벤트 [ScriptDragDropBlock]을 추가 
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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


'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Focus 이동 
'========================================================================================================
Sub vspdData_LeaveCell(Col, Row, NewCol, NewRow, Cancel)
    
 '   frm1.vspdData.OperationMode = 3             
	
End Sub


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col'
	Select Case Col'
    Case C_ERROR_POP
            Call OpenCode("", C_ERROR_POP, Row)
    End Select    
End Sub


'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim yyyymm, Job_cd, strYYYYMM
	Dim strYear,strMonth,strDay 
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth   
	yyyymm = FilterVar(strYYYYMM, "''", "S")

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_JOB_CD
	Job_cd = FilterVar(frm1.vspdData.Text, "''", "S")
	
'	MsgBox yyyymm & " ### " & Job_cd

	Select Case iWhere
	    Case C_ERROR_POP
	        arrParam(0) = "ERROR목록팝업"										' 팝업 명칭 
	    	arrParam(1) = "A_JOB_ERROR"													' TABLE 명칭 
	    	arrParam(2) = strCode                   								' Code Condition
	    	arrParam(3) = ""														' Name Cindition
	    	arrParam(4) = " YYYYMM = " & yyyymm & " and JOB_CD = " & Job_cd         ' Where Condition

	    	arrParam(5) = "ERROR목록" 											' TextBox 명칭 
	
	    	arrField(0) = "ERROR_SEQ"														' Field명(0)	    	
	    	arrField(1) = "ERROR_CONTENTS"														' Field명(0)	    	
	    	
	    	arrHeader(0) = "ERROR_SEQ"	   		    							' Header명(0)	    	    
	    	arrHeader(1) = "ERROR설명"	   		    							' Header명(1)	    	
	End Select
    
'	arrRet = window.showModalDialog("../GB/GCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
'		"dialogWidth=615px; dialogHeight=450px; center: Yes; help: No; resizable: YES; status: No;")

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else

       	ggoSpread.Source = frm1.vspdData

	End If	

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>월차결산진행현황</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>결산년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5953ma1_fpLoanDtFr_txtYyyymm.js'></script>&nbsp;</TD>
			            		<TD CLASS="TD5" NOWRAP>사업장</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtMajorCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup()">
									<INPUT TYPE="Text" NAME="txtMajorName" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="사업장명">
			            		</TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5953ma1_vaSpread1_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
					    <TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpChk(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage(1)">월차결산작업</a></TD>
						<TD WIDTH=10>&nbsp;</TD>

					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>

		</TD>
		
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtJumpFlag"		TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

