<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5957MA1
'*  4. Program Name         : 유가증권정보조회 
'*  5. Program Desc         : 유가증권정보조회 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/25
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : LEE KANG YOUNG
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


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>


<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'========================================================================================================

Const BIZ_PGM_ID      = "a5959mb1.asp"						           '☆: Biz Logic ASP Name


'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_SECURITY_CD		
Dim C_SECURITY_NM		
Dim C_SECURITY_TYPE		
Dim C_DOC_CUR           
Dim C_BUY_AMT           
Dim C_LOC_BUY_AMT       
Dim C_PRICE_AMT         
Dim C_LOC_PRICE_AMT		
Dim C_CNT				
Dim C_PRICE_SUM         
Dim C_LOC_PRICE_SUM		
Dim C_CALCU_YN			
Dim C_GL_NO_YN			
    


Const COOKIE_SPLIT      = 4877	                                      'Cookie Split String

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop

'========================================================================================================
Sub InitSpreadPosVariables()
	C_SECURITY_CD		= 1                                                  'Column ant for Spread Sheet 
	C_SECURITY_NM		= 2
	C_SECURITY_TYPE		= 3
	C_DOC_CUR           = 4
	C_BUY_AMT           = 5
	C_LOC_BUY_AMT       = 6
	C_PRICE_AMT         = 7
	C_LOC_PRICE_AMT		= 8
	C_CNT				= 9
	C_PRICE_SUM         = 10
	C_LOC_PRICE_SUM		= 11
	C_CALCU_YN			= 12
	C_GL_NO_YN			= 13
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"
			strTemp = ReadCookie("SecuCode")
			Call WriteCookie("SecuCode", "")
			
			If strTemp = "" then Exit Function

			frm1.txtSecurityCd.value = strTemp
	
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("SecuCode", "")
				Exit Function 
			End If
					
			Call MainQuery()
		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   Dim rdoGiFlagT, rdoYiFlagT
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
	If frm1.rdoGiFlag(0).checked Then
		rdoGiFlagT="A"
	ELSEIF frm1.rdoGiFlag(1).checked THEN
		rdoGiFlagT="Y"
	ELSEIF frm1.rdoGiFlag(2).checked THEN
		rdoGiFlagT="N"
	End If
   
	If frm1.rdoYiFlag(0).checked Then
	 	rdoYiFlagT="A"
	Elseif frm1.rdoYiFlag(1).checked Then
		rdoYiFlagT="Y"
	Elseif frm1.rdoYiFlag(2).checked Then
		rdoYiFlagT="N"	
	End if		
   
    lgKeyStream       = Frm1.txtSecurityCd.Value								& Parent.gColSep       'You Must append one character(Parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtSecurity_TypeCd.value				& Parent.gColSep                    
	lgKeyStream       = lgKeyStream & frm1.txtMajorCd.value						& Parent.gColSep                     
	lgKeyStream       = lgKeyStream & rdoGiFlagT								& Parent.gColSep     
	lgKeyStream       = lgKeyStream & rdoYiFlagT								& Parent.gColSep                           
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021103",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_GL_NO_YN + 1                                                  ' ☜:☜: Add 1 to Maxcols
	    .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
        .ColHidden = True           
       
		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A")
        
'       Call AppendNumberPlace("6","2","0")
                              'ColumnPosition   Header			Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
       ggoSpread.SSSetEdit   C_SECURITY_CD	,	"증권코드"    ,15       ,               ,      ,20
       ggoSpread.SSSetEdit   C_SECURITY_NM    ,"증권명"      ,18       ,               ,      ,30
	   ggoSpread.SSSetEdit   C_SECURITY_TYPE    ,"증권종류"  ,18       ,               ,      ,30
	   ggoSpread.SSSetEdit   C_DOC_CUR          ,"거래통화"  ,10       ,               ,      ,15        ,2
	   ggoSpread.SSSetFloat  C_BUY_AMT          , "취득금액"    ,15  ,"A"   ,ggStrIntegeralPart   ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	   ggoSpread.SSSetFloat  C_LOC_BUY_AMT      , "취득금액(자국)"  ,15  ,Parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart  ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec 
	   ggoSpread.SSSetFloat  C_PRICE_AMT        , "액면금액"      ,15  ,"A"    ,ggStrIntegeralPart  ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetFloat  C_LOC_PRICE_AMT    , "액면금액(자국)"  ,15  ,Parent.ggAmtOfMoneyNo   ,ggStrIntegeralPart  ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetFloat  C_CNT    , "매수"   ,15  ,Parent.ggQtyNo   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetFloat  C_PRICE_SUM        ,"총금액"    ,15   ,"A"  ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec       
       ggoSpread.SSSetFloat  C_LOC_PRICE_SUM    ,"총금액(자국)"    ,15   ,Parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
       ggoSpread.SSSetEdit   C_CALCU_YN    ,"이자"   ,10       ,                  ,      ,18         ,2
       ggoSpread.SSSetEdit   C_GL_NO_YN     ,"승인"    ,10         ,              ,     ,18        ,2                         

	   .ReDraw = true
	   
	   call ggoSpread.MakePairsColumn(C_SECURITY_CD,C_SECURITY_TYPE)
	   call ggoSpread.MakePairsColumn(C_BUY_AMT,C_LOC_BUY_AMT)
	   call ggoSpread.MakePairsColumn(C_PRICE_AMT,C_LOC_PRICE_AMT)
	   call ggoSpread.MakePairsColumn(C_PRICE_SUM,C_LOC_PRICE_SUM)
	   
	
       Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1    
       .vspdData.ReDraw = False
       ggoSpread.SSSetRequired    C_SID      , pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_SNm      , pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_AddressNm, pvStartRow, pvEndRow
       .vspdData.ReDraw = True
    End With
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
			C_SECURITY_CD			= iCurColumnPos(1)
			C_SECURITY_NM			= iCurColumnPos(2)
			C_SECURITY_TYPE			= iCurColumnPos(3)    
			C_DOC_CUR				= iCurColumnPos(4)
			C_BUY_AMT				= iCurColumnPos(5)
			C_LOC_BUY_AMT			= iCurColumnPos(6)
			C_PRICE_AMT				= iCurColumnPos(7)
			C_LOC_PRICE_AMT			= iCurColumnPos(8)
			C_CNT					= iCurColumnPos(9)
			C_PRICE_SUM				= iCurColumnPos(10)
			C_LOC_PRICE_SUM			= iCurColumnPos(11)
			C_CALCU_YN				= iCurColumnPos(12)
			C_GL_NO_YN				= iCurColumnPos(13)
    End Select    
    
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
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


'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
	Call InitVariables
    Call SetDefaultVal
	
	Call SetToolbar("1100000000001111")     
	Call CookiePage("FORM_LOAD")                                            '☆: Developer must customize
	frm1.txtSecurityCd.focus	
'	Call CookiePage (0)                                                              '☜: Check Cookie
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															  '☜: Processing is NG
    Err.Clear                                                                     '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			          '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										  '⊙: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables															  '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									          '⊙: This function check indispensable field
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery = False Then                                                       '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
  
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
   
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
    On Error Resume Next															'☜: If process fails
    Err.Clear							
    
   If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		

End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False	                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True	                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
    Call InitSpreadSheet()      
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1 , -1 ,C_DOC_CUR ,C_BUY_AMT ,   "A" ,"Q","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1 , -1 ,C_DOC_CUR ,C_PRICE_AMT ,   "A" ,"Q","X","X")
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1 , -1 ,C_DOC_CUR ,C_PRICE_SUM ,   "A" ,"Q","X","X")
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			         '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    if LayerShowHide(1) = False then
	   Exit Function
	end if                                                    '☜: Show Processing Message

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
    
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	lgIntFlgMode      = Parent.OPMD_UMODE                                                   '⊙: Indicates that current mode is Create mode
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

  
	Call SetToolbar("1100000000011111")	                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

End Sub



'========================================================================================================
Function OpenPopup(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strItemGroupCd
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
	
	Select Case iRequried
	Case 1												
		arrParam(0) = "증권 팝업"							
		arrParam(1) = "A_SECURITY"						
		arrParam(2) = Trim(frm1.txtSecurityCd.value)		
		arrParam(3) = "" 		            		<%' Name Cindition%>
		arrParam(4) = ""							<%' Where Condition%>

'		arrParam(3) = Trim(frm1.txtminornm.value)
'		arrParam(4) = "MAJOR_CD='S0001'"					
		arrParam(5) = "증권코드"					
		
	    arrField(0) = "SECURITY_CD"						
	    arrField(1) = "SECURITY_NM"					
	    
	    arrHeader(0) = "증권코드"					
	    arrHeader(1) = "증권명"					

	Case 2					
		arrParam(0) = "증권종류 팝업"														
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtSecurity_TypeCd.value)			
		arrParam(3) = ""
		arrParam(4) = "MAJOR_CD=" & FilterVar("A1031", "''", "S") & " "
		arrParam(5) = "증권종류"								
	
		arrField(0) = "MINOR_CD"									
		arrField(1) = "MINOR_NM"									
		
	    arrHeader(0) = "증권종류코드"								
	    arrHeader(1) = "증권종류"							
	Case 3
		arrParam(0) = "사업장 팝업"							
		arrParam(1) = "B_BIZ_AREA"					<%' TABLE 명칭 %>
		arrParam(2) = frm1.txtMajorCd.value			<%' Code Condition%>
		arrParam(3) = "" 		            		<%' Name Cindition%>
		arrParam(4) = ""							<%' Where Condition%>
		arrParam(5) = "사업장"			
	
		arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
		arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>
    
		arrHeader(0) = "사업장코드"				<%' Header명(0)%>
		arrHeader(1) = "사업장명"				<%' Header명(1)%>
    
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
			Select Case iRequried
			Case 1
				.txtSecurityCd.focus
			Case 2
				.txtSecurity_TypeCd.focus
			Case 3
				.txtMajorCd.focus
			End Select
		End With
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iRequried)
	End If
End Function

'========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iRequried)
	With frm1
		Select Case iRequried
		Case 1
			.txtSecurityCd.focus
			.txtSecurityCd.value = arrRet(0)
			.txtSecurityNm.value = arrRet(1)
		Case 2
			.txtSecurity_TypeCd.focus
			.txtSecurity_TypeCd.value = arrRet(0)
			.txtSecurity_TypeNm.value = arrRet(1)
		Case 3
			.txtMajorCd.focus
			.txtMajorCd.value = arrRet(0)
			.txtMajorName.value = arrRet(1)
		End Select
	End With
End Function



'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
   
   Call SetPopupMenuItemInf("0000011111")   
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
    

End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 특정 column를 click할때 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
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
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End if
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
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>유가증권정보조회</font></td>
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
			            		<TD CLASS="TD5" NOWRAP>증권</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtSecurityCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="증권코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup 1">
									<INPUT TYPE="Text" NAME="txtSecurityNm" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="증권명">
			            		</TD>
			            		<TD CLASS="TD5" NOWRAP>증권종류</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtSecurity_TypeCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="증권종류코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup 2">
									<INPUT TYPE="Text" NAME="txtSecurity_TypeNm" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="증권종류">
			            		</TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>사업장</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtMajorCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup 3">
									<INPUT TYPE="Text" NAME="txtMajorName" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="사업장명">
			            		</TD>
			            		<TD CLASS=TD5 NOWRAP>이자계산여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoYiFlag" id="rdoYiAll" VALUE="A" tag = "11" CHECKED>
										<LABEL FOR="rdoYiAll">전체</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoYiFlag" id="rdoYiYes" VALUE="Y" tag = "11">
										<LABEL FOR="rdoYiYes">계산</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoYiFlag" id="rdoYiNo" VALUE="N" tag = "11">
										<LABEL FOR="rdoYiNo">미계산</LABEL></TD> 
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>&nbsp;&nbsp;</TD>
			            		<TD CLASS="TD6" NOWRAP></TD>
			            		<TD CLASS=TD5 NOWRAP>승인여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoGiFlag" id="rdoGiAll" VALUE="A" tag = "11" CHECKED>
										<LABEL FOR="rdoGiAll">전체</LABEL>&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoGiFlag" id="rdoGiYes" VALUE="Y" tag = "11">
										<LABEL FOR="rdoGiYes">승인</LABEL>&nbsp;&nbsp;	
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoGiFlag" id="rdoGiNo" VALUE="N" tag = "11">
										<LABEL FOR="rdoGiNo">미승인</LABEL></TD>		
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
									<script language =javascript src='./js/a5959ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

