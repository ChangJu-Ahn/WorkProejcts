
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 담당자등록
'*  3. Program ID           : B81105MA1.asp
'*  4. Program Name         : B81105MA1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/01/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Wol san
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

Const BIZ_PGM_ID = "B81105MB1.asp"												<%'비지니스 로직 ASP명 %>
Dim C_Item_acct_cd
Dim C_Item_acct_nm 

Dim C_Item_kind_cd
Dim C_Item_kind_nm

Dim C_REMARK
Dim C_Itemp_p
Dim C_SeqNo
Dim C_Item_d
Dim C_Item_ver
Dim C_UsdRate
Dim C_Scope_Average
Dim C_Item_r
Dim C_Item_t
Dim C_Item_g
Dim C_Item_q
Dim C_USER_ID
Dim C_USER_Popup
Dim C_USER_NAME
Dim lgStrGbn
Dim lgRowTemp


<% EndDate= GetSvrDate %>

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lgStrPrevToKey

Sub InitSpreadPosVariables()

	C_USER_ID		  = 1
	C_USER_Popup      = 2
	C_USER_NAME       = 3
	
    C_Item_acct_cd    = 1 
    C_Item_acct_nm    = 2 
    C_Item_kind_cd    = 3
    C_Item_kind_nm    = 4
    C_Item_r		  = 5
    C_Item_t		  = 6
    C_Item_g		  = 7
    C_Item_q		  = 8
    C_REMARK          = 9
   
  
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevToKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgStrGbn="A"
End Sub

Sub SetDefaultVal()
    Dim strYear
    Dim strMonth
    Dim strDay
    
    Call ExtractDateFrom("<%= EndDate %>",parent.gServerDateFormat , parent.gServerDateType      ,strYear,strMonth,strDay)
	//frm1.txtValidDt.Year  = strYear
	//frm1.txtValidDt.Month = strMonth
	
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet(ByVal pvSpdNo)
   

    Call initSpreadPosVariables()  
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
			With frm1.vspdData1
	
		ggoSpread.Source = frm1.vspdData1	
		'patch version
		 ggoSpread.Spreadinit "V20050822",,parent.gAllowDragDropSpread    
		 
			.ReDraw = false

		 .MaxCols = C_Item_q + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		 .Col = .MaxCols														'☆: 사용자 별 Hidden Column
		 .ColHidden = True    
		 .MaxRows = 0
		 ggoSpread.ClearSpreadData
		 ggoSpread.SSSetEdit C_Item_acct_cd,	"", 10,,,8,2 
		 ggoSpread.SSSetEdit C_Item_acct_nm,	"품목계정", 12,,,8,2 
		 ggoSpread.SSSetEdit C_Item_kind_cd,	"ITEM_KIND", 10,,,8,2 
		 ggoSpread.SSSetEdit C_Item_kind_nm,	"품목구분", 15,,,20,2 
		 ggoSpread.SSSetCheck  C_Item_r ,		"접수검토", 10, ,"",1
		 ggoSpread.SSSetCheck  C_Item_t,		"기술검토", 10,,,1
		 ggoSpread.SSSetCheck  C_Item_g,		"구매검토",  10,,,1
		 ggoSpread.SSSetCheck  C_Item_q,		"품질검토",  10,,,1
		 
		 Call ggoSpread.SSSetColHidden(C_Item_acct_cd,C_Item_acct_cd,True)
		 Call ggoSpread.SSSetColHidden(C_Item_kind_cd,C_Item_kind_cd,True)
	
		.ReDraw = true
	 
		 Call SetSpreadLock 
		  
		 End With
    end if
    
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
   
		With frm1.vspdData2
		
		ggoSpread.Source = frm1.vspdData2	
	   'patch version
	    ggoSpread.Spreadinit "V20050701",,parent.gAllowDragDropSpread    
		.ReDraw = false
	    .MaxCols = C_REMARK + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
	    .MaxRows = 0
	    ggoSpread.ClearSpreadData
	    ggoSpread.SSSetEdit   C_USER_ID,	"사용자 ID", 10,,,13,2
	    ggoSpread.SSSetButton C_USER_Popup
	    ggoSpread.SSSetEdit   C_USER_NAME,	"사용자명", 15,,,8,2 
	    ggoSpread.SSSetEdit   C_Item_kind_nm,"TEMP", 15,,,8,2 '자리수를맞추기위한TEMP
		ggoSpread.SSSetCheck  C_Item_r ,	"접수검토", 10, ,"",1
		ggoSpread.SSSetCheck  C_Item_t,		"기술검토", 10,,,1
		ggoSpread.SSSetCheck  C_Item_g,		"구매검토",  10,,,1
		ggoSpread.SSSetCheck  C_Item_q,		"품질검토",  10,,,1
		ggoSpread.SSSetEdit   C_REMARK,		"비고", 22,,,100
		
		Call ggoSpread.SSSetColHidden(C_Item_kind_nm,C_Item_kind_nm,True)
	
		.ReDraw = true
	
	    End With
	  end if 
    
End Sub

Sub SetSpreadLock()
    Dim j
    With frm1
    ggoSpread.Source = frm1.vspdData1
    .vspdData1.ReDraw = False
    ggoSpread.SpreadLock -1, -1
	.vspdData1.ReDraw = True

	 ggoSpread.Source = frm1.vspdData2
	 .vspdData2.ReDraw = False
	
	 ggoSpread.SpreadLock C_USER_ID, -1, C_USER_ID 
	 ggoSpread.SpreadLock C_USER_Popup, -1, C_USER_Popup 
	 
	 ggoSpread.SpreadLock C_USER_NAME, -1, C_USER_NAME
	 ggoSpread.SpreadLock C_ITEM_R, -1, C_ITEM_R
	 ggoSpread.SpreadLock C_ITEM_t, -1, C_ITEM_T
	 ggoSpread.SpreadLock C_ITEM_G, -1, C_ITEM_G
	 ggoSpread.SpreadLock C_ITEM_Q, -1, C_ITEM_Q
	 ggoSpread.SpreadUnLock C_REMARK, -1, C_REMARK
	.vspdData2.ReDraw = True
	
    End With
    
    SetSpreadColor1 frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow 
  
End Sub

Sub SetSpreadColor1 (ByVal pvStartRow, ByVal pvEndRow)
Dim j
	
 With frm1
	.vspdData2.ReDraw = False
	   for j=C_ITEM_R to  C_ITEM_Q
         ggoSpread.SpreadUnLock j, pvStartRow, pvEndRow
		if GetSpreadvalue(frm1.vspdData1,j,frm1.vspdData1.ActiveRow,"X","X")="1" then
			ggoSpread.SpreadUnLock j, pvStartRow, pvEndRow	
		else
			ggoSpread.SpreadLock j, -1, J
		end if
       next 
     
    .vspdData2.ReDraw = True
    
    End With
End Sub

Sub SetSpreadColor (ByVal pvStartRow, ByVal pvEndRow)
Dim j
	
 With frm1
	.vspdData2.ReDraw = False
	
	 ggoSpread.SpreadLock 3, -1, 3
	 ggoSpread.SSSetRequired 1,	pvStartRow, pvEndRow
	 
	 // 200.03 수정 
	   for j=C_ITEM_R to  C_ITEM_Q
         ggoSpread.SpreadUnLock j, pvStartRow, pvEndRow
		if GetSpreadvalue(frm1.vspdData1,j,frm1.vspdData1.ActiveRow,"X","X")="1" then
			ggoSpread.SpreadUnLock j, pvStartRow, pvEndRow	
		else
			ggoSpread.SpreadLock j, -1, J
		end if
       next 
     
    .vspdData2.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Item_acct_cd         = iCurColumnPos(1) 
            C_Item_acct_nm         = iCurColumnPos(2) 
            C_Item_kind_cd         = iCurColumnPos(4)
            C_Item_kind_nm         = iCurColumnPos(5)

    End Select    
End Sub



Function OpenUser(pValue,colPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim cd_temp
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrParam(0) = "사용자정보 POPUP"						<%' 팝업 명칭 %>
	arrParam(1) = "Z_USR_MAST_REC"	
	arrParam(2) = pValue	<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "USR_KIND='U'"	<%' Where Condition%>
	arrParam(5) = "사용자 ID"					<%' 조건필드의 라벨 명칭 %>
    arrField(0) = "USR_ID"						<%' Field명(0)%>
    arrField(1) = "USR_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "사용자 ID"							<%' Header명(0)%>
    arrHeader(1) = "사용자명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData2.Col = colPos-1
			.vspdData2.Text = arrRet(0)
			.vspdData2.Col = colPos+1
			.vspdData2.Text = arrRet(1)
		end with
	End If	
	
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    Call InitSpreadSheet("*")                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call SetDefaultVal
    
    Call fncQuery()
    Call SetToolbar("1100000000000011")										<%'버튼 툴바 제어 %>
   
End Sub


Sub vspdData2_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub


Sub vspdData1_Click(ByVal Col, ByVal Row)

      Call SetPopupMenuItemInf("0000011111") 
    
    gMouseClickStatus = "SPC"   
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

	lgStrGbn="B"
	if  Col=2 or Col=4 then
		frm1.vspdData2.maxRows = 0
		Call SetToolbar("1100110100011111")
		call DbQuery()
  
	end if
	frm1.hCol.value = col
	frm1.hRow.value = Row
	
      lgStrGbn="A"
	  lgRowTemp=Row
	
End Sub


Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101011111") 
    gMouseClickStatus = "SP2C"   
    Set gActiveSpdSheet = frm1.vspdData2
   
   
   	    ggoSpread.Source = frm1.vspdData2
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
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



Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows = 0 Then
        Exit Sub
    End If
	
End Sub

Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub  
Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    
  

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	           
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
    Call InitSpreadSheet("*")      
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData2 
		If Row > 0 And Col = C_USER_Popup Then
		    .Row = Row
		    .Col = C_USER_Popup
		    Call OpenUser(GetSpreadvalue(frm1.vspdData2,Col-1,Row,"X","X"),Col )
		
		End If
    End With
End Sub



Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) _
	And Not(lgStrPrevKey = "" And lgStrPrevToKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
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
    ggoSpread.Source = frm1.vspdData2
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
    
<%  '-----------------------
    'Query function call area
    '----------------------- %>
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
    FncQuery = True															
    
End Function

Function FncSave() 
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR     '⊙: Check contents area
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

    If Frm1.vspdData2.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData2

 
	With frm1.vspdData2
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
			
    		ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow    

			.ReDraw = True
		End If
	End With

    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData2	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
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
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		
		.vspdData2.ReDraw = False
        ggoSpread.InsertRow ,imRow
        
        SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
        
        
		For iRow = .vspdData2.ActiveRow to .vspdData2.ActiveRow + imRow - 1
		    .vspdData2.Row = iRow
		Next
		.vspdData2.ReDraw = True
    End With
  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData2 
    	.focus
    	ggoSpread.Source = frm1.vspdData2 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 

    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
	Set gActiveElement = document.activeElement 
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData2	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
   
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	Call LayerShowHide(1)
	Dim strVal
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
   
    	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜:
		strVal = strVal & "&item_acct=" & GetSpreadText(frm1.vspdData1,C_Item_acct_cd,frm1.vspdData1.ActiveRow,"X","X")						'☜: 
		strVal = strVal & "&item_kind=" & GetSpreadText(frm1.vspdData1,C_Item_kind_cd,frm1.vspdData1.ActiveRow,"X","X")						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
		strVal = strVal & "&lgStrGbn=" & lgStrGbn
    Else	
    	
    	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&item_acct=" & GetSpreadText(frm1.vspdData1,C_Item_acct_cd,frm1.vspdData1.ActiveRow,"X","X")						'☜: 
		strVal = strVal & "&item_kind=" & GetSpreadText(frm1.vspdData1,C_Item_kind_cd,frm1.vspdData1.ActiveRow,"X","X")						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
		strVal = strVal & "&lgStrGbn=" & lgStrGbn
    End If
  
 
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
   
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>
	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspddata1.focus
	End If
	CALL SetSpreadLock
	Set gActiveElement = document.activeElement
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, k, strTxt
	DIm strCur, strToCur
	Dim YearMonthFormat
	Dim strYYYYMM
	Dim strApplyDt
	Dim ColSep,RowSep
	
    DbSave = False
                                                              
    ColSep = parent.gColSep               
	RowSep = parent.gRowSep    
    Call LayerShowHide(1)
    On Error Resume Next                                                   <%'☜: Protect system from crashing%>

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
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타    
    For lRow = 1 To .vspdData2.MaxRows
    
		    .vspdData2.Row = lRow
		    .vspdData2.Col = 0
		    
		    Select Case .vspdData2.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
			End Select			

		    Select Case .vspdData2.Text
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 
					strVal = strVal & Trim(GetSpreadText(.vspdData1,C_Item_acct_cd,.vspdData1.ActiveRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData1,C_Item_kind_cd,.vspdData1.ActiveRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_USER_ID,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_Item_r,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_Item_t,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_Item_g,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_Item_q,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData2,C_REMARK,lRow,"X","X")) & ColSep  & lRow & ColSep & RowSep
		            lGrpCnt = lGrpCnt + 1
		           

		        Case ggoSpread.DeleteFlag							'☜: 삭제 
						
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep					'☜: U=Update
					strDel = strDel & Trim(GetSpreadText(.vspdData1,C_Item_acct_cd,.vspdData1.ActiveRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData1,C_Item_kind_cd,.vspdData1.ActiveRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData2,C_USER_ID,lRow,"X","X")) & ColSep & lRow & ColSep & RowSep
  		            lGrpCnt = lGrpCnt + 1
  		            
		    End Select
		Next
	
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal

	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	
End Function

Sub txtValidDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtValidDt.Action = 7
        Call SetFocusToDocument("M")   
        frm1.txtValidDt.Focus
    End If
End Sub

Sub txtValidDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

Function TransToYYYYMM(strDate)
	Dim strYear
    Dim strMonth
    Dim strDay
    
   ' ON Error Resume Next
	
	Call ExtractDateFrom(strDate ,parent.gDateFormatYYYYMM , parent.gComDateType ,strYear, strMonth, strDay)
	
	strDate = strYear & strMonth
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strAspMnuMnunm")%></font></td>
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="50%">
									<script language =javascript src='./js/b81105ma1_vaSpread1_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/b81105ma1_vaSpread2_vspdData2.js'></script>
								</TD>
							</TR>
							
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B81101MB1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="hGridGbn" value="A" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">

<INPUT TYPE=HIDDEN NAME="hCol" tag="24">
<INPUT TYPE=HIDDEN NAME="hRow" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

