
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81101MA1.asp
'*  4. Program Name         : B81101MA1.asp
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

Const BIZ_PGM_ID = "B81101MB1.asp"												<%'비지니스 로직 ASP명 %>
Dim C_Item_acct_cd
Dim C_Item_acct_nm 

Dim C_Item_kind_cd
Dim C_Item_kind_nm
Dim C_Item_Acct_Popup
Dim C_Item_Kind_Popup
Dim C_Division1
Dim C_Division2
Dim C_Division3
Dim C_ITEM_CREATE
Dim C_SeqNo
Dim C_Item_d
Dim C_Item_ver
Dim C_UsdRate
Dim C_Scope_Average
Dim C_TOTAL
Dim C_Item_r
Dim C_Item_t
Dim C_Item_g
Dim C_Item_q
DIM C_CREATE_ITEM_CD
DIM C_CREATE_ITEM


<% EndDate= GetSvrDate %>


<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim lgStrPrevToKey

Sub InitSpreadPosVariables()


    C_Item_acct_cd   = 1 
    C_Item_acct_nm    = 2 
    C_Item_acct_Popup = 3 
    C_Item_kind_cd    = 4
    C_Item_kind_nm    = 5
    C_Item_kind_Popup = 6
    C_Division1       = 7
    C_Division2		  = 8
    C_Division3       = 9
    C_SeqNo           = 10
    C_Item_d          = 11
    C_Item_ver        = 12
    C_TOTAL			  = 13
    C_Item_r		  = 14
    C_Item_t		  = 15
    C_Item_g		  = 16
    C_Item_q		  = 17
    C_CREATE_ITEM_CD  = 18
    C_CREATE_ITEM     = 19
    
  
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
End Sub

Sub SetDefaultVal()
   
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

Sub InitSpreadSheet()
    

    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20050301",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_CREATE_ITEM + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
    ggoSpread.SSSetEdit   C_Item_acct_cd,		"", 10,,,8,2 
    ggoSpread.SSSetEdit   C_Item_acct_nm,		"품목계정", 10,,,8,2 
    ggoSpread.SSSetButton C_Item_Acct_Popup
    ggoSpread.SSSetEdit   C_Item_kind_cd,		"", 10,,,8,2 
    ggoSpread.SSSetEdit   C_Item_kind_nm,		"품목구분", 12,,,20,2 
    ggoSpread.SSSetButton C_Item_kind_Popup
    ggoSpread.SSSetFloat  C_Division1,		"대분류", 9, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
    ggoSpread.SSSetFloat  C_Division2,		"중분류", 9, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
	ggoSpread.SSSetFloat  C_Division3,		"소분류", 9, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
	ggoSpread.SSSetFloat  C_SeqNo,			"Serial No.",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","10"
	ggoSpread.SSSetFloat  C_Item_d,			"파생번호",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","10"
	ggoSpread.SSSetFloat  C_Item_ver,		"이슈구분",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","10"
	ggoSpread.SSSetFloat  C_Total,			"Total",7,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"1","10"
	ggoSpread.SSSetCheck  C_Item_r ,		"접수", 7, ,"",1
	ggoSpread.SSSetCheck  C_Item_t,			"기술", 7,,,1
	ggoSpread.SSSetCheck  C_Item_g,			"구매",  7,,,1
	ggoSpread.SSSetCheck  C_Item_q,			"품질",  7,,,1
	ggoSpread.SSSetCombo  C_CREATE_ITEM_CD,	"",4
	ggoSpread.SSSetCombo  C_CREATE_ITEM,	"생성시점",10
	
	
    Call ggoSpread.SSSetColHidden(C_Item_acct_cd,C_Item_acct_cd,True)
    Call ggoSpread.SSSetColHidden(C_Item_kind_cd,C_Item_kind_cd,True)
    Call ggoSpread.SSSetColHidden(C_CREATE_ITEM_CD,C_CREATE_ITEM_CD,True)
    
  
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_Item_acct_nm, -1, C_Item_acct_nm
	ggoSpread.SpreadLock C_Item_Acct_Popup, -1, C_Item_Acct_Popup
	ggoSpread.SpreadLock C_Item_kind_nm, -1, C_Item_kind_nm
	ggoSpread.SpreadLock C_ITEM_R, -1, C_ITEM_R
	
	ggoSpread.SpreadLock C_TOTAL, -1, C_TOTAL
	ggoSpread.SpreadLock C_Item_Kind_Popup, -1, C_Item_Kind_Popup

	.vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired C_Item_acct_nm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_Item_kind_nm, pvStartRow, pvEndRow
    ggoSpread.SpreadLock C_TOTAL, -1, C_TOTAL
    ggoSpread.SpreadLock C_ITEM_R, -1, C_ITEM_R
	
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Item_acct_cd         = iCurColumnPos(1) 
            C_Item_acct_nm         = iCurColumnPos(2) 
            C_Item_Acct_Popup	   = iCurColumnPos(3)
            C_Item_kind_cd         = iCurColumnPos(4)
            C_Item_kind_nm         = iCurColumnPos(5)
            C_Item_Acct_Popup      = iCurColumnPos(6)
            C_Division1       = iCurColumnPos(7)
            C_Division2		  = iCurColumnPos(8)
            C_Division3       = iCurColumnPos(9)
            C_SeqNo        = iCurColumnPos(10)
            C_Item_d     = iCurColumnPos(11)
            C_Item_ver    = iCurColumnPos(12)
       
    End Select    
End Sub

Sub InitSpreadComboBox()
	ggoSpread.SetCombo "*" & vbTab & "/", C_Division3
End Sub


Sub InitComboBox2()
    Dim    iCodeArr
    Dim    iNameArr
    
     ggoSpread.SetCombo "R" & vbtab & "T" & vbtab & "P"& vbtab & "Q" , C_CREATE_ITEM_CD
     ggoSpread.SetCombo "접수" & vbtab & "기술" & vbtab & "구매" & vbtab & "품질", C_CREATE_ITEM
 
End Sub


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
 
    With frm1.vspdData
		.Row = Row
		Select Case Col
		    Case C_CREATE_ITEM
		        .Col = Col
		        intIndex = .Value 
				.Col = C_CREATE_ITEM_CD
				.Value = intIndex
		
		End Select
    End With

     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

Function OpenAcct(iWhere,colPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim cd_temp

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	
	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = "품목계정 팝업"						<%' 팝업 명칭 %>
		arrParam(1) = "B_MINOR"						<%' TABLE 명칭 %>
	
		arrParam(2) = Trim(frm1.vspdData.Text)	<%' Code Condition%>
		arrParam(4) = " MAJOR_CD = 'P1001' "	<%' Where Condition%>
		arrParam(5) = "품목계정"					<%' 조건필드의 라벨 명칭 %>
		arrParam(3) = ""								<%' Name Cindition%>
	
		arrField(0) = "MINOR_CD"						<%' Field명(0)%>
		arrField(1) = "MINOR_NM"					<%' Field명(1)%>
    
		arrHeader(0) = "품목계정"							<%' Header명(0)%>
		arrHeader(1) = "품목계정명"							<%' Header명(1)%>
    
	Else 'spread
		arrParam(0) = "품목구분 팝업"						<%' 팝업 명칭 %>
		arrParam(1) = "B_MINOR"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.vspdData.Text)	<%' Code Condition%>
		arrParam(4) = " MAJOR_CD = 'Y1001' "	<%' Where Condition%>
		arrParam(5) = "품목구분"					<%' 조건필드의 라벨 명칭 %>
		arrParam(3) = ""								<%' Name Cindition%>
	
		arrField(0) = "MINOR_CD"						<%' Field명(0)%>
		arrField(1) = "MINOR_NM"					<%' Field명(1)%>
    
		arrHeader(0) = "품목구분"							<%' Header명(0)%>
		arrHeader(1) = "품목구분명"							<%' Header명(1)%>
    
	End If
	
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData.Col = colPos
			.vspdData.Text = arrRet(0)
			.vspdData.Col = colPos+1
			.vspdData.Text = arrRet(1)
		end with
	End If	
	
End Function

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    Call SetDefaultVal
    Call SetToolbar("11000000000111")										<%'버튼 툴바 제어 %>
    Call fncQuery()
    Call InitComboBox2()
  
   
End Sub


Sub getTotal()
	dim tempSum
	dim k

	tempSum=0

	With frm1.vspdData
		for k=c_division1 to C_Item_ver
			.Col= k
			tempSum = cint(tempSum) + cInt(.text)
		next
		.Col = C_TOTAL
		.text = tempSum
		
	end With
end Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
	
	call getTotal()
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
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		If Row > 0 And Col = C_Item_Acct_Popup Then
		    .Row = Row
		    .Col = C_Item_Acct_cd

		    Call OpenAcct(0,.Col )
		ElseIf Row > 0 And Col = C_Item_Kind_Popup Then
		    .Row = Row
		    .Col = C_Item_Kind_cd

		    Call OpenAcct(1,.Col )
		End If
    End With
End Sub



Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
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
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR     '⊙: Check contents area
       Exit Function
    End If
    
   
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    if chk_item()<>true then	Exit Function
    if chkItemLen()<>true then	Exit Function
    if totalChk()<>true then	Exit Function
    if chkCreateItem<> true then exit function
    If DbSave = False Then Exit Function                                        '☜: Save db data
    FncSave = True                                                          
    
End Function


function chk_item()
	

	Dim i,j
	Dim item_acct_cd,item_kind_cd
	chk_item=false
	
	with frm1 	
		for i=1 to .vspdData.maxRows
		.vspdData.Row = i
		.vspdData.Col = 0
			Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag	
	
				item_acct_cd=Trim(GetSpreadText(.vspdData,C_Item_acct_cd,i,"X","X"))  
				if item_acct_cd="" then
		
					Call DisplayMsgBox("970029","X","품목계정","X")
					.vspdData.Col=C_Item_acct_nm
					.vspdData.action=0
					.vspdData.focus()
					exit function
					
		
				end if
				
				item_kind_cd=Trim(GetSpreadText(.vspdData,C_Item_kind_cd,i,"X","X")) 
				if item_kind_cd="" then
		
						Call DisplayMsgBox("970029","X","품목구분","X")
						.vspdData.Col=C_Item_kind_nm
						.vspdData.action=0
						.vspdData.focus()
						exit function
						
		
					end if
				 
			
		    end Select 
		next
	end With
	
	chk_item=true
	
	
		
end Function

Function chkItemLen()
	Dim i,j
	chkItemLen=false
	with frm1 	
		for i=1 to .vspdData.maxRows
		.vspdData.Row = i
		.vspdData.Col = 0
			Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag	
	
				for j=7 to 	12
	
		         IF j=7 or j=8 or j=9 or j=10 then
					if cInt(GetSpreadvalue(.vspdData,j,i,"X","X")) = 0 then
			 			Call  DisplayMsgBox("970022", vbOKOnly, "길이", "0") '0보다 커야 합니다. 
							.vspdData.row=i
							.vspdData.col=j
							.vspdData.action=0
							exit Function	
						end if
				  end if	
					if cInt(GetSpreadvalue(.vspdData,j,i,"X","X")) > 5 then
			 		Call  DisplayMsgBox("127928", vbOKOnly, "길이 5", "X") '초과할 수 없습니다. 
						.vspdData.row=i
						.vspdData.col=j
						.vspdData.action=0
						exit Function	
					end if
				next
		    end Select 
		next
	end With
	
	chkItemLen=true
End Function


Function totalChk()
	Dim i,totLength
	totalChk=false
	
   with frm1 	
	for i=1 to frm1.vspdData.maxRows
	.vspdData.Row = i
	.vspdData.Col = 0
		
		Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag	
	
				totLength=cInt(GetSpreadText(.vspdData,C_TOTAL,i,"X","X"))
				if  totLength > 18 THEN
					call DisplayMsgBox("127928", vbOKOnly, "길이 18", "X") '초과할 수 없습니다. 
					.vspdData.col=C_TOTAL
					.vspdData.row=i
					.vspdData.Action = 0
					Set gActiveElement = document.activeElement  
					exit function
				end if
		end select		
			
	next
	End With
	totalChk=true
End Function

Function chkCreateItem()
dim i
dim temp

	chkCreateItem=false
	
	for i=1 to frm1.vspdData.maxRows
		with  frm1
			temp = getSpreadValue(.vspdData,C_CREATE_ITEM,i,"X","X")
			if temp=1 then 
				if getSpreadText(.vspdData,c_item_t,i,"X","X")<>"1" then
				
					call DisplayMsgBox("970029", vbOKOnly, "생성시점", "X") 
					.vspdData.col=C_CREATE_ITEM
					.vspdData.row=i
					.vspdData.action = 0 
					exit function
				end if
			end if
			
			if temp=2 then 
				if getSpreadText(.vspdData,c_item_g,i,"X","X")<>"1" then
				
					call DisplayMsgBox("970029", vbOKOnly, "생성시점", "X") 
					.vspdData.col=C_CREATE_ITEM
					.vspdData.row=i
					.vspdData.action = 0 
					exit function
				end if
			end if
			
			if temp=3 then 
				if getSpreadText(.vspdData,c_item_q,i,"X","X")<>"1" then
				
					call DisplayMsgBox("970029", vbOKOnly, "생성시점", "X") 
					.vspdData.col=C_CREATE_ITEM
					.vspdData.row=i
					.vspdData.action = 0 
					exit function
				end if
			end if
			
		end with
	next
	
	chkCreateItem=true
end Function

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
			.ReDraw = False
			
    		ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow    
			'Key field clear
			SetSpreadValue frm1.vspdData,C_Item_acct_cd,.ActiveRow,"","",""
		SetSpreadValue frm1.vspdData,C_Item_acct_nm,.ActiveRow,"","",""
		SetSpreadValue frm1.vspdData,C_Item_kind_cd,.ActiveRow,"","",""
		SetSpreadValue frm1.vspdData,C_Item_kind_nm,.ActiveRow,"","",""

			.ReDraw = True
		End If
	End With

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
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
		    .vspdData.Row = iRow
		    .vspdData.Col = C_CREATE_ITEM_CD
		    .vspdData.value=1
		    .vspdData.Col = C_CREATE_ITEM
		    .vspdData.value=1
		  SetSpreadValue frm1.vspdData,C_ITEM_R,iRow,"1","1","1"
		  SetSpreadValue frm1.vspdData,C_ITEM_T,iRow,"1","1","1"
		Next
	
		.vspdData.ReDraw = True
    End With
    
  
  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
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

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal

    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
   
    	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
    Else	
    	
    	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
		
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
	
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, k, strTxt
	DIm strCur, strToCur
	Dim YearMonthFormat
	Dim strYYYYMM
	Dim strApplyDt
	Dim ColSep,RowSep,item_acct_cd,item_kind_cd
	
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
    For lRow = 1 To .vspdData.MaxRows
    
	
		
		.vspdData.Row = lRow
		.vspdData.Col = 0
		
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
			    Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
			End Select			

			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'☜: 신규, 수정 
			        
					strVal = strVal & uCase(Trim(GetSpreadText(.vspdData,C_Item_acct_cd,lRow,"X","X"))) & ColSep
					strVal = strVal & uCase(Trim(GetSpreadText(.vspdData,C_Item_kind_cd,lRow,"X","X"))) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Division1,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Division2,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Division3,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_CREATE_ITEM_CD,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_d,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_ver,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_r,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_t,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_g,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_Item_q,lRow,"X","X")) & RowSep
					
			        lGrpCnt = lGrpCnt + 1

			    Case ggoSpread.DeleteFlag							'☜: 삭제 
						   	
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep					'☜: U=Update
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Item_acct_cd,lRow,"X","X")) & ColSep
					strDel = strDel & Trim(GetSpreadText(.vspdData,C_Item_kind_cd,lRow,"X","X")) & ColSep & lRow & ColSep & RowSep
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
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
    
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

'========================================================================================
' Function Name : CheckCell
' Function Desc : 각 항목의 자리수를 배열에 담아놓기 
'========================================================================================
function CheckLength()
	Dim i
	CheckLength=false
	With frm1
	
	ggoSpread.source = frm1.vspdData

      For i = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = i
		    .vspdData.Col =0

		    Select Case .vspdData.Text 
		        Case ggoSpread.InsertFlag		'☜: 신규, 수정 
					if cint(GetSpreadText(.vspdData,C_LEN,i,"X","X"))<> cInt(len(GetSpreadText(.vspdData,C_CLASS_CD,i,"X","X"))) then
						Call DisplayMsgBox("970029","X","자릿수","X")
						Call SetToolBar("111011110011111")
						
						.vspdData.focus()	
						.vspdData.Action=0
						exit Function
					end if
		    end select 
	next 
	
	End With
	
		    
	CheckLength=true
End function

	
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/b81101ma1_vaSpread1_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

