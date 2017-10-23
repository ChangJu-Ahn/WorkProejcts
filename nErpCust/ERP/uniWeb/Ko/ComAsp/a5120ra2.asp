
<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5120RA1
'*  4. Program Name         : 
'*  5. Program Desc         : Ado query Sample with DBAgent(Multi + Multi)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2002/02/23
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Park Shim Seo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :  2002/11/25 : ASP Standard for Include improvement
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../inc/incSvrCcm.inc"  -->
<!-- #Include file="../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAMain.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAEvent.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAOperation.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliVariables.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../inc/incImage.js">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incEB.vbs">						</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">



Option Explicit 
	
Dim arrParent
Dim arrParam	
'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)
	
top.document.title = PopupParent.gActivePRAspName
	

Const C_HEAD = 0
Const C_MASTER = 1
Const C_DETAIL = 2
	

Const BIZ_PGM_ID        = "a5120rb21.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a5120rb22.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID2       = "a5120rb23.asp"                         '☆: Biz logic spread sheet for #2

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_MaxKey		        = 6                                    '☆☆☆☆: Max key value

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgDocCur                          
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgPageNo_C
Dim lgSortKey_A
Dim lgSortKey_B
Dim lgSortKey_C
Dim lgIsOpenPop


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgBlnFlgChgValue	= False
    lgIntFlgMode		= PopupParent.OPMD_CMODE

    lgPageNo_A			= ""
    lgSortKey_A			= 1

    lgPageNo_B			= ""
    lgSortKey_B			= 1

	lgPageNo_C			= ""
    lgSortKey_C			= 1

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	frm1.hHqBrchNo.value  = arrParam(0)
	
'	frm1.txtRefNo.value = arrParam(1)
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="./LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
End Sub


'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	
	Call SetZAdoSpreadSheet("A5120RA1_HDR", "S", "C", "V20030210", PopupParent.C_SORT_DBAGENT, frm1.vspdData0, C_MaxKey, "X", "X")
	
	Call SetZAdoSpreadSheet("A5120RA1", "S", "A", "V20030210", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	
	Call SetZAdoSpreadSheet("A5120RA1_DTL", "S", "B", "V20021108", PopupParent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")
	
	Call SetSpreadLock ("C")    
	Call SetSpreadLock ("A")
	Call SetSpreadLock ("B")
    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock( iOpt )
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLockWithOddEvenRowColor()
          .vspdData.ReDraw = True
       End With
    ELSEIF iOpt = "C" Then   
		With frm1
          .vspdData0.ReDraw = False
          ggoSpread.Source = .vspdData0 
          ggoSpread.SpreadLockWithOddEvenRowColor()
          .vspdData0.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
			ggoSpread.SpreadLockWithOddEvenRowColor()
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub


'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
   Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
   Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    
   Call ggoOper.LockField(Document, "N")                                      ' ⊙: Lock  Suitable  Field

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()

    '--------- Developer Coding Part (End  ) ----------------------------------------------------------

'	If GetDucCur() Then
		
'		Call CurFormatNumericOCX()
'		Call CurFormatNumSprSheet()
'	End If

	Call FncQuery()

End Sub


'========================================================================================================
'   Event Name : GetDucCur()
'   Event Desc :
'========================================================================================================
'Function GetDucCur()
'	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
'   Dim strBizAreaCd, strBizAreaNm
'    Dim strSelect
'    Dim strFrom
'    Dim strWhere
'    Dim arrTemp
'    
'    GetDucCur = False
'    strSelect	= "isnull(doc_cur,'')"
'    strFrom		= "a_gl_item"
'    strWhere	= "gl_no='" & frm1.txtGlNo.value & "'" 
'    
'    If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then'
'		arrTemp		= split(lgF0, Chr(11))
'		lgDocCur	= arrTemp(0) 		
'	if Trim(lgDocCur) = "" Then
'		GetDucCur = False
'	Else
'		GetDucCur = True
'	End If					
'End If
	
'nd Function


'========================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    
     If Trim(frm1.hHqBrchNo.value) = "" Then
		Call DisplayMsgBox("113100", "X", "X", "X")
		Call CancelClick()
		Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
	frm1.vspdData0.MaxRows = 0                                                      '☜: Protect system from crashing
    Call DbQuery(C_HEAD)															'☜: Query db data

    FncQuery = True	
End Function



'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    Dim iColumnLimit2
    
    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit  =3
       
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars = PopupParent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    End If   
	
	'----------------------------------------
	' Spread가 두개일 경우 2번째 Spread
	'----------------------------------------
	
	
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = 4
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = PopupParent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    End If   
    
End Function


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo)
	Dim intRetCd

	StrEbrFile = "a5121ma1"
	
	frm1.vspddata0.col = 2	
	frm1.vspddata0.row = Frm1.vspddata0.ActiveRow	'1
	
	VarDateFr = UniConvDateToYYYYMMDD(Trim(frm1.vspddata0.text), PopupParent.gDateFormat,"")	
	VarDateTo = UniConvDateToYYYYMMDD(Trim(frm1.vspddata0.text), PopupParent.gDateFormat,"")	
' 회계전표의 key는 temp_GL_NO이기 때문에 temp_GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	
	frm1.vspddata0.col = 1					
	frm1.vspddata0.row = Frm1.vspddata0.ActiveRow	'1	
	varGlNoFr = Trim(frm1.vspddata0.value)			
	frm1.vspddata0.row = frm1.vspddata0.maxrows
	varGlNoTo = Trim(frm1.vspddata0.value)
	
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varGlPutType
    Dim StrEbrFile
    Dim intRetCd
	
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	varGlPutType ="%"

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo)
	
	
'    On Error Resume Next
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|GlPutType|" & varGlPutType	
	StrUrl = StrUrl & "|OrgChangeId|" & PopupParent.gChangeOrgId
	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	'On Error Resume Next 
    
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varGlPutType
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If
	varGlPutType ="%"
	
	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo)
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
   
    StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|GlPutType|" & varGlPutType	
	StrUrl = StrUrl & "|OrgChangeId|" & PopupParent.gChangeOrgId
 
	Call FncEBRPreview(ObjName,StrUrl)

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
Function DbQuery(pDirect) 
	Dim strVal

    DbQuery = False
    
    Err.Clear 
    
    Select Case pDirect
		Case  C_HEAD
			
			Call LayerShowHide(1)
			    
			With frm1
			
			'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------		    
				strVal = BIZ_PGM_ID & "?txtHqBrchNo=" & Trim(.hHqBrchNo.value)
				
			'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
			    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
			    strVal = strVal & "&lgPageNo="       & lgPageNo_C                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("C")         
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("C")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("C"))
				strVal = strVal & "&txtMaxRows=" & .vspdData0.MaxRows
			        
			    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

			End With
			
		Case  C_MASTER           
		
			FRM1.vspdData.MaxRows  = 0   
			Call LayerShowHide(1)			
			With frm1			
			'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
				strVal = BIZ_PGM_ID1 & "?txtGlNo=" & Trim(GetKeyPosVal("C", 1))				
		'		------------ 개발자 coding part(실행로직,End)------------------------------------------------
			    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
			    strVal = strVal & "&lgPageNo="       & lgPageNo_A                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			        
			    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

			End With
    
		Case C_Detail
			frm1.vspdData2.MaxRows = 0 
			Call LayerShowHide(1)
			    
			With frm1
			'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------
				strVal = BIZ_PGM_ID2 & "?txtGlNo=" & Trim(GetKeyPosVal("A", 1))
				strVal = strVal & "&txtSeq=" & Trim(GetKeyPosVal("A", 2))
			'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
			    strVal = strVal & "&lgPageNo="       & lgPageNo_B                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")         
				'strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
			    
			    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
			End With
		
	End Select 
    DbQuery = True   

End Function


'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk( iOpt)											 '☆: 조회 성공후 실행로직 
	
	
	Dim lngRows
	Dim strTableid
	Dim strColid
	Dim strColNm
	Dim strMajorCd
	Dim strNmwhere
	Dim arrVal

	Const C_Tableid		= 8
	Const C_Colid		= 9
	Const C_ColNm		= 10
	Const C_MajorCd		= 14
	Const C_CtrlVal		= 4
	Const C_CtrlValNm	= 6


    lgIntFlgMode     = PopupParent.OPMD_UMODE
	If iOpt = 0 Then
       Call vspdData0_Click(1,1)
       frm1.vspdData0.focus
	end if
	
	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
		       
	End If
    
    For lngRows = 1 To frm1.vspdData2.Maxrows
		frm1.vspddata2.row = lngRows	
		frm1.vspddata2.col = C_Tableid 
		IF Trim(frm1.vspddata2.text) <> "" Then

			frm1.vspddata2.col = C_Tableid
			strTableid = frm1.vspddata2.text
			frm1.vspddata2.col = C_Colid
			strColid = frm1.vspddata2.text
			frm1.vspddata2.col = C_ColNm
			strColNm = frm1.vspddata2.text	
			frm1.vspddata2.col = C_MajorCd
			strMajorCd = frm1.vspddata2.text

			frm1.vspddata2.col = C_CtrlVal

			strNmwhere = strColid & " =  '" & frm1.vspddata2.text & "' "
			IF Trim(strMajorCd) <> "" Then
				strNmwhere = strNmwhere & " AND MAJOR_CD = '" & strMajorCd & "' "
			End IF
			IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
				frm1.vspddata2.col = C_CtrlValNm
				arrVal = Split(lgF0, Chr(11))
				If Ubound(arrVal, 1) <> - 1 Then 
					frm1.vspddata2.text = arrVal(0)
				End If
			End IF
		End IF
	Next


	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
End Function


'========================================================================================================
'	Name : OpenConItemCd()
'	Description : Item PopUp
'========================================================================================================
Function OpenConItemCd()


End Function

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenSortPopup()

	Dim arrRet	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
	arrRet = window.showModalDialog("./ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If

End Function

'========================================================================================================
'   Event Name : vspdData0_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData0_Click( Col,  Row)
    Dim ii

	gMouseClickStatus = "SPC"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData0
        If lgSortKey_C = 1 Then
            ggoSpread.SSSort, lgSortKey_C
            lgSortKey_C = 2
        Else
            ggoSpread.SSSort, lgSortKey_C
            lgSortKey_C = 1
        End If    
        Exit Sub
    End If

	If Col < 1 Then Exit Sub
	
	Call SetSpreadColumnValue("C",frm1.vspdData0,Col,Row)
    Call DbQuery(C_MASTER)
End Sub

'========================================================================================================
'   Event Name : vspdData0_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData0_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    ggoSpread.Source = frm1.vspdData0
		    If lgSortKey_C = 1 Then
		        ggoSpread.SSSort, lgSortKey_C
		        lgSortKey_C = 2
		    Else
		        ggoSpread.SSSort, lgSortKey_C
		        lgSortKey_C = 1
		    End If    
		    Exit Sub
		End If

		If Col < 1 Then Exit Sub
		
		Call SetSpreadColumnValue("C",frm1.vspdData0,Col,NewRow)
		Call DbQuery(C_MASTER)
    End If
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click( Col,  Row)
    Dim ii

	gMouseClickStatus = "SPC"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If

	If Col < 1 Then Exit Sub
	
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    Call DbQuery(C_DETAIL)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    ggoSpread.Source = frm1.vspdData
		    If lgSortKey_A = 1 Then
		        ggoSpread.SSSort, lgSortKey_A
		        lgSortKey_A = 2
		    Else
		        ggoSpread.SSSort, lgSortKey_A
		        lgSortKey_A = 1
		    End If    
		    Exit Sub
		End If

		If Col < 1 Then Exit Sub
		
		Call SetSpreadColumnValue("A",frm1.vspdData,Col,NewRow)
		Call DbQuery(C_DETAIL)
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click( Col,  Row)
    Dim ii
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If
        Exit Sub
    End If

	gMouseClickStatus = "SP2C"

End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'           Call DisableToolBar(PopupParent.TBC_QUERY)
           If DbQuery(C_MASTER) = False Then
'              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'           Call DisableToolBar(PopupParent.TBC_QUERY)
           If DbQuery(C_DETAIL) = False Then
'              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
    
End Sub


'=======================================================================================================
'   Event Name : fpdtToEnterDt_KeyDown(keycode, shift)
'   Event Desc : 
'=======================================================================================================
Sub fpdtToEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'차변금액 

		ggoOper.FormatFieldByObjectOfCur .txtDrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'대변금액 
		ggoOper.FormatFieldByObjectOfCur .txtCrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

End Sub

</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### %>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<script language =javascript src='./js/a5120ra2_vspdData0_vspdData0.js'></script>
					</TD>
				</TR>
				<TR>
					<TD>
						<script language =javascript src='./js/a5120ra2_vspdData_vspdData.js'></script>
					</TD>
				</TR>
				<TR HEIGHT="30%">
					<TD>
						<script language =javascript src='./js/a5120ra2_vspdData2_vspdData2.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>				
				<TR>
				<TD >&nbsp;&nbsp;<BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
				                 <BUTTON NAME="bttnPrint"	 CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				<TD ALIGN=RIGHT> <IMG SRC="../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)" >	 </IMG>&nbsp;
								 <IMG SRC="../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/zpConfig.gif',1)" ></IMG>&nbsp;
                                 <IMG SRC="../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)">	 </IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>		
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hHqBrchNo" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>