<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM211QB101
'*  4. Program Name         : 멀티컴퍼니B/L조회-멀티 
'*  5. Program Desc         : 멀티컴퍼니B/L조회-멀티 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	
	Dim iStrBlNo
	
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
    Dim lgDataExist
    Dim lgPageNo
    Dim sRow
    Dim lglngHiddenRows
    Dim lgStrPrevKeyM
    DIM MaxRow2
    Dim MaxCount
    
    Dim arrRsVal(11)
	Const C_SHEETMAXROWS_D  = 100
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo       = UNICInt(Trim(Request("lgPageNo1")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgStrPrevKeyM  = UNICInt(Trim(Request("lgStrPrevKeyM")),0)
	lgDataExist    = "No"
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	lgStrPrevKey   = Request("lgStrPrevKey")
	sRow           = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))

	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,0)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "MM211QA102" 											' header
  	      
    iStrBlNo = Trim(Request("txtBlNo"))    
                  
	UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtBlNo"))), " " , "SNM") & "' "

	
	    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
'		Call DisplayMsgBox("172400", vbOKOnly, iStrPrNo, "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgPageNo1				=  """ & lgPageNo	 & """" & vbCr	
    
    Response.Write "	.lgStrPrevKeyM(" & sRow - 1 & ") = """ & lgStrPrevKeyM & """" & vbCr
    Response.Write "    .lglngHiddenRows(" & sRow - 1 & ") = """ & MaxRow2 & """" & vbCr  
    Response.Write "    .DbQueryOk2(" & MaxCount & ")" & vbCr
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr        

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    DIM i

	Const M_BL_DTL_BL_NO			=	0
	Const M_BL_DTL_BL_SEQ			=	1
	Const M_BL_DTL_ITEM_CD	     	=	2
	Const B_ITEM_ITEM_NM            =	3
	Const B_ITEM_SPEC               =	4
	Const M_BL_DTL_QTY              =	5
	Const M_BL_DTL_UNIT             =	6
	Const M_BL_DTL_PRICE            =	7
	Const M_BL_DTL_DOC_AMT          =	8
	Const M_BL_DTL_LOC_AMT          =	9
	Const M_BL_DTL_PO_NO            =	10
	Const M_BL_DTL_PO_SEQ_NO        =	11

    lgDataExist    = "Yes"
	MaxRow2 = 0	
    iLoopCount = 0
    i = 0
	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		MaxRow2     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    

        
	'----- 레코드셋 칼럼 순서 ----------
	'A.ITEM_CD, (SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = A.ITEM_CD) ITEM_NM, 
	'(SELECT SPEC FROM B_ITEM WHERE ITEM_CD = A.ITEM_CD) SPEC, A.QTY, A.UNIT, A.PRICE,
	'A.DOC_AMT, A.LOC_AMT, A.PO_NO, A.PO_SEQ_NO
	'-----------------------------------    
	
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_BL_DTL_ITEM_CD))	 									'
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_ITEM_NM))                                    '
		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_BL_DTL_ITEM_CD))	 									'
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_ITEM_NM))                                    '
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_SPEC))                                   '품목          

		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_BL_DTL_QTY), ggQty.DecPoint,0)                                                     '품목명        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_BL_DTL_UNIT))                                                         '품목규격      
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_BL_DTL_PRICE), ggAmtOfMoney.DecPoint,0)                   '수량          
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_BL_DTL_DOC_AMT), ggAmtOfMoney.DecPoint,0)                                        '단위          
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_BL_DTL_LOC_AMT), ggAmtOfMoney.DecPoint,0)       '단가           
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_BL_DTL_PO_NO))                                       '부가세유형    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_BL_DTL_PO_SEQ_NO))                                    '부가세유형명  

        
        iRowStr = iRowStr & Chr(11) & sRow
		iRowStr = iRowStr & Chr(11) & Trim(ConvSPChars(MaxRow2 + iLoopCount))
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             


        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
        Else
           'lgStrPrevKeyM = lgStrPrevKeyM + 1
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
        i = i + 1
   Loop

    If iLoopCount-1 < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
		lgPageNo = ""
       'lgStrPrevKeyM = ""
    End If

    MaxRow2 = MaxRow2 + iLoopCount 
    MaxCount = iLoopCount
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub



'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>