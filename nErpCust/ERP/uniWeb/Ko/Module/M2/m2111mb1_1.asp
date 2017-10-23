<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MB1_1
'*  4. Program Name         : 구매요청등록 
'*  5. Program Desc         : 구매요청번호별 업체지정내역 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/12/06
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim lgstrData
	Dim iStrPoNo
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
    Dim lgDataExist
    Dim lgPageNo
    
    dim lgPageNo1
    Dim lgStrPrevKeyM
    
    Dim lglngHiddenRows
    Dim lRow
    
	
	DIM MaxRow2
	
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call SubBizQueryMulti


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))

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
	Dim arrVal(3)
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,0)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "m2111ma201" 											' header

	strval = ""
	
	If Len(Request("txtPrno")) Then
		strVal = " " & FilterVar(Trim(Request("txtPrno")), " " , "S") & " "		
	End If

    UNIValue(0,UBound(UNIValue,2)) = strval			  '	UNISqlId(0)의 마지막 ?에 입력됨	
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
	If Not(rs0.EOF Or rs0.BOF) Then
		Call  MakeSpreadSheetData()
	Else
		lgPageNo = ""
        rs0.Close
        Set rs0 = Nothing
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    lgDataExist    = "Yes"
  
	If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
	Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1

        iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sppl_cd"))	    
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_nm"))				 
		iRowStr = iRowStr & Chr(11) & UNIConvNum(rs0("quota_rate"),0)				         							
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("apportion_qty"),ggQty.DecPoint,0)			
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("pur_plan_dt"))		
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pur_grp"))	                   
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pur_grp_nm"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount    
		
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop

	lgstrData  = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If

    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing
End Sub

%>

<Script Language=vbscript>
    With parent
		If "<%=lgDataExist%>" = "Yes" Then
		   .ggoSpread.Source  = .frm1.vspdData2
		   .ggoSpread.SSShowData "<%=lgstrData%>"          '☜ : Display data
		   .lgPageNo2      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		   .DbQueryOk2
		End If  
	End with
</Script>	
 	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
