<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111pb4
'*  4. Program Name         : 발주번호별진행조회 
'*  5. Program Desc         : 발주번호별진행조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgStrPrevKey                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
Dim ICount  		                                        '   Count for column index
Dim strPurGrpCd												'	구매그룹 
Dim strPurGrpCdFrom 										
Dim strBpCd                                               '   공급처 
Dim strBpCdFrom
Dim strPoFrDt                                               '   발주일 
Dim strPoToDt
Dim strItemCd                                               '   품목 
Dim strItemCdFrom
Dim strPoNo                                              '   Po No
Dim strPoNoFrom
Dim strPoType                                               '   발주형태 
Dim strPoTypeFrom	
Dim strCfmFlg								                '   단가구분 
Dim strCfmFlgFrom	
Dim arrRsVal(11)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
Dim lgPageNo
Dim lgDataExist
Dim iFrPoint
iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "QB")
	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

     Call  TrimData()                                                     '☜ : Parent로 부터의 데이타 가공 
     Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
     call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
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
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(1,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "m3111qa302"
								
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

     UNIValue(0,1)  = UCase(Trim(strPoNoFrom))			'---발주번호    
     
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
     
     UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By 조건 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
    
    '============================= 추가된 부분 =====================================================================
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		if Request("Query_Msg_Flg") = "T" then
			'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
         	Call HTMLFocus("Parent.Frm1.vspdData",I_MKSCRIPT)

		end if
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     
    '---Po No
    If Len(Trim(Request("txtPoNo"))) Then
    	strPoNo	= " " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & " "
    	strPoNoFrom = strPoNo
    Else
    	strPoNo	= "''"
    	strPoNoFrom = "" & FilterVar("%%", "''", "S") & ""
    End If

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData2
         Parent.frm1.vspdData2.Redraw = True
         .ggoSpread.SSShowData "<%=lgstrData%>", "F"                  '☜ : Display data
                  
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",8),"C", "Q" ,"X","X")	'발주금액 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",16),"A", "Q" ,"X","X")	'발주금액 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",18),"D", "Q" ,"X","X")	'환율 
         Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows, Parent.Parent.gCurrency , Parent.GetKeyPos("B",19),"A", "Q" ,"X","X")	'발주자국금액				'매입자국금액 
         
         .lgPageNo2			=  "<%=lgPageNo%>"               '☜ : Next next data tag
  		 
         .DbQueryOk(2)
         Parent.frm1.vspdData2.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
