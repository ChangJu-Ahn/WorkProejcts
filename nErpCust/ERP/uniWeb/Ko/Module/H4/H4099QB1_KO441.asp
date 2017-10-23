<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
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
Dim arrRsVal(11)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
Dim iFrPoint
iFrPoint=0
Dim lgPageNo
Dim lgDataExist
'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "QB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

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
	Dim strSQL
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,3)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "h4099qa1_KO441"
 																		  'Reusage is Recommended
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	strSQL = ""     
    If Len(Trim(Request("txtValidDt"))) Then
			strSQL = strSQL & " AND convert(varchar(7),DILIG_DT,20) =  " & FilterVar(Trim(UCase(Request("txtValidDt"))), " " , "S") & " "
    End If
    If Len(Trim(Request("txtDeptNm"))) Then
			strSQL = strSQL & " AND DEPT_NM like  " & FilterVar(Trim(Request("txtDeptNm"))&"%", "''", "S") & ""
    End If
    If Len(Trim(Request("txtEmpNo"))) Then
			strSQL = strSQL & " AND EMP_NO =  " & FilterVar(Trim(UCase(Request("txtEmpNo"))), " " , "S") & " "
    End If


     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
	 UNIValue(0,2)  = FilterVar(Request("org_change_id"),"''","S")
 	 UNIValue(0,1)  = FilterVar(Request("txtValidDt")&"-01","''","S")
     UNIValue(0,3) = Trim(strSQL)	
	 	 
          
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

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
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         Parent.frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=lgstrData%>"             '☜ : Display data
                  '         
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
		 				 
         .DbQueryOk
         Parent.frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
