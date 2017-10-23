<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3			'☜ : DBAgent Parameter 선언 
    Dim lgstrData															'☜ : data for spreadsheet data
    Dim lgFromWhere
    Dim lgPageNo
    Dim lgMaxCount
    Dim lgLngMaxRow
    Dim lgDataExist
    Dim ColNm
    Dim ColId
    Dim i
    
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

    Dim lgMaxColCnt			'컬럼 개수
    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd
    
    lgSelectList = Request("StrSelect_RUN")
    lgFromWhere  = ""
        
    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)			'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 1000											'☜ : 한번에 가져올수 있는 데이타 건수 
	
    lgMaxColCnt	= Request("txtColCnt")

    Call FixUNISQLData()
    
    Call QueryData()
   
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
 
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,1)

    UNISqlId(0) = "B1Z010MB1"

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	UNIValue(0,1) = lgFromWhere

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs( gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	Exit Sub
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        %>
		<Script Language=vbscript>
		Call parent.DbQueryOk
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        If lgPageNo = "1" Then Call SetConditionData()
        Call WriteResult()
    End If  
End Sub

Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   For i=0 To CINT(rs0.Fields.Count) - 1
		ColNm = ColNm & rs0.Fields.item(i).Name & Chr(11)
		ColId = ColId & i + 1 & Chr(11)
   Next
        
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
    	
		For ColCnt = 0 To lgMaxColCnt  - 1'UBound(lgSelectListDT) - 1 
				iRowStr = iRowStr & Chr(11) & FormatRsString("ED",rs0(ColCnt))
		Next

        If iLoopCount < lgMaxCount Then
'        		Response.Write iLoopCount & "@"
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
'           		Response.Write lgMaxCount & "#"
        Else
	       lgPageNo = lgPageNo + 1
'	       		Response.Write lgPageNo & "$"
           Exit Do
 '          		Response.Write lgPageNo & "%"
        End If
  '      		Response.Write lgPageNo & "^"
        rs0.MoveNext        		
	Loop
	
	'Call DisplayMsgBox("x", vbInformation, lgstrData & "==S1", "FASDFADS1111", I_MKSCRIPT)

    If iLoopCount < lgMaxCount Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
 	Response.Write ".vspdData.Redraw = False " & vbCr      
 	Response.Write "parent.SetColNm_New """ & ColNm & """,""" & ColId & """" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData  & """ ,""F""" & vbCr
	Response.Write ".lgPageNo.value	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
 	Response.Write ".vspdData.Redraw = True " & vbCr      
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub


%>


