<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111RB1.ASP
'*  4. Program Name         : 납품처상세정보 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/17
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : SON BUM YEOL
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%   

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "*", "NOCOOKIE", "PB")   
Call LoadBNumericFormatB("Q","*","NOCOOKIE","PB")

On Error Resume Next

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '☜ : DBAgent Parameter 선언 
Dim rs1
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim BlankchkFlg

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strShipToParty	                                                       
'----------------------- 추가된 부분 ----------------------------------------------------------------------
Dim arrRsVal(1)							'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
'----------------------------------------------------------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
    Call HideStatusWnd 
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'☜:
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = 30							                          '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = " ORDER BY A.STP_INFO_NO DESC "
    'lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist    = "No"
	
	Call TrimData()
	Call FixUNISQLData()
	Call QueryData()
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
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
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                              '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    
	Dim strVal
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(2,2)

    UNISqlId(0) = "S4111RA101"									'* : 데이터 조회를 위한 SQL문 만듬 
    UNISqlId(1) = "B1261MA802"			'납품처 
	
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
	UNIValue(1,0)  = UCase(Trim(strShipToParty))
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
    
    
	If Trim(Request("txtShipToParty")) <> "" Then
		strVal = strVal & " AND A.SHIP_TO_PARTY = " & FilterVar(Request("txtShipToParty"), "''", "S") & "  "
	End If
	


	If Trim(Request("txtReceiver")) <> "" Then
		strVal = strVal & " AND A.RECEIVER LIKE " & FilterVar("%" & Trim(Request("txtReceiver")) & "%", "''", "S") & ""	
	Else
		strVal = strVal & ""
	End If


    UNIValue(0,1) = strVal   

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    
	Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
	Dim FalsechkFlg
    
    FalsechkFlg = False
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtShipToParty")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
	    %>
            <Script language=vbs>
            Parent.frm1.txtShipToParty.focus    
            </Script>
        <%		   		
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
		
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		    %>
                <Script language=vbs>
                Parent.frm1.txtReceiver.focus    
                </Script>
            <%
		    Exit Sub
		' 이 위치에 있는 Response.End 를 삭제하여야 함. Client 단에서 Name을 모두 뿌려준 후에 Response.End 를 기술함.
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

 '---납품처 

    If Len(Trim(Request("txtShipToParty"))) Then
    	strShipToParty = " " & FilterVar(Request("txtShipToParty"), "''", "S") & " "
    	
    Else
    	strShipToParty = "''"
    End If



End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtShipToPartyNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	

	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHShipToParty.value	 =  "<%=ConvSPChars(Request("txtShipToParty"))%>"
				.frm1.txtHReceiver.value	 =  "<%=ConvSPChars(Request("txtReceiver"))%>" 	
			End If
			.ggoSpread.Source    = .frm1.vspdData
			.frm1.vspdData.Redraw = False 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"          '☜: Display data 
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.frm1.vspdData.Redraw = True
			.DbQueryOk
		
		End with
	
	End If   
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

