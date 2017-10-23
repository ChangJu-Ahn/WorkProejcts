<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111RB2.ASP
'*  4. Program Name         : 운송정보참조 
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
    lgMaxCount     = 30							                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = " ORDER BY A.TRANS_INFO_NO "
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
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(1,2)

    UNISqlId(0) = "S4112RA101"									'* : 데이터 조회를 위한 SQL문 만듬 

	
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
'	UNIValue(1,0)  = UCase(Trim(strShipToParty))
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
    
    

    

	If Trim(Request("txtTransCo")) <> "" Then
		strVal = strVal & " AND A.TRANS_CO LIKE " & FilterVar("%" & Trim(Request("txtTransCo")) & "%", "''", "S") & ""	
	Else
		strVal = strVal & ""
	End If


	If Trim(Request("txtSender")) <> "" Then
		strVal = strVal & " AND A.SENDER LIKE " & FilterVar("%" & Trim(Request("txtSender")) & "%", "''", "S") & ""	
	Else
		strVal = strVal & ""
	End If


	If Trim(Request("txtVehicleNo")) <> "" Then
		strVal = strVal & " AND A.VEHICLE_NO LIKE " & FilterVar("%" & Trim(Request("txtVehicleNo")) & "%", "''", "S") & ""	
	Else
		strVal = strVal & ""
	End If



	If Trim(Request("txtDriver")) <> "" Then
		strVal = strVal & " AND A.DRIVER LIKE " & FilterVar("%" & Trim(Request("txtDriver")) & "%", "''", "S") & ""	
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
                Parent.frm1.txtTransCo.focus    
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


End Sub


%>
<Script Language=vbscript>
     	

	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHTransCo.value	 =  "<%=ConvSPChars(Request("txtTransCo"))%>"
				.frm1.txtHSender.value	 =  "<%=ConvSPChars(Request("txtSender"))%>" 	
				.frm1.txtHVehicleNo.value	 =  "<%=ConvSPChars(Request("txtVehicleNo"))%>"
				.frm1.txtHDriver.value	 =  "<%=ConvSPChars(Request("txtDriver"))%>" 	

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

