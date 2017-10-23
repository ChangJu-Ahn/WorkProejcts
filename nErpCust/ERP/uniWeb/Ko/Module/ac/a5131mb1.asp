<!-- #Include file="../../inc/IncServer.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next



Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                        '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim lgtxtAccountYear
Dim lgtxtBizArea
Dim lgtxtdeptcd
Dim lgtxtMaxRows

Dim biz_area_nm
Dim cost_nm
Dim dept_nm


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo       = Request("lgPageNo")                               '☜ : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
	
	lgtxtAccountYear		= FilterVar(Request("txtAccountYear"), "''", "S")
	lgtxtBizArea		= FilterVar(Trim(Request("txtBizArea")),"''" ,"S")
	lgtxtdeptcd			= FilterVar(Request("txtdeptcd"), "''", "S")
	lgtxtMaxRows		= Request("txtMaxRows")
  
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""


    If Len(Trim(lgPageNo))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
          
       End If   
    Else   
       lgPageNo = 0
    End If   

   
    'rs0에 대한 결과 
    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1
    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
    'rs1에 대한 결과 
    IF NOT (rs1.EOF or rs1.BOF) then
	    biz_area_nm = rs1("biz_area_nm")
    END IF
    rs1.Close
    Set rs1 = Nothing
    
   'rs2에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
		dept_nm = rs1("dept_nm")
    END IF
    rs2.Close
    Set rs2 = Nothing



End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(4)                                                    '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,5)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "a5131QA101"
    UNISqlId(1) = "ABIZNM"
    UNISqlId(2) = "ADEPTNM"
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    'rs0에 대한 Value값 setting    
    UNIValue(0,0) = lgSelectList  
    UNIValue(0,1)  = " " & FilterVar(gChangeOrgId, "''", "S") & ""
  	UNIValue(0,2)  = lgtxtAccountYear
	
	
	IF lgtxtdeptcd = "''" then
		UNIValue(0,3)  = ""
	Else 
		UNIValue(0,3)  = " AND A.DEPT_CD = " & lgtxtdeptcd 
	end if

	IF lgtxtBizArea = "''" then
		UNIValue(0,4)  = ""
	Else 
		UNIValue(0,4)  = " AND A.BIZ_AREA_CD = " & lgtxtBizArea
	end if	 
	
    'rs1에 대한 Value값 setting
	UNIValue(1,0) = lgtxtBizArea
	
	'rs2에 대한 Value값 setting
	IF lgtxtdeptcd = "''" then
		UNIValue(2,0)  = "" & FilterVar("XXXXX", "''", "S") & " "				'입력된 값이 없을때 더미값을 넘겨준다 
	Else 
		UNIValue(2,0)  = lgtxtdeptcd
	End if
	UNIValue(2,1) = " " & FilterVar(gChangeOrgId, "''", "S") & ""
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
       
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMsgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>

<Script Language=vbscript>
 
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
					.Frm1.htxtAccountYear.Value		= .Frm1.txtAccountYear.text					
					.Frm1.htxtBizArea.Value			= .Frm1.txtBizArea.Value
					.Frm1.htxtdeptcd.Value			= .Frm1.txtdeptcd.Value
			End If
       
        'Show multi spreadsheet data from this line       
        .ggoSpread.Source	= .frm1.vspdData      
        .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
        .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
       
       																	'☜: 화면 처리 ASP 를 지칭함 

		.frm1.txtBizAreaNm.value		= "<%=biz_area_nm%>"
		.frm1.txtdeptnm.value			= "<%=dept_nm%>"
		
	   End With
       
       
       Parent.DbQueryOk
    Else

	End if

</Script>	

