<%

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%						
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
' 아래 함수는 비지니스 로직 시작되는 시점에서 호출해 주세요..
Call HideStatusWnd		
On Error Resume Next	
Err.Clear    
                                                           '☜: Clear Error status
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2, rs3                         '☜ : DBAgent Parameter 선언 
	Dim lgstrData                                                              '☜ : data for spreadsheet data
	Dim lgStrPrevKey                                                           '☜ : 이전 값 
	Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
	Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

	dim arrCondi
	Dim strCond
    Dim strArno
    
    Dim iPrevEndRow
    Dim iEndRow   
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
	
    iPrevEndRow = 0
    iEndRow = 0
	
	arrCondi = split(Request("txtParentArno"), chr(11))
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
    
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
	
	    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    
    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1

 

    Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			IF ColCnt = 9  then '반제금액 셋팅 
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),"")
			Else					
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			ENd if		
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	

	rs0.Close
    Set rs0 = Nothing 
End Sub
 '----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "a3125RA2"
	UNISqlId(1) = "ABPNM"
	UNISqlId(2) = "ABPNM"
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(Request("htxtBpCd"), "''", "S") & "" )
    UNIValue(2,0) = UCase(" " & FilterVar(Request("htxtDealBpCd"), "''", "S") & "" ) 
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If UCase(Trim(Request("htxtBpCd"))) <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtBpNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
			End With
		</Script>
<%		
		Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtBpNm.value = ""
			End With
		</Script>
<%		
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs1.Close
		    Set rs1 = Nothing
			Exit sub
		End IF
	rs1.Close
	Set rs1 = Nothing
	End If   
        
	If UCase(Trim(Request("htxtDealBpCd"))) <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtdealBpNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
			End With
		</Script>
<%		
		Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtdealBpNm.value = ""
			End With
		</Script>
<%		
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
	rs2.Close
	Set rs2 = Nothing
	End If
	
    
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If


    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 
Sub TrimData()
	Dim ii
	arrCondi = split(Request("htxtParentArno"), chr(11))
	
	strArno = ""
	
	
	If UBound(arrCondi) > 0 Then	
		For ii = 0 To UBound(arrCondi)-1
	
			If ii<>UBound(arrCondi)-1 Then
				strArno= strArno & FilterVar(arrCondi(ii),null,"S") & ","
			Else
				strArno= strArno & FilterVar(arrCondi(ii),null,"S") 
			End If
		Next
	Else
		strArno="''"
	End If
	
	strCond=""


	strCond = " AND A.ar_dt >=  " & FilterVar(UNIConvDate(Request("htxtArDt")), "''", "S") & "" 
	strCond = strCond & " AND A.ar_dt <=  " & FilterVar(UNIConvDate(Request("htxtToArDt")), "''", "S") & "" 
	strCond = strCond & " AND A.doc_cur = " & FilterVar(UCase(Request("htxtDocCur")), "''", "S")
	
	
	If "" & UCase(Trim(Request("htxtBpCd"))) <> "" Then
		strCond = strCond & " AND A.pay_bp_cd = " & FilterVar(UCase(Request("htxtBpCd")), "''", "S")
	End If
		
	If "" & UCase(Trim(Request("htxtBizCd"))) <> "" Then			
		strCond = strCond & " AND A.biz_area_cd = " & FilterVar(UCase(Request("htxtBizCd")), "''", "S") 
	End If
	
	If "" & UCase(Trim(Request("htxtDealBpCd"))) <> "" Then			
		strCond = strCond & " AND A.deal_bp_cd = " & FilterVar(UCase(Request("htxtDealBpCd")), "''", "S")
	End If
	
	If "" & UCase(Trim(Request("htxtArNo"))) <> "" Then			
		strCond = strCond & " AND A.ar_no = " & FilterVar(UCase(Request("htxtArNo")), "''", "S")
	End If



	strCond= strCond & "AND A.ar_no NOt IN (" & strArno & ")"
				
End Sub


%>

	
<Script Language="vbscript">   
	With parent

 	   .ggoSpread.Source	= .frm1.vspdData
	   .frm1.vspdData.Redraw		= False
       .ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

       Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",4),.GetKeyPos("A",6),"A", "I" ,"X","X")
       Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",4),.GetKeyPos("A",8),"A", "I" ,"X","X")
       Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",4),.GetKeyPos("A",10),"A", "I" ,"X","X")
        .frm1.lgPageNo.value      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.frm1.vspdData.Redraw		= True
		.DbQueryOk
 
	
	End With
</Script>


