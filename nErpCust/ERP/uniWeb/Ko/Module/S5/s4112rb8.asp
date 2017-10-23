<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4112RA8
'*  4. Program Name         : 출하내역현황 참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'*                            -2002/12/20 : Get방식 --> Post방식으로 변경 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													

On Error Resume Next

Call LoadBasisGlobalInf()
'Call HideStatusWnd

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList													       '☜ : select 대상목록 
Dim lgSelectListDT														   '☜ : 각 필드의 데이타 타입	
'Dim strMode																   '☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
'Dim arrRsVal(1)												 			'☜ : QueryData()실행시 레코드셋을 배열로 받을때 사용 
																	    	'☜ : 받을 레코드셋의 갯수만큼 배열 크기 선언			
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Call HideStatusWnd 

'strMode        = Request("txtMode")

'Select Case strMode

'Case CStr(UID_M0001)

lgStrPrevKey   = Request("txt_lgStrPrevKey")                               '☜ : Next key flag
lgMaxCount     = 100										                   '☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("txt_lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("txt_lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("txt_lgTailList")                                 '☜ : Order by value

Call TrimData()
Call FixUNISQLData()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""


    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
						iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
  
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
																		  '아래에 보면 화면단에서 넣어 주는 query시 where조건임을 알 수 있다.	
    Redim UNISqlId(0)                                                        '☜: SQL ID 저장을 위한 영역확보 
																		  '조회화면에서 필요한 query조건문들의 영역(Statements table에 있음)
    Redim UNIValue(0,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S4112RA801"  ' main query(spread sheet에 뿌려지는 query statement)
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																		  '	UNISqlId(0)의 첫번째 ?에 입력됨				
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtConDnNo")) Then
		strVal = " " & FilterVar(Request("txtConDnNo"), "''", "S") & " "
	Else
		strVal = ""
	End If
	
    
    UNIValue(0,1) = strVal    '	UNISqlId(0)의 두번째 ?에 입력됨	
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'☜:ADO 객체를 생성 
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End     

    Else   

        Call  MakeSpreadSheetData()

    End If
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>

<Script Language=vbscript>
    With parent
       .frm1.vspdData.Redraw = False
        .ggoSpread.Source    = .frm1.vspdData
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"			 '☜: Display data 
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"		 '☜: set next data tag 
       .frm1.vspdData.Redraw = True
        <% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
		If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,NewTop) And .lgStrPrevKey <> "" Then
			.DbQuery
		Else
			.DbQueryOk
		End If
   	End with
</Script>	
 	
<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
'End Select
%>
