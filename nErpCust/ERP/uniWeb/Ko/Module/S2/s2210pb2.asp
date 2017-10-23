<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2210PB2
'*  4. Program Name         : 배분율 미등록 현황 Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                           '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList													       '☜ : select 대상목록 
Dim lgSelectListDT														   '☜ : 각 필드의 데이타 타입	
Dim lgDataExist
Dim lgPageNo
   
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
Call HideStatusWnd 

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)				'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = CInt(30)											'☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgTailList     = Request("lgTailList")                                 '☜ : Order by value
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgDataExist      = "No"

Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
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
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub
'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
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
	    Redim UNIValue(0,4)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

	    UNISqlId(0) = "S2210PA201"  ' main query(spread sheet에 뿌려지는 query statement)
	    
	    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																			  '	UNISqlId(0)의 첫번째 ?에 입력됨				
	    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
		UNIValue(0,1) = Trim(Request("txtFrSpPeriod"))
		UNIValue(0,2) = Trim(Request("txtToSpPeriod"))
		
		If Len(Trim(Request("txtSalesGrp"))) Then
			UNIValue(0,3) = " AND SALES_GRP =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""
		End If
	  
	     '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	    UNIValue(0,UBound(UNIValue, 2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
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
		
		Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'☜:ADO 객체를 생성 
	    
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)

	    Set lgADF   = Nothing    
	    iStr = Split(lgstrRetMsg,gColSep)
		
		
	    If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If    
	    
	    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
	        rs0.Close
	        Set rs0 = Nothing
	        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)	'No Data Found!!
	    Else   
	        Call  MakeSpreadSheetData()
			Call  SetConditionData()
	    End If
	   
	End Sub
%>   

<Script Language=vbscript>
	With parent.frm1
	
		<%If lgDataExist = "Yes" Then%>
	    'Show multi spreadsheet data from this line
		parent.ggoSpread.Source = .vspdData 
		parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"
		parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
        parent.DbQueryOk
		<%End If%>
	End With
</Script>	
<%
Response.End
%>
