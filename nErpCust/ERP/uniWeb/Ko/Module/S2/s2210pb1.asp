<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2210PB1
'*  4. Program Name         : 품목 Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/12/27
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

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = CInt(30)                           '☜ : 한번에 가져올수 있는 데이타 건수 
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgTailList     = Request("lgTailList")                                 '☜ : Order by value
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgDataExist    = "No"

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

	    Dim iStrVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
																			  '아래에 보면 화면단에서 넣어 주는 query시 where조건임을 알 수 있다.	
		Dim iStrPlantCd
	    Redim UNISqlId(0)                                                        '☜: SQL ID 저장을 위한 영역확보 
	    Redim UNIValue(0,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

		iStrPlantCd = Trim(Request("txtPlantCd"))
		If iStrPlantCd = "" Then
			UNISqlId(0) = "S2210PA101"  ' main query(spread sheet에 뿌려지는 query statement)
		Else
			UNISqlId(0) = "S2210PA102"
		End If
	    
	    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																			  '	UNISqlId(0)의 첫번째 ?에 입력됨				
	    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

		iStrVal = ""
		
		If Len(Trim(Request("txtItemCd"))) Then
			iStrVal = iStrVal & " AND IT.ITEM_CD LIKE  " & FilterVar(Request("txtItemCd") & "%", "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtItemNm"))) Then
			iStrVal = iStrVal & " AND IT.ITEM_NM LIKE  " & FilterVar("%" & Request("txtItemNm") & "%", "''", "S") & ""
		End If

		If Len(Trim(Request("txtItemGroupCd"))) Then
			iStrVal = iStrVal & " AND IT.ITEM_GROUP_CD =  " & FilterVar(Request("txtItemGroupCd"), "''", "S") & ""
		End If

		If Len(Trim(Request("txtItemAcct"))) Then
			iStrVal = iStrVal & " AND IT.ITEM_ACCT =  " & FilterVar(Request("txtItemAcct"), "''", "S") & ""
		End If
		
		If Len(Trim(Request("txtItemSpec"))) Then
			iStrVal = iStrVal & " AND IT.SPEC LIKE  " & FilterVar("%" & Request("txtItemSpec") & "%", "''", "S") & ""
		End If
		
		If iStrPlantCd = "" Then		
			If Len(Trim(Request("txtFromDt"))) Then
				iStrVal = iStrVal & " AND IT.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
			End If
		
			If Len(Trim(Request("txtToDt"))) Then
				iStrVal = iStrVal & " AND IT.VALID_TO_DT >=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
			End If
		Else
			If Len(Trim(Request("txtFromDt"))) Then
				iStrVal = iStrVal & " AND ITP.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
			End If
		
			If Len(Trim(Request("txtToDt"))) Then
				iStrVal = iStrVal & " AND ITP.VALID_TO_DT >=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
			End If
			
			iStrVal = iStrVal & " AND ITP.PLANT_CD =  " & FilterVar(iStrPlantCd, "''", "S") & ""
		End If
		 UNIValue(0,1) = iStrVal				'UNISqlId(0)의 두번째 ?에 입력됨	
	  
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
			%>
			<Script Language=vbscript>
			parent.frm1.txtConItemCd.focus
			</Script>	
			<%
	    Else   
	        Call  MakeSpreadSheetData()
			Call  SetConditionData()
	    End If
	   
	End Sub

If lgDataExist = "Yes" Then
%>   

<Script Language=vbscript>
	With parent.frm1
		'Set condition data to hidden area
		' "1" means that this query is first and next data exists	
			<%If lgPageNo = "1" Then %>
			.txtHItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
			.txtHItemNm.value		= "<%=ConvSPChars(Request("txtItemNm"))%>"
			.txtHItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
			.txtHItemAcct.value		= "<%=Request("txtItemAcct")%>"
			.txtHItemSpec.value		= "<%=Request("txtItemSpec")%>"
			.txtHPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
			<%End If%>
	    'Show multi spreadsheet data from this line
		parent.ggoSpread.Source = .vspdData 
		parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"
		parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
        parent.DbQueryOk
	End With
</Script>	
<%End If%>
