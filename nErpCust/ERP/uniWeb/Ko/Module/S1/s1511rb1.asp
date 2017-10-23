<%
'======================================================================================================
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 								     										*
'*  3. Program ID           : S1511RB1																	*
'*  4. Program Name         : 품목참조										         					*
'*  5. Program Desc         : 품목그룹별품먹구성비등록을 위한 품목참조									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2002/05/08																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : Cho inkuk																	*
'* 10. Modifier (Last)      : 																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->

<%

On Error Resume Next

	Call LoadBasisGlobalInf()

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0			   '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo   


   Call HideStatusWnd 
     
   lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
   lgMaxCount       = CInt(50)											'☜ : 한번에 가져올수 있는 데이타 건수 
   lgSelectList     = Request("lgSelectList")
   lgTailList       = Request("lgTailList")
   lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
   lgDataExist      = "No"

   Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
   Call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
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

    If iLoopCount < lgMaxCount Then                                     '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close															'☜: Close recordset object
    Set rs0 = Nothing													'☜: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "S1511RA101"    
     
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "


    If Len(Request("txtItemGroup")) Then
		strVal = strVal & " A.ITEM_GROUP_CD =  " & FilterVar(Request("txtItemGroup"), "''", "S") & " "
	Else
		strVal = ""
	End If
	
	If Len(Request("txtItem")) Then
		strVal = strVal & " AND A.ITEM_CD LIKE " & FilterVar("%" & Trim(Request("txtItem")) & "%", "''", "S") & " "	
	End If

	If Len(Request("txtItemNm")) Then
		strVal = strVal & " AND A.ITEM_NM LIKE " & FilterVar("%" & Trim(Request("txtItemNm")) & "%", "''", "S") & " "
	End If

	If Len(Request("txtItemSpec")) Then
		strVal = strVal & " AND A.SPEC LIKE " & FilterVar("%" & Trim(Request("txtItemSpec")) & "%", "''", "S") & " "
	End If

	
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

    UNIValue(0,1) = strVal  
   
'================================================================================================================   
   
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    'UNIValue(0,UBound(UNIValue,2)) = " " & "Order By ISNULL(B.ITEM_RATE, 0) DESC, A.ITEM_CD ASC, A.ITEM_NM ASC"			'☜: 표준적용대신 입력 
    UNILock = DISCONNREAD :	UNIFlag = "1"						'☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
 
    If rs0.EOF And rs0.BOF Then
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
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.HItem.value		= "<%=ConvSPChars(Request("txtItem"))%>"
		    parent.frm1.HItemNm.value	= "<%=ConvSPChars(Request("txtItemNm"))%>"			
		    parent.frm1.txtHItemSpec.value	= "<%=ConvSPChars(Request("txtItemSpec"))%>"			
       End If
       'Show multi spreadsheet data from this line       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>"          '☜ : Display data
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
    End If   
</Script>	
