<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출관리 
'*  3. Program ID           : s5111qb4
'*  4. Program Name         : 품목그룹별월매출현황 조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/10/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
On Error Resume Next                                                                         
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
	lgstrData = chr(11) & rs0.GetString(,,chr(11), chr(11) &chr(12) &chr(11))
	lgStrData = Left(lgStrData, len(lgStrData) - 1)

    lgPageNo = ""

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
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,12)

    UNISqlId(0) = "S5111QA501"
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = " " & FilterVar(UNIConvDate(Request("txtConFromDt")), "''", "S") & ""
	UNIValue(0,2) = " " & FilterVar(UNIConvDate(Request("txtConToDt")), "''", "S") & ""
	UNIValue(0,3) = " " & FilterVar(Request("txtBlFlag"), "''", "S") & ""				' B/L Flag
	UNIValue(0,4) = " " & FilterVar(Request("txtQueryData"), "''", "S") & ""
	UNIValue(0,5) = Trim(Request("txtConItemGroupLvl"))
	
	If Len(Trim(Request("txtConItemGroup"))) Then
		UNIValue(0,6) = " " & FilterVar(Request("txtConItemGroup"), "''", "S") & ""
	Else
		UNIValue(0,6) = "Default"
	End If
	UNIValue(0,7) = "Default"											' Usage Flag	
	UNIValue(0,8) = " " & FilterVar(Request("txtRootKey"), "''", "S") & ""				' Root Org.
	'UNIValue(0,9) = "'" & Trim(Request("txtOrgSuffix")) & "'"			' Org. Suffix
	UNIValue(0,9) = " " & FilterVar(Request("txtGrpSuffix"), "''", "S") & ""			' Grp. Suffix
	UNIValue(0,10) = "" & FilterVar("N", "''", "S") & " "												' Rate Flag

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
   lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        %>
		<Script Language=vbscript>
		Call parent.SetFocusToDocument("M")
		parent.frm1.txtConFromDt.focus
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If  
End Sub

%>

<Script Language=vbscript>
	On Error Resume Next
	Dim iArrRows, iArrCols
	Dim iIntRowCnt, iIntColCnt
	Dim iIntRow, iIntCol
	
    With parent.frm1
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area

			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source		= .vspdDataH 
			parent.ggoSpread.SSShowDataByClip "<%=ConvSPChars(lgstrData)%>"								'☜: Display data 

			parent.lgPageNo_A			=  "<%=lgPageNo%>"							'☜: Next next data tag
			parent.DbQueryOk
		End If
	End with
</Script>	
