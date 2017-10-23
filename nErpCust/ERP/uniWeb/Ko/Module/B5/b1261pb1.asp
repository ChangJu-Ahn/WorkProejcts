<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : S1261PB1
'*  4. Program Name         : 거래처팝업 
'*  5. Program Desc         : 거래처정보의 거래처팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Cho inkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2000/12/09
'*                            2001/12/18  Date 표준적용 
'*							  2002/04/12 ADO 변환 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2      '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort 종류 

Dim strBiz_grp_nm											  ' 영업그룹명 
Dim strPur_grp_nm
Dim BlankchkFlg 											  ' 구매그룹명 

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(30)             '☜ : 한번에 가져올수 있는 데이타 건수 
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
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBiz_grp_nm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBiz_grp")) And BlankchkFlg = False  Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	
			BlankchkFlg = True
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strPur_grp_nm =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPur_grp")) And BlankchkFlg = False  Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	
			BlankchkFlg = True
		End If			
    End If     
     
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "B1261PA101"
    UNISqlId(1) = "s0000qa005"					'영업그룹명 
    UNISqlId(2) = "s0000qa019"					'구매그룹명    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtBp_cd")) Then
		strVal = "AND A.BP_CD LIKE " & FilterVar("%" & Trim(UCase(Request("txtBp_cd"))) & "%", "''", "S")	
	Else
		strVal = ""
	End If

	If Len(Request("txtBp_nm")) Then
		strVal = strVal & " AND A.BP_NM LIKE " & FilterVar("%" & Trim(UCase(Request("txtBp_nm"))) & "%", "''", "S")			
	End If		
		   
	If Len(Request("txtBiz_grp")) Then
		strVal = strVal & " AND A.BIZ_GRP = " & FilterVar(Trim(UCase(Request("txtBiz_grp"))), " " , "S") & " "		
		arrVal(0) = Trim(Request("txtBiz_grp"))
	End If		
    
 	If Len(Request("txtPur_grp")) Then
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPur_grp"))), " " , "S") & " "		
		arrVal(1) = Trim(Request("txtPur_grp")) 
	End If
	
	If Trim(Request("txtRadio2")) = "C" Or Trim(Request("txtRadio2")) = "S" Then
		strVal = strVal & " AND A.BP_TYPE LIKE  " & FilterVar("%" & Trim(Request("txtRadio2")) & "%", "''", "S") & ""		
	End If
	
	If Trim(Request("txtRadio3")) = "Y" Or Trim(Request("txtRadio3")) = "N" Then
		strVal = strVal & " AND A.USAGE_FLAG = " & FilterVar(Request("txtRadio3"), "''", "S") & ""		
	End If   	
	
	If Len(Request("txtOwnRgstN")) Then
		strVal = strVal & " AND A.BP_RGST_NO = " & FilterVar(Request("txtOwnRgstN"), "''", "S") & " "		
	End If   	
	
	
	UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal(0), " " , "S")									'영업그룹 
    UNIValue(2,0) = FilterVar(arrVal(1), " " , "S")									'구매그룹    
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))					  '☜: 표준적용대신 입력 
    UNILock = DISCONNREAD :	UNIFlag = "1"										  '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    
    Call  SetConditionData()

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And BlankchkFlg  =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>
<Script Language=vbscript>
    With parent
		.frm1.txtSales_grp_nm.value	= "<%=ConvSPChars(strBiz_grp_nm)%>" 
		.frm1.txtPur_grp_nm.value	= "<%=ConvSPChars(strPur_grp_nm)%>"
'		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.HBp_cd.value	= "<%=ConvSPChars(Request("txtBp_cd"))%>"
				.frm1.HBp_nm.value	= "<%=ConvSPChars(Request("txtBp_nm"))%>"
				.frm1.HBiz_grp.value= "<%=ConvSPChars(Request("txtBiz_grp"))%>"
				.frm1.HPur_grp.value= "<%=ConvSPChars(Request("txtPur_grp"))%>"				
				.frm1.HRadio2.value	= "<%=Request("txtRadio2")%>"
				.frm1.HRadio3.value	= "<%=Request("txtRadio3")%>"					
				.frm1.HOwn_Rgst_N.value	= "<%=ConvSPChars(Request("txtOwnRgstN"))%>"					
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '☜: Display data 																					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
