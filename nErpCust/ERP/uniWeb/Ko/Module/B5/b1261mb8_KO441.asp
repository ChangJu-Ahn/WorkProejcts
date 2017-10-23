<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B1261MB8
'*  4. Program Name         : 거래처조회 
'*  5. Program Desc         : 거래처조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : kim hyung suk
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'**************************************************************************************
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

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '☜ : DBAgent Parameter 선언 
Dim rs1, rs2 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strtxtBpcdFrom	                                                       
Dim strtxtBpcdTo	                                                           

'----------------------- 추가된 부분 ----------------------------------------------------------------------
Dim arrRsVal(3)								'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array

'----------------------- 추가된 부분 ----------------------------------------------------------------------
Const C_SHEETMAXROWS_D  = 100                '☆: Server에서 한번에 fetch할 최대 데이타 건수 

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Call HideStatusWnd 
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'☜:
    lgStrPrevKey   = Request("lgStrPrevKey")								'☜ : Next key flag
	lgMaxCount     = CInt(C_SHEETMAXROWS_D)									'☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")								'☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")									'☜ : Orderby value
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
			IF ColCnt < 12 Then
            	iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
            Else
								'Call ServerMesgBox(Instr(rs0(ColCnt), chr(32)) , vbInformation, I_MKSCRIPT)            
            	iRowStr = iRowStr & Chr(11) & fnGetString(rs0(ColCnt))
        	End If
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        		'Call ServerMesgBox(rs0(12) , vbInformation, I_MKSCRIPT)	'>>air
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub


'++++++++++++++++++++++++++++++++++++
' 문자열변환
'++++++++++++++++++++++++++++++++++++
Function fnGetString(iStr)

	If not IsNull(iStr) And iStr <> "" Then
		iStr = Replace(iStr, "''", "'")
		iStr = replace(iStr,"&lt;","<")
		iStr = replace(iStr,"&gt;",">")
		iStr = replace(iStr,"&amp;","&")
		'iStr = replace(iStr, Chr(13)&Chr(10), "_")
		'iStr = replace(iStr, Chr(32), "_")
		
		fnGetString = iStr
	End if
	
End Function


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "B1261MA801KO441"		'>>air							'* : 데이터 조회를 위한 SQL문 만듬 
	
	UNISqlId(1) = "B1261MA802"			'거래처 
	UNISqlId(2) = "B1261MA802"			'거래처 
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
     
	UNIValue(1,0)  = UCase(Trim(strtxtBpcdFrom))
    UNIValue(2,0)  = UCase(Trim(strtxtBpcdTo))
    
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = ""
    
	If Trim(Request("txtBp_cdFrom")) <> "" Then
		strVal = strVal& " A.BP_CD >= " & FilterVar(Trim(UCase(Request("txtBp_cdFrom"))), " " , "S") & "  AND A.BP_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	Else
		strVal = strVal& " A.BP_CD >='' AND A.BP_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If

	If Trim(Request("txtRadioType")) <> "" Then
		strVal = strVal& " AND A.BP_TYPE LIKE  " & FilterVar("%" & Trim(UCase(Request("txtRadioType"))) & "%", "''", "S") & " "
	Else
		strVal = strVal& " AND A.BP_TYPE >='' AND A.BP_TYPE <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If
	
	If Trim(Request("txtBp_cdTo")) <> "" Then
		strVal = strVal& " AND A.BP_CD >='' AND A.BP_CD <=  " & FilterVar(Trim(UCase(Request("txtBp_cdTo"))), " " , "S") & " "
	Else
		strVal = strVal& " AND A.BP_CD >='' AND A.BP_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If
	
	If Trim(Request("txtRadioFlag")) <> "" Then
		strVal = strVal& " AND A.USAGE_FLAG >= " & FilterVar(UCase(Request("txtRadioFlag")), "''", "S") & " AND A.USAGE_FLAG <=  " & FilterVar(UCase(Request("txtRadioFlag")), "''", "S") & ""
	Else
		strVal = strVal& " AND A.USAGE_FLAG >='' AND A.USAGE_FLAG <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If
  		
    UNIValue(0,1) = strVal   
	
'================================================================================================================   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1,rs2) '* : Record Set 의 갯수 조정 
    
    Set lgADF   = Nothing

    
    iStr = Split(lgstrRetMsg,gColSep)
    
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtBp_cdFrom")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "거래처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True

            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBp_cdFrom.focus    
                </Script>
            <%     	       	
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtBp_cdTo")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "거래처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBp_cdTo.focus    
                </Script>
            <%	       
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	
		
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF  And BlankchkFlg =  False Then
		    Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBp_cdFrom.focus    
                </Script>
            <%		    
    
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

	'---거래처 
    If Len(Trim(Request("txtBp_cdFrom"))) Then
    	strtxtBpcdFrom = " " & FilterVar(Trim(Request("txtBp_cdFrom")), " " , "S") & " "
    	
    Else
    	strtxtBpcdFrom = "''"
    End If
    '---거래처 
    If Len(Trim(Request("txtBp_cdTo"))) Then
    	strtxtBpcdTo = " " & FilterVar(Trim(Request("txtBp_cdTo")), " " , "S") & " "
    Else
    	strtxtBpcdTo = "''"
    End If
		

End Sub

'response.write lgstrData
%>
<Script Language=vbscript>
    parent.frm1.txtBp_nmFrom.value	= "<%=ConvSPChars(arrRsVal(1))%>" 	
	parent.frm1.txtBp_nmTo.value	= "<%=ConvSPChars(arrRsVal(3))%>" 	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HBp_cdFrom.value	= "<%=ConvSPChars(Request("txtBp_cdFrom"))%>"
				.frm1.HBp_cdTo.value	= "<%=ConvSPChars(Request("txtBp_cdTo"))%>"
				.frm1.HRadioFlag.value	= "<%=Request("txtRadioFlag")%>"
				.frm1.HRadioType.value	= "<%=Request("txtRadioType")%>"
			End If
						
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"          '☜: Display data 
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.DbQueryOk
		
		End with
	
	End If   
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
