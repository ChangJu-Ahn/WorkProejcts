<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2211QB1
'*  4. Program Name         : 판매계획확정정보조회 
'*  5. Program Desc         : 판매계획확정정보조회 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2001/04/18
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>

<BODY bgColor=White><!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next
	
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                    '☜ : DBAgent Parameter 선언 
    Dim lgstrData                                                              '☜ : data for spreadsheet data
    Dim lgStrPrevKey                                                           '☜ : 이전 값 
    Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgConFlag
    Dim lgPageNo
    Dim lgConStep
    Dim lgConSalesGrp
    Dim lgConSpPeriod,lgStrOrgNm,lgConSpPeriodDesc
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgConFlag     = Request("cboConFlag")                                 '☜ : Orderby value
    lgDataExist    = "No"
    lgConStep  = Trim(Request("cboConStep"))
    lgConSalesGrp= Replace(Trim(Request("txtConSalesGrp")), "'", "''")
	lgConSpPeriod  = Replace(Trim(Request("txtConSpPeriod")), "'", "''")
	

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
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CInt(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CInt(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2) 
    Redim UNIValue(2,3)                                                         '☜: SQL ID 저장을 위한 영역확보 
    DIM strWhere
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
                                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
   strWhere="" 
   
	'영업그룹 or 공장 
	select case lgConFlag
    case "G"
		UNISqlId(0) = "S2211QA101" '영업그룹'
		
		If lgConStep <> "" Then
			strWhere = " and a.SP_STEP='" & lgConStep & "'"
		End if	
		
		if lgConSalesGrp <> "" Then
			UNISqlId(1) = "s0000qa005"							' 영업그룹 
			UNIValue(1,0) = FilterVar(lgConSalesGrp, "''", "S")	
			
		    strWhere = strWhere & " and a.SALES_GRP='" & lgConSalesGrp & "'"
        end if    
		
		if lgConSpPeriod <> "" Then
			UNISqlId(2) = "S0000QA029"							'계획기간 
			UNIValue(2,0) = FilterVar(Request("cboConSpType"), "''", "S")	 
			UNIValue(2,1) = FilterVar(lgConSpPeriod, "''", "S")	
        	strWhere = strWhere & " and a.FR_SP_PERIOD >='" & lgConSpPeriod & "'"
        End if
		
	Case "P"
		UNISqlId(0) = "S2211QA102" '공장'
		If lgConStep <> "" Then
			strWhere = " and a.SP_STEP='" & lgConStep & "'"
		End if	
		
		if lgConSalesGrp <> "" Then
		    UNISqlId(1) = "122700sab"							' 공장 
			UNIValue(1,0) = FilterVar(lgConSalesGrp, "''", "S")	
		    strWhere = strWhere & " and a.PLANT_CD='" & lgConSalesGrp & "'"
        end if    
        
		if lgConSpPeriod <> "" Then
			UNISqlId(2) = "S0000QA029"							'계획기간 
			UNIValue(2,0) = FilterVar(Request("cboConSpType") , "''", "S")	
			UNIValue(2,1) = FilterVar(lgConSpPeriod, "''", "S")	
        	strWhere = strWhere & " and a.FR_SP_PERIOD >='" & lgConSpPeriod & "'"
        End if
		
	end select	
	
	strWhere = strWhere & " AND a.SP_TYPE = '" & Request("cboConSpType")  & "'"
    
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(0,1)  = strWhere
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
       
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
	
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If  
    
    
   	' 영업그룹의 존재여부 
	lgStrOrgNm = ""   
	
    If  UNIValue(1,0) <> "" Then
  
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing
			
			select case lgConFlag
			case "G"
				Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			case "P"
				Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			end select 	
				%>
				<Script language=VBScript>
				Parent.frm1.txtConSalesGrp.focus  
			
				</Script>
			
				<%     	
			Exit Sub
		Else
			lgStrOrgNm = rs1(1)		' 영업그룹명 
			
			%>
			<Script language=VBScript>
				parent.frm1.txtConSalesGrpNm.Value = "<%=ConvSPChars(lgStrOrgNm)%>"
			</Script>
			<%
			
		End If
	Else
		%>
		<Script language=VBScript>
			parent.frm1.txtConSalesGrpNm.Value = ""
		</Script>
		<%
    End If
	'계획기간의 존재여부    
    lgConSpPeriodDesc=""
    If  UNIValue(2,0) <> "" Then
  
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing
			%>
			<Script language=VBScript>
			parent.frm1.txtConSpPeriodDesc.Value = ""
			</Script>
			<%     	
		Else
			lgConSpPeriodDesc = rs2(1)		' 계획기간 
			%>
			<Script language=VBScript>
				parent.frm1.txtConSpPeriodDesc.Value = "<%=ConvSPChars(lgConSpPeriodDesc)%>"
			</Script>
			<%
		End If
	Else
		%>
		<Script language=VBScript>
			parent.frm1.txtConSpPeriodDesc.Value = ""
		</Script>
		<%
    End If
     
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
			
        rs0.Close
        Set rs0 = Nothing
        %>
		<Script language=VBScript>
			Parent.frm1.cboConSpType.focus    
		</Script>
		<%     
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

%>

<Script Language=vbscript>
With Parent
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          .Frm1.hcboConStep.Value      = "<%=Request("cboConSpStep")%>"                  'For Next Search
          .Frm1.hcboConSpType.Value  = "<%=Request("cboConSpType")%>"
          .Frm1.htxtConSalesGrp.Value  = "<%=Request("txtConSalesGrp")%>"
          .Frm1.htxtConSpPeriod.Value  = "<%=Request("txtConSpPeriod")%>"
       End If
       
       'Show multi spreadsheet data from this line
		If "<%=lgConFlag%>" = "G" Then       
			.ggoSpread.Source  = .frm1.vspdData
		Else
			.ggoSpread.Source  = .frm1.vspdData2
		End If
		
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '☜ : Display data
		.lgPageNo			 =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.DbQueryOk
    End If
End With
</Script>	

