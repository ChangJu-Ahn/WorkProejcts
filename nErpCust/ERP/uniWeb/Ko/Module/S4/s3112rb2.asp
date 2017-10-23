<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3112rb2.asp																*
'*  4. Program Name         : 수주내역참조(Local L/C내역등록에서)										*
'*  5. Program Desc         : 수주내역참조(Local L/C내역등록에서)										*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2001/04/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'*                            2. 2002/07/13 : Ado 변환													*
'********************************************************************************************************

Response.Expires = -1							'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")   
Call LoadBNumericFormatB("I","*","NOCOOKIE","RB")

On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1          '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim strItemNm
   Dim lgCurrency
   Dim BlankchkFlg
   
	Const C_SHEETMAXROWS_D  = 30                                          '☆: Fetch max count at once
'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
   Dim iFrPoint
   iFrPoint=0

'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

    Call HideStatusWnd 
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"
	
	If Len(Request("txtCurrency")) Then
		lgCurrency	 = Request("txtCurrency")
	Else
		lgCurrency   = gCurrency
	End If	

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
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
        strItemNm =  rs1(1)
        rs1.Close
        Set rs1 = Nothing
    Else
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtItem")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			BlankchkFlg  =  True
		 %>
            <Script language=vbs>
            parent.frm1.txtItem.focus    
            </Script>
         <%		       						
		End If
	End If   	
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------


Sub FixUNISQLData()

    Dim strVal														  '☜:UNISqlId(0)에 들어가는 입력변수																		  '아래에 보면 화면단에서 넣어 주는 query시 where조건임을 알 수 있다.	
    Dim arrVal(1)														  '☜: 화면에서 팝업하여 query
    
																		  '아래에 보면 UNISqlId(1),UNISqlId(2), UNISqlId(3)의 where조건임을 알 수 있다.
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
																		  '조회화면에서 필요한 query조건문들의 영역(Statements table에 있음)
    Redim UNIValue(1,3)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S3112RA201"  ' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "s0000qa001"  ' popup에 의하여 query시 


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
							                                          '☜: Select list'	UNISqlId(0)의 첫번째 ?에 입력됨	
	UNIValue(0,0) = lgSelectList					'	UNISqlId(1)의 첫번째 ?에 입력됨	
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strVal = " "
	
	arrVal(0) = " " & FilterVar(Request("txtApplicant"), "''", "S") & " "

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " and S_SO_HDR.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If
	
	If Len(Request("txtItem")) Then		
		strVal = strVal & " and s_so_dtl.item_cd = " & FilterVar(Trim(Request("txtItem")), "" , "S") & " "
		arrVal(1) = Trim(Request("txtItem"))
	End If	
	
	If Len(Request("txtSONo")) Then		
		strVal = strVal & " and s_so_hdr.so_no = " & FilterVar(Request("txtSONo"), "''", "S") & " "
		
	End If	
	
	If Len(Request("txtCurrency")) Then		
		strVal = strVal & " and  S_SO_HDR.cur = " & FilterVar(Request("txtCurrency"), "''", "S") & " "		
	End If		  
	
	If Len(Request("txtRadio")) Then		
		strVal = strVal & " and  S_SO_HDR.lc_flag = " & FilterVar(Request("txtRadio"), "''", "S") & ""		
	End If	

        If Len(Request("txtTrackingNo")) Then		
		strVal = strVal & " and  S_SO_DTL.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "		
	End If		  	  

    UNIValue(0,1) = arrVal(0)    '	UNISqlId(0)의 두번째 ?에 입력됨	
    UNIValue(0,2) = strVal    '	UNISqlId(0)의 세번째 ?에 입력됨	
    UNIValue(1,0) = FilterVar(Trim(Request("txtItem")), " " , "S") '	UNISqlId(1)의 첫번째 ?에 입력됨	   
            
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

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    
    Call  SetConditionData()
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then         
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		 %>
            <Script language=vbs>
            parent.frm1.txtItem.focus    
            </Script>
         <%		    
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
	End If
    
End Sub

'수주번호가 없을 때 조회를 하지 않도록 해야한다.
%>

<Script Language=vbscript>

    parent.frm1.txtItemNm.Value	= "<%=ConvSPChars(strItemNm)%>"   

    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.txtHItem.value	= "<%=ConvSPChars(Request("txtItem"))%>"
                        parent.frm1.txtHTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
			parent.frm1.txtHSONo.value		= "<%=ConvSPChars(Request("txtSONo"))%>"			
       End If
       'Show multi spreadsheet data from this line
       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       
       parent.frm1.vspdData.Redraw = False
	   parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
					
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=lgCurrency%>",Parent.GetKeyPos("A",8),"C", "Q" ,"X","X")
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=lgCurrency%>",Parent.GetKeyPos("A",9),"A", "Q" ,"X","X")		
	   
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
       parent.frm1.vspdData.Redraw = True
    End If   
</Script>
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
