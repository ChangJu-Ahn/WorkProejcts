<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>


<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<% 
	Call LoadBasisGlobalInf() 

	On Error Resume Next														'☜: 
	Err.clear

	Dim lgOpModeCRUD
 																'☆ : 입력/수정용 ComProxy Dll 사용 변수
    Call HideStatusWnd                                                               '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)	
		
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
            Call SubBizQuery()
        
        Case CStr(UID_M0002)                                                         '☜: Save,Update
      		Call SubFileDownLoad()          
        
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubFileDownLoad2()
          
    End Select
    
'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
Sub SubFileDownLoad()
	On Error Resume Next														'☜: 
	Err.clear
	dim strFilePath                                                         '☜: Protect system from crashing
    
	Call HideStatusWnd
	strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
	If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
		strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
	End If
	strFilePath = strFilePath  & "files/u2000/"
	strFilePath = strFilePath & Request("txtFileName")
	%>
	<SCRIPT LANGUAGE=VBSCRIPT>
		Dim SF
		On Error Resume Next

		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%=strFilePath%>")
		Set SF = Nothing
	</SCRIPT>
<%


		strFilePath = strFilePath  & "files/u2000/"
		strFilePath = strFilePath & Request("txtFileName")
		
		
End Sub

'--------------------------------------------------------------------------------------------------------
'					            File DownLoad(With xinSoft solution)
'--------------------------------------------------------------------------------------------------------
Sub SubFileDownLoad2()
	On Error Resume Next    
	Err.Clear                                                               '☜: Protect system from crashing

	Dim xdn

	Call HideStatusWnd

	set xdn = Server.CreateObject("Xionsoft.XionFileDownLoad")
	strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/template/files/" & gCompany & "/"

	'Call ServerMesgBox(strFilePath, vbInformation, I_MKSCRIPT)

	'다음 두 줄 임시 COMMENT 2001/1/27
	xdn.DownFromFile strFilePath & Request("txtFileName")
	xdn.OnEndPage

	set xdn = nothing
	Response.End
End Sub


'--------------------------------------------------------------------------------------------------------
'					         
'--------------------------------------------------------------------------------------------------------

Sub SubBizQuery()

		On Error Resume Next														'☜: 
		Err.clear
		
		Dim ex1_file_name
		Dim ex2_return_code    
    
    	const I1_biz_area_cd			= 0
		const I2_start_issued_dt		= 1
		const I3_end_issued_dt			= 2
		const I4_report_issued_dt		= 3
		const I5_file_path_lef_supplied = 4	
		const I6_singoGubun				= 5
		const I7_yearmonth				= 6
		const I8_cnt					= 7	
		const I9_docamt					= 8
		const I10_locamt				= 9
		const I10_chkYn                 = 10
				
		Dim strGubun
		Dim arrValue
		Redim arrValue(I10_chkYn) 
		dim strFilePath
		Dim iPAVG012
		
		arrValue(I1_biz_area_cd)			= UCase(Trim(Request("txtBizAreaCD")))
		arrValue(I2_start_issued_dt)		= replace(UNIConvDate(Request("txtIssueDt1")),"-","")
		arrValue(I3_end_issued_dt)			= replace(UNIConvDate(Request("txtIssueDt2")),"-","")
		arrValue(I4_report_issued_dt)		= replace(UNIConvDate(Request("txtReportDt")),"-","")
		strFilePath							= Server.MapPath("../../files/u2000") & "\"
		arrValue(I5_file_path_lef_supplied) = strFilePath
		arrValue(I6_singoGubun)				= Trim(Request("txtSingoGubun"))
		arrValue(I7_yearmonth)				= Trim(Request("txtYearMonth"))
		arrValue(I8_cnt)					= UNIConvNum(Trim(Request("txtCnt")),0)
		arrValue(I9_docamt)					= UNIConvNum(Trim(Request("txtDocAmt")),0)
		arrValue(I10_locamt)				= UNIConvNum(Trim(Request("txtLocAmt")),0)
		arrValue(I10_chkYn)				    =  Request("chkYn")
		
		
		Set iPAVG012 = Server.CreateObject("PAVG012.cbExportRptDiskSvrEab")
		
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		
		If Err.Number <> 0 Then
			Set iPAVG012 = Nothing																'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙: 에러내용, 메세지타입, 스크립트유형
			Call HideStatusWnd
			Response.End																		'☜: Process End
		End If
		If Trim(gStrGlobalCollection) = "" Then
			Call DisplayMsgBox("127310", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
			Exit Sub			
		End If
		  
   
		call iPAVG012.EAB_ExportRpt_DISK_SVR(gStrGlobalCollection,arrValue,ex1_file_name,ex2_return_code) 
		If CheckSYSTEMError(Err, True) = True Then					
			   Set iPAVG012 = Nothing
			   Exit Sub
		End If    
    
		Select Case ex2_return_code
			Case "A"	
			Case "B"	
			Case "C"	
			Case "D"	
			Case "Z"	' 결과화일 다운로드 작업
			Call HideStatusWnd
			Set iPAVG012 = Nothing

%>
			<SCRIPT LANGUAGE=VBSCRIPT>
				parent.frm1.txtFileName.value = "<%=ex1_file_name%>"
				parent.subExportDiskOK("<%=ex1_file_name%>")
			</SCRIPT>
<%				
		End Select

		Call HideStatusWnd
		Set iPAVG012 = Nothing															'☜: Unload Comproxy

		Response.End   
       
End Sub
%>
	

