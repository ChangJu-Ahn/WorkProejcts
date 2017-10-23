<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb3
'*  4. Program Name         : 고정자산취득내역등록 
'*  5. Program Desc         : 고정자산별 취득내역을 삭제 
'*  6. Comproxy List        : +As0041ManageSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : 김희정 
'* 10. Modifier (Last)      : 김희정 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%	
On Error Resume Next													
Call HideStatusWnd														

    Call LoadBasisGlobalInf()
    
    '-- Common --
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
'    lgOpModeCRUD      = Request("txtMode")       
    
'-------------------------
' 변수, 상수 선언 
'-------------------------
	Dim iPAAG010																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

	'Import Variant
	Dim I3_a_asset_acq
	Dim I4_ief_supplied
	Dim E1_a_asset_master
	Dim E3_a_asset_acq
	Dim IG2_import_itm_grp

	'Import Const
	'View Name : import a_asset_acq
	Public Const A504_I3_acq_no = 0
	Public Const A504_I3_acq_fg = 2
	Public Const A504_I3_ap_no = 18
	Public Const A504_I3_gl_no = 19

	'View Name : import_mode_fg ief_supplied
	Public Const A504_I4_select_char = 0

	'View Name : export a_asset_master
	Public Const A504_E1_asst_no = 0

	'View Name : export a_asset_acq
	Public Const A504_E3_acq_no = 0    

'-------------------------   
' 업무 처리 
'-------------------------
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then											'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
	ElseIf Request("txtAcqNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'삭제 조건값이 비어있습니다!           
		Response.End 
	End If


	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqMngSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If    
	
	Redim I3_a_asset_acq(30)
	Redim I4_ief_supplied(0)
	
    I3_a_asset_acq(A504_I3_acq_no) = Trim(Request("txtAcqNo"))
    I3_a_asset_acq(A504_I3_acq_fg) = Request("cboAcqFg")
    I3_a_asset_acq(A504_I3_gl_no)  = Trim(Request("txtGLNo"))
    I3_a_asset_acq(A504_I3_ap_no)  = Trim(Request("txtApNo"))
    
	I4_ief_supplied(A504_I4_select_char) = "D"
	
	E1_a_asset_master = Request("txtSpread_m")		'Master Data Spread
	IG2_import_itm_grp = Request("txtSpread_i")		'취득상세내역 Spread
		
	call iPAAG010.AS0021_ACQ_MANAGE_SVR(gStrGloBalCollection, _
										, , I3_a_asset_acq, I4_ief_supplied, E1_a_asset_master, IG2_import_itm_grp, , , , , _
										, E3_a_asset_acq)
            
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
    Response.Write " parent.DbDeleteOk()		" & vbCr
    Response.Write " </Script>					" & vbCr
%>