<%@ page language="java" contentType="text/html; charset=utf-8" pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf"%>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>체결 계약서 등록서</H2>
	<form id="frm" name="frm" enctype="multipart/form-data">
		<table class="board_view">
			<colgroup>
				<col width="20%">
				<col width="20%" />
				<col width="20%" />
				<col width="20%" />
				<col width="20%" />
			</colgroup>
			<tbody>
				<tr>
					<th scope="row">사업부</th>
					<td colspan="3"><input type="text" id="BU_NM" name="BU_NM" class="pop_up" readonly="readonly"></input> <input type="hidden" id="BU_CODE"
						name="BU_CODE"
					></td>
					<th scope="row">해지조건</th>
					<td colspan="1"><select id="EXPIRE_CONDITION" name="EXPIRE_CONDITION">
					</select></td>
				</tr>
				<tr>
					<th scope="row">고객사(업체)명</th>
					<td><input type="text" id="BP_NM1" name="BP_NM1" class="pop_up"></input> <input type="hidden" id="BP_CD1" name="BP_CD1"></td>
					<th scope="row">고객사(업체)명2</th>
					<td><input type="text" id="BP_NM2" name="BP_NM2" class="pop_up"></input> <input type="hidden" id="BP_CD2" name="BP_CD2"></td>
					<th scope="row">구분</th>
					<td><select id="BP_TYPE" name="BP_TYPE"></select></td>
				</tr>
				<tr>
					<th scope="row">계약서 명</th>
					<td colspan="5"><input type="text" id="CONTRACT_NM" name="CONTRACT_NM" style="width: 90%"></input></td>
				</tr>
				<tr>
					<th scope="row">계약 구분</th>
					<td><input type="text" id="CONTRACT_TYPENM" name="CONTRACT_TYPENM" class="pop_up" readonly="readonly"></input> <input type="hidden"
						id="CONTRACT_TYPE" name="CONTRACT_TYPE"
					></td>
					<th scope="row">목적사업(제품)</th>
					<td colspan="3"><input type="text" id="PURPOSE" name="PURPOSE" style="width: 90%"></input></td>
				</tr>
				<tr>
					<th scope="row">효력발생일</th>
					<td><input type="text" id="VALID_DT" name="VALID_DT" class="DATE" readonly="readonly"></input></td>
					<th scope="row">기간만료일</th>
					<td><input type="text" id="EXPIRE_DT" name="EXPIRE_DT" class="DATE" maxlength="10" onkeypress="auto_date_format(event,this)"
						onkeyup="auto_date_format(event,this)" readonly="readonly"
					></input></td>
					<th scope="row">자동연장</th>
					<td><select id="EXTEND_FLAG" name="EXTEND_FLAG">
							<option value="Y">Y</option>
							<option value="N">N</option>
					</select> <select id="EXTEND_TERM" name="EXTEND_TERM">
					</select></td>
				</tr>
				<tr>
					<th scope="row">부속계약서</th>
					<td colspan="3"><input type="text" id="P_CONTRACT" name="P_CONTRACT" style="width: 90%"></input></td>
					<th scope="row">해지통지기간</th>
					<td><select id="EXPIRE_TERM" name="EXPIRE_TERM"></select></td>
				</tr>
				<tr>
					<th scope="row">비고</th>
					<td colspan="5"><textarea id="REMARK" name="REMARK" rows="3" cols="100"></textarea></td>
				</tr>
				<tr>
					<th scope="row">
						<div>파일 첨부</div>
						<div>
							<p id="file">
								<input type="image" id="btn_add" src="../images/btn_add.png" alt="추가" />
							</p>
							<p id="file">
								<input type="image" id="btn_delete" src="../images/btn_delete.png" alt="삭제" />
							</p>
						</div>
					</th>
					<td colspan="5">
						<div id="fileDiv">
							<input type="file" id="file" name="file_0">
						</div>
					</td>
				</tr>
			</tbody>
		</table>
		<br />
		<div style="width: 60%; list-style: inside; text-align: left; margin: auto;">
			<a href="#this" class="btn" id="write">작성하기</a> <a href="#this" class="btn" id="list">목록으로</a>
		</div>
	</form>
	<img alt="Loading" id="imgLoading" src="../images/loading_spinner.gif" class="loadingSpinner_Layer">
	<%@ include file="/WEB-INF/include/include-body.jspf"%>
	<script type="text/javascript">
		var g_id = "";
 		var cnt = 1;
		var userid = "<sec:authentication property='principal.username'/>";

		$(document).ready(
				function() {
					$.ajax({
						url : "../sample/openinitBox.do",
						type : "post",
						dataType : "json",
						success : function(responseData) {

							console.log(responseData.list);

							$.each(responseData.list, function(index, value) {
								if (value.HIGH_CODE == "C") {
									$("#BP_TYPE").append(
											"<option value='"+ value.CODE +"'>"
													+ value.CODE_NM
													+ "</option>");
								} else if (value.HIGH_CODE == "D") {
									$("#EXPIRE_CONDITION").append(
											"<option value='"+ value.CODE +"'>"
													+ value.CODE_NM
													+ "</option>");
								} else if (value.HIGH_CODE == "F") {
									if (index == 0) {
										$("#EXTEND_TERM").append(
												"<option value='"+ value.CODE +"'>"
														+ value.CODE_NM
														+ "</option>");
									} else
										$("#EXTEND_TERM").append(
												"<option value='"+ value.CODE +"'>"
														+ value.CODE_NM
														+ "</option>");
								} else if (value.HIGH_CODE == "G") {
									$("#EXPIRE_TERM").append(
											"<option value='"+ value.CODE +"'>"
													+ value.CODE_NM
													+ "</option>");
								}
							});
						}
					});

					$("#list").on("click", function(e) { //목록으로 버튼
						e.preventDefault();
						fn_openBoardList();
					});

					$("#write").on("click", function(e) { //작성하기 버튼
						e.preventDefault();
						if ($("#write").attr("disabled") != "disabled")	fn_insertBoard();
					});

					$('input[class=pop_up]').on("click", function(e) {
						e.preventDefault();
						fn_popup(e.target.id);
					});

					$.datepicker.regional['ko'] = {
						closeText : '닫기',
						prevText : '이전달',
						nextText : '다음달',
						currentText : '오늘',
						monthNames : [ '1월(JAN)', '2월(FEB)', '3월(MAR)',
								'4월(APR)', '5월(MAY)', '6월(JUN)', '7월(JUL)',
								'8월(AUG)', '9월(SEP)', '10월(OCT)', '11월(NOV)',
								'12월(DEC)' ],
						monthNamesShort : [ '1월', '2월', '3월', '4월', '5월', '6월',
								'7월', '8월', '9월', '10월', '11월', '12월' ],
						dayNames : [ '일', '월', '화', '수', '목', '금', '토' ],
						dayNamesShort : [ '일', '월', '화', '수', '목', '금', '토' ],
						dayNamesMin : [ '일', '월', '화', '수', '목', '금', '토' ],
						weekHeader : 'Wk',
						dateFormat : 'yy-mm-dd',
						firstDay : 0,
						isRTL : false,
						showMonthAfterYear : true,
						yearSuffix : ''
					};
					$.datepicker.setDefaults($.datepicker.regional['ko']);

					$('input[class=DATE]').datepicker({
						changeMonth : true,
						changeYear : true,
						showButtonPanel : true
					});

					$('input[id=btn_add]').on("click", function(e) {
						e.preventDefault();
						if (cnt > 4) {
							Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
							{
								msg : "파일은 최대 5건 첨부 가능합니다."
							});
						} else {
							fn_addFile(e.target.id);
						}
					});

					$('input[id=btn_delete]').on("click", function(e) {
						e.preventDefault();

						if (cnt <= 1) {
							Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
							{
								msg : "파일은 한개 이상 첨부 가능합니다."
							});
						} else {
							var file = "file_" + cnt;
							fn_deleteFile(file);
						}
					});

					$('#EXPIRE_CONDITION').change(
							function(e) {
								e.preventDefault();
								var str = "";
								$("#EXPIRE_CONDITION option:selected").each(
										function(e) {
											str = $(this).val();
										});
								fn_view(str);
							});

					$('#EXTEND_FLAG').change(
							function(e) {
								e.preventDefault();
								$("#EXTEND_FLAG option:selected").each(
										function(e) {
											if ($(this).val() == "Y") {
												$('#EXTEND_TERM').attr("style",
														"visibility: visible");
											} else {
												$('#EXTEND_TERM').attr("style",
														"visibility: hidden");
											}
										});
							});
					
					//기간만료일은 효력발생일보다 이전일 수 없음
					$('#EXPIRE_DT').bind('change', function(){
						var vValid_dt = $('#VALID_DT').val();
						var vExpird_dt = $('#EXPIRE_DT').val(); 
						
						if(vValid_dt.length == 0){
							//alert('효력발생일이 먼저 입력되어야 합니다.');
							Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
					   	    		 {
					   	    		     msg: "효력발생일이 먼저 입력되어야 합니다."
					   	    		 });
							$('#EXPIRE_DT').val("");
							return;
						}
						
						if(vValid_dt > vExpird_dt){
							//alert('기간만료일은 효력발생일 이후여야 합니다.');
							Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "기간만료일은 효력발생일 이후여야 합니다."
			   	    		 });
							$('#EXPIRE_DT').val("");
							return;							
						}
					});

				});

		function fn_view(id) {
			if (id == "D0001") {
				$("#VALID_DT").attr("style", "background-color: white");
				$("#EXPIRE_DT").attr("style", "background-color: white");
				$("#VALID_DT").attr("disabled", false);
				$("#EXPIRE_DT").attr("disabled", false);
				$("#VALID_DT").val("");
				$("#EXPIRE_DT").val("");
			}
			else{
				$("#VALID_DT").attr("style", "background-color: white");
				$("#EXPIRE_DT").attr("style", "background-color: #f0f0f0");
				$("#VALID_DT").attr("disabled", false);
				$("#EXPIRE_DT").attr("disabled", true);
				$("#VALID_DT").val("");
				$("#EXPIRE_DT").val("");
			}
		}

		function fn_popup(id) {
			g_id = id;

			if (g_id == "CONTRACT_TYPENM" || g_id == "BU_NM") //16.07.20, ahncj : 사업부도 계층으로 관리하고 싶다는 현업의 요청으로 수정 ()	
				var popUrl = "../sample/openGubunPopup.do?id=" + g_id + ""; //팝업창에 출력될 페이지 URL
			else
				var popUrl = "../sample/openPopup.do?id=" + g_id + ""; //팝업창에 출력될 페이지 URL

			var popOption = "width=370, height=360, resizable=no, scrollbars=yes, status=no;";
			var newWindow = window.open(popUrl, "", popOption);
			newWindow.focus();
		}

		function fn_addFile(e) {
			cnt++;
			var str = "<input type='file' id='file' name='file_"+cnt+"'/>";
			$("#fileDiv").append(str);
		}

		function fn_deleteFile(e) {
			console.log(cnt);
			$('input[name=' + e + ']').remove();
			cnt--;
		}

		function getReturnValue(returnValue) {
			if (g_id == "BU_NM") {
				$("#BU_NM").val(returnValue[1]);
				$("#BU_CODE").val(returnValue[0]);
			} else if (g_id == "BP_NM1") {
				$("#BP_NM1").val(returnValue[1]);
				$("#BP_CD1").val(returnValue[0]);
			} else if (g_id == "BP_NM2") {
				$("#BP_NM2").val(returnValue[1]);
				$("#BP_CD2").val(returnValue[0]);
			} else if (g_id == "CONTRACT_TYPENM") {
				$("#CONTRACT_TYPENM").val(returnValue[1]);
				$("#CONTRACT_TYPE").val(returnValue[0]);
			}
		}

		function auto_date_format(e, id) {
			
			var num_arr = [ 97, 98, 99, 100, 101, 102, 103, 104, 105, 96, 48,
					49, 50, 51, 52, 53, 54, 55, 56, 57 ];

			var key_code = (e.which) ? e.which : e.keyCode;

			if (num_arr.indexOf(Number(key_code)) != -1) {
				var len = id.value.length;
				if (len == 4)
					id.value += "-";
				if (len == 7)
					id.value += "-";
			}
		}

		function fn_openBoardList() {
			var comSubmit = new ComSubmit();
			comSubmit.setUrl("<c:url value='/sample/openBoardList.do' />");
			comSubmit.submit();
		}

		function fn_insertBoard() {
			var comSubmit = new ComSubmit("frm");
			var valid = fn_validCheck();

			if (valid == true){
				$("#imgLoading").show();
				$("#write").attr("disabled",true);
				
				comSubmit.setUrl("<c:url value='/sample/insertBoard.do' />");
				comSubmit.addParam("userid", userid);
				comSubmit.submit();
			}

		}
		
		function fn_validCheck(){

			if ($("#BU_CODE").val().length < 1) {
				//alert("사업부를 선택 해주세요.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
		   	    		 {
		   	    		     msg: "사업부를 선택 바랍니다."
		   	    		 });
				$("#BU_NM").attr("style", "border: 1px solid #ff0000;");
			} 
			else if ($("#BP_CD1").val().length < 1) {
				//alert("고객사명을 선택 해주세요.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "고객사명을 선택 바랍니다."
			   	    		 });
				$("#BP_NM1").attr("style", "border: 1px solid #ff0000;");
			} 
			else if ($("#CONTRACT_NM").val().length < 1) {
				//alert("계약서명을 작성 해주세요.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "계약서명을 작성 바랍니다."
			   	    		 });
				$("#CONTRACT_NM").focus();
			} 
			else if ($("#CONTRACT_TYPE").val().length < 1) {
				//alert("계약구분을 선택 해주세요.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "계약구분을 선택 바랍니다."
			   	    		 });
				$("#CONTRACT_TYPENM").attr("style", "border: 1px solid #ff0000;");
			} 
			else if ($("#PURPOSE").val().length < 1) {
				//alert("목적사업을 작성 해주세요.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "목적사업을 작성 바랍니다."
			   	    		 });
				$("#PURPOSE").focus();
			} 
			else if (!$("#VALID_DT").attr("disabled") == true && $("#VALID_DT").val().length < 1) {
					//alert("효력발생일 선택 해주세요.");
					Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "효력발생일을 선택 바랍니다."
			   	    		 });
					$("#VALID_DT").attr("style", "border: 1px solid #ff0000;");
			} 
			else if (!$("#EXPIRE_DT").attr("disabled") == true && $("#EXPIRE_DT").val() < 1) {
					//alert("기간만료일 선택 해주세요.");
					Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "기간만료일을 선택 바랍니다."
			   	    		 });
					$("#EXPIRE_DT").attr("style", "border: 1px solid #ff0000;");
			} 
			else {
				if ($("#EXTEND_FLAG option:selected").val() == "N")
					$('#EXTEND_TERM option:selected').val("");

				return true;
			}
		}
		
	</script>
</body>
</html>