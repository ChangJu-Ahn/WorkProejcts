<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>계약 변경/해지합의서</H2>
	<form id="frm" name="frm" enctype="multipart/form-data">
        <table class="board_view" >
            <colgroup>
                <col width="20%">
                <col width="20%"/>
                <col width="20%"/>
                <col width="20%"/>
            </colgroup>
            <tbody>
	            <tr>
	            	<th scope="row">계약구분</th>
	            	<td colspan="3" style="">     				
						<input type="radio" id = "modify" name="GUNUN" onclick="fn_viewReset(this)" style="display: inline;">변경
						<input type="radio" id = "end" name="GUNUN"  onclick="fn_viewReset(this)" style="display: inline;">해지							
	            	</td>
	            </tr>
                <tr class = "trID" style="visibility: hidden;">
                    <th scope="row">고객사(업체)명</th>
                    <td><input type="text" id="BP_NM1" name="BP_NM1" readonly="readonly" style="background-color: #f0f0f0;"></input>
                    <input type = "hidden" id = "BP_CD1" name="BP_CD1">
                    </td>
                    <th scope="row">고객사(업체)명2</th>
                    <td><input type="text" id="BP_NM2" name="BP_NM2" readonly="readonly" style="background-color: #f0f0f0;"></input>
                    <input type = "hidden" id = "BP_CD2" name="BP_CD2">
                    </td>
                </tr>
                <tr class = "trID" style="visibility: hidden;">
                    <th scope="row">원계약서</th>
                    <td colspan="3"><input type="text" id="CONTRACT_NM" name="CONTRACT_NM" style="width: 85%" class="pop_up"></input>
                    <input type = "hidden" id = "CONTRACT_NO" name="CONTRACT_NO"></td>
                </tr>
                <tr class = "trID" style="visibility: hidden;">
                    <th scope="row">효력발생일</th>
                    <td><input type="text" id="VALID_DT" name="VALID_DT" class="DATE" readonly="readonly"></input></td>
                    <th scope="row">기간만료일</th>
                    <td><input type="text" id="EXPIRE_DT" name="EXPIRE_DT" class="DATE" maxlength="10" onkeypress="auto_date_format(event,this)" onkeyup="auto_date_format(event,this)"></input></td>
                </tr>
                <tr class = "trID" style="visibility: hidden;">
                    <th scope="row">변경 내용</th>
                    <td><input type="text" id="MODIFY_CONTENT" name="MODIFY_CONTENT"></input></td>
                    <th scope="row">해지일</th>
                    <td><input type="text" id="END_DT" name="END_DT" class="DATE" readonly="readonly"></input></td>
                </tr>
                <tr class = "trID" style="visibility: hidden;">
                    <th scope="row">비고</th>
                    <td colspan="3"><textarea id ="REMARK" name ="REMARK" rows="3" cols="100" ></textarea></td>
                </tr>
                <tr class = "trID" style="visibility: hidden;">
                   <th scope="row">
						<div>파일 첨부</div>
						<div>
							<p id = "file">
								<input type="image" id ="btn_add" src="../images/btn_add.png" alt="추가" />
							</p>
							<p id = "file">
								<input type="image" id ="btn_delete"  src="../images/btn_delete.png" alt="삭제" />
							</p>
						</div>
					</th>
					<td colspan="5">
						<div id="fileDiv">
							<input type="hidden" id="FILE_SEQ" name="FILE_SEQ" />
							<INPUT TYPE="FILE" ID="FILE" NAME="FILE_0">		
				        </div>
					</td>
                </tr>
            </tbody>
        </table>
        <br/>
        <div id="btnDiv" style="display: none">
	        <div style="width: 60%;list-style: inside; text-align: left;margin: auto;">
		   		<a href="#this" class="btn" id="write" >작성하기</a>
	        	<a href="#this" class="btn" id="list" >목록으로</a>
	    	</div>
    	</div>
    	<div style="display: none">
			<input type = "hidden" id = "userid" name="userid">
    	</div>
    </form>
    <img alt="Loading" id="imgLoading" src="../images/loading_spinner.gif" class="loadingSpinner_Layer">
    <%@ include file="/WEB-INF/include/include-body.jspf" %> 
    <script type="text/javascript">
        //var g_id = "";
        var cnt = 1;
        var userid = "<sec:authentication property='principal.username'/>";
        
        $(document).ready(function(){               

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
       	
       	  $("#userid").val(userid.trim());
       	  
             $("#list").on("click", function(e){ //목록으로 버튼
                 e.preventDefault();
                 fn_openBoardList();
             });
              
             $("#write").on("click", function(e){ 
			  var retnCheck = false;  
                 
           	  e.preventDefault();
           	  retnCheck = fn_validCheck();
                 
           	  if(retnCheck == true)
			  	fn_modifyBoard();
             });
             
             $('input[class=pop_up]').on("click", function(e){ 
                 e.preventDefault();
                 //fn_popup(e.target.id);
                 fn_popup();
             });
             
             $('input[id=btn_add]').on("click", function(e) {
				e.preventDefault();
				fn_addFile(e.target.id);
		  });
			
		 $('input[id=btn_delete]').on("click", function(e) { 
			e.preventDefault();
			
			if(cnt < 1){
				//alert("파일은 한개 이상 첨부 가능합니다.");
				Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
		   	    		 {
		   	    		     msg: "파일은 한개 이상 첨부 가능합니다."
		   	    		 });
				return;
			}
			else{
				var file = "file_"+cnt;
				fn_deleteFile(file);
			}
		 });
             
             $.datepicker.regional['ko'] = {
				  closeText: '닫기',
				  prevText: '이전달',
				  nextText: '다음달',
				  currentText: '오늘',
				  monthNames: ['1월(JAN)','2월(FEB)','3월(MAR)','4월(APR)','5월(MAY)','6월(JUN)',
				  '7월(JUL)','8월(AUG)','9월(SEP)','10월(OCT)','11월(NOV)','12월(DEC)'],
				  monthNamesShort: ['1월','2월','3월','4월','5월','6월',
				  '7월','8월','9월','10월','11월','12월'],
				  dayNames: ['일','월','화','수','목','금','토'],
				  dayNamesShort: ['일','월','화','수','목','금','토'],
				  dayNamesMin: ['일','월','화','수','목','금','토'],
				  weekHeader: 'Wk',
				  dateFormat: 'yy-mm-dd',
				  firstDay: 0,
				  isRTL: false,
				  showMonthAfterYear: true,
				  yearSuffix: ''};
				  $.datepicker.setDefaults($.datepicker.regional['ko']);
				 
				  $('input[class=DATE]').datepicker({ 
					   changeMonth: true,
					   changeYear: true,
					   showButtonPanel: true
				  });
                             
        });
        
        function fn_validCheck(){
        	
           	if($("input[id=modify]").is(":checked") == true)
           	{
           		if ($("#CONTRACT_NO").val().length < 1){
           			//alert("원계약서가 선택되지 않았습니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "원 계약서가 선택되지 않았습니다."
			   	    		 });
           			$("#CONTRACT_NO").focus();
           		}
           		else if ($("#VALID_DT").val().length < 1){
           			//alert("효력발생일을 입력 바랍니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "효력발생일을 입력 바랍니다."
			   	    		 });
           			$("#VALID_DT").focus();
           		}
           		else if ($("#EXPIRE_DT").val().length < 1){
           			//alert("기간만료일이 입력되어야 합니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "기간만료일이 입력되어야 합니다."
			   	    		 });
           			$("#EXPIRE_DT").focus();
           		}
           		else if ($("#MODIFY_CONTENT").val().length < 1){
           			//alert("변경내용이 입력되어야 합니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "변경내용이 입력되어야 합니다."
			   	    		 });
           			$("#MODIFY_CONTENT").focus();
           		}
           		else
           			return true;
           		
           	}
           	else if($("input[id=end]").is(":checked") == true)
           	{
           		if ($("#CONTRACT_NO").val().length < 1){
           			//alert("원계약서가 선택되지 않았습니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "원 계약서가 선택되지 않았습니다."
			   	    		 });
           			$("#CONTRACT_NO").focus();
           		}
           		else if ($("#END_DT").val().length < 1){
           			//alert("해지일이 입력되어야 합니다.");
           			Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
			   	    		 {
			   	    		     msg: "해지일이 입력되어야 합니다."
			   	    		 });
           			$("#END_DT").focus();
           		}
           		else
           			return true;
           	}
        }
        
        function fn_viewReset(input){
        	var id = input.id;
        	//var cnt = $("tbody > tr").length;
        	$("#btnDiv").attr("style", "display:inline;")
        	
			$("#VALID_DT").val("");
       		$("#EXPIRE_DT").val("");
       		$("#MODIFY_CONTENT").val("");

       		fn_view(id);
        }
        
        function fn_view(id){
        	
	    	if(id == "modify"){
	    		/* for(var i = 1; i < cnt; i++){
	    			$("tbody > tr").eq(i).attr("style","visibility: visible");
	    		} 
	    		$("input[name=file]").attr("style","visibility: visible");
	    		*/
	          	$(".trID").attr("style","visibility: visible");
	    		        		
	    		$("#VALID_DT").attr("style","background-color: white");
	    		$("#EXPIRE_DT").attr("style","background-color: white");
	    		$("#END_DT").attr("style","background-color: #f0f0f0");
	    		$("#MODIFY_CONTENT").attr("style","background-color: white");
	    		
	    		$("#VALID_DT").attr("disabled",false);
	    		$("#EXPIRE_DT").attr("disabled",false);
	    		$("#END_DT").attr("disabled",true);
	    		$("#MODIFY_CONTENT").attr("disabled",false);
	
	    	}else if(id == "end"){
	    		/* for(var i = 1; i < cnt; i++){
	    			$("tbody > tr").eq(i).attr("style","visibility: visible");
	    		}
	    		
	    		//$("tbody > tr").eq(6).attr("style","visibility: hidden");
	    		//$("input[name=file]").attr("style","visibility: hidden"); */
	    		
	    		$(".trID").attr("style","visibility: visible");
	    		
	    		$("#VALID_DT").attr("style","background-color: #f0f0f0");
	    		$("#EXPIRE_DT").attr("style","background-color: #f0f0f0");
	    		$("#END_DT").attr("style","background-color: white");
	    		$("#MODIFY_CONTENT").attr("style","background-color: #f0f0f0");
	    		
	    		$("#VALID_DT").attr("disabled",true);
	    		$("#EXPIRE_DT").attr("disabled",true);
	    		$("#END_DT").attr("disabled",false);
	    		$("#MODIFY_CONTENT").attr("disabled",true);
	    	}
	    	
        }
        
        function fn_popup(){
       
        	var popUrl = "../sample/openContractPopup.do?userid=" + userid;
			var popOption = "width=450, height=360, resizable=no, scrollbars=yes, status=no;";
			var newWindow = window.open(popUrl,"",popOption);
			newWindow.focus();
			
        }
        
        function fn_addFile(e){
			cnt++;
			var str = "<input type='file' id='file' name='file_"+cnt+"'/>";
            $("#fileDiv").append(str);
            console.log(cnt);
        }
		
		function fn_deleteFile(e){
			console.log(cnt);
			$('input[name='+e+']').remove();
            cnt--;
        }
        
         function getReturnValue(returnValue) {
        	for(var i = 0 ; i < returnValue.length ; i ++){
        		$("#"+returnValue[i].split("=")[0].trim()+"").val(returnValue[i].split("=")[1].trim());
        		
        		/*if(returnValue[i].split("=")[0].trim() == "ORIGINAL_FILE_NAME"){
        			console.log(returnValue[i].split("=")[0].trim());
        			$("#"+returnValue[i].split("=")[0].trim()+"").html("기존 파일 : "+returnValue[i].split("=")[1].trim());
        		}*/
        	}
        	/*
        	if($("input[name=file]").val() != "" || $("input[name=file]").val() != null){
        		$("input[name=file]").val("");
        	}*/
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
		
        function fn_openBoardList(){
            var comSubmit = new ComSubmit();
            comSubmit.setUrl("<c:url value='/sample/openBoardList.do' />");
            comSubmit.submit();
        }
         
        function fn_modifyBoard(){
            
            var comSubmit = new ComSubmit("frm");
            
           	if($("input[id=modify]").is(":checked") == true){
           		$("#imgLoading").show();
				$("#write").attr("disabled",true);
           		
           		comSubmit.setUrl("<c:url value='/sample/modifyBoard.do' />");
           		comSubmit.submit();
           	}else if($("input[id=end]").is(":checked") == true)
           	{
           		$("#imgLoading").show();
				$("#write").attr("disabled",true);
           		
           		comSubmit.setUrl("<c:url value='/sample/endBoard.do' />");
           		comSubmit.submit();
           	}else
           	{
           		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
		   	    		 {
		   	    		     msg: "계약구분을 선택 바랍니다."
		   	    		 });
           		//alert("계약구분을 선택 해주세요.");
           	}
            
        }
    </script>
</body>
</html>
