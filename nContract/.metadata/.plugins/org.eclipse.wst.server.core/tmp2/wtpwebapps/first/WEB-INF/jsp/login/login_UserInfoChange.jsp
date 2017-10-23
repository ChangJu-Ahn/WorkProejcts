<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%@ include file="/WEB-INF/include/include-header.jspf"%>

<!-- 로그인만 특별하게 CSS, JS 관리 -->
<link rel="stylesheet" type="text/css" href="../css/login/style.css">
<link rel="stylesheet" type="text/css" href="../css/login/reset.css" >
<link rel="stylesheet" href="../css/lobibox.min.css"/>

<script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
<script src="../js/common.js" charset="utf-8"></script>
<script src="../js/lobibox.min.js"></script>

</head>
<body>
<c:choose>
	<c:when test = "${change == 'true'}">
		<script type="text/javascript">
		 Lobibox.alert("info", //AVAILABLE TYPES: "error", "info", "success", "warning"
	    		 {
	    		     msg: "패스워드를 정상적으로 변경하였습니다.",
	    		     callback: function () {self.close();}
	    		 });
		 </script>
	</c:when>
	<c:when test = "${change == 'false'}">
		<script type="text/javascript">
		 Lobibox.alert("error", //AVAILABLE TYPES: "error", "info", "success", "warning"
	    		 {
	    		     msg: "패스워드가 변경되지 않았습니다.</br>개인정보를 확인하시길 바랍니다.",
	    		     callback: function () {self.close();}
	    		 });
		 </script>
	 </c:when>
	 <c:otherwise>
	<!-- 정상로직일 경우 하단부의 form생성 후 비밀번호 변경창을 생성 함 -->
	 </c:otherwise>
</c:choose>
<div style="display: none;">
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
</div>
<form id="frm" name="frm" enctype="multipart/form-data">
	<div class="container">
		<div class="card alt">
			<div class="toggle"></div>
			<h1 class="title">
				패스워드 변경
			</h1>
			<!-- <form> -->
				<div class="input-container">
					<input type="text" id="USER_NO" name="USER_NO"/> 
					<label for="Username">ID(사번)</label>
					<div class="bar"></div>
				</div>
				<div class="input-container">
					<input type="password" id="USER_PWD" name="USER_PWD" /> 
					<label for="Password">비밀번호</label>
					<div class="bar"></div>
				</div>
				<div class="input-container">
					<input type="password" id="REPEAT_PASSWORD" name="REPEAT_PASSWORD" />
					<label for="Repeat Password">비밀번호 확인</label>
					<div class="bar"></div>
				</div>
				<div class="input-container">
					<input type="password" id="CHANGE_PASSWORD" name="CHANGE_PASSWORD" />
					<label for="Repeat Password">변경 비밀번호</label>
					<div class="bar"></div>
				</div>
				<div class="button-container">
					<button id="change">
						<span>변경</span>
					</button>
				</div>
		</div>
	</div>
</form>
<script type="text/javascript">
	var userId = "<sec:authentication property='principal.username'/>";	
	
	$(document).ready(function() {
		fn_initSetting();
		
		$("#change").on('click', function(e){
   			e.preventDefault();
   			fn_openChkChangePasswrd();
   		});
		
		//user_no의 input box를  readonly처리할 경우 css가 겹치는 현상이 발생 됨,
		//그러기에 user_no쪽으로 포커스가 이동 될 경우 알람창 및 사번초기화를 진행
		$('#USER_NO').focus(function(){
			Lobibox.alert("info", //AVAILABLE TYPES: "error", "info", "success", "warning"
		    		 {
		    		     msg: "사번은 변경할 수 없습니다.",
		    		     callback: function () {
		    		    	 	$('#USER_NO').val(userId.toString());	
		    		    	 	$('#USER_PWD').focus();
		    		    	 }
		    		 });
		});
		
	});
	
	function fn_initSetting()
	{
		$('.container').stop().addClass('active'); //자등으로 화면생성
		$('#USER_NO').val(userId.toString());
	}
	
    function fn_openChkChangePasswrd(){
    	if($("input[id = 'USER_NO']").val().length == 0){
    		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
   	    		 {
   	    		     msg: "[ID(사번)]란을 입력해주세요."
   	    		 });
    	}else if($("input[id = 'USER_PWD']").val().length == 0){
    		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
   	    		 {
   	    		     msg: "[비밀번호]란을 입력해주세요."
   	    		 });
    	}else if($("input[id = 'REPEAT_PASSWORD']").val().length == 0){ 
    		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
      	    		 {
      	    		     msg: "[비밀번호 확인]란을 입력해주세요."
      	    		 });
    	}else if($("input[id = 'CHANGE_PASSWORD']").val().length < 5){
    		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
      	    		 {
      	    		     msg: "변경될 비밀번호는 최소 6자리 이상입니다."
      	    		 });
    	}
    	else{
     		 var comSubmit = new ComSubmit("frm");
     		 
             comSubmit.setUrl("<c:url value='/admin/initUserInfo.do'/>");
             comSubmit.addParam("type", "U");			//진행 타입 -> I: 초기화(initial), U: 수정(update)
             
             comSubmit.submit();   
    	}                     
    }
</script>
</body>
</html>