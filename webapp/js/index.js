var page = {
	data : {

	},
	common : {
		/**
		 * 选择文件
		 * @return {[type]} [description]
		 */
		selectFile : function() {
			$('.files').remove();
			var input = '<input type="file" class="files">';
			$('body').append(input);
			$('.files').click();
			page.common.getFilesInfo();
		},
		/**
		 * 获取文件信息
		 * @return {[type]} [description]
		 */
		getFilesInfo : function() {
			$('.files').on('change', function() {
				var files = $('.files').prop("files");
				page.common.createAudio(files[0].name);
				console.log(files)
				/*if(files.length == 0 ) {
					alert("请选择文件");
					return;
				}else {
					var reader = new FileReader();
					reader.readAsText(files[0], "UTF-8");
					reader.onload = function(evt){ //读取完文件之后会回来这里
				        var fileString = evt.target;
				        console.log(fileString)
			    	}
				}*/
			})
		},
		/**
		 * 目前在chrome下只能播放audio文件夹下的音频文件
		 * [根据选择的文件创建音乐播放]
		 * @param  {[type]} filename [文件名]
		 * @return {[type]}     [description]
		 */
		createAudio : function(filename) {
			var audio = '<audio controls="controls">'
				  	+'<source src="./audio/'+filename+'" type="audio/ogg">'
				  	+'<source src="./audio/'+filename+'" type="audio/mpeg">'
					+'Your browser does not support the audio element.'
					+'</audio>';
			$('body').append(audio);
		},
		//点击插入音乐，弹出modal 如果url不为空，则可进行修改，删除和试听操作，如果url为空，则只能进行选择文件添加操作。  回调函数返回执行操作后的url
		getModule : function(url, callback) {
			$('#myModal').removeClass('fade').show();
			page.event.openFileSelect();
			page.common.closeModal();
			if(url == "") {

			}
			
		},
		//关闭modal 不执行任何操作
		closeModal : function() {
			$('.close').on('click', function() {
				$('#myModal').addClass('fade').hide();
			})
		}
	},
	event : {
		openFileSelect : function() {
			$('.openFileSelect').on('click', function() {
				page.common.selectFile();
			})
		}
	}
}


$(document).ready(function() {
	
	
	
	if (!(window.File || window.FileReader || window.FileList || window.Blob)) {
    	alert('你妈喊你换Chrome浏览器啦');
	}

	$('.insertAudio').on('click', function(e) {
		page.common.getModule()
	})
})