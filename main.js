// global.$ = $;

// var nwGui = require('nw.gui');

var initMenu = function () {
  // try {
  //   var nwGui = require('nw.gui');
  //   var nativeMenuBar = new nwGui.Menu({type: "menubar"});
  //   if (process.platform == "darwin") {
  //     nativeMenuBar.createMacBuiltin && nativeMenuBar.createMacBuiltin("C1");
  //   }
  //   nwGui.Window.get().menu = nativeMenuBar;
  // } catch (error) {
  //   console.error(error);
  //   setTimeout(function () { throw error }, 1);
  // }
}

$(function() {

  initMenu();

  var sc = new Date();
  var $time = $('#time');

  var reader;
  var progress = document.querySelector('.percent');

  function updateTime() {
    $time.text(((new Date()) - sc)/1000);
    setTimeout(updateTime, 1000);
  }

  updateTime();


  $('#calcBtn').click(function (e) {
    if ( !($('#selectFile').val() && $('#sYear').val() && $('#sTerm').val()) ) {
      $('#calcAlertWrap').removeClass('mh0');
      setTimeout(function() {
        $('#calcAlertWrap').addClass('mh0');
      }, 1000);
    }
  });

  $("#selectBtn").change(function (e) {
    $("#selectFile").val($(this).val());
    if ($(this).val()) {
      $('#loadingWrap').removeClass('mh0');
      setTimeout(function(){
        handleFileSelect(e);
      }, 250);
    }
  });

  function abortRead() {
    reader.abort();
  }

  function errorHandler(evt) {
    switch(evt.target.error.code) {
      case evt.target.error.NOT_FOUND_ERR:
        alert('파일이 없습니다.');
        break;
      case evt.target.error.NOT_READABLE_ERR:
        alert('파일을 읽을 수 없습니다.');
        break;
      case evt.target.error.ABORT_ERR:
        break; // noop
      default:
        alert('파일을 읽는 도중 에러가 발생했습니다.');
    };
  }

  function updateProgress(evt) {
    // evt is an ProgressEvent.
    if (evt.lengthComputable) {
      var percentLoaded = Math.round((evt.loaded / evt.total) * 100);
      // Increase the progress bar length.
      if (percentLoaded < 100) {
        progress.style.width = percentLoaded + '%';
        progress.textContent = percentLoaded + '%';
      }
    }
  }

  function handleFileSelect(evt) {
    // Reset progress indicator on new file selection.
    progress.style.width = '0%';
    progress.textContent = '0%';

    reader = new FileReader();
    reader.onerror = errorHandler;
    reader.onprogress = updateProgress;
    reader.onabort = function(e) {
      alert('File read cancelled');
    };
    reader.onloadstart = function(e) {
      document.getElementById('progress_bar').className = 'loading';
    };
    reader.onload = function(e) {
      $('#loadingWrap').addClass('mh0');

      calc(e.target.result);
      // Ensure that the progress bar displays 100% at the end.
      progress.style.width = '100%';
      progress.textContent = '100%';
      setTimeout("document.getElementById('progress_bar').className='';", 2000);
    }

    // Read in the image file as a binary string.
    reader.readAsBinaryString(evt.target.files[0]);
  }

  function calc(data) {
    var workbook = XLSX.read(data, {type: 'binary'});
    var wsInput = workbook.Sheets['Input_보유'];
    var wsInputA1 = wsInput['A1'];
  }
});
