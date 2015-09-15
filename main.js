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

  $('#calcBtn').click(function (e) {
    if ( !($('#selectFile').val() && $('#sYear').val() && $('#sTerm').val()) ) {
      $('#calcAlertWrap').removeClass('mh0');
      setTimeout(function() {
        $('#calcAlertWrap').addClass('mh0');
      }, 1000);
    }
  })

  $("#selectBtn").change(function (e) {
    $("#selectFile").val($(this).val());
    if ($(this).val()) {
      $('#loadingWrap').removeClass('mh0');
      setTimeout(function(){
        handleFile(e);
      }, 250);
    }
  });

  function handleFile(e) {
    $('#loadingWrap').removeClass('mh0');
    var files = e.target.files;
    var i,f;
    for (i = 0, f = files[i]; i != files.length; ++i) {
      var reader = new FileReader();
      var name = f.name;
      reader.onload = function(e) {
        setTimeout(function() {
          $('#loadingWrap').addClass('mh0');
        }, 500);
        var data = e.target.result;

        var workbook = XLSX.read(data, {type: 'binary'});
        
        /* Get worksheet */
        var wsInput = workbook.Sheets['Input_보유'];

        /* Find desired cell */
        var wsInputA1 = wsInput['A1'];

      };
      reader.readAsBinaryString(f);
    }
  }
});
