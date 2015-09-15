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

  var reader;
  var progress = document.querySelector('.percent');

  var workbook;
  var totalLines;
  var resultM = [[
      '계약년', '계약년월', '경과년', 'I/N', '납입기간', '인풋라인', 'H/R', 
      '기간', '년월', 'TSN', '보유한도금액', '위험률', '할인율', 'Lapse',
      '할인후출재보험료', '재보험금'
    ]];
  var resultA = [];
  var resultR = [];
  var resultN = [];

  function updateTime() {
    $time.text(((new Date()) - sc)/1000);
    setTimeout(updateTime, 1000);
  }

  // updateTime();


  $('#calcBtn').click(function (e) {
    if ( !($('#selectFile').val() && $('#sYear').val() && $('#sTerm').val()) ) {
      $('#calcAlertWrap').removeClass('mh0');
      setTimeout(function() {
        $('#calcAlertWrap').addClass('mh0');
      }, 1000);

      return;
    }

    calc();
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

  $('#exportXl').click(function(e) {
    function datenum(v, date1904) {
      if(date1904) v+=1462;
      var epoch = Date.parse(v);
      return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    }
     
    function sheet_from_array_of_arrays(data, opts) {
      var ws = {};
      var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
      for(var R = 0; R != data.length; ++R) {
        for(var C = 0; C != data[R].length; ++C) {
          if(range.s.r > R) range.s.r = R;
          if(range.s.c > C) range.s.c = C;
          if(range.e.r < R) range.e.r = R;
          if(range.e.c < C) range.e.c = C;
          var cell = {v: data[R][C] };
          if(cell.v == null) continue;
          var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
          
          if(typeof cell.v === 'number') cell.t = 'n';
          else if(typeof cell.v === 'boolean') cell.t = 'b';
          else if(cell.v instanceof Date) {
            cell.t = 'n'; cell.z = XLSX.SSF._table[14];
            cell.v = datenum(cell.v);
          }
          else cell.t = 's';
          
          ws[cell_ref] = cell;
        }
      }
      if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
      return ws;
    }
     
    /* original data */
    var data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]
    var ws_name = "macro";
     
    function Workbook() {
      if(!(this instanceof Workbook)) return new Workbook();
      this.SheetNames = [];
      this.Sheets = {};
    }
     
    var wb = new Workbook(), ws = sheet_from_array_of_arrays(resultM);
     
    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});

    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "result.xlsx");
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

  function updateProgress(line) {
    // evt is an ProgressEvent.
    var percentLoaded = Math.round((line / totalLines) * 100);
    // Increase the progress bar length.
    if (percentLoaded < 100) {
      progress.style.width = percentLoaded + '%';
      progress.textContent = percentLoaded + '%';
    }
  }

  function handleFileSelect(evt) {
    reader = new FileReader();
    reader.onerror = errorHandler;
    reader.onabort = function(e) {
      alert('File read cancelled');
    };
    reader.onload = function(e) {
      parseFileAsync(e.target.result, {type: 'binary'}, function(wb) {
        $('#loadingWrap').addClass('mh0');
        workbook = wb;
      });
    }
    reader.readAsBinaryString(evt.target.files[0]);
  }

  function calc(data) {
    var sc = new Date();
    var $time = $('#time');
    progress.style.width = '0%';
    progress.textContent = '0%';
    document.getElementById('progress_bar').className = 'loading';

    var 
      wsSp1 = workbook.Sheets['가정1_Lapse']
      ,wsSp2 = workbook.Sheets['가정2_Retention']
      ,wsSp3 = workbook.Sheets['가정3_위험율']
      ,wsSp4 = workbook.Sheets['가정4_할인율']
      ,wsSp5 = workbook.Sheets['가정5_TSN율']
      ,wsSp6 = workbook.Sheets['가정6_재보험금지급율']
      ,wsInput = workbook.Sheets['Input_보유']
      ,cc
      ,sYear = parseInt($('#sYear').val(), 10)
      ,sTerm = parseInt($('#sTerm').val(), 10)
      ,line = 2
      ,loop = true
      ,cell
      ,tYm
      ,ty
      ,sex
      ,age
      ,rrCode
      ,rgCode
      ,ggYear
      ,niGigan
      ,tM
      ,sM
      ,bHr = true
      ,hr
      ,tsn
      ,boyu
      ,rr
      ,rs
      ,i
      ,ii
      ,lpsYear
      ,lps
      ,rVal
      ,hVal
      ,jVal
      ,wsJson = XLSX.utils.sheet_to_json(wsInput, {
      header: 1,
      raw: true
    });

    totalLines = wsJson.length - 1;

    while((wsInput['K' + line]) && (cell = wsInput['K' + line]) && cell.v && loop) {
      updateProgress(line);

      tYm = cell.v;
      tY = parseInt((tYm+'').substr(0,4), 10);
      sex = wsInput['Q' + line].v;
      age = wsInput['R' + line].v;
      rrCode = wsInput['S' + line].v;
      rgCode = wsInput['A' + line].v;
      ggYear = wsInput['L' + line].v;
      niGigan = wsInput['H' + line].v;

      for(i = 0; i <= sTerm; i++) {
        sYearI = sYear + i;
        ggYearI = ggYear + i;
        tM = tY % 10;
        sM = sYearI % 10;

        if ( tY < sYearI && tM === sM ) {
          bHr = true;
        }

        if (bHr === true) {
          hr = 'R';
        } else {
          hr = 'H';
        }

        //TSN
        ii = 2;
        while(wsSp5['A' + ii]) {
          if (wsSp5['A' + ii].v === ggYearI) {
            if (wsSp5['B' + ii].v === niGigan) {
              tsn = wsSp5['C' + ii].v;
              break;
            }
          }
          ii++;
        }

        //보유한도금액
        ii = 2;
        while(wsSp2['A' + ii]) {
          if (wsSp2['A' + ii].v <= tYm) {
            if (wsSp2['B' + ii].v >= tYm) {
              if (wsSp2['C' + ii].v <= ggYearI) {
                if (wsSp2['D' + ii].v >= ggYearI) {
                  boyu = wsSp2['E' + ii].v;
                  break;
                }
              }
            }
          }
          ii++;
        }

        //위험률
        ii = 2;
        while(wsSp3['A' + ii]) {
          rr = 0;
          if (wsSp3['A' + ii].v === rrCode) {
            if (wsSp3['B' + ii].v === sex) {
              if (wsSp3['C' + ii].v === (age + i)) {
                rr = wsSp3['D' + ii].v
                break;
              }
            }
          }
          ii++;
        }

        //할인율
        ii = 2;
        while(wsSp4['A' + ii]) {
          if (wsSp4['A' + ii].v === rgCode) {
            if (wsSp4['B' + ii].v === tY) {
              if (wsSp4['C' + ii].v <= ggYearI) {
                if (wsSp4['D' + ii].v >= ggYearI) {
                  rs = wsSp4['E' + ii].v;
                  break;
                }
              }
            } 
          }
          ii++;
        }

        //Lapse
        if (wsInput['J' + line].v > i) {
          lpsYear = 99;
        } else if (ggYearI > 10) {
          lpsYear = 10;
        } else {
          lpsYear = ggYearI;
        }

        ii = 2;
        while(wsSp1['A' + ii]) {
          if (wsSp1['A' + ii].v === lpsYear) {
            lps = wsSp1['B' + ii].v
            break;
          }
          ii++;
        }

        //할인후출재보험료
        //준비금
        rVal = wsInput['T' + line].v;

        if (rVal !== 0) {
          hVal = wsInput['M' + line].v - rVal;
        } else {
          hVal = wsInput['M' + line].v *
            (wsInput['O' + line].v + wsInput['P' + line].v * tsn) /
            (wsInput['O' + line].v + wsInput['P' + line].v);
        }

        hVal *= Math.max((wsInput['N' + line].v - boyu) / wsInput['N' + line].v, 0) *
          rr / 60000 * ( 1 - rs ) * wsInput['C' + line].v * ( 1 - lps ) *
          (wsInput['I' + line].v > i?1:0)

        //재보험금
        ii = 2;
        while(wsSp6['A' + ii]) {
          if (wsSp6['A' + ii].v === sex) {
            if (wsSp6['B' + ii].v <= ggYearI) {
              if (wsSp6['C' + ii].v >= ggYearI) {
                jVal = wsSp6['D' + ii].v * hVal;
                break;
              }
            }
          }
          ii++;
        }

        if (line <= 100) {
          if (i === 0) {
            resultM.push([tY, tYm, ggYear, wsInput['B'+line].v, niGigan, line, hr, i, sYearI, tsn, boyu, rr, rs, lps, hVal, jVal]);
          } else {
            resultM.push([null, null, null, null, null, null, hr, i, sYearI, tsn, boyu, rr, rs, lps, hVal, jVal]);
          }
        }
        
      }

      bHr = false;
      hr = '';

      if (i > 1000) loop = false;
      line++;
    }

    progress.style.width = '100%';
    progress.textContent = '100% (' + ((new Date()) - sc)/1000 + '초)';
    setTimeout("document.getElementById('progress_bar').className='';", 2000);
  }

  function parseFileAsync(mixed, options, callback) {
      var wb = XLSX.read(mixed, options);
      callback(wb);
  }
});
