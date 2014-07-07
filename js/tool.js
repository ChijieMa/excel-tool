/**
 * Created by dongxuehai on 14-5-20.
 */

'use strict';

var gui = require('nw.gui');
var win = gui.Window.get();
var fs = require('fs');
var mysql = require('mysql');
var async = require('async');
var path = require('path');
var xlsx = require('node-xlsx');
var configFile = path.join(path.dirname(process.execPath), 'config.json');
var host, port, user, passwd, dbname, excel, tabPre, dbPre, connection;
var checkedArr = [];

var loadExcelPath = function() {
  var excelPath = $("#excel-path")[0].files[0].path;
  $("#excel").val(excelPath);
};

function msg(type, msg) {
  var info = "<p>" + msg + "</p>";
  if (type === 1) {
    $("#errMsg").append(info);
  }
  if (type === 2) {
    $("#msg").append(info);
  }
}

function checkInput(opts) {
  opts = opts || {};
  host = $("#host").val();
  port = $("#port").val();
  user = $("#user").val();
  passwd = $("#passwd").val();
  dbname = $("#dbname").val();
  excel = $("#excel").val();
  $("#errMsg").html('');
  $("#msg").html('');
  if (host.length <= 0) {
    msg(1, "MYSQL数据库主机不能为空!");
    return false;
  }
  if (port.length <= 0) {
    msg(1, "MYSQL数据库端口不能为空!");
    return false;
  }
  if (user.length <= 0) {
    msg(1, "MYSQL数据库账号不能为空!");
    return false;
  }
  if (passwd.length <= 0) {
    msg(1, "MYSQL数据库密码不能为空!");
    return false;
  }
  if (!opts.dbname && dbname.length <= 0) {
    msg(1, "MYSQL数据库库名不能为空!");
    return false;
  }
  if (!opts.excel && excel.length <= 0) {
    msg(1, "请先指定EXCEL目录!");
    return false;
  }

  if (!opts.excel && !fs.existsSync(excel)) {
    msg(1, "EXCEL目录不存在！");
    return false;
  }

  if (!opts.excel && !fs.statSync(excel).isDirectory()) {
    msg(1, "EXCEL目录不是目录！");
    return false;
  }
  return true;
}

function checkedList() {
  checkedArr = [];
  $("#tbs :checkbox").each(function() {
    if ($(this).is(":checked")) {
      checkedArr.push($(this).val());
    }
  });
}

function connectDb() {
  connection = mysql.createConnection({
    host: host,
    port: port,
    user: user,
    password: passwd,
    database: dbname
  });
}

var importExcel = function() {
  if (!checkInput()) {
    return false;
  }
  checkedList();

  connectDb();

  var backExcelPath;

  async.waterfall([
    function(cb) {
      msg(2, "正在执行...");
      $("#importExcel").attr('disabled', 'true');
      connection.connect(cb);
    },
    function(conInfo, cb) {
      var d = dateFormat(new Date(), 'yyyyMMddhhmmss');
      backExcelPath = path.join(path.dirname(process.execPath), d);
      fs.exists(backExcelPath, function(exists) {
        if (!exists) {
          fs.mkdir(backExcelPath, cb);
        }
      });
    },
    function(cb) {
      fs.readdir(excel, cb);
    },
    function(files, cb) {
      async.each(files, function(file, callback) {
        if (tabPre && file.indexOf(tabPre) !== 0) {
          callback(null);
          return;
        }
        var lid = file.lastIndexOf('.');
        if (lid <= 0 || file.substring(lid + 1) !== 'xlsx') {
          callback(null);
          return;
        }
        var tabname = file.substring(0, lid);

        if (checkedArr.length > 0 && checkedArr.indexOf(tabname) === -1) {
          callback(null);
          return;
        }

        async.waterfall([
          function(cbk) {
            exportSingleExcel(tabname, backExcelPath, cbk);
          },
          function(cbk) {
            connection.query("truncate table " + tabname, function(err) {
              if (err) {
                cbk(err);
                return;
              }
              cbk(null);
            });
          },
          function(cbk) {
            var xlsxFile = path.join(excel, file);
            var obj = xlsx.parse(xlsxFile);
            var dataArr = obj.worksheets[0].data;
            var i = 0, keyLen = 0;
            dataArr.shift();
            //删除第一行注释
            var keys = dataArr.shift();
            if (dataArr.length <= 0) {
              cbk(null);
              return;
            }
            var keyArr = [];
            keyLen = keys.length;
            for (i = 0; i < keyLen; i++) {
              keyArr.push(keys[i].value);
            }
            var valArr = [];
            async.each(dataArr, function(data, callbk) {
              valArr = [];
              for (i = 0; i < keyLen; i++) {
                if (data[i]) {
                  var t = isNull(data[i].value) ? '' : data[i].value;
                  valArr.push(mysql.escape(t));
                } else {
                  valArr.push('');
                }
              }
              var field = "`" + keyArr.join('`,`') + "`";
              var val = valArr.join(",");
              var sql = "insert into " + tabname + "(" + field + ") values(" + val + ")";
              connection.query(sql, function(err) {
                if (err) {
                  msg(1, sql);
                  callbk(err);
                  return;
                }
                callbk(null);
              });
            }, cbk);
          }
        ], callback);
      }, cb);
    }
  ], function(err) {
    connection.end();
    $("#importExcel").removeAttr('disabled');
    if (err) {
      msg(1, "执行错误：" + err);
      msg(2, "执行中有错误！");
      return;
    }
    msg(2, "执行完毕...");
  });
};

var exportExcel = function() {
  if (!checkInput()) {
    return false;
  }

  checkedList();
  connectDb();

  async.waterfall([
    function (cb) {
      $("#exportExcel").attr('disabled', 'true');
      msg(2, '正在执行...');
      connection.connect(cb);
    }, function(conInfo, cb) {
      var sql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES where TABLE_SCHEMA='" + dbname + "'";
      connection.query(sql, function(err, rows) {
        if (err) {
          cb(err);
          return;
        }
        cb(null, rows);
      });
    }, function (tbls, cb) {
      async.each(tbls, function(tbl, callback) {
        var tblName = tbl.TABLE_NAME;
        if (tabPre && tblName.indexOf(tabPre) !== 0) {
          callback(null);
          return;
        }

        if (checkedArr.length > 0 && checkedArr.indexOf(tblName) === -1) {
          callback(null);
          return;
        }
        exportSingleExcel(tblName, excel, callback);
      }, cb);
    }
  ], function(err) {
    connection.end();
    $("#exportExcel").removeAttr('disabled');
    if (err) {
      msg(1, "执行错误：" + err);
      msg(2, "执行中有错误!");
      return;
    }
    msg(2, '执行完毕...');
  });
};

/**
 * 导出单个excel文件
 * @param tableName
 * @param cb
 */
function exportSingleExcel(tableName, excelPath, cb) {
  var sqlArr = [
    "select COLUMN_NAME,COLUMN_COMMENT from information_schema.columns where" +
      " table_schema='" + dbname + "' and table_name='" + tableName + "' order by ORDINAL_POSITION asc",
    "select * from " + tableName
  ];
  async.map(sqlArr, function(sql, ncb) {
    connection.query(sql, function(err, rows) {
      if (err) {
        msg(1, sql);
        ncb(err);
        return;
      }
      ncb(null, rows);
    });
  }, function(err, results) {
    if (err) {
      cb(err);
      return;
    }
    var fieldArr = results[0];
    var fieldLen = fieldArr.length;
    var xlsxArr = [], fields = [], comments = [];
    for (var i = 0; i < fieldLen; i++) {
      comments.push(fieldArr[i].COLUMN_COMMENT);
      fields.push(fieldArr[i].COLUMN_NAME);
    }
    xlsxArr.push(comments);
    xlsxArr.push(fields);
    var dataArr = results[1];
    var dataLen = dataArr.length;
    var data;
    for (var j = 0; j < dataLen; j++) {
      data = [];
      for (var k = 0; k < fieldLen; k++) {
        data.push({"value": dataArr[j][fields[k]] !== null ? dataArr[j][fields[k]] : '', "formatCode": "General"});
      }
      xlsxArr.push(data);
    }
    var obj = {"worksheets": [
      {"data": xlsxArr}
    ]};
    try {
      var file = xlsx.build(obj);
      fs.writeFileSync(path.join(excelPath, tableName + '.xlsx'), file, 'binary');
    } catch (e) {
      msg(1, tableName);
      cb(e);
      return;
    }
    cb(null);
  });
}

function dateFormat(date, format) {
  var o = {
    "M+": date.getMonth() + 1, //月份
    "d+": date.getDate(), //日
    "h+": date.getHours(), //小时
    "m+": date.getMinutes(), //分
    "s+": date.getSeconds(), //秒
    "q+": Math.floor((date.getMonth() + 3) / 3), //季度
    "S": date.getMilliseconds() //毫秒
  };
  if (/(y+)/.test(format)) {
    format = format.replace(RegExp.$1, (date.getFullYear() + "").substr(4 - RegExp.$1.length));
  }
  for (var k in o) {
    if (new RegExp("(" + k + ")").test(format)) {
      format = format.replace(RegExp.$1, (RegExp.$1.length === 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    }
  }
  return format;
}

function isNull(arg1){
  return !arg1 && arg1!==0 && typeof arg1!=="boolean"?true:false;
}


function updateTab() {
  if (!checkInput({excel: true})) {
    return false;
  }

  connectDb();

  var tabArr = [];

  async.waterfall([
    function(cb) {
      $("#updateTab").attr('disabled', 'true');
      connection.connect(cb);
    },
    function(conInfo, cb) {
      var sql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES where TABLE_SCHEMA='" + dbname + "'";
      connection.query(sql, function(err, rows) {
        if (err) {
          cb(err);
          return;
        }
        cb(null, rows);
      });
    },
    function (tbls, cb) {
      async.each(tbls, function(tbl, callback) {
        var tblName = tbl.TABLE_NAME;
        if (tabPre && tblName.indexOf(tabPre) !== 0) {
          callback(null);
          return;
        }
        tabArr.push(tblName);
        callback(null);
      }, cb);
    }
  ], function(err) {
    connection.end();
    $("#updateTab").removeAttr('disabled');
    if (err) {
      $("#tbs").html(err);
      return;
    }
    var tbsHtml = '';
    tabArr.forEach(function(tab) {
      tbsHtml += "<p><input type='checkbox'value='" + tab + "'>" + tab + "</p>";
    });
    $('#tbs').html(tbsHtml);
  });
}

$(document).ready(function(){
  $("#closeButton").mouseover(function() {
    $(this).attr("src","./img/close_hover.png");
  });

  $("#closeButton").mouseout(function() {
    $(this).attr("src","./img/close.png");
  });

  $("#closeButton").click(function() {
    win.close();
  });

  if (fs.existsSync(configFile)) {
    var config = require(configFile);
    $('#host').val(config.host);
    $('#port').val(config.port);
    $('#user').val(config.user);
    $('#passwd').val(config.passwd);
    $('#dbname').val(config.dbname);
    tabPre = config.prefix;
    dbPre = config.dbPre;
    //updateDb();
    updateTab();
  }

  win.show();
});



