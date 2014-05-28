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
var host, port, user, passwd, dbname, excel, prefix, connection;

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

function checkInput() {
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
  if (dbname.length <= 0) {
    msg(1, "MYSQL数据库库名不能为空!");
    return false;
  }
  if (excel.length <= 0) {
    msg(1, "请先指定EXCEL目录!");
    return false;
  }

  if (!fs.existsSync(excel)) {
    msg(1, "EXCEL目录不存在！");
    return false;
  }

  if (!fs.statSync(excel).isDirectory()) {
    msg(1, "EXCEL目录不是目录！");
    return false;
  }
  return true;
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

  connectDb();

  async.waterfall([
    function(cb) {
      msg(2, "正在执行...");
      $("#importExcel").attr('disabled', 'true');
      connection.connect(cb);
    },
    function(conInfo, cb) {
      fs.readdir(excel, cb);
    },
    function(files, cb) {
      async.each(files, function(file, callback) {
        if (prefix && file.indexOf(prefix) !== 0) {
          callback(null);
          return;
        }
        var lid = file.lastIndexOf('.');
        if (lid <= 0 || file.substring(lid + 1) !== 'xlsx') {
          return;
        }
        var tabname = file.substring(0, lid);
        async.waterfall([
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
            var keyArr = [];
            keyLen = keys.length;
            for (i = 0; i < keyLen; i++) {
              keyArr.push(keys[i].value);
            }
            var valArr = [];
            async.each(dataArr, function(data, callbk) {
              valArr = [];
              for (i = 0; i < keyLen; i++) {
                valArr.push(data[i].value);
              }
              var field = "`" + keyArr.join('`,`') + "`";
              var val = "'" + valArr.join("','") + "'";
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
        if (prefix && tblName.indexOf(prefix) !== 0) {
          callback(null);
          return;
        }
        var sqlArr = [
          "select COLUMN_NAME,COLUMN_COMMENT from information_schema.columns where" +
            " table_schema='" + dbname + "' and table_name='" + tblName + "' order by ORDINAL_POSITION asc",
          "select * from " + tblName
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
            callback(err);
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
            fs.writeFileSync(path.join(excel, tblName + '.xlsx'), file, 'binary');
          } catch (e) {
            msg(1, tblName);
            callback(e);
            return;
          }

          callback(null);
        });
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
    prefix = config.prefix;
  }

  win.show();
});



