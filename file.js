var xlsx = require('node-xlsx');
var fs = require('fs');

//写入Excel数据
try {
  //excel数据
  var excelData = [];
  //表1
  {
    //添加数据
    var addInfo = {};
    //名称
    addInfo.name = "用户表";
    //数据数组
    addInfo.data = [
      ["英文标题", "日文标题", "主图图片", "图片文件夹序号", '价格', '型号', "商品描述", "颜色", "尺码", "详情描述", "日元", "原价格（美金）"
      ],
    ];

    //添加数据
    addInfo.data.push([10000, "张三", "男", 15]);
    addInfo.data.push([10001, "李四", "男", 40]);

    //添加数据
    excelData.push(addInfo);
    console.log()
  }

  // 写xlsx
  var buffer = xlsx.build(excelData);
  //写入数据
  fs.writeFileSync('./data.xls', buffer, function (err) {
    if (err) {
      throw err;
    }
    //输出日志
    console.log('Write to xls has finished');
  });
}
catch (e) {
  //输出日志
  console.log("excel写入异常,error=%s", e.stack);
}