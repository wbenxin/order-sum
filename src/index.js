#!/usr/bin/env node

const fs = require('fs');
const readline = require('readline');
const xlsx = require('xlsx');
const iconv = require('iconv-lite');

console.log("请输入订单文件的完整路径：");
const rl = readline.createInterface(process.stdin, process.stdout);
rl.on('line', (line) => {
  if (line.trim() === 'break') {
    rl.close();
    return;
  }

  let workbook = xlsx.readFile(line); //workbook就是xls文档对象
  let sheetNames = workbook.SheetNames; //获取表明
  let sheet = workbook.Sheets[sheetNames[0]]; //通过表明得到表对象
  var data = xlsx.utils.sheet_to_json(sheet); //通过工具将表对象的数据读出来并转成json

  // 买家会员名,总金额,收货人姓名,收货地址 ,宝贝标题 
  let map = {};
  data.forEach(item => {
    let key = item['宝贝标题 '] + ',' + item['买家会员名'] + ',' + item['收货人姓名'] + ',' + item['收货地址 '];
    if (map.hasOwnProperty(key)) {
      map[key] += Number(item.总金额);
    } else {
      map[key] = Number(item.总金额);
    }
  })

  let csv = "宝贝标题,买家会员名,收货人姓名,收货地址,总金额\n";
  Object.keys(map).forEach(key => {
    let csv_line = key + ',' + map[key] + '\n';
    csv += csv_line;
  });

  fs.writeFile(line + '.csv', iconv.encode(csv, 'gbk'), () => {
    console.log('已将汇总数据保存到：' + line + '.csv');

    setTimeout(() => {
      process.exit();
    }, 1000);
  });

});