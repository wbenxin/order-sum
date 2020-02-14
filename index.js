#!/usr/bin/env node

const fs = require('fs');
const readline = require('readline');
const xlsx = require('node-xlsx');
const iconv = require('iconv-lite');

console.log("请输入订单文件的完整路径：");
const rl = readline.createInterface(process.stdin, process.stdout);
rl.on('line', (line) => {
  if (line.trim() === 'break') {
    rl.close();
    return;
  }

  let sheets = xlsx.parse(line);
  // sheets[0].data[0] = ["订单编号","买家会员名","买家支付宝账号","支付单号","支付详情","买家应付货款","买家应付邮费","买家支付积分","总金额","返点积分","买家实际支付金额","买家实际支付积分","订单状态","买家留言","收货人姓名","收货地址 ","运送方式","联系电话 ","联系手机","订单创建时间","订单付款时间 ","宝贝标题 ","宝贝种类 ","物流单号 ","物流公司","订单备注","宝贝总数量","店铺Id","店铺名称","订单关闭原因","卖家服务费","买家服务费","发票抬头","是否手机订单","分阶段订单信息","特权订金订单id","是否上传合同照片","是否上传小票","是否代付","定金排名","修改后的sku","修改后的收货地址","异常信息","天猫卡券抵扣","集分宝抵扣","是否是O2O交易","新零售交易类型","新零售导购门店名称","新零售导购门店id","新零售发货门店名称","新零售发货门店id","退款金额","预约门店","确认收货时间","打款商家金额","含应开票给个人的个人红包","是否信用购","体验期结束时间","前N有礼","配送类型","主订单编号"]
  let data = sheets[0].data;

  // 买家会员名,总金额,收货人姓名,收货地址 ,宝贝标题 : 1,8,14,15,21
  let sheet1 = { name: 'sheet1', data: [['宝贝标题', '买家会员名', '收货人姓名', '收货地址', '总金额']] };
  let pay_map = new Map();
  let sheet_map = new Map();
  for (let i = 1; i < data.length; i++) {
    let key = data[i][21] + ',' + data[i][1] + ',' + data[i][14] + ',' + data[i][15];
    if (pay_map.has(key)) {
      pay_map.set(key, pay_map.get(key) + Number(data[i][8]));
    } else {
      pay_map.set(key, Number(data[i][8]));
      sheet1.data.push([data[i][21], data[i][1], data[i][14], data[i][15], key]);
    }

    // 按会员名分别筛选到单独sheet中
    if (sheet_map.has(data[i][1])) {
      sheet_map.get(data[i][1]).push([data[i][0], Number(data[i][8]), data[i][14], data[i][15], data[i][19]]);
    } else {
      sheet_map.set(data[i][1], [[data[i][0], Number(data[i][8]), data[i][14], data[i][15], data[i][19]]]);
    }
  }

  // 更新总金额
  for (let i = 1; i < sheet1.data.length; i++) {
    let key = sheet1.data[i][4];
    sheet1.data[i][4] = pay_map.get(key);
  }

  let out = [sheet1];
  sheet_map.forEach((value, key) => {
    out.push({ name: key, data: value });
  });
  let buffer = xlsx.build(out, { '!cols': [{ wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }] });
  fs.writeFile(line + ".xlsx", buffer, () => {
    console.log('已将汇总数据保存到：' + line + '.xlsx');

    setTimeout(() => {
      process.exit();
    }, 1000);
  });
});