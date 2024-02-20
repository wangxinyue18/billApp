// // const 铁塔账单文件表路径 = fileData['铁塔账单文件表'];
// // const workbook = new ExcelJs.Workbook();
// // const 铁塔账单文件表文件流 = fs.createReadStream(铁塔账单文件表路径);
// // workbook.xlsx.read(铁塔账单文件表文件流).then(excelTemp => {
// //   excelTemp.getWorksheet('Sheet1')
// //   console.log('🍑', '文件read完成')
// // })
// // // 铁塔账单文件表文件流.on('data', async (data) => {
// // //   const excelTemp = await workbook.xlsx.load(data);
// // //   console.log('🍑', excelTemp);
// // // })
// // // 设置结束事件处理程序
// // 铁塔账单文件表文件流.on('end', () => {
// //   console.log('文件读取完成！');
// // });

// // // 设置错误事件处理程序
// // 铁塔账单文件表文件流.on('error', (err) => {
// //   console.error(`发生了错误：${err}`);
// // });

//  // 传输订单文件😀😀😀
//  const odtranmissorsheets = xlsxData.总订单文件表[1].data
//  let odtransmisslist = []
//  let odtransnum = 0
//  const odtransmisstitle = odtranmissorsheets[0];
//  odtranmissorsheets.forEach((item, index) => {
//      // console.log(item)
//      // console.log(index)
//      if (index == 0 || index == 1 || index == 2) {
//          return
//      }
//      else {
//          odtransmisslist.push(
//              {
//                  [odtransmisstitle[0]]: item[0],
//                  [odtransmisstitle[1]]: item[1],
//                  [odtransmisstitle[2]]: item[2],
//                  [odtransmisstitle[42]]: item[42],
//                  [odtransmisstitle[69]]: item[69]
//              }
//          )

//          odtransnum = odtransnum + 1
//          // console.log('❀❀ '+odtransmisslist)
//      }
//  })
//  // console.log(odtransmisslist)
//  console.log('🍌传输订单文件数（已筛选）：' + odtransnum)
//  // const transmiss2 = xlsx.parse("D:/typescript/demo/accountbill/transmission.xlsx", {
//  //   cellDates: true,
//  // });
//  //终止文件😀😀😀
//  const forbiden2 = xlsxData.终止文件表
//  let forbidnum = 0
//  let forbidensheets = forbiden2[0].data
//  let forbidlist = []
//  const forbidtitle = forbidensheets[0]
//  forbidensheets.forEach((item, index) => {
//      if (index == 0) {
//          return
//      }
//      else {
//          forbidlist.push({
//              [forbidtitle[4]]: item[4]

//          })
//          forbidnum = forbidnum + 1
//      }
//  })
//  console.log("🍑终止文件订单数目： " + forbidnum)
//  //传输账单文件
//  const transmiss2 = xlsxData.传输账单文件表
//  let transmisssheets = transmiss2[0].data
//  let transmisslists = []
//  const transmisstitle = transmisssheets[1]
//  let transnum = 0
//  transmisssheets.forEach((item, index) => {
//      if (index == 1 || index == 0) {
//          return
//      }
//      else if (item[8] != undefined) {
//          transmisslists.push({
//              [transmisstitle[2]]: item[2],
//              [transmisstitle[8]]: item[8],
//              [transmisstitle[19]]: item[19],
//              [transmisstitle[20]]: item[20],
//              [transmisstitle[27]]: item[27],
//              [transmisstitle[28]]: item[28],
//              [transmisstitle[38]]: item[38],

//          })
//          transnum = transnum + 1
//      }
//  })
//  transmisslists.forEach((item, index) => {
//      if ((item.运营商 == '移动' && item.运营商区县 == '双流县') || (item.运营商 == '移动' && item.运营商区县 == '龙泉驿区') || (item.运营商 == '移动' && item.运营商区县 == '天府新区')) {
//          item.运营商 = '天府移动'
//      }
//  })
//  console.log('🍓传输账单订单数:' + transnum)

//  // 从订单文件向账单传输进行对比😀😀😀
//  let numcsz = 0
//  let numcsy = 0
//  for (let i = 0; i < odtransnum; i++) {
//      let numtj4 = 0
//      let numtj5 = 0
//      for (let j = 0; j < transnum; j++) {
//          if (odtransmisslist[i].订单号 != transmisslists[j].需求确认单编号) {
//              numtj4 = numtj4 + 1
//          }
//          else if (odtransmisslist[i].订单号 == transmisslists[j].需求确认单编号) {
//              //正常订单数目
//              numcsz = numcsz + 1
//          }
//      }
//      if (numtj4 == transnum) {
//          // console.log('存在可能异常订单号：'+titlelist[i].订单号)

//          for (let k = 0; k < forbidnum; k++) {
//              if (odtransmisslist[i].订单号 == forbidlist[k].订单编号) {
//                  // console.log('终止文件存在正常订单号：' + titlelist[i].订单号)
//                  numcsz = numcsz + 1
//              }
//              else if (odtransmisslist[i].订单号 != forbidlist[k].订单编号) {
//                  numtj5 = numtj5 + 1
//              }
//              if (numtj5 == forbidnum) {
//                  // console.log('异常账号' + odtransmisslist[i].订单号 + '原因：在详单里面，但是不在账单里面')
//                  numcsy = numcsy + 1
//              }
//          }
//      }
//  }
//  //从传输订单文件向订单文件传输
//  for (let j1 = 0; j1 < transnum; j1++) {
//      let numtj3 = 0
//      for (let i1 = 0; i1 < odtransnum; i1++) {
//          if (transmisslists[j1].需求确认单编号 != odtransmisslist[i1].订单号) {
//              numtj3 = numtj3 + 1
//          }
//          else if (transmisslists[j1].需求确认单编号 == odtransmisslist[i1].订单号) {
//              // numcsz = numcsz + 1
//          }
//      }
//      if (numtj3 == odtransnum) {
//          // console.log('异常订单' + transmisslists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
//          numcsy = numcsy + 1
//      }
//  }
//  console.log("正常订单数：（按照订单文件为基准）" + numcsz)
//  console.log("异常订单数：（账单文件＋订单文件）" + numcsy)
//  // console.log(transmisslists)
//  let yidong1 = 0
//  let tfyidong1 = 0
//  let liantong1 = 0
//  let dianxing1 = 0

//  let stocksf1a = 0
//  let stocksf2a = 0
//  let stocksf3a = 0
//  let stocksf4a = 0
//  let sumt1 = 0
//  let sumt2 = 0
//  let sumt3 = 0
//  let sumt4 = 0
//  let sumt5 = 0
//  let sumt6 = 0
//  let sumt7 = 0
//  let sumt8 = 0
//  // 传输只有存量没得新增
//  transmisslists.forEach((item, index) => {

//      if (item.运营商 == '移动') {
//          yidong1 = yidong1 + 1
//          stocksf1a = stocksf1a + 1
//          sumt1 = parseFloat(sumt1 + item.产品服务费合计1 + item.产品服务费合计2)
//          sumt2 = parseFloat(sumt2 + item.维护费1 + item.维护费2)
//      }
//      else if (item.运营商 == '天府移动') {
//          tfyidong1 = tfyidong1 + 1
//          stocksf2a = stocksf2a + 1
//          sumt3 = parseFloat(sumt3 + item.产品服务费合计1 + item.产品服务费合计2)
//          sumt4 = parseFloat(sumt4 + item.维护费1 + item.维护费2)
//      }
//      else if (item.运营商 == '联通') {
//          liantong1 = liantong1 + 1
//          stocksf3a = stocksf3a + 1
//          sumt5 = parseFloat(sumt5 + item.产品服务费合计1 + item.产品服务费合计2)
//          sumt6 = parseFloat(sumt6 + item.维护费1 + item.维护费2)
//      }
//      else if (item.运营商 == '电信') {
//          dianxing1 = dianxing1 + 1
//          stocksf4a = stocksf4a + 1
//          sumt7 = parseFloat(sumt7 + item.产品服务费合计1 + item.产品服务费合计2)
//          sumt8 = parseFloat(sumt8 + item.维护费1 + item.维护费2)
//      }
//  })
//  console.log(yidong1 + ' ' + sumt1 + '  ' + sumt2)
//  console.log(tfyidong1 + ' ' + sumt3 + '  ' + sumt4)
//  console.log(liantong1 + ' ' + sumt5 + '  ' + sumt6)
//  console.log(dianxing1 + ' ' + sumt7 + '  ' + sumt8)

//  console.log('❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️  ❤️  ')

//  // let detailbillsheets = xlsxData.详单文件表[0].data
//  // console.log(detailbillsheets + '❀')
//  // let detailtitle = detailbillsheets[0]
//  // let detaillists = []
//  // console.log(detailbillsheets)
//  let allbill = sumt1 + sumt3 + sumt5 + sumt7
//  // detailbillsheets.forEach((item, index) => {
//  //     if (index == 1) {
//  //         detaillists.push({
//  //             [detailtitle[1]]: transnum,
//  //             [detailtitle[2]]: yidong1,
//  //             [detailtitle[3]]: 0,
//  //             [detailtitle[4]]: tfyidong1,
//  //             [detailtitle[5]]: 0,
//  //             [detailtitle[6]]: liantong1,
//  //             [detailtitle[7]]: 0,
//  //             [detailtitle[8]]: dianxing1,
//  //             [detailtitle[9]]: 0
//  //         })
//  //     }
//  //     else if (index == 3) {
//  //         detaillists.push({
//  //             [detailtitle[1]]: allbill - (sumt2 + sumt4 + sumt6 + sumt8),
//  //             [detailtitle[2]]: sumt1 - sumt2,
//  //             [detailtitle[3]]: 0,
//  //             [detailtitle[4]]: sumt3 - sumt4,
//  //             [detailtitle[5]]: 0,
//  //             [detailtitle[6]]: sumt5 - sumt6,
//  //             [detailtitle[7]]: 0,
//  //             [detailtitle[8]]: sumt7 - sumt8,
//  //             [detailtitle[9]]: 0
//  //         })
//  //     }
//  //     else if (index == 2) {
//  //         detaillists.push({
//  //             [detailtitle[1]]: allbill,
//  //             [detailtitle[2]]: sumt1,
//  //             [detailtitle[3]]: 0,
//  //             [detailtitle[4]]: sumt3,
//  //             [detailtitle[5]]: 0,
//  //             [detailtitle[6]]: sumt5,
//  //             [detailtitle[7]]: 0,
//  //             [detailtitle[8]]: sumt7,
//  //             [detailtitle[9]]: 0
//  //         })
//  //     }
//  //     if (index != 0 && index != 1 && index != 2 && index != 3)
//  //         detaillists.push({
//  //             [detailtitle[1]]: 0,
//  //             [detailtitle[2]]: 0,
//  //             [detailtitle[3]]: 0,
//  //             [detailtitle[4]]: 0,
//  //             [detailtitle[5]]: 0,
//  //             [detailtitle[6]]: 0,
//  //             [detailtitle[7]]: 0,
//  //             [detailtitle[8]]: 0,
//  //             [detailtitle[9]]: 0

//  //         })
//  // })
//  // // const Jsondata = JSON.stringify(detaillists)
//  // // const filePath = 'D:/typescript/demo/accountbill/data.json';
//  // fs.writeFileSync(filePath, Jsondata);
//  // console.log(`已将对象数组保存到${filePath}`);


//  // fs.readFile('D:/typescript/demo/accountbill/data.json', 'utf8', (err, data) => {
//  //     if (err) throw err;
//  //     const json = JSON.parse(data);
//  //     const jsonArray = [];
//  //     json.forEach(function (item) {
//  //         let temp = {
//  //             '传输小计': item.传输小计,
//  //             '成都移动存量': item.成都移动存量,
//  //             '成都移动新增': item.成都移动新增,
//  //             '天府移动存量': item.天府移动存量,
//  //             '天府移动新增': item.天府移动新增,
//  //             '电信存量': item.电信存量,
//  //             '电信新增': item.电信新增,
//  //             '联通存量': item.联通存量,
//  //             '联通新增': item.联通新增,
//  //         }
//  //         jsonArray.push(temp);
//  //     });

//  //     let xls = json2xls(jsonArray);

//  //     fs.writeFileSync('D:/typescript/demo/accountbill/newdetailorder.xlsx', xls, 'binary');
//  //     console.log('文件已经保存成功')
//  // })

//  console.log('\^o^/\^o^/\^o^/\^o^/\^o^/\^o^/')



//  // 订单文件室分😀😀😀
//  let odinnersheet2 = xlsxData.总订单文件表[0].data
//  let odinnerrlist = []
//  // 获取标题行
//  const orderinnertitle = odinnersheet2[2];
//  // console.log(ordertitle)
//  let odnum = 0
//  odinnersheet2.forEach((item, index) => {
//      // console.log(item)
//      // console.log(index)
//      if (index == 0 || index == 1) {
//          return
//      }
//      else if (item[0] != undefined && item[1] == '已起租' && item[95] != '0.00') {
//          odinnerrlist.push(
//              {
//                  [orderinnertitle[0]]: item[0],
//                  [orderinnertitle[1]]: item[1],
//                  [orderinnertitle[2]]: item[2],
//                  [orderinnertitle[95]]: item[95],
//              }
//          )
//          odnum = odnum + 1
//      }
//  })
//  console.log('室分订单文件数（已筛选）' + odnum)

//  //账期订单文件😀😀😀
//  const buildinnfile = xlsxData.室分账单文件表
//  let binum = 0
//  let buildinnsheet = buildinnfile[0].data
//  // console.log(buildinnsheet)
//  let buildinnlist = []
//  const buildinntitle = buildinnsheet[0]
//  buildinnsheet.forEach((item, index) => {
//      if (index == 0) {
//          return
//      }
//      else if (item[2] == '移动' && (item[76] == ' 天府新区' || item[76] == ' 双流县' || item[76] == '龙泉驿区')) {
//          item[2] = '天府移动'

//      }
//      else if (item[8] != undefined) {
//          buildinnlist.push({
//              [buildinntitle[2]]: item[2],
//              [buildinntitle[8]]: item[8],
//              [buildinntitle[59]]: item[59],
//              [buildinntitle[56]]: item[56],
//              [buildinntitle[57]]: item[57],
//              [buildinntitle[58]]: item[58],
//              [buildinntitle[67]]: item[67],
//              [buildinntitle[73]]: item[73],
//              [buildinntitle[74]]: item[74],
//              [buildinntitle[75]]: item[75],
//              [buildinntitle[76]]: item[76],
//              [buildinntitle[41]]: item[41],
//              [buildinntitle[42]]: item[42],
//              [buildinntitle[31]]: item[31],
//              [buildinntitle[32]]: item[32],

//          })
//          binum = binum + 1
//      }

//  })
//  buildinnlist.forEach((item, index) => {
//      if ((item.运营商 == '移动' && item.运营商区县 == '双流县') || (item.运营商 == '移动' && item.运营商区县 == '龙泉驿区') || (item.运营商 == '移动' && item.运营商区县 == '天府新区')) {
//          item.运营商 = '天府移动'
//      }
//  })
//  console.log('室分账单订单数' + binum)


//  // 将室分订单文件和账单传输进行对比😀😀😀
//  let numcfz = 0
//  let numcfy = 0

//  for (let i = 0; i < odnum; i++) {
//      let num8 = 0
//      let num9 = 0
//      for (let j = 0; j < binum; j++) {
//          if (odinnerrlist[i].订单号 != buildinnlist[j].需求确认单编号) {
//              num8 = num8 + 1
//          }
//          else if (odinnerrlist[i].订单号 == buildinnlist[j].需求确认单编号) {
//              numcfz = numcfz + 1
//              // console.log('正常订单号：' + odinnerrlist[i].订单号)
//          }
//      }
//      if (num8 == binum) {
//          for (let k = 0; k < forbidnum; k++) {
//              if (odinnerrlist[i].订单号 != forbidlist[k].订单编号) {
//                  num9 = num9 + 1
//              }
//              else if (odinnerrlist[i].订单号 == forbidlist[k].订单编号) {
//                  numcfz = numcfz + 1
//              }
//          }
//          if (num9 == forbidnum) {
//              // console.log('存在异常账号：' + odinnerrlist[i].订单号)
//              numcfy = numcfy + 1
//          }
//      }
//  }

//  //从账单文件向订单文件遍历订单是否异常
//  for (let j1 = 0; j1 < binum; j1++) {
//      let numtj3 = 0
//      for (let i1 = 0; i1 < odnum; i1++) {
//          if (buildinnlist[j1].需求确认单编号 != odinnerrlist[i1].订单号) {
//              numtj3 = numtj3 + 1
//          }
//          else if (buildinnlist[j1].需求确认单编号 == buildinnlist[i1].订单号) {
//              // numcfz = numcfz + 1
//          }
//      }
//      if (numtj3 == odnum) {
//          // console.log('异常订单' + buildinnlist[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
//          numcfy = numcfy + 1
//      }
//  }
//  console.log("正常订单数：（按照订单文件为基准）" + numcfz)
//  console.log("异常订单数：（账单文件＋订单文件）" + numcfy)
//  // console.log(buildinnlist)

//  let yidong = 0
//  let tfyidong = 0
//  let liantong = 0
//  let dianxing = 0

//  let stocksf1 = 0
//  let stocksf11 = 0
//  let stocksf2 = 0
//  let stocksf22 = 0
//  let stocksf3 = 0
//  let stocksf33 = 0
//  let stocksf4 = 0
//  let stocksf44 = 0
//  let sum1 = 0
//  let sum2 = 0
//  let sum3 = 0
//  let sum4 = 0
//  let sum5 = 0
//  let sum6 = 0
//  let sum7 = 0
//  let sum8 = 0
//  let sum9 = 0
//  let sum10 = 0
//  let sum1b = 0
//  let sum2b = 0
//  let sum3b = 0
//  let sum4b = 0
//  let sum5b = 0
//  let sum6b = 0
//  let sum7b = 0
//  let sum8b = 0
//  let sum9b = 0
//  let sum10b = 0
//  let sum1c = 0
//  let sum2c = 0
//  let sum3c = 0
//  let sum4c = 0
//  let sum5c = 0
//  let sum6c = 0
//  let sum7c = 0
//  let sum8c = 0
//  let sum9c = 0
//  let sum10c = 0
//  let sum1d = 0
//  let sum2d = 0
//  let sum3d = 0
//  let sum4d = 0
//  let sum5d = 0
//  let sum6d = 0
//  let sum7d = 0
//  let sum8d = 0
//  let sum9d = 0
//  let sum10d = 0
//  let test = 0
//  // console.log(buildinnlist)
//  //申明数组
//  buildinnlist.forEach((item, index) => {

//      if (item.运营商 == '移动') {
//          if (item.产品服务费与上月相比是否变化a == '存量') {
//              stocksf1 = stocksf1 + 1
//              test = parseFloat(item.产品服务费合计1 + test)
//              sum1 = parseInt(item.产品服务费合计1 + item.产品服务费合计2 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费2 + sum1)
//              sum3 = parseInt(sum3 + item.维护费折扣后金额1 + item.维护费折扣后金额2)//正常
//              sum5 = parseInt(sum5 + item.场地费折扣后金额1 + item.场地费折扣后金额2)//正常
//              sum7 = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum7)
//              sum9 = parseInt(item.产品服务费合计0 + item.产品服务费合计2 + sum9)
//          }
//          else if (item.产品服务费与上月相比是否变化a == '新增') {
//              stocksf11 = stocksf11 + 1
//              sum2 = parseInt(item.产品服务费合计1 + sum2)
//              sum4 = parseInt(sum4 + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum6 = parseInt(sum6 + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum8 = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum8)
//              sum10 = parseInt(item.产品服务费合计0 + item.产品服务费合计2 + sum10)
//          }
//          yidong = yidong + 1
//      }
//      else if (item.运营商 == '天府移动') {
//          if (item.产品服务费与上月相比是否变化a == '存量') {
//              stocksf2 = stocksf2 + 1
//              sum1b = parseInt(item.产品服务费合计1 + sum1b)
//              // console.log(item.罚责赠费合计)
//              sum3b = parseInt(sum3b + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum5b = parseInt(sum5b + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum7b = parseInt(sum7b + item.油机发电服务费1 + item.油机发电服务费2)
//              sum9b = parseInt(sum9b + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          else if (item.产品服务费与上月相比是否变化a == '新增') {
//              stocksf22 = stocksf22 + 1
//              sum2b = parseInt(sum2b + item.产品服务费合计1)
//              sum4b = parseInt(sum4b + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum6b = parseInt(sum6b + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum8b = parseInt(sum8b + item.油机发电服务费1 + item.油机发电服务费2)
//              sum10b = parseInt(sum10b + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          tfyidong = tfyidong + 1
//      }
//      else if (item.运营商 == '联通') {
//          if (item.产品服务费与上月相比是否变化a == '存量') {
//              stocksf3 = stocksf3 + 1
//              sum1c = parseInt(item.产品服务费合计1 + sum1c)
//              sum3c = parseInt(sum3c + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum5c = parseInt(sum5c + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum7c = parseInt(sum7c + item.油机发电服务费1 + item.油机发电服务费2)
//              sum9c = parseInt(sum9c + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          else if (item.产品服务费与上月相比是否变化a == '新增') {
//              stocksf33 = stocksf33 + 1
//              sum2c = parseInt(item.产品服务费合计1 + sum2c)
//              sum4c = parseInt(sum4c + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum6c = parseInt(sum6c + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum8c = parseInt(sum8c + item.油机发电服务费1 + item.油机发电服务费2)
//              sum10c = parseInt(sum10c + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          liantong = liantong + 1
//      }
//      else if (item.运营商 == '电信') {
//          if (item.产品服务费与上月相比是否变化a == '存量') {
//              stocksf4 = stocksf4 + 1
//              sum1d = parseInt(item.产品服务费合计1 + sum1d)
//              sum3d = parseInt(sum3d + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum5d = parseInt(sum5d + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum7d = parseInt(sum7d + item.油机发电服务费1 + item.油机发电服务费2)
//              sum9d = parseInt(sum9d + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          else if (item.产品服务费与上月相比是否变化a == '新增') {
//              stocksf44 = stocksf44 + 1
//              sum2d = parseInt(item.产品服务费合计1 + sum2d)
//              sum4d = parseInt(sum4d + item.维护费折扣后金额1 + item.维护费折扣后金额2)
//              sum6d = parseInt(sum6d + item.场地费折扣后金额1 + item.场地费折扣后金额2)
//              sum8d = parseInt(sum8d + item.油机发电服务费1 + item.油机发电服务费2)
//              sum10d = parseInt(sum10d + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          dianxing = dianxing + 1
//      }
//  })
//  // console.log(test)
//  // console.log(test - sum3 - sum5)
//  console.log(yidong + '  ' + stocksf1 + ' ' + test + '  ' + (test - sum3 - sum5) + ' ' + sum3 + '  ' + sum5 + '  ' + sum7 + ' ' + sum9)
//  console.log(yidong + '  ' + stocksf11 + '  ' + sum2 + ' ' + (sum2 - sum4 - sum6) + '  ' + sum4 + '  ' + sum6 + '  ' + sum8 + '  ' + sum10)
//  console.log(tfyidong + '  ' + stocksf2 + '  ' + sum1b + '  ' + (sum1b - sum3b - sum5b) + ' ' + sum3b + '  ' + sum5b + '  ' + sum7b + '  ' + sum9b)
//  console.log(tfyidong + '  ' + stocksf22 + '  ' + sum2b + '  ' + (sum2b - sum4b - sum6b) + '  ' + sum4b + '  ' + sum6b + '  ' + sum8b + '  ' + sum10b)
//  console.log(liantong + '  ' + stocksf3 + '  ' + sum1c + '  ' + (sum1c - sum3c - sum5c) + ' ' + sum3c + '  ' + sum5c + '  ' + sum7c + '  ' + sum9c)
//  console.log(liantong + '  ' + stocksf33 + '   ' + sum2c + '  ' + (sum2c - sum4c - sum6c) + '  ' + sum4c + '  ' + sum6c + '  ' + sum8c + '  ' + sum10c)
//  console.log(dianxing + '   ' + stocksf4 + '  ' + sum1d + '  ' + (sum1d - sum3d - sum5d) + ' ' + sum3d + '  ' + sum5d + '  ' + sum7d + '  ' + sum9d)
//  console.log(dianxing + ' ' + stocksf44 + '  ' + sum2d + '  ' + (sum2d - sum4d - sum6d) + '  ' + sum4d + '  ' + sum6d + '  ' + sum8d + '  ' + sum10d)




//  console.log('\^o^/\^o^/\^o^/\^o^/\^o^/')

//  //微站
//  const microfile = xlsxData.微站账单文件表
//  let microsheet = microfile[0].data
//  const microtitle = microsheet[0]
//  let ordermicro = xlsxData.总订单文件表[2].data
//  const ordermicrotitle = ordermicro[2]
//  let microOdlists = []
//  let microlists = []
//  let ordernum = 0
//  let micronum = 0
//  //遍历微站订单已筛选订单
//  ordermicro.forEach((item, index) => {
//      if (index == 0 || index == 1 || index == 2) {
//          return
//      }
//      else if (item[0] != undefined && item[1] == '已起租' && item[50] != '0.00' && item[87] != '已暂停出账') {
//          microOdlists.push({
//              [ordermicrotitle[1]]: item[1],
//              [ordermicrotitle[2]]: item[2],

//          })
//          ordernum = ordernum + 1

//      }
//  })
//  // console.log(microOdlists)
//  //遍历微站账单订单
//  microsheet.forEach((item, index) => {
//      if (index == 0) {
//          return
//      }
//      else {
//          microlists.push({
//              [microtitle[9]]: item[9],
//              [microtitle[2]]: item[2],
//              [microtitle[21]]: item[21],
//              [microtitle[22]]: item[22],
//              [microtitle[25]]: item[25],
//              [microtitle[26]]: item[26],
//              [microtitle[52]]: item[52],
//              [microtitle[53]]: item[53],
//              [microtitle[54]]: item[54],
//              [microtitle[55]]: item[55],
//              [microtitle[69]]: item[69],
//              [microtitle[70]]: item[70],
//          })
//      }
//      micronum = micronum + 1
//  })

//  console.log("微站订单文件数（已筛选）：" + ordernum)
//  console.log("微站账单订单数：" + micronum)
//  //从订单文件向账单文件
//  let numz = 0
//  let numy = 0
//  for (let i = 0; i < ordernum; i++) {
//      let numtj = 0
//      let numtj2 = 0
//      for (let j = 0; j < micronum; j++) {
//          if (microOdlists[i].订单号 != microlists[j].需求确认单编号) {
//              numtj = numtj + 1
//          }
//          else if (microOdlists[i].订单号 == microlists[j].需求确认单编号) {
//              numz = numz + 1
//              // console.log('正常订单'+microOdlists[i].订单号)
//          }
//      }
//      if (numtj == micronum) {

//          for (let k = 0; k < forbidnum; k++) {
//              if (microOdlists[i].订单号 == forbidlist[k].订单编号) {
//                  numz = numz + 1
//                  //  console.log('正常订单'+microOdlists[i].订单号)
//              }
//              else if (microOdlists[i].订单号 != forbidlist[k].订单编号) {
//                  numtj2 = numtj2 + 1
//              }
//          }
//          if (numtj2 == forbidnum) {
//              numy = numy + 1
//              // console.log('异常账号' + microOdlists[i].订单号 + '原因：在详单里面，但是不在账单里面')
//          }
//      }
//  }

//  //从账单文件向订单文件遍历订单是否异常
//  for (let j1 = 0; j1 < micronum; j1++) {
//      let numtj3 = 0
//      for (let i1 = 0; i1 < ordernum; i1++) {
//          if (microlists[j1].需求确认单编号 != microOdlists[i1].订单号) {
//              numtj3 = numtj3 + 1
//          }
//          else if (microlists[j1].需求确认单编号 == microOdlists[i1].订单号) {
//              // numz = numz + 1
//              // console.log('正常订单' + microlists[j1].需求确认单编号)


//          }
//      }
//      if (numtj3 == ordernum) {
//          // console.log('异常订单' + microlists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
//          numy = numy + 1
//      }
//  }
//  console.log("正常订单数：（按照订单文件为基准）" + numz)
//  console.log("异常订单数：（账单文件＋订单文件）" + numy)

//  //算新增和存量
//  let numxz1 = 0
//  let numxz2 = 0
//  let numxz3 = 0
//  let numxz4 = 0
//  let numcl1 = 0
//  let numcl2 = 0
//  let numcl3 = 0
//  let numcl4 = 0
//  let money1 = 0
//  let money2 = 0
//  let money3 = 0
//  let money4 = 0
//  let money5 = 0
//  let money6 = 0
//  let money7 = 0
//  let money8 = 0
//  let repare1 = 0
//  let repare2 = 0
//  let repare3 = 0
//  let repare4 = 0
//  let repare5 = 0
//  let repare6 = 0
//  let repare7 = 0
//  let repare8 = 0
//  let placer1 = 0
//  let placer2 = 0
//  let placer3 = 0
//  let placer4 = 0
//  let placer5 = 0
//  let placer6 = 0
//  let placer7 = 0
//  let placer8 = 0
//  let oilw1 = 0
//  let oilw2 = 0
//  let oilw3 = 0
//  let oilw4 = 0
//  let oilw5 = 0
//  let oilw6 = 0
//  let oilw7 = 0
//  let oilw8 = 0
//  let callw1 = 0
//  let callw2 = 0
//  let callw3 = 0
//  let callw4 = 0
//  let callw5 = 0
//  let callw6 = 0
//  let callw7 = 0
//  let callw8 = 0
//  let fff = 0
//  // console.log(microlists)
//  console.log(microlists[0])
//  microlists.forEach((item, index) => {
//      if (item.产品服务费合计1 < 0 && parseInt(item.产品服务费合计2) == 0) {
//          item.产品服务费与上月相比是否变化 = '新增'
//      }
//      if (item.运营商 == '移动') {
//          if (item.产品服务费与上月相比是否变化 == '新增') {
//              numxz1 = numxz1 + 1
//              money1 = parseInt(money1 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare1 = parseInt(repare1 + item.维护费1 + item.维护费2)
//              placer1 = parseInt(placer1 + item.场地费1 + item.场地费2)
//              oilw2 = parseInt(oilw2 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw1 = parseInt(callw1 + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          else if (item.产品服务费与上月相比是否变化 == '存量') {
//              numcl1 = numcl1 + 1
//              money2 = parseInt(money2 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare2 = parseInt(repare2 + item.维护费1 + item.维护费2)
//              placer2 = parseInt(placer2 + item.场地费1 + item.场地费2)
//              oilw1 = parseInt(oilw1 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw2 = parseInt(callw2 + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//      }
//      else if (item.运营商 == '天府移动') {
//          if (item.产品服务费与上月相比是否变化 == '新增') {
//              numxz2 = numxz2 + 1
//              money3 = parseInt(money3 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare3 = parseInt(repare3 + item.维护费1 + item.维护费2)
//              placer3 = parseInt(placer3 + item.场地费1 + item.场地费2)
//              oilw3 = parseInt(oilw3 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw3 = parseInt(callw3 + item.产品服务费合计0 + item.产品服务费合计2)
//          }
//          else if (item.产品服务费与上月相比是否变化 == '存量') {
//              numcl2 = numcl2 + 1
//              money4 = parseInt(money4 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare4 = parseInt(repare4 + item.维护费1 + item.维护费2)
//              placer4 = parseInt(placer4 + item.场地费1 + item.场地费2)
//              oilw4 = parseInt(oilw4 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw4 = parseInt(oilw4 + item.油机发电服务费1 + item.油机发电服务费2)
//          }
//      }

//      else if (item.运营商 == '电信') {
//          if (item.产品服务费与上月相比是否变化 == '新增') {
//              numxz3 = numxz3 + 1
//              money5 = parseInt(money5 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare5 = parseInt(repare5 + item.维护费1 + item.维护费2)
//              placer5 = parseInt(placer5 + item.场地费1 + item.场地费2)
//              oilw5 = parseInt(oilw5 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw5 = parseInt(oilw5 + item.油机发电服务费1 + item.油机发电服务费2)
//          }
//          else if (item.产品服务费与上月相比是否变化 == '存量') {
//              numcl3 = numcl3 + 1
//              money6 = parseInt(money6 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare6 = parseInt(repare6 + item.维护费1 + item.维护费2)
//              placer6 = parseInt(placer6 + item.场地费1 + item.场地费2)
//              oilw6 = parseInt(oilw6 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw6 = parseInt(oilw6 + item.油机发电服务费1 + item.油机发电服务费2)
//          }
//      }
//      else if (item.运营商 == '联通') {
//          if (item.产品服务费与上月相比是否变化 == '新增') {
//              numxz4 = numxz4 + 1
//              money7 = parseInt(money7 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare7 = parseInt(repare7 + item.维护费1 + item.维护费2)
//              placer7 = parseInt(placer7 + item.场地费1 + item.场地费2)
//              oilw7 = parseInt(oilw7 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw7 = parseInt(oilw7 + item.油机发电服务费1 + item.油机发电服务费2)
//          }
//          else if (item.产品服务费与上月相比是否变化 == '存量') {
//              numcl4 = numcl4 + 1
//              money8 = parseInt(money8 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
//              repare8 = parseInt(repare8 + item.维护费1 + item.维护费2)
//              placer8 = parseInt(placer8 + item.场地费1 + item.场地费2)
//              oilw8 = parseInt(oilw8 + item.油机发电服务费1 + item.油机发电服务费2)
//              callw8 = parseInt(oilw8 + item.油机发电服务费1 + item.油机发电服务费2)
//          }
//      }


//  })
//  console.log(numxz1 + ' ' + money1 + ' ' + (money1 - repare1 - placer1) + ' ' + repare1 + ' ' + placer1 + ' ' + oilw2 + ' ' + callw1)
//  console.log(numcl1 + ' ' + money2 + ' ' + (money2 - repare2 - placer2) + ' ' + repare2 + ' ' + placer2 + ' ' + oilw1 + ' ' + callw2)
//  console.log(numxz2 + ' ' + money3 + ' ' + (money3 - repare3 - placer3) + ' ' + repare3 + ' ' + placer3 + ' ' + oilw3 + ' ' + callw3)
//  console.log(numcl2 + ' ' + money4 + ' ' + (money4 - repare4 - placer4) + ' ' + repare4 + ' ' + placer4 + ' ' + oilw4 + ' ' + callw4)
//  console.log(numxz3 + ' ' + money5 + ' ' + (money5 - repare5 - placer5) + ' ' + repare5 + ' ' + placer5 + ' ' + oilw5 + ' ' + callw5)
//  console.log(numcl3 + ' ' + money6 + ' ' + (money6 - repare6 - placer6) + ' ' + repare6 + ' ' + placer6 + ' ' + oilw6 + ' ' + callw6)
//  console.log(numxz4 + ' ' + money7 + ' ' + (money7 - repare7 - placer7) + ' ' + repare7 + ' ' + placer7 + ' ' + oilw7 + ' ' + callw7)
//  console.log(numcl4 + ' ' + money8 + ' ' + (money8 - repare8 - placer8) + ' ' + repare8 + ' ' + placer8 + ' ' + oilw8 + ' ' + callw8)
//  console.log("❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️")

//  let detailbillsheets = xlsxData.详单文件表[0].data
//  let detailtitle = detailbillsheets[0]
//  let detaillists = []
//  // console.log(detailbillsheets)
//  detailbillsheets.forEach((item, index) => {
//      if (index == 1) {
//          detaillists.push({
//              [detailtitle[1]]: transnum,
//              [detailtitle[2]]: yidong1,
//              [detailtitle[3]]: 0,
//              [detailtitle[4]]: tfyidong1,
//              [detailtitle[5]]: 0,
//              [detailtitle[6]]: liantong1,
//              [detailtitle[7]]: 0,
//              [detailtitle[8]]: dianxing1,
//              [detailtitle[9]]: 0,
//              [detailtitle[10]]: binum,
//              [detailtitle[11]]: stocksf1,
//              [detailtitle[12]]: stocksf11,
//              [detailtitle[13]]: stocksf2,
//              [detailtitle[14]]: stocksf22,
//              [detailtitle[15]]: stocksf3,
//              [detailtitle[16]]: stocksf33,
//              [detailtitle[17]]: stocksf4,
//              [detailtitle[18]]: stocksf44,
//              [detailtitle[19]]: 0,
//              [detailtitle[20]]: binum,
//              [detailtitle[21]]: stocksf1,
//              [detailtitle[22]]: stocksf11,
//              [detailtitle[23]]: stocksf2,
//              [detailtitle[24]]: stocksf22,
//              [detailtitle[25]]: stocksf3,
//              [detailtitle[26]]: stocksf33,
//              [detailtitle[27]]: stocksf4,
//              [detailtitle[28]]: stocksf44
//          })
//      }
//      else if (index == 2) {
//          detaillists.push({
//              [detailtitle[1]]: allbill,
//              [detailtitle[2]]: sumt1,
//              [detailtitle[3]]: 0,
//              [detailtitle[4]]: sumt3,
//              [detailtitle[5]]: 0,
//              [detailtitle[6]]: sumt5,
//              [detailtitle[7]]: 0,
//              [detailtitle[8]]: sumt7,
//              [detailtitle[9]]: 0,
//              [detailtitle[10]]: (test - sum3 - sum5) + (sum2 - sum4 - sum6) + (sum1b - sum3b - sum5b) + (sum2b - sum4b - sum6b) + (sum1c - sum3c - sum5c) + (sum2c - sum4c - sum6c) + (sum1d - sum3d - sum5d) + (sum2d - sum4d - sum6d),
//              [detailtitle[11]]: test - sum3 - sum5,
//              [detailtitle[12]]: sum2 - sum4 - sum6,
//              [detailtitle[13]]: sum1b - sum3b - sum5b,
//              [detailtitle[14]]: sum2b - sum4b - sum6b,
//              [detailtitle[15]]: sum1c - sum3c - sum5c,
//              [detailtitle[16]]: sum2c - sum4c - sum6c,
//              [detailtitle[17]]: sum1d - sum3d - sum5d,
//              [detailtitle[18]]: sum2d - sum4d - sum6d,
//              [detailtitle[19]]: 0,
//              [detailtitle[20]]: binum,
//              [detailtitle[21]]: stocksf1,
//              [detailtitle[22]]: stocksf11,
//              [detailtitle[23]]: stocksf2,
//              [detailtitle[24]]: stocksf22,
//              [detailtitle[25]]: stocksf3,
//              [detailtitle[26]]: stocksf33,
//              [detailtitle[27]]: stocksf4,
//              [detailtitle[28]]: stocksf44
//          })
//      }
//      else if (index == 3) {
//          detaillists.push({
//              [detailtitle[1]]: allbill - (sumt2 + sumt4 + sumt6 + sumt8),
//              [detailtitle[2]]: sumt1 - sumt2,
//              [detailtitle[3]]: 0,
//              [detailtitle[4]]: sumt3 - sumt4,
//              [detailtitle[5]]: 0,
//              [detailtitle[6]]: sumt5 - sumt6,
//              [detailtitle[7]]: 0,
//              [detailtitle[8]]: sumt7 - sumt8,
//              [detailtitle[9]]: 0,
//              [detailtitle[10]]: sum3 + sum4 + sum3b + sum4b + sum3c + sum4c + sum3d + sum4d,
//              [detailtitle[11]]: sum3,
//              [detailtitle[12]]: sum4,
//              [detailtitle[13]]: sum3b,
//              [detailtitle[14]]: sum4b,
//              [detailtitle[15]]: sum3c,
//              [detailtitle[16]]: sum4c,
//              [detailtitle[17]]: sum3d,
//              [detailtitle[18]]: sum4d,
//              [detailtitle[19]]: 0,
//              [detailtitle[20]]: binum,
//              [detailtitle[21]]: stocksf1,
//              [detailtitle[22]]: stocksf11,
//              [detailtitle[23]]: stocksf2,
//              [detailtitle[24]]: stocksf22,
//              [detailtitle[25]]: stocksf3,
//              [detailtitle[26]]: stocksf33,
//              [detailtitle[27]]: stocksf4,
//              [detailtitle[28]]: stocksf44
//          })
//      }
//      else if (index == 4)
//          detaillists.push({
//              [detailtitle[10]]: sum5 + sum6 + sum5b + sum6b + sum5c + sum6c + sum5d + sum6d,
//              [detailtitle[11]]: sum5,
//              [detailtitle[12]]: sum6,
//              [detailtitle[13]]: sum5b,
//              [detailtitle[14]]: sum6b,
//              [detailtitle[15]]: sum5c,
//              [detailtitle[16]]: sum6c,
//              [detailtitle[17]]: sum5d,
//              [detailtitle[18]]: sum6d

//          })
//      else if (index == 5) {
//          detaillists.push({
//              [detailtitle[10]]: sum5 + sum6 + sum5b + sum6b + sum5c + sum6c + sum5d + sum6d,
//              [detailtitle[11]]: sum5,
//              [detailtitle[12]]: sum6,
//              [detailtitle[13]]: sum5b,
//              [detailtitle[14]]: sum6b,
//              [detailtitle[15]]: sum5c,
//              [detailtitle[16]]: sum6c,
//              [detailtitle[17]]: sum5d,
//              [detailtitle[18]]: sum6d

//          })
//      }
//  })

//  const Jsondata = JSON.stringify(detaillists)
//  const filePath = 'D:/typescript/demo/accountbill/data2.json';
//  fs.writeFileSync(filePath, Jsondata);
//  console.log(`已将对象数组保存到${filePath}`);


//  fs.readFile('D:/typescript/demo/accountbill/data2.json', 'utf8', (err, data) => {
//      if (err) throw err;
//      const json = JSON.parse(data);
//      const jsonArray = [];
//      json.forEach(function (item) {
//          let temp = {
//              '传输小计': item.传输小计,
//              '成都移动存量': item.成都移动存量1,
//              '成都移动新增': item.成都移动新增1,
//              '天府移动存量': item.天府移动存量1,
//              '天府移动新增': item.天府移动新增1,
//              '电信存量': item.电信存量1,
//              '电信新增': item.电信新增1,
//              '联通存量': item.联通存量1,
//              '联通新增': item.联通新增1,
//              '室分小计': item.传输小计,
//              '成都移动存量': item.成都移动存量2,
//              '成都移动新增': item.成都移动新增2,
//              '天府移动存量': item.天府移动存量2,
//              '天府移动新增': item.天府移动新增2,
//              '电信存量': item.电信存量2,
//              '电信新增': item.电信新增2,
//              '联通存量': item.联通存量2,
//              '联通新增': item.联通新增2,
//              '室分小计': item.微站小计,
//              '成都移动存量': item.成都移动存量3,
//              '成都移动新增': item.成都移动新增3,
//              '天府移动存量': item.天府移动存量3,
//              '天府移动新增': item.天府移动新增3,
//              '电信存量': item.电信存量3,
//              '电信新增': item.电信新增3,
//              '联通存量': item.联通存量3,
//              '联通新增': item.联通新增3,
//          }
//          jsonArray.push(temp);
//      });

//      let xls = json2xls(jsonArray);

//      fs.writeFileSync('D:/typescript/demo/accountbill/newdetailorder2.xlsx', xls, 'binary');
//      console.log('文件已经保存成功')
//  })

//  console.log('❤❤❤❤❤❤❤❤')
