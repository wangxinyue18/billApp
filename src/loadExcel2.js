// const xlsx2 = require('xlsx');
// //铁塔账单文件
// // let biggerFilePath = fileData.铁塔账单文件表
// let biggerFilePath1 = 'D:/typescript/demo/accountbill/towerorder.xlsx'
// let biggerFilePath2 = 'D:/typescript/demo/accountbill/towerbill.xlsx'
// let biggerFilePath3 = 'D:/typescript/demo/accountbill/forbidden.xlsx'
// async function loadExcel(pathname, sheetNames) {
//     const dense_wb = xlsx2.readFile(pathname, { dense: true });
//     return (sheetNames ? sheetNames : dense_wb.SheetNames).reduce((pre, curr) => {
//         if (!curr) return pre;
//         const sheet = dense_wb.Sheets[curr];
//         pre[curr] = xlsx2.utils.sheet_to_json(sheet, {
//             raw: true
//         });
//         return pre;
//     }, {});
// }
// const date = new Date().valueOf();
// loadExcel(biggerFilePath1).then(data => {
//     const sheetsNames = Object.keys(data);
//     sheetsNames.forEach(name => {
//         // console.log('🍎sheetsName', name, data[name].length, data['铁塔订单']);
//     });
//     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
// })
// loadExcel(biggerFilePath2).then(data => {
//     const sheetsNames = Object.keys(data);
//     sheetsNames.forEach(name => {
//         // console.log('🍎sheetsName', name, data[name].length, data['铁塔订单']);
//     });
//     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
// })
// loadExcel(biggerFilePath3).then(data => {
//     const sheetsNames = Object.keys(data);
//     sheetsNames.forEach(name => {
//         // console.log('🍎sheetsName', name, data[name].length, data['终止订单表']);
//     });
//     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
// })

// async function main() {
//     const data1 = await loadExcel(biggerFilePath1)
//     const data2 = await loadExcel(biggerFilePath2)
//     const data3 = await loadExcel(biggerFilePath3)
//     // console.log('🍑',data1['铁塔订单'])
//     // console.log('🍑',data2['towerbill1'])
//     // console.log('🍑',data3['终止订单表'])
// //     //铁塔订单文件处理
// //     let odTowersheet = data1['铁塔订单']
// //     // let odTowerlist = []
// //     let odtowernum = 0
// //     odTowersheet.forEach((item, index) => {
// //         if (index == 0) {
// //             return
// //         }
// //         else if (item.已暂停出账 != '已暂停计费') {
// //             odtowernum = odtowernum + 1
// //         }
// //     })
// //     console.log('铁塔订单数目：' + odtowernum)
// //     //铁塔账单文件处理
// //     let towerSheet = data2['towerbill1']
// //     let towernum = 0
// //     towerSheet.forEach((item, index) => {
// //         towernum = towernum + 1
// //         if (index == 0) {
// //             return
// //         }
// //         else if (item.运营商 == '移动' && (item.原产权方 == ' 天府新区' || item.原产权方 == ' 双流县' || item.原产权方 == '龙泉驿区')) {
// //             item.原产权方 = '天府移动'

// //         }
// //     })
// //     console.log('铁塔账单数目：' + towernum)
// //     //终止文件处理
// //     let forbidenSheet = data3['终止订单表']
// //     let forbidenlist = []
// //     let forbidennum = 0
// //     forbidenSheet.forEach((item, index) => {
// //         if (index == 0) {
// //             return
// //         }
// //         else if (item.审批状态 == '运营商审批成功') {
// //             forbidennum = forbidennum + 1
// //             forbidenlist.push({
// //                 item
// //             })
// //         }
// //     })
// //     console.log('终止文件数目：' + forbidennum)
// //     // console.log(forbidenlist)


// //     // 从订单文件向账单传输进行对比😀😀😀
// //     let numtower1 = 0
// //     let numtower2 = 0
// //     for (let i = 0; i < odtowernum; i++) {
// //         let numtw4 = 0
// //         let numtw5 = 0
// //         for (let j = 0; j < towernum; j++) {
// //             if (odTowersheet[i].订单号 != towerSheet[j].需求确认单编号) {
// //                 numtw4 = numtw4 + 1
// //             }
// //             else if (odTowersheet[i].订单号 == towerSheet[j].需求确认单编号) {
// //                 //正常订单数目
// //                 numtower1 = numtower1 + 1
// //             }
// //         }
// //         if (numtw4 == towernum) {
// //             // console.log('存在可能异常订单号：'+titlelist[i].订单号)

// //             for (let k = 0; k < forbidennum; k++) {
// //                 if (odTowersheet[i].订单号 == forbidenlist[k].订单编号) {
// //                     // console.log('终止文件存在正常订单号：' + titlelist[i].订单号)
// //                     numtower1 = numtower1 + 1
// //                 }
// //                 else if (odTowersheet[i].订单号 != forbidenlist[k].订单编号) {
// //                     numtw5 = numtw5 + 1
// //                 }
// //                 if (numtw5 == forbidennum) {
// //                     // console.log('异常账号' + odtransmisslist[i].订单号 + '原因：在详单里面，但是不在账单里面')
// //                     numtower2 = numtower2 + 1
// //                 }
// //             }
// //         }
// //     }
// //     //从传输订单文件向订单文件传输
// //     for (let j1 = 0; j1 < towernum; j1++) {
// //         let numtw3 = 0
// //         for (let i1 = 0; i1 < odtowernum; i1++) {
// //             if (towerSheet[j1].需求确认单编号 != odTowersheet[i1].订单号) {
// //                 numtw3 = numtw3 + 1
// //             }
// //             else if (towerSheet[j1].需求确认单编号 == odTowersheet[i1].订单号) {
// //                 // numcsz = numcsz + 1
// //             }
// //         }
// //         if (numtw3 == odtowernum) {
// //             // console.log('异常订单' + transmisslists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
// //             numtower2 = numtower2 + 1
// //         }
// //     }
// //     console.log("正常订单数：（按照订单文件为基准）" + numtower1)
// //     console.log("异常订单数：（账单文件＋订单文件）" + numtower2)



// //     let yidongt = 0
// //     let tfyidongt = 0
// //     let liantongt = 0
// //     let dianxingt = 0

// //     let stocksf1t = 0
// //     let stocksf11t = 0
// //     let stocksf2t = 0
// //     let stocksf22t = 0
// //     let stocksf3t = 0
// //     let stocksf33t = 0
// //     let stocksf4t = 0
// //     let stocksf44t = 0
// //     let sum1t = 0
// //     let sum2t = 0
// //     let sum3t = 0
// //     let sum4t = 0
// //     let sum5t = 0
// //     let sum6t = 0
// //     let sum7t = 0
// //     let sum8t = 0
// //     let sum9t = 0
// //     let sum10t = 0
// //     let sum1bt = 0
// //     let sum2bt = 0
// //     let sum3bt = 0
// //     let sum4bt = 0
// //     let sum5bt = 0
// //     let sum6bt = 0
// //     let sum7bt = 0
// //     let sum8bt = 0
// //     let sum9bt = 0
// //     let sum10bt = 0
// //     let sum1ct = 0
// //     let sum2ct = 0
// //     let sum3ct = 0
// //     let sum4ct = 0
// //     let sum5ct = 0
// //     let sum6ct = 0
// //     let sum7ct = 0
// //     let sum8ct = 0
// //     let sum9ct = 0
// //     let sum10ct = 0
// //     let sum1dt = 0
// //     let sum2dt = 0
// //     let sum3dt = 0
// //     let sum4dt = 0
// //     let sum5dt = 0
// //     let sum6dt = 0
// //     let sum7dt = 0
// //     let sum8dt = 0
// //     let sum9dt = 0
// //     let sum10dt = 0
// //     let testt = 0
// //     // console.log(buildinnlist)
// //     //申明数组
// //     towerSheet.forEach((item, index) => {

// //         if (item.运营商 == '移动') {
// //             if (item.产品服务费与上月相比是否变化 == '存量') {
// //                 stocksf1t = stocksf1t + 1
// //                 testt = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + testt)
// //                 // sum1t = parseInt(item.产品服务费合计1 + item.产品服务费合计2 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费2 + sum1t)
// //                 sum3t = parseInt(sum3t + item.维护费折扣后金额1 + item.维护费折扣后金额2)//正常
// //                 sum5t = parseInt(sum5t + item.场地费折扣后金额1 + item.场地费折扣后金额2)//正常
// //                 sum7t = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum7t)
// //                 sum9t = parseInt(item.产品服务费合计1 + item.产品服务费合计3 + sum9t)
// //             }
// //             else if (item.产品服务费与上月相比是否变化 == '新增') {
// //                 stocksf11t = stocksf11t + 1
// //                 sum2t = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum2t)
// //                 sum4t = parseInt(sum4t + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum6t = parseInt(sum6t + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum8t = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum8t)
// //                 sum10t = parseInt(item.产品服务费合计1 + item.产品服务费合计3 + sum10t)
// //             }
// //             yidongt = yidongt + 1
// //         }
// //         else if (item.运营商 == '天府移动') {
// //             if (item.产品服务费与上月相比是否变化 == '存量') {
// //                 stocksf2t = stocksf2t + 1
// //                 sum1bt = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum1bt)
// //                 // console.log(item.罚责赠费合计)
// //                 sum3bt = parseInt(sum3bt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum5bt = parseInt(sum5bt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum7bt = parseInt(sum7bt + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum9bt = parseInt(sum9bt + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             else if (item.产品服务费与上月相比是否变化 == '新增') {
// //                 stocksf22t = stocksf22t + 1
// //                 sum2bt = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum2bt)
// //                 sum4bt = parseInt(sum4bt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum6bt = parseInt(sum6bt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum8bt = parseInt(sum8bt + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum10bt = parseInt(sum10bt + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             tfyidongt = tfyidongt + 1
// //         }
// //         else if (item.运营商 == '联通') {
// //             if (item.产品服务费与上月相比是否变化 == '存量') {
// //                 stocksf3t = stocksf3t + 1
// //                 sum1ct = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum1ct)
// //                 sum3ct = parseInt(sum3ct + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum5ct = parseInt(sum5ct + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum7ct = parseInt(sum7ct + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum9ct = parseInt(sum9ct + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             else if (item.产品服务费与上月相比是否变化== '新增') {
// //                 stocksf33t = stocksf33t + 1
// //                 sum2ct = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum2ct)
// //                 sum4ct = parseInt(sum4ct + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum6ct = parseInt(sum6ct + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum8ct = parseInt(sum8ct + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum10ct = parseInt(sum10ct + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             liantongt = liantongt + 1
// //         }
// //         else if (item.运营商 == '电信') {
// //             if (item.产品服务费与上月相比是否变化== '存量') {
// //                 stocksf4t = stocksf4t + 1
// //                 sum1dt = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum1dt)
// //                 sum3dt = parseInt(sum3dt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum5dt = parseInt(sum5dt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum7dt = parseInt(sum7dt + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum9dt = parseInt(sum9dt + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             else if (item.产品服务费与上月相比是否变化 == '新增') {
// //                 stocksf44t = stocksf44t + 1
// //                 sum2dt = parseFloat(item.产品服务费合计2 + item.产品服务费合计3 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费1 + sum2dt)
// //                 sum4dt = parseInt(sum4dt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
// //                 sum6dt = parseInt(sum6dt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
// //                 sum8dt = parseInt(sum8dt + item.油机发电服务费1 + item.油机发电服务费2)
// //                 sum10dt = parseInt(sum10dt + item.产品服务费合计1 + item.产品服务费合计3)
// //             }
// //             dianxingt = dianxingt + 1
// //         }
// //     })
// //     console.log(testt)
// //     console.log(testt - sum3t - sum5t)
// //     console.log(yidongt + '  ' + stocksf1t + ' ' + testt + '  ' + (testt - sum3t - sum5t) + ' ' + sum3t + '  ' + sum5t + '  ' + sum7t + ' ' + sum9t)
// //     console.log(yidongt + '  ' + stocksf11t + '  ' + sum2t + ' ' + (sum2t - sum4t - sum6t) + '  ' + sum4t + '  ' + sum6t + '  ' + sum8t + '  ' + sum10t)
// //     console.log(tfyidongt + '  ' + stocksf2t + '  ' + sum1bt + '  ' + (sum1bt - sum3bt - sum5bt) + ' ' + sum3bt + '  ' + sum5bt + '  ' + sum7bt + '  ' + sum9bt)
// //     console.log(tfyidongt + '  ' + stocksf22t + '  ' + sum2bt + '  ' + (sum2bt - sum4bt - sum6bt) + '  ' + sum4bt + '  ' + sum6bt + '  ' + sum8bt + '  ' + sum10bt)
// //     console.log(liantongt + '  ' + stocksf3t + '  ' + sum1ct + '  ' + (sum1ct - sum3ct - sum5ct) + ' ' + sum3ct + '  ' + sum5ct + '  ' + sum7ct + '  ' + sum9ct)
// //     console.log(liantongt + '  ' + stocksf33t + '   ' + sum2ct + '  ' + (sum2ct - sum4ct - sum6ct) + '  ' + sum4ct + '  ' + sum6ct + '  ' + sum8ct + '  ' + sum10ct)
// //     console.log(dianxingt + '   ' + stocksf4t + '  ' + sum1dt + '  ' + (sum1dt - sum3dt - sum5dt) + ' ' + sum3dt + '  ' + sum5dt + '  ' + sum7dt + '  ' + sum9dt)
// //     console.log(dianxingt + ' ' + stocksf44t + '  ' + sum2dt + '  ' + (sum2dt - sum4dt - sum6dt) + '  ' + sum4dt + '  ' + sum6dt + '  ' + sum8dt + '  ' + sum10dt)
// // }
// // main()

// }

