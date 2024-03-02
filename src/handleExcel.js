'use strict'

import installExtension, { VUEJS3_DEVTOOLS } from 'electron-devtools-installer'
import xlsx from 'node-xlsx'
const electron = require('electron');
const fs = require('fs');
const path = require('path');
const json2xls = require('json2xls')
import { transmissHandle } from './transmissHandle.js'
import { loadExcel } from './loadExcel2.js'
import { Console } from 'console';
const xlsx2 = require('xlsx');

export async function handleExcel(fileData) {

    const xlsxData = Object.keys(fileData).reduce((pre, curr) => {
        // console.log('🍎', filePath)
        if (!curr) return pre;
        const filePath = fileData[curr];
        if (curr == '铁塔账单文件表') {
            return pre;
        }
        else {
            const orderarr = xlsx.parse(fs.readFileSync(filePath), {
                cellDate: true
            });
            pre[curr] = orderarr;
        }

        return pre;
    }, {});
    console.log(fileData)
    console.log('上传的文件数目： ' + Object.keys(xlsxData).length + ' 个')
    // console.log(xlsxData)


    console.log('🍑🍑🍑🍑🍑🍑🍑🍑🍑')

    // 传输订单文件😀😀😀
    const odtranmissorsheets = xlsxData.总订单文件表[1].data
    let odtransmisslist = []
    let odtransnum = 0
    const odtransmisstitle = odtranmissorsheets[0];
    odtranmissorsheets.forEach((item, index) => {
        // console.log(item)
        // console.log(index)
        if (index == 0 || index == 1 || index == 2) {
            return
        }
        else {
            odtransmisslist.push(
                {
                    [odtransmisstitle[0]]: item[0],
                    [odtransmisstitle[1]]: item[1],
                    [odtransmisstitle[2]]: item[2],
                    [odtransmisstitle[42]]: item[42],
                    [odtransmisstitle[69]]: item[69]
                }
            )

            odtransnum = odtransnum + 1
            // console.log('❀❀ '+odtransmisslist)
        }
    })
    console.log(odtransmisslist)
    console.log('🍌传输订单文件数（已筛选）：' + odtransnum)
    // const transmiss2 = xlsx.parse("D:/typescript/demo/accountbill/transmission.xlsx", {
    //   cellDates: true,
    // });
    //终止文件😀😀😀
    const forbiden2 = xlsxData.终止文件表
    let forbidnum = 0
    let forbidensheets = forbiden2[0].data
    let forbidlist = []
    const forbidtitle = forbidensheets[0]
    forbidensheets.forEach((item, index) => {
        if (index == 0) {
            return
        }
        else {
            forbidlist.push({
                [forbidtitle[4]]: item[4]

            })
            forbidnum = forbidnum + 1
        }
    })
    console.log("🍑终止文件订单数目： " + forbidnum)
    //传输账单文件
    const transmiss2 = xlsxData.传输账单文件表
    let transmisssheets = transmiss2[0].data
    let transmisslists = []
    const transmisstitle = transmisssheets[1]
    let transnum = 0
    transmisssheets.forEach((item, index) => {
        if (index == 1 || index == 0) {
            return
        }
        else if (item[8] != undefined) {
            transmisslists.push({
                [transmisstitle[2]]: item[2],
                [transmisstitle[8]]: item[8],
                [transmisstitle[19]]: item[19],
                [transmisstitle[20]]: item[20],
                [transmisstitle[27]]: item[27],
                [transmisstitle[28]]: item[28],
                [transmisstitle[38]]: item[38],

            })
            transnum = transnum + 1
        }
    })
    transmisslists.forEach((item, index) => {
        if ((item.运营商 == '移动' && item.运营商区县 == '双流县') || (item.运营商 == '移动' && item.运营商区县 == '龙泉驿区') || (item.运营商 == '移动' && item.运营商区县 == '天府新区')) {
            item.运营商 = '天府移动'
        }
    })
    console.log('🍓传输账单订单数:' + transnum)

    // 从订单文件向账单传输进行对比😀😀😀
    let numcsz = 0
    let numcsy = 0
    for (let i = 0; i < odtransnum; i++) {
        let numtj4 = 0
        let numtj5 = 0
        for (let j = 0; j < transnum; j++) {
            if (odtransmisslist[i].订单号 != transmisslists[j].需求确认单编号) {
                numtj4 = numtj4 + 1
            }
            else if (odtransmisslist[i].订单号 == transmisslists[j].需求确认单编号) {
                //正常订单数目
                numcsz = numcsz + 1
            }
        }
        if (numtj4 == transnum) {
            // console.log('存在可能异常订单号：'+titlelist[i].订单号)

            for (let k = 0; k < forbidnum; k++) {
                if (odtransmisslist[i].订单号 == forbidlist[k].订单编号) {
                    // console.log('终止文件存在正常订单号：' + titlelist[i].订单号)
                    numcsz = numcsz + 1
                }
                else if (odtransmisslist[i].订单号 != forbidlist[k].订单编号) {
                    numtj5 = numtj5 + 1
                }
                if (numtj5 == forbidnum) {
                    // console.log('异常账号' + odtransmisslist[i].订单号 + '原因：在详单里面，但是不在账单里面')
                    numcsy = numcsy + 1
                }
            }
        }
    }
    //从传输订单文件向订单文件传输
    for (let j1 = 0; j1 < transnum; j1++) {
        let numtj3 = 0
        for (let i1 = 0; i1 < odtransnum; i1++) {
            if (transmisslists[j1].需求确认单编号 != odtransmisslist[i1].订单号) {
                numtj3 = numtj3 + 1
            }
            else if (transmisslists[j1].需求确认单编号 == odtransmisslist[i1].订单号) {
                // numcsz = numcsz + 1
            }
        }
        if (numtj3 == odtransnum) {
            // console.log('异常订单' + transmisslists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
            numcsy = numcsy + 1
        }
    }
    console.log("正常订单数：（按照订单文件为基准）" + numcsz)
    console.log("异常订单数：（账单文件＋订单文件）" + numcsy)
    // console.log(transmisslists)
    let yidong1 = 0
    let tfyidong1 = 0
    let liantong1 = 0
    let dianxing1 = 0

    let stocksf1a = 0
    let stocksf2a = 0
    let stocksf3a = 0
    let stocksf4a = 0
    let sumt1 = 0
    let sumt2 = 0
    let sumt3 = 0
    let sumt4 = 0
    let sumt5 = 0
    let sumt6 = 0
    let sumt7 = 0
    let sumt8 = 0
    // 传输只有存量没得新增
    transmisslists.forEach((item, index) => {

        if (item.运营商 == '移动') {
            yidong1 = yidong1 + 1
            stocksf1a = stocksf1a + 1
            sumt1 = parseFloat(sumt1 + item.产品服务费合计1 + item.产品服务费合计2)
            sumt2 = parseFloat(sumt2 + item.维护费1 + item.维护费2)
        }
        else if (item.运营商 == '天府移动') {
            tfyidong1 = tfyidong1 + 1
            stocksf2a = stocksf2a + 1
            sumt3 = parseFloat(sumt3 + item.产品服务费合计1 + item.产品服务费合计2)
            sumt4 = parseFloat(sumt4 + item.维护费1 + item.维护费2)
        }
        else if (item.运营商 == '联通') {
            liantong1 = liantong1 + 1
            stocksf3a = stocksf3a + 1
            sumt5 = parseFloat(sumt5 + item.产品服务费合计1 + item.产品服务费合计2)
            sumt6 = parseFloat(sumt6 + item.维护费1 + item.维护费2)
        }
        else if (item.运营商 == '电信') {
            dianxing1 = dianxing1 + 1
            stocksf4a = stocksf4a + 1
            sumt7 = parseFloat(sumt7 + item.产品服务费合计1 + item.产品服务费合计2)
            sumt8 = parseFloat(sumt8 + item.维护费1 + item.维护费2)
        }
    })
    console.log(yidong1 + ' ' + sumt1 + '  ' + sumt2)
    console.log(tfyidong1 + ' ' + sumt3 + '  ' + sumt4)
    console.log(liantong1 + ' ' + sumt5 + '  ' + sumt6)
    console.log(dianxing1 + ' ' + sumt7 + '  ' + sumt8)

    console.log('❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️  ❤️  ')

    let detailbillsheets = xlsxData.详单文件表[0].data
    let detailtitle = detailbillsheets[0]
    console.log(detailtitle + '❀')
    let detaillists = []
    let allbill = sumt1 + sumt3 + sumt5 + sumt7
    detailbillsheets.forEach((item, index) => {
        if (index == 1) {
            detaillists.push({
                [detailtitle[27]]: (yidong1 + tfyidong1 + liantong1 + dianxing1),
                [detailtitle[28]]: yidong1,
                [detailtitle[29]]: 0,
                [detailtitle[30]]: tfyidong1,
                [detailtitle[31]]: 0,
                [detailtitle[32]]: liantong1,
                [detailtitle[33]]: 0,
                [detailtitle[34]]: dianxing1,
                [detailtitle[35]]: 0
            })
        }
        else if (index == 2) {
            detaillists.push({
                [detailtitle[27]]: allbill,
                [detailtitle[28]]: sumt1,
                [detailtitle[29]]: 0,
                [detailtitle[30]]: sumt3,
                [detailtitle[31]]: 0,
                [detailtitle[32]]: sumt5,
                [detailtitle[33]]: 0,
                [detailtitle[34]]: sumt7,
                [detailtitle[35]]: 0
            })
        }
        else if (index == 3) {
            detaillists.push({
                [detailtitle[27]]: allbill - (sumt2 + sumt4 + sumt6 + sumt8),
                [detailtitle[28]]: sumt1 - sumt2,
                [detailtitle[29]]: 0,
                [detailtitle[30]]: sumt3 - sumt4,
                [detailtitle[31]]: 0,
                [detailtitle[32]]: sumt5 - sumt6,
                [detailtitle[33]]: 0,
                [detailtitle[34]]: sumt7 - sumt8,
                [detailtitle[35]]: 0
            })
        }

        if (index != 0 && index != 1 && index != 2 && index != 3)
            detaillists.push({
                [detailtitle[27]]: 0,
                [detailtitle[28]]: 0,
                [detailtitle[29]]: 0,
                [detailtitle[30]]: 0,
                [detailtitle[31]]: 0,
                [detailtitle[32]]: 0,
                [detailtitle[33]]: 0,
                [detailtitle[34]]: 0,
                [detailtitle[35]]: 0

            })
    })
    console.log(detaillists)
    const Jsondata = JSON.stringify(detaillists)
    const filePath = 'D:/typescript/demo/accountbill/data.json';
    fs.writeFileSync(filePath, Jsondata);
    console.log(`已将对象数组保存到${filePath}`);


    fs.readFile('D:/typescript/demo/accountbill/data.json', 'utf8', (err, data) => {
        if (err) throw err;
        const json = JSON.parse(data);
        const jsonArray = [];
        json.forEach(function (item) {
            let temp = {
                '传输小计': item.传输小计,
                '成都移动存量': item.成都移动存量,
                '成都移动新增': item.成都移动新增,
                '天府移动存量': item.天府移动存量,
                '天府移动新增': item.天府移动新增,
                '电信存量': item.电信存量,
                '电信新增': item.电信新增,
                '联通存量': item.联通存量,
                '联通新增': item.联通新增,
            }
            jsonArray.push(temp);
        });

        let xls = json2xls(jsonArray);

        fs.writeFileSync('D:/typescript/demo/accountbill/transmissionbill.xlsx', xls, 'binary');
        console.log('文件已经保存成功🍌')
    })

    console.log('\^o^/\^o^/\^o^/\^o^/\^o^/\^o^/')



    // 订单文件室分😀😀😀
    let odinnersheet2 = xlsxData.总订单文件表[0].data
    let odinnerrlist = []
    // 获取标题行
    const orderinnertitle = odinnersheet2[2];
    // console.log(ordertitle)
    let odnum = 0
    odinnersheet2.forEach((item, index) => {
        // console.log(item)
        // console.log(index)
        if (index == 0 || index == 1) {
            return
        }
        else if (item[0] != undefined && item[1] == '已起租' && item[95] != '0.00') {
            odinnerrlist.push(
                {
                    [orderinnertitle[0]]: item[0],
                    [orderinnertitle[1]]: item[1],
                    [orderinnertitle[2]]: item[2],
                    [orderinnertitle[95]]: item[95],
                }
            )
            odnum = odnum + 1
        }
    })
    console.log('室分订单文件数（已筛选）' + odnum)

    //账期订单文件😀😀😀
    const buildinnfile = xlsxData.室分账单文件表
    let binum = 0
    let buildinnsheet = buildinnfile[0].data
    // console.log(buildinnsheet)
    let buildinnlist = []
    const buildinntitle = buildinnsheet[0]
    buildinnsheet.forEach((item, index) => {
        if (index == 0) {
            return
        }
        else if (item[2] == '移动' && (item[76] == ' 天府新区' || item[76] == ' 双流县' || item[76] == '龙泉驿区')) {
            item[2] = '天府移动'

        }
        else if (item[8] != undefined) {
            buildinnlist.push({
                [buildinntitle[2]]: item[2],
                [buildinntitle[8]]: item[8],
                [buildinntitle[59]]: item[59],
                [buildinntitle[56]]: item[56],
                [buildinntitle[57]]: item[57],
                [buildinntitle[58]]: item[58],
                [buildinntitle[67]]: item[67],
                [buildinntitle[73]]: item[73],
                [buildinntitle[74]]: item[74],
                [buildinntitle[75]]: item[75],
                [buildinntitle[76]]: item[76],
                [buildinntitle[41]]: item[41],
                [buildinntitle[42]]: item[42],
                [buildinntitle[31]]: item[31],
                [buildinntitle[32]]: item[32],

            })
            binum = binum + 1
        }

    })
    buildinnlist.forEach((item, index) => {
        if ((item.运营商 == '移动' && item.运营商区县 == '双流县') || (item.运营商 == '移动' && item.运营商区县 == '龙泉驿区') || (item.运营商 == '移动' && item.运营商区县 == '天府新区')) {
            item.运营商 = '天府移动'
        }
    })
    console.log('室分账单订单数' + binum)


    // 将室分订单文件和账单传输进行对比😀😀😀
    let numcfz = 0
    let numcfy = 0

    for (let i = 0; i < odnum; i++) {
        let num8 = 0
        let num9 = 0
        for (let j = 0; j < binum; j++) {
            if (odinnerrlist[i].订单号 != buildinnlist[j].需求确认单编号) {
                num8 = num8 + 1
            }
            else if (odinnerrlist[i].订单号 == buildinnlist[j].需求确认单编号) {
                numcfz = numcfz + 1
                // console.log('正常订单号：' + odinnerrlist[i].订单号)
            }
        }
        if (num8 == binum) {
            for (let k = 0; k < forbidnum; k++) {
                if (odinnerrlist[i].订单号 != forbidlist[k].订单编号) {
                    num9 = num9 + 1
                }
                else if (odinnerrlist[i].订单号 == forbidlist[k].订单编号) {
                    numcfz = numcfz + 1
                }
            }
            if (num9 == forbidnum) {
                // console.log('存在异常账号：' + odinnerrlist[i].订单号)
                numcfy = numcfy + 1
            }
        }
    }

    //从账单文件向订单文件遍历订单是否异常
    for (let j1 = 0; j1 < binum; j1++) {
        let numtj3 = 0
        for (let i1 = 0; i1 < odnum; i1++) {
            if (buildinnlist[j1].需求确认单编号 != odinnerrlist[i1].订单号) {
                numtj3 = numtj3 + 1
            }
            else if (buildinnlist[j1].需求确认单编号 == buildinnlist[i1].订单号) {
                // numcfz = numcfz + 1
            }
        }
        if (numtj3 == odnum) {
            // console.log('异常订单' + buildinnlist[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
            numcfy = numcfy + 1
        }
    }
    console.log("正常订单数：（按照订单文件为基准）" + numcfz)
    console.log("异常订单数：（账单文件＋订单文件）" + numcfy)
    // console.log(buildinnlist)

    let yidong = 0
    let tfyidong = 0
    let liantong = 0
    let dianxing = 0

    let stocksf1 = 0
    let stocksf11 = 0
    let stocksf2 = 0
    let stocksf22 = 0
    let stocksf3 = 0
    let stocksf33 = 0
    let stocksf4 = 0
    let stocksf44 = 0
    let sum1 = 0
    let sum2 = 0
    let sum3 = 0
    let sum4 = 0
    let sum5 = 0
    let sum6 = 0
    let sum7 = 0
    let sum8 = 0
    let sum9 = 0
    let sum10 = 0
    let sum1b = 0
    let sum2b = 0
    let sum3b = 0
    let sum4b = 0
    let sum5b = 0
    let sum6b = 0
    let sum7b = 0
    let sum8b = 0
    let sum9b = 0
    let sum10b = 0
    let sum1c = 0
    let sum2c = 0
    let sum3c = 0
    let sum4c = 0
    let sum5c = 0
    let sum6c = 0
    let sum7c = 0
    let sum8c = 0
    let sum9c = 0
    let sum10c = 0
    let sum1d = 0
    let sum2d = 0
    let sum3d = 0
    let sum4d = 0
    let sum5d = 0
    let sum6d = 0
    let sum7d = 0
    let sum8d = 0
    let sum9d = 0
    let sum10d = 0
    let test = 0
    // console.log(buildinnlist)
    //申明数组
    buildinnlist.forEach((item, index) => {

        if (item.运营商 == '移动') {
            if (item.产品服务费与上月相比是否变化a == '存量') {
                stocksf1 = stocksf1 + 1
                test = parseFloat(item.产品服务费合计1 + test)
                sum1 = parseInt(item.产品服务费合计1 + item.产品服务费合计2 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费2 + sum1)
                sum3 = parseInt(sum3 + item.维护费折扣后金额1 + item.维护费折扣后金额2)//正常
                sum5 = parseInt(sum5 + item.场地费折扣后金额1 + item.场地费折扣后金额2)//正常
                sum7 = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum7)
                sum9 = parseInt(item.产品服务费合计0 + item.产品服务费合计2 + sum9)
            }
            else if (item.产品服务费与上月相比是否变化a == '新增') {
                stocksf11 = stocksf11 + 1
                sum2 = parseInt(item.产品服务费合计1 + sum2)
                sum4 = parseInt(sum4 + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum6 = parseInt(sum6 + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum8 = parseInt(item.油机发电服务费1 + item.油机发电服务费2 + sum8)
                sum10 = parseInt(item.产品服务费合计0 + item.产品服务费合计2 + sum10)
            }
            yidong = yidong + 1
        }
        else if (item.运营商 == '天府移动') {
            if (item.产品服务费与上月相比是否变化a == '存量') {
                stocksf2 = stocksf2 + 1
                sum1b = parseInt(item.产品服务费合计1 + sum1b)
                // console.log(item.罚责赠费合计)
                sum3b = parseInt(sum3b + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum5b = parseInt(sum5b + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum7b = parseInt(sum7b + item.油机发电服务费1 + item.油机发电服务费2)
                sum9b = parseInt(sum9b + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化a == '新增') {
                stocksf22 = stocksf22 + 1
                sum2b = parseInt(sum2b + item.产品服务费合计1)
                sum4b = parseInt(sum4b + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum6b = parseInt(sum6b + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum8b = parseInt(sum8b + item.油机发电服务费1 + item.油机发电服务费2)
                sum10b = parseInt(sum10b + item.产品服务费合计0 + item.产品服务费合计2)
            }
            tfyidong = tfyidong + 1
        }
        else if (item.运营商 == '联通') {
            if (item.产品服务费与上月相比是否变化a == '存量') {
                stocksf3 = stocksf3 + 1
                sum1c = parseInt(item.产品服务费合计1 + sum1c)
                sum3c = parseInt(sum3c + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum5c = parseInt(sum5c + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum7c = parseInt(sum7c + item.油机发电服务费1 + item.油机发电服务费2)
                sum9c = parseInt(sum9c + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化a == '新增') {
                stocksf33 = stocksf33 + 1
                sum2c = parseInt(item.产品服务费合计1 + sum2c)
                sum4c = parseInt(sum4c + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum6c = parseInt(sum6c + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum8c = parseInt(sum8c + item.油机发电服务费1 + item.油机发电服务费2)
                sum10c = parseInt(sum10c + item.产品服务费合计0 + item.产品服务费合计2)
            }
            liantong = liantong + 1
        }
        else if (item.运营商 == '电信') {
            if (item.产品服务费与上月相比是否变化a == '存量') {
                stocksf4 = stocksf4 + 1
                sum1d = parseInt(item.产品服务费合计1 + sum1d)
                sum3d = parseInt(sum3d + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum5d = parseInt(sum5d + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum7d = parseInt(sum7d + item.油机发电服务费1 + item.油机发电服务费2)
                sum9d = parseInt(sum9d + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化a == '新增') {
                stocksf44 = stocksf44 + 1
                sum2d = parseInt(item.产品服务费合计1 + sum2d)
                sum4d = parseInt(sum4d + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                sum6d = parseInt(sum6d + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                sum8d = parseInt(sum8d + item.油机发电服务费1 + item.油机发电服务费2)
                sum10d = parseInt(sum10d + item.产品服务费合计0 + item.产品服务费合计2)
            }
            dianxing = dianxing + 1
        }
    })
    // console.log(test)
    // console.log(test - sum3 - sum5)
    console.log(yidong + '  ' + stocksf1 + ' ' + test + '  ' + (test - sum3 - sum5) + ' ' + sum3 + '  ' + sum5 + '  ' + sum7 + ' ' + sum9)
    console.log(yidong + '  ' + stocksf11 + '  ' + sum2 + ' ' + (sum2 - sum4 - sum6) + '  ' + sum4 + '  ' + sum6 + '  ' + sum8 + '  ' + sum10)
    console.log(tfyidong + '  ' + stocksf2 + '  ' + sum1b + '  ' + (sum1b - sum3b - sum5b) + ' ' + sum3b + '  ' + sum5b + '  ' + sum7b + '  ' + sum9b)
    console.log(tfyidong + '  ' + stocksf22 + '  ' + sum2b + '  ' + (sum2b - sum4b - sum6b) + '  ' + sum4b + '  ' + sum6b + '  ' + sum8b + '  ' + sum10b)
    console.log(liantong + '  ' + stocksf3 + '  ' + sum1c + '  ' + (sum1c - sum3c - sum5c) + ' ' + sum3c + '  ' + sum5c + '  ' + sum7c + '  ' + sum9c)
    console.log(liantong + '  ' + stocksf33 + '   ' + sum2c + '  ' + (sum2c - sum4c - sum6c) + '  ' + sum4c + '  ' + sum6c + '  ' + sum8c + '  ' + sum10c)
    console.log(dianxing + '   ' + stocksf4 + '  ' + sum1d + '  ' + (sum1d - sum3d - sum5d) + ' ' + sum3d + '  ' + sum5d + '  ' + sum7d + '  ' + sum9d)
    console.log(dianxing + ' ' + stocksf44 + '  ' + sum2d + '  ' + (sum2d - sum4d - sum6d) + '  ' + sum4d + '  ' + sum6d + '  ' + sum8d + '  ' + sum10d)

    console.log((test - sum3 - sum5))
    let detailbillsheets2 = xlsxData.详单文件表[0].data
    let detailtitle2 = detailbillsheets2[0]
    // console.log(detailtitle + '❀')
    let detaillists2 = []
    detailbillsheets2.forEach((item, index) => {
        // console.log(index)
        if (index == 0) {
            console.log("😊")
            detaillists2.push({
                [detailtitle2[18]]: yidong + tfyidong + liantong + dianxing,
                [detailtitle2[19]]: stocksf1,
                [detailtitle2[20]]: stocksf11,
                [detailtitle2[21]]: stocksf2,
                [detailtitle2[22]]: stocksf22,
                [detailtitle2[23]]: stocksf3,
                [detailtitle2[24]]: stocksf33,
                [detailtitle2[25]]: stocksf4,
                [detailtitle2[26]]: stocksf44
            })
        }
        else if (index == 1) {
            console.log("🤞")
            detaillists2.push({
                [detailtitle2[18]]: test + sum2 + sum1b + sum2b + sum1c + sum2c + sum1d + sum2d,
                [detailtitle2[19]]: test,
                [detailtitle2[20]]: sum2,
                [detailtitle2[21]]: sum1b,
                [detailtitle2[22]]: sum2b,
                [detailtitle2[23]]: sum1c,
                [detailtitle2[24]]: sum2c,
                [detailtitle2[25]]: sum1d,
                [detailtitle2[26]]: sum2d
            })
        }
        else if (index == 2) {
            console.log("👌")
            detaillists2.push({
                [detailtitle2[18]]: (test - sum3 - sum5) + (sum2 - sum4 - sum6) + (sum1b - sum3b - sum5b) + (sum2b - sum4b - sum6b) + (sum1c - sum3c - sum5c) + (sum2c - sum4c - sum6c) + (sum1d - sum3d - sum5d) + (sum2d - sum4d - sum6d),
                [detailtitle2[19]]: (test - sum3 - sum5),
                [detailtitle2[20]]: (sum2 - sum4 - sum6),
                [detailtitle2[21]]: (sum1b - sum3b - sum5b),
                [detailtitle2[22]]: (sum2b - sum4b - sum6b),
                [detailtitle2[23]]: (sum1c - sum3c - sum5c),
                [detailtitle2[24]]: (sum2c - sum4c - sum6c),
                [detailtitle2[25]]: (sum1d - sum3d - sum5d),
                [detailtitle2[26]]: (sum2d - sum4d - sum6d)
            })
        }

        else if (index == 3) {
            console.log("❀")
            detaillists2.push({
                [detailtitle2[18]]: sum3 + sum4 + sum3b + sum4b + sum3c + sum4c + sum3d + sum4d,
                [detailtitle2[19]]: sum3,
                [detailtitle2[20]]: sum4,
                [detailtitle2[21]]: sum3b,
                [detailtitle2[22]]: sum4b,
                [detailtitle2[23]]: sum3c,
                [detailtitle2[24]]: sum4c,
                [detailtitle2[25]]: sum3d,
                [detailtitle2[26]]: sum4d

            })
        }
        else if (index == 4) {
            console.log("🌹")
            detaillists2.push({
                [detailtitle2[18]]: sum5 + sum6 + sum5b + sum6b + sum5c + sum6c + sum5d + sum5d,
                [detailtitle2[19]]: sum5,
                [detailtitle2[20]]: sum6,
                [detailtitle2[21]]: sum5b,
                [detailtitle2[22]]: sum6b,
                [detailtitle2[23]]: sum5c,
                [detailtitle2[24]]: sum6c,
                [detailtitle2[25]]: sum5d,
                [detailtitle2[26]]: sum6d

            })


        }
        else if (index == 5) {
            console.log("🍊")
            detaillists2.push({
                [detailtitle2[18]]: sum7 + sum8 + sum7b + sum8b + sum7c + sum8c + sum7d + sum8d,
                [detailtitle2[19]]: sum7,
                [detailtitle2[20]]: sum8,
                [detailtitle2[21]]: sum7b,
                [detailtitle2[22]]: sum8b,
                [detailtitle2[23]]: sum7c,
                [detailtitle2[24]]: sum8c,
                [detailtitle2[25]]: sum7d,
                [detailtitle2[26]]: sum8d

            })
        }
        else if (index == 6) {
            console.log("🍑")
            detaillists2.push({
                [detailtitle2[18]]: 0,
                [detailtitle2[19]]: 0,
                [detailtitle2[20]]: 0,
                [detailtitle2[21]]: 0,
                [detailtitle2[22]]: 0,
                [detailtitle2[23]]: 0,
                [detailtitle2[24]]: 0,
                [detailtitle2[25]]: 0,
                [detailtitle2[26]]: 0

            })
        }
    })
    console.log(detaillists2)
    const Jsondata2 = JSON.stringify(detaillists2)
    const filePath2 = 'D:/typescript/demo/accountbill/data2.json';
    fs.writeFileSync(filePath2, Jsondata2);
    console.log(`已将对象数组保存到${filePath2}`);


    fs.readFile('D:/typescript/demo/accountbill/data2.json', 'utf8', (err, data) => {
        if (err) throw err;
        const json = JSON.parse(data);
        const jsonArray = [];
        json.forEach(function (item) {
            let temp = {
                '室分小计': item.室分小计,
                '成都移动存量': item.成都移动存量,
                '成都移动新增': item.成都移动新增,
                '天府移动存量': item.天府移动存量,
                '天府移动新增': item.天府移动新增,
                '联通存量': item.联通存量,
                '联通新增': item.联通新增,
                '电信存量': item.电信存量,
                '电信新增': item.电信新增,
            }
            jsonArray.push(temp);
        });

        let xls = json2xls(jsonArray);

        fs.writeFileSync('D:/typescript/demo/accountbill/build.xlsx', xls, 'binary');
        console.log('文件已经保存成功🍑')
    })

    console.log('\^o^/\^o^/\^o^/\^o^/\^o^/\^o^/')


    //微站
    const microfile = xlsxData.微站账单文件表
    let microsheet = microfile[0].data
    const microtitle = microsheet[0]
    let ordermicro = xlsxData.总订单文件表[2].data
    const ordermicrotitle = ordermicro[2]
    let microOdlists = []
    let microlists = []
    let ordernum = 0
    let micronum = 0
    //遍历微站订单已筛选订单
    ordermicro.forEach((item, index) => {
        if (index == 0 || index == 1 || index == 2) {
            return
        }
        else if (item[0] != undefined && item[1] == '已起租' && item[50] != '0.00' && item[87] != '已暂停出账') {
            microOdlists.push({
                [ordermicrotitle[1]]: item[1],
                [ordermicrotitle[2]]: item[2],

            })
            ordernum = ordernum + 1

        }
    })
    // console.log(microOdlists)
    //遍历微站账单订单
    microsheet.forEach((item, index) => {
        if (index == 0) {
            return
        }
        else {
            microlists.push({
                [microtitle[9]]: item[9],
                [microtitle[2]]: item[2],
                [microtitle[21]]: item[21],
                [microtitle[22]]: item[22],
                [microtitle[25]]: item[25],
                [microtitle[26]]: item[26],
                [microtitle[52]]: item[52],
                [microtitle[53]]: item[53],
                [microtitle[54]]: item[54],
                [microtitle[55]]: item[55],
                [microtitle[69]]: item[69],
                [microtitle[70]]: item[70],
            })
        }
        micronum = micronum + 1
    })

    console.log("微站订单文件数（已筛选）：" + ordernum)
    console.log("微站账单订单数：" + micronum)
    //从订单文件向账单文件
    let numz = 0
    let numy = 0
    for (let i = 0; i < ordernum; i++) {
        let numtj = 0
        let numtj2 = 0
        for (let j = 0; j < micronum; j++) {
            if (microOdlists[i].订单号 != microlists[j].需求确认单编号) {
                numtj = numtj + 1
            }
            else if (microOdlists[i].订单号 == microlists[j].需求确认单编号) {
                numz = numz + 1
                // console.log('正常订单'+microOdlists[i].订单号)
            }
        }
        if (numtj == micronum) {

            for (let k = 0; k < forbidnum; k++) {
                if (microOdlists[i].订单号 == forbidlist[k].订单编号) {
                    numz = numz + 1
                    //  console.log('正常订单'+microOdlists[i].订单号)
                }
                else if (microOdlists[i].订单号 != forbidlist[k].订单编号) {
                    numtj2 = numtj2 + 1
                }
            }
            if (numtj2 == forbidnum) {
                numy = numy + 1
                // console.log('异常账号' + microOdlists[i].订单号 + '原因：在详单里面，但是不在账单里面')
            }
        }
    }

    //从账单文件向订单文件遍历订单是否异常
    for (let j1 = 0; j1 < micronum; j1++) {
        let numtj3 = 0
        for (let i1 = 0; i1 < ordernum; i1++) {
            if (microlists[j1].需求确认单编号 != microOdlists[i1].订单号) {
                numtj3 = numtj3 + 1
            }
            else if (microlists[j1].需求确认单编号 == microOdlists[i1].订单号) {
                // numz = numz + 1
                // console.log('正常订单' + microlists[j1].需求确认单编号)


            }
        }
        if (numtj3 == ordernum) {
            // console.log('异常订单' + microlists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
            numy = numy + 1
        }
    }
    console.log("正常订单数：（按照订单文件为基准）" + numz)
    console.log("异常订单数：（账单文件＋订单文件）" + numy)

    //算新增和存量
    let numxz1 = 0
    let numxz2 = 0
    let numxz3 = 0
    let numxz4 = 0
    let numcl1 = 0
    let numcl2 = 0
    let numcl3 = 0
    let numcl4 = 0
    let money1 = 0
    let money2 = 0
    let money3 = 0
    let money4 = 0
    let money5 = 0
    let money6 = 0
    let money7 = 0
    let money8 = 0
    let repare1 = 0
    let repare2 = 0
    let repare3 = 0
    let repare4 = 0
    let repare5 = 0
    let repare6 = 0
    let repare7 = 0
    let repare8 = 0
    let placer1 = 0
    let placer2 = 0
    let placer3 = 0
    let placer4 = 0
    let placer5 = 0
    let placer6 = 0
    let placer7 = 0
    let placer8 = 0
    let oilw1 = 0
    let oilw2 = 0
    let oilw3 = 0
    let oilw4 = 0
    let oilw5 = 0
    let oilw6 = 0
    let oilw7 = 0
    let oilw8 = 0
    let callw1 = 0
    let callw2 = 0
    let callw3 = 0
    let callw4 = 0
    let callw5 = 0
    let callw6 = 0
    let callw7 = 0
    let callw8 = 0
    let fff = 0
    microlists.forEach((item, index) => {
        if (item.产品服务费合计1 < 0 && parseInt(item.产品服务费合计2) == 0) {
            item.产品服务费与上月相比是否变化 = '新增'
        }
        if (item.运营商 == '移动') {
            if (item.产品服务费与上月相比是否变化 == '新增') {
                numxz1 = numxz1 + 1
                money1 = parseInt(money1 + item.产品服务费合计1)
                repare1 = parseInt(repare1 + item.维护费1 + item.维护费2)
                placer1 = parseInt(placer1 + item.场地费1 + item.场地费2)
                oilw2 = parseInt(oilw2 + item.油机发电服务费1 + item.油机发电服务费2)
                callw1 = parseInt(callw1 + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化 == '存量') {
                numcl1 = numcl1 + 1
                money2 = parseInt(money2 + item.产品服务费合计1)
                repare2 = parseInt(repare2 + item.维护费1 + item.维护费2)
                placer2 = parseInt(placer2 + item.场地费1 + item.场地费2)
                oilw1 = parseInt(oilw1 + item.油机发电服务费1 + item.油机发电服务费2)
                callw2 = parseInt(callw2 + item.产品服务费合计0 + item.产品服务费合计2)
            }
        }
        else if (item.运营商 == '天府移动') {
            if (item.产品服务费与上月相比是否变化 == '新增') {
                numxz2 = numxz2 + 1
                money3 = parseInt(money3 + item.产品服务费合计1)
                repare3 = parseInt(repare3 + item.维护费1 + item.维护费2)
                placer3 = parseInt(placer3 + item.场地费1 + item.场地费2)
                oilw3 = parseInt(oilw3 + item.油机发电服务费1 + item.油机发电服务费2)
                callw3 = parseInt(callw3 + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化 == '存量') {
                numcl2 = numcl2 + 1
                money4 = parseInt(money4 + item.产品服务费合计1)
                repare4 = parseInt(repare4 + item.维护费1 + item.维护费2)
                placer4 = parseInt(placer4 + item.场地费1 + item.场地费2)
                oilw4 = parseInt(oilw4 + item.油机发电服务费1 + item.油机发电服务费2)
                callw4 = parseInt(callw4 + item.产品服务费合计0 + item.产品服务费合计2)
            }
        }

        else if (item.运营商 == '电信') {
            if (item.产品服务费与上月相比是否变化 == '新增') {
                numxz3 = numxz3 + 1
                money5 = parseInt(money5 + item.产品服务费合计1)
                repare5 = parseInt(repare5 + item.维护费1 + item.维护费2)
                placer5 = parseInt(placer5 + item.场地费1 + item.场地费2)
                oilw5 = parseInt(oilw5 + item.油机发电服务费1 + item.油机发电服务费2)
                callw5 = parseInt(callw5 + item.油机发电服务费1 + item.油机发电服务费2)
            }
            else if (item.产品服务费与上月相比是否变化 == '存量') {
                numcl3 = numcl3 + 1
                money6 = parseInt(money6 + item.产品服务费合计1 + item.产品服务费合计2 + item.油机发电服务费1 + item.油机发电服务费2)
                repare6 = parseInt(repare6 + item.维护费1 + item.维护费2)
                placer6 = parseInt(placer6 + item.场地费1 + item.场地费2)
                oilw6 = parseInt(oilw6 + item.油机发电服务费1 + item.油机发电服务费2)
                callw6 = parseInt(callw6 + item.产品服务费合计0 + item.产品服务费合计2)
            }
        }
        else if (item.运营商 == '联通') {
            if (item.产品服务费与上月相比是否变化 == '新增') {
                numxz4 = numxz4 + 1
                money7 = parseInt(money7 + item.产品服务费合计1)
                repare7 = parseInt(repare7 + item.维护费1 + item.维护费2)
                placer7 = parseInt(placer7 + item.场地费1 + item.场地费2)
                oilw7 = parseInt(oilw7 + item.油机发电服务费1 + item.油机发电服务费2)
                callw7 = parseInt(callw7 + item.产品服务费合计0 + item.产品服务费合计2)
            }
            else if (item.产品服务费与上月相比是否变化 == '存量') {
                numcl4 = numcl4 + 1
                money8 = parseInt(money8 + item.产品服务费合计1)
                repare8 = parseInt(repare8 + item.维护费1 + item.维护费2)
                placer8 = parseInt(placer8 + item.场地费1 + item.场地费2)
                oilw8 = parseInt(oilw8 + item.油机发电服务费1 + item.油机发电服务费2)
                callw8 = parseInt(callw8 + item.产品服务费合计0 + item.产品服务费合计2)
            }
        }


    })
    console.log(numxz1 + ' ' + money1 + ' ' + (money1 - repare1 - placer1) + ' ' + repare1 + ' ' + placer1 + ' ' + oilw2 + ' ' + callw1)
    console.log(numcl1 + ' ' + money2 + ' ' + (money2 - repare2 - placer2) + ' ' + repare2 + ' ' + placer2 + ' ' + oilw1 + ' ' + callw2)
    console.log(numxz2 + ' ' + money3 + ' ' + (money3 - repare3 - placer3) + ' ' + repare3 + ' ' + placer3 + ' ' + oilw3 + ' ' + callw3)
    console.log(numcl2 + ' ' + money4 + ' ' + (money4 - repare4 - placer4) + ' ' + repare4 + ' ' + placer4 + ' ' + oilw4 + ' ' + callw4)
    console.log(numxz3 + ' ' + money5 + ' ' + (money5 - repare5 - placer5) + ' ' + repare5 + ' ' + placer5 + ' ' + oilw5 + ' ' + callw5)
    console.log(numcl3 + ' ' + money6 + ' ' + (money6 - repare6 - placer6) + ' ' + repare6 + ' ' + placer6 + ' ' + oilw6 + ' ' + callw6)
    console.log(numxz4 + ' ' + money7 + ' ' + (money7 - repare7 - placer7) + ' ' + repare7 + ' ' + placer7 + ' ' + oilw7 + ' ' + callw7)
    console.log(numcl4 + ' ' + money8 + ' ' + (money8 - repare8 - placer8) + ' ' + repare8 + ' ' + placer8 + ' ' + oilw8 + ' ' + callw8)
    console.log("❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️ ❤️")


    let detailbillsheets3 = xlsxData.详单文件表[0].data
    let detailtitle3 = detailbillsheets3[0]
    let detaillists3 = []
    console.log(detailbillsheets3)
    detailbillsheets3.forEach((item, index) => {
        if (index == 0) {
            detaillists3.push({
                [detailtitle3[9]]: numxz1 + numcl1 + numxz2 + numcl2 + numxz3 + numcl3 + numxz4 + numcl4,
                [detailtitle3[10]]: numxz1,
                [detailtitle3[11]]: numcl1,
                [detailtitle3[12]]: numxz2,
                [detailtitle3[13]]: numcl2,
                [detailtitle3[14]]: numxz3,
                [detailtitle3[15]]: numcl3,
                [detailtitle3[16]]: numxz4,
                [detailtitle3[17]]: numcl4,

            })
        }
        else if (index == 1) {
            detaillists3.push({
                [detailtitle3[9]]: money1 + money2 + money3 + money4 + money5 + money6 + money7 + money8,
                [detailtitle3[10]]: money1,
                [detailtitle3[11]]: money2,
                [detailtitle3[12]]: money3,
                [detailtitle3[13]]: money4,
                [detailtitle3[14]]: money5,
                [detailtitle3[15]]: money6,
                [detailtitle3[16]]: money7,
                [detailtitle3[17]]: money8,

            })
        }
        else if (index == 2) {
            detaillists3.push({
                [detailtitle3[9]]: (money1 - repare1 - placer1) + (money2 - repare2 - placer2) + (money3 - repare3 - placer3) + (money4 - repare4 - placer4) + (money5 - repare5 - placer5) + (money6 - repare6 - placer6) + (money7 - repare7 - placer7) + (money8 - repare8 - placer8),
                [detailtitle3[10]]: (money1 - repare1 - placer1),
                [detailtitle3[11]]: (money2 - repare2 - placer2),
                [detailtitle3[12]]: (money3 - repare3 - placer3),
                [detailtitle3[13]]: (money4 - repare4 - placer4),
                [detailtitle3[14]]: (money5 - repare5 - placer5),
                [detailtitle3[15]]: (money6 - repare6 - placer6),
                [detailtitle3[16]]: (money7 - repare7 - placer7),
                [detailtitle3[17]]: (money8 - repare8 - placer8),

            })
        }

        else if (index == 3) {
            detaillists3.push({
                [detailtitle3[9]]: repare1 + repare2 + repare3 + repare4 + repare5 + repare6 + repare7 + repare8,
                [detailtitle3[10]]: repare1,
                [detailtitle3[11]]: repare2,
                [detailtitle3[12]]: repare3,
                [detailtitle3[13]]: repare4,
                [detailtitle3[14]]: repare5,
                [detailtitle3[15]]: repare6,
                [detailtitle3[16]]: repare7,
                [detailtitle3[17]]: repare8,


            })
        }
        else if (index == 4) {
            detaillists3.push({
                [detailtitle3[9]]: placer1 + placer2 + placer3 + placer4 + placer5 + placer6 + placer7 + placer8,
                [detailtitle3[10]]: placer1,
                [detailtitle3[11]]: placer2,
                [detailtitle3[12]]: placer3,
                [detailtitle3[13]]: placer4,
                [detailtitle3[14]]: placer5,
                [detailtitle3[15]]: placer6,
                [detailtitle3[16]]: placer7,
                [detailtitle3[17]]: placer8,


            })

        }
        else if (index == 5) {
            detaillists3.push({
                [detailtitle3[9]]: 0,
                [detailtitle3[10]]: 0,
                [detailtitle3[11]]: 0,
                [detailtitle3[12]]: 0,
                [detailtitle3[13]]: 0,
                [detailtitle3[14]]: 0,
                [detailtitle3[15]]: 0,
                [detailtitle3[16]]: 0,
                [detailtitle3[17]]: 0,


            })
        }
        else if (index == 6) {
            detaillists3.push({
                [detailtitle3[9]]: 0,
                [detailtitle3[10]]: 0,
                [detailtitle3[11]]: 0,
                [detailtitle3[12]]: 0,
                [detailtitle3[13]]: 0,
                [detailtitle3[14]]: 0,
                [detailtitle3[15]]: 0,
                [detailtitle3[16]]: 0,
                [detailtitle3[17]]: 0

            })
        }
    })
    console.log(detaillists3)
    const Jsondata3 = JSON.stringify(detaillists3)
    const filePath3 = 'D:/typescript/demo/accountbill/data3.json';
    fs.writeFileSync(filePath3, Jsondata3);
    console.log(`已将对象数组保存到${filePath3}`);


    fs.readFile('D:/typescript/demo/accountbill/data3.json', 'utf8', (err, data) => {
        if (err) throw err;
        const json = JSON.parse(data);
        const jsonArray = [];
        json.forEach(function (item) {
            let temp = {
                '微站小计': item.微站小计,
                '成都移动存量': item.成都移动存量,
                '成都移动新增': item.成都移动新增,
                '天府移动存量': item.天府移动存量,
                '天府移动新增': item.天府移动新增,
                '电信存量': item.电信存量,
                '电信新增': item.电信新增,
                '联通存量': item.联通存量,
                '联通新增': item.联通新增,

            }
            jsonArray.push(temp);
        });

        let xls = json2xls(jsonArray);

        fs.writeFileSync('D:/typescript/demo/accountbill/excelbill.xlsx', xls, 'binary');
        console.log('文件已经保存成功🍓')
    })
    //铁塔账单文件
    let biggerFilePath1 = fileData.铁塔订单文件表
    let biggerFilePath2 = fileData.铁塔账单文件表
    let biggerFilePath3 = fileData.终止文件表
    async function loadExcel(pathname, sheetNames) {
        const dense_wb = xlsx2.read(fs.readFileSync(pathname))
        return (sheetNames ? sheetNames : dense_wb.SheetNames).reduce((pre, curr) => {
            if (!curr) return pre;
            const sheet = dense_wb.Sheets[curr];
            pre[curr] = xlsx2.utils.sheet_to_json(sheet, {
                raw: true
            });
            return pre;
        }, {});
    }
    // const date = new Date().valueOf();
    // const data1 = await loadExcel(biggerFilePath1).then(data => {
    //     const sheetsNames = Object.keys(data);
    //     sheetsNames.forEach(name => {
    //         console.log('🍎sheetsName', name, data[name].length, data['铁塔订单']);
    //     });
    //     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
    // })
    // console.log(data1, '🍑')
    // loadExcel(biggerFilePath2).then(data => {
    //     const sheetsNames = Object.keys(data);
    //     sheetsNames.forEach(name => {
    //         console.log('🍎sheetsName', name, data[name].length, data['铁塔订单']);
    //     });
    //     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
    // })
    // loadExcel(biggerFilePath3).then(data => {
    //     const sheetsNames = Object.keys(data);
    //     sheetsNames.forEach(name => {
    //         console.log('🍎sheetsName', name, data[name].length, data['终止订单表']);
    //     });
    //     console.log('🍎🍎🍎一共耗时', ((new Date().valueOf()) - date) / 1000)
    // })

    async function main() {
        const data1 = await loadExcel(biggerFilePath1)
        const data2 = await loadExcel(biggerFilePath2)
        const data3 = await loadExcel(biggerFilePath3)
        //铁塔订单文件处理
        let odTowersheet = data1['铁塔订单']
        let odtowertitle = odTowersheet
        let odTowerlist = []
        let odtowernum = 0
        // console.log('🍉', odTowersheet)
        odTowersheet.forEach((item, index) => {
            if (index == 0) {
                return
            }
            else if (item.已暂停出账 != '已暂停计费' && item.订单状态 == '已起租' && item.总费用 != '0.00') {
                odtowernum = odtowernum + 1
                odTowerlist.push({
                    [odtowertitle[0]]: item.订单状态,
                    [odtowertitle[1]]: item.订单号,
                    [odtowertitle[2]]: item.总费用,
                    [odtowertitle[3]]: item.已暂停出账,
                })
            }
        })
        console.log(odTowerlist)
        console.log('铁塔订单数目：' + odtowernum)
        //铁塔账单文件处理
        let towerSheet = data2['towerbill1']
        let towernum = 0
        towerSheet.forEach((item, index) => {
            towernum = towernum + 1
            if (index == 0) {
                return
            }
            else if (item.运营商 == '移动' && (item.运营商区县 == '天府新区' || item.运营商区县 == '双流县' || item.运营商区县 == '龙泉驿区')) {
                item.运营商 = '天府移动'
                // towerSheet.push({
                //     item
                // })
            }
        })
        console.log(towerSheet[4])
        console.log('铁塔账单数目：' + towernum)
        //终止文件处理
        let forbidenSheet = data3['终止订单表']
        let forbidenlist = []
        let forbidennum = 0
        forbidenSheet.forEach((item, index) => {
            if (index == 0) {
                return
            }
            else if (item.审批状态 == '运营商审批成功') {
                forbidennum = forbidennum + 1
                forbidenlist.push({
                    item
                })
            }
        })
        console.log('终止文件数目：' + forbidennum)
        // console.log(forbidenlist)


        // 从订单文件向账单传输进行对比😀😀😀
        let numtower1 = 0
        let numtower2 = 0
        for (let i = 0; i < odtowernum; i++) {
            let numtw4 = 0
            let numtw5 = 0
            for (let j = 0; j < towernum; j++) {
                if (odTowerlist[i].订单号 != towerSheet[j].需求确认单编号) {
                    numtw4 = numtw4 + 1
                }
                else if (odTowerlist[i].订单号 == towerSheet[j].需求确认单编号) {
                    //正常订单数目
                    numtower1 = numtower1 + 1
                    // console.log('正常订单号：'+odTowerlist[i].订单号)
                }
            }
            if (numtw4 == towernum) {
                // console.log('存在可能异常订单号：'+titlelist[i].订单号)

                for (let k = 0; k < forbidennum; k++) {
                    if (odTowerlist[i].订单号 == forbidenlist[k].订单编号) {
                        // console.log('终止文件存在正常订单号：' + titlelist[i].订单号)
                        numtower1 = numtower1 + 1
                        // console.log('正常订单号：'+odTowerlist[i].订单号)
                    }
                    else if (odTowerlist[i].订单号 != forbidenlist[k].订单编号) {
                        numtw5 = numtw5 + 1
                    }

                }
                if (numtw5 == forbidennum) {
                    // console.log('异常账号' + odtransmisslist[i].订单号 + '原因：在详单里面，但是不在账单里面')
                    numtower2 = numtower2 + 1
                    // console.log('異常賬號🍑：',odTowerlist[i].訂單號)
                }
            }
        }
        //从传输订单文件向订单文件传输
        for (let j1 = 0; j1 < towernum; j1++) {
            let numtw3 = 0
            for (let i1 = 0; i1 < odtowernum; i1++) {
                if (towerSheet[j1].需求确认单编号 != odTowerlist[i1].订单号) {
                    numtw3 = numtw3 + 1
                }
                else if (towerSheet[j1].需求确认单编号 == odTowerlist[i1].订单号) {
                    // numcsz = numcsz + 1
                }
            }
            if (numtw3 == odtowernum) {
                // console.log('异常订单' + transmisslists[j1].需求确认单编号 + '原因：出账，但是不在详单里面')
                numtower2 = numtower2 + 1
            }
        }
        console.log("正常订单数：（按照订单文件为基准）" + numtower1)
        console.log("异常订单数：（账单文件＋订单文件）" + numtower2)



        let yidongt = 0
        let tfyidongt = 0
        let liantongt = 0
        let dianxingt = 0

        let stocksf1t = 0
        let stocksf11t = 0
        let stocksf2t = 0
        let stocksf22t = 0
        let stocksf3t = 0
        let stocksf33t = 0
        let stocksf4t = 0
        let stocksf44t = 0
        let sum1t = 0
        let sum2t = 0
        let sum3t = 0
        let sum4t = 0
        let sum5t = 0
        let sum6t = 0
        let sum7t = 0
        let sum8t = 0
        let sum9t = 0
        let sum10t = 0
        let sum1bt = 0
        let sum2bt = 0
        let sum3bt = 0
        let sum4bt = 0
        let sum5bt = 0
        let sum6bt = 0
        let sum7bt = 0
        let sum8bt = 0
        let sum9bt = 0
        let sum10bt = 0
        let sum1ct = 0
        let sum2ct = 0
        let sum3ct = 0
        let sum4ct = 0
        let sum5ct = 0
        let sum6ct = 0
        let sum7ct = 0
        let sum8ct = 0
        let sum9ct = 0
        let sum10ct = 0
        let sum1dt = 0
        let sum2dt = 0
        let sum3dt = 0
        let sum4dt = 0
        let sum5dt = 0
        let sum6dt = 0
        let sum7dt = 0
        let sum8dt = 0
        let sum9dt = 0
        let sum10dt = 0
        let testt = 0
        let sum1et = 0
        let sum2et = 0
        let sum3et = 0
        let sum4et = 0
        let sum5et = 0
        let sum6et = 0
        let sum7et = 0
        let sum8et = 0
        //申明数组
        towerSheet.forEach((item, index) => {

            if (item.运营商 == '移动') {
                if (item.产品服务费与上月相比是否变化 == '存量') {
                    stocksf1t = stocksf1t + 1
                    sum1t = parseInt(sum1t + item.产品服务费合计1 + + item.油机发电服务费1 + + item.罚责赠费合计)
                    // sum1t = parseInt(item.产品服务费合计1 + item.产品服务费合计2 + item.罚责赠费合计 + item.油机发电服务费1 + item.油机发电服务费2 + sum1t)
                    sum3t = parseInt(sum3t + item.维护费折扣后金额1 + item.维护费折扣后金额2)//正常
                    sum5t = parseInt(sum5t + item.场地费折扣后金额1 + item.场地费折扣后金额2)//正常
                    sum7t = parseInt(sum7t + item.油机发电服务费1 + item.油机发电服务费2)
                    sum9t = parseInt(sum9t + item.产品服务费合计0 + item.产品服务费合计2)
                    sum1et = parseInt(sum1et + item.罚责赠费合计)
                }
                else if (item.产品服务费与上月相比是否变化 == '新增') {
                    stocksf11t = stocksf11t + 1
                    sum2t = parseInt(sum2t + item.产品服务费合计1 + + item.油机发电服务费1 + item.罚责赠费合计)
                    sum4t = parseInt(sum4t + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum6t = parseInt(sum6t + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum8t = parseInt(sum8t + item.油机发电服务费1 + item.油机发电服务费2)
                    sum10t = parseInt(sum10t + item.产品服务费合计0 + item.产品服务费合计2)
                    sum2et = parseInt(sum2et + item.罚责赠费合计)
                }
                yidongt = yidongt + 1
            }
            else if (item.运营商 == '天府移动') {
                if (item.产品服务费与上月相比是否变化 == '存量') {
                    stocksf2t = stocksf2t + 1
                    sum1bt = parseInt(sum1bt + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    // console.log(item.罚责赠费合计)
                    sum3bt = parseInt(sum3bt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum5bt = parseInt(sum5bt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum7bt = parseInt(sum7bt + item.油机发电服务费1 + item.油机发电服务费2)
                    sum9bt = parseInt(sum9bt + item.产品服务费合计0 + item.产品服务费合计2)
                    sum3et = parseInt(sum3et + item.罚责赠费合计)
                }
                else if (item.产品服务费与上月相比是否变化 == '新增') {
                    stocksf22t = stocksf22t + 1
                    sum2bt = parseInt(sum2bt + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    sum4bt = parseInt(sum4bt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum6bt = parseInt(sum6bt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum8bt = parseInt(sum8bt + item.油机发电服务费1 + item.油机发电服务费2)
                    sum10bt = parseInt(sum10bt + item.产品服务费合计0 + item.产品服务费合计2)
                    sum4et = parseInt(sum4et + item.罚责赠费合计)
                }
                tfyidongt = tfyidongt + 1
            }
            else if (item.运营商 == '联通') {
                if (item.产品服务费与上月相比是否变化 == '存量') {
                    stocksf3t = stocksf3t + 1
                    sum1ct = parseInt(sum1ct + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    sum3ct = parseInt(sum3ct + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum5ct = parseInt(sum5ct + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum7ct = parseInt(sum7ct + item.油机发电服务费1 + item.油机发电服务费2)
                    sum9ct = parseInt(sum9ct + item.产品服务费合计0 + item.产品服务费合计2)
                    sum5et = parseInt(sum5et + item.罚责赠费合计)
                }
                else if (item.产品服务费与上月相比是否变化 == '新增') {
                    stocksf33t = stocksf33t + 1
                    sum2ct = parseInt(sum2ct + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    sum4ct = parseInt(sum4ct + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum6ct = parseInt(sum6ct + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum8ct = parseInt(sum8ct + item.油机发电服务费1 + item.油机发电服务费2)
                    sum10ct = parseInt(sum10ct + item.产品服务费合计0 + item.产品服务费合计2)
                    sum6et = parseInt(sum6et + item.罚责赠费合计)
                }
                liantongt = liantongt + 1
            }
            else if (item.运营商 == '电信') {
                if (item.产品服务费与上月相比是否变化 == '存量') {
                    stocksf4t = stocksf4t + 1
                    sum1dt = parseInt(sum1dt + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    sum3dt = parseInt(sum3dt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum5dt = parseInt(sum5dt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum7dt = parseInt(sum7dt + item.油机发电服务费1 + item.油机发电服务费2)
                    sum9dt = parseInt(sum9dt + item.产品服务费合计0 + item.产品服务费合计2)
                    sum7et = parseInt(sum7et + item.罚责赠费合计)
                }
                else if (item.产品服务费与上月相比是否变化 == '新增') {
                    stocksf44t = stocksf44t + 1
                    sum2dt = parseInt(sum2dt + item.产品服务费合计1 + item.油机发电服务费1 + item.罚责赠费合计)
                    sum4dt = parseInt(sum4dt + item.维护费折扣后金额1 + item.维护费折扣后金额2)
                    sum6dt = parseInt(sum6dt + item.场地费折扣后金额1 + item.场地费折扣后金额2)
                    sum8dt = parseInt(sum8dt + item.油机发电服务费1 + item.油机发电服务费2)
                    sum10dt = parseInt(sum10dt + item.产品服务费合计0 + item.产品服务费合计2)
                    sum8et = parseInt(sum8et + item.罚责赠费合计)
                }
                dianxingt = dianxingt + 1
            }
        })
        console.log(yidongt + '  ' + stocksf1t + ' ' + sum1t + '  ' + (testt - sum3t - sum5t - sum1et) + ' ' + sum3t + '  ' + sum5t + '  ' + sum7t + ' ' + sum1et + ' ' + sum9t)
        console.log(yidongt + '  ' + stocksf11t + '  ' + sum2t + ' ' + (sum2t - sum4t - sum6t - sum2et) + '  ' + sum4t + '  ' + sum6t + '  ' + sum8t + '  ' + sum2et + ' ' + sum10t)
        console.log(tfyidongt + '  ' + stocksf2t + '  ' + sum1bt + '  ' + (sum1bt - sum3bt - sum5bt - sum3et) + ' ' + sum3bt + '  ' + sum5bt + '  ' + sum7bt + '  ' + sum3et + ' ' + sum9bt)
        console.log(tfyidongt + '  ' + stocksf22t + '  ' + sum2bt + '  ' + (sum2bt - sum4bt - sum6bt - sum4et) + '  ' + sum4bt + '  ' + sum6bt + '  ' + sum8bt + '  ' + sum4et + ' ' + sum10bt)
        console.log(liantongt + '  ' + stocksf3t + '  ' + sum1ct + '  ' + (sum1ct - sum3ct - sum5ct - sum5et) + ' ' + sum3ct + '  ' + sum5ct + '  ' + sum7ct + '  ' + sum5et + ' ' + sum9ct)
        console.log(liantongt + '  ' + stocksf33t + '   ' + sum2ct + '  ' + (sum2ct - sum4ct - sum6ct - sum6et) + '  ' + sum4ct + '  ' + sum6ct + '  ' + sum8ct + '  ' + sum6et + ' ' + sum10ct)
        console.log(dianxingt + '   ' + stocksf4t + '  ' + sum1dt + '  ' + (sum1dt - sum3dt - sum5dt - sum7et) + ' ' + sum3dt + '  ' + sum5dt + '  ' + sum7dt + '  ' + sum7et + ' ' + sum9dt)
        console.log(dianxingt + ' ' + stocksf44t + '  ' + sum2dt + '  ' + (sum2dt - sum4dt - sum6dt - sum8et) + '  ' + sum4dt + '  ' + sum6dt + '  ' + sum8dt + '  ' + sum8et + ' ' + sum10dt)

        let detailbillsheets4 = xlsxData.详单文件表[0].data
        let detailtitle4 = detailbillsheets4[0]
        let detaillists4 = []
        console.log(detailtitle4[0])
        detailbillsheets4.forEach((item, index) => {
            console.log(index)
            if (index == 0) {
                detaillists4.push({
                    [detailtitle4[0]]: (yidongt + tfyidongt + liantongt + dianxingt),
                    [detailtitle4[1]]: stocksf1t,
                    [detailtitle4[2]]: stocksf11t,
                    [detailtitle4[3]]: stocksf2t,
                    [detailtitle4[4]]: stocksf22t,
                    [detailtitle4[5]]: stocksf3t,
                    [detailtitle4[6]]: stocksf33t,
                    [detailtitle4[7]]: stocksf4t,
                    [detailtitle4[8]]: stocksf44t
                })
            }
            else if (index == 1) {
                detaillists4.push({
                    [detailtitle4[0]]: (sum1t + sum2t + sum1bt + sum2bt + sum1ct + sum2ct + sum1dt + sum2dt),
                    [detailtitle4[1]]: sum1t,
                    [detailtitle4[2]]: sum2t,
                    [detailtitle4[3]]: sum1bt,
                    [detailtitle4[4]]: sum2bt,
                    [detailtitle4[5]]: sum1ct,
                    [detailtitle4[6]]: sum2ct,
                    [detailtitle4[7]]: sum1bt,
                    [detailtitle4[8]]: sum2dt,
                })
            }
            else if (index == 2) {
                detaillists4.push({
                    [detailtitle4[0]]: ((sum1t - sum3t - sum5t + sum1et) + (sum2t - sum4t - sum6t - +sum2et) + (sum1bt - sum3bt - sum5bt + sum3et) + (sum2bt - sum4bt - sum6bt + sum4et) + (sum1ct - sum3ct - sum5ct + sum5et) + (sum2ct - sum4ct - sum6ct + sum6et) + (sum1dt - sum3dt - sum5dt + sum7et) + (sum2dt - sum4dt - sum6dt + sum8et)),
                    [detailtitle4[1]]: (sum1t - sum3t - sum5t + sum1et),
                    [detailtitle4[2]]: (sum2t - sum4t - sum6t + sum2et),
                    [detailtitle4[3]]: (sum1bt - sum3bt - sum5bt + sum3et),
                    [detailtitle4[4]]: (sum2bt - sum4bt - sum6bt + sum4et),
                    [detailtitle4[5]]: (sum1ct - sum3ct - sum5ct + sum5et),
                    [detailtitle4[6]]: (sum2ct - sum4ct - sum6ct + sum6et),
                    [detailtitle4[7]]: (sum1dt - sum3dt - sum5dt + sum7et),
                    [detailtitle4[8]]: (sum2dt - sum4dt - sum6dt + sum8et),
                })
            }

            else if (index == 3) {
                detaillists4.push({
                    [detailtitle4[0]]: sum3t + sum4t + sum3bt + sum4bt + sum3ct + sum4ct + sum3dt + sum4dt,
                    [detailtitle4[1]]: sum3t,
                    [detailtitle4[2]]: sum4t,
                    [detailtitle4[3]]: sum3bt,
                    [detailtitle4[4]]: sum4bt,
                    [detailtitle4[5]]: sum3ct,
                    [detailtitle4[6]]: sum4ct,
                    [detailtitle4[7]]: sum3dt,
                    [detailtitle4[8]]: sum4dt
                })
            }
            else if (index == 4) {
                detaillists4.push({
                    [detailtitle4[0]]: sum5t + sum6t + sum5bt + sum6bt + sum5ct + sum6ct + sum5dt + sum6dt,
                    [detailtitle4[1]]: sum5t,
                    [detailtitle4[2]]: sum6t,
                    [detailtitle4[3]]: sum5bt,
                    [detailtitle4[4]]: sum6bt,
                    [detailtitle4[5]]: sum5ct,
                    [detailtitle4[6]]: sum6ct,
                    [detailtitle4[7]]: sum5dt,
                    [detailtitle4[8]]: sum6dt
                })

            }
            else if (index == 5) {
                detaillists4.push({
                    [detailtitle4[0]]: sum7t + sum8t + sum7bt + sum8bt + sum7ct + sum8ct + sum7dt + sum8dt,
                    [detailtitle4[1]]: sum7t,
                    [detailtitle4[2]]: sum8t,
                    [detailtitle4[3]]: sum7bt,
                    [detailtitle4[4]]: sum8bt,
                    [detailtitle4[5]]: sum7ct,
                    [detailtitle4[6]]: sum8ct,
                    [detailtitle4[7]]: sum7dt,
                    [detailtitle4[8]]: sum8dt
                })
            }
            else if (index == 6) {
                detaillists4.push({
                    [detailtitle4[0]]: sum1et + sum2et + sum3et + sum4et + sum5et + sum6et + sum7et + sum8et,
                    [detailtitle4[1]]: sum1et,
                    [detailtitle4[2]]: sum2et,
                    [detailtitle4[3]]: sum3et,
                    [detailtitle4[4]]: sum4et,
                    [detailtitle4[5]]: sum5et,
                    [detailtitle4[6]]: sum6et,
                    [detailtitle4[7]]: sum7et,
                    [detailtitle4[8]]: sum8et
                })
            }
            else if (index == 7) {
                detaillists4.push({
                    [detailtitle4[0]]: sum9t + sum10t + sum9bt + sum10bt + sum9ct + sum10ct + sum9dt + sum10dt,
                    [detailtitle4[1]]: sum9t,
                    [detailtitle4[2]]: sum10t,
                    [detailtitle4[3]]: sum9bt,
                    [detailtitle4[4]]: sum10bt,
                    [detailtitle4[5]]: sum9ct,
                    [detailtitle4[6]]: sum10ct,
                    [detailtitle4[7]]: sum9dt,
                    [detailtitle4[8]]: sum10dt
                })
            }
        })
        console.log(detaillists4)
        const Jsondata4 = JSON.stringify(detaillists4)
        const filePath4 = 'D:/typescript/demo/accountbill/data4.json';
        fs.writeFileSync(filePath4, Jsondata4);
        console.log(`已将对象数组保存到${filePath4}`);


        fs.readFile('D:/typescript/demo/accountbill/data4.json', 'utf8', (err, data) => {
            if (err) throw err;
            const json = JSON.parse(data);
            const jsonArray = [];
            json.forEach(function (item) {
                let temp = {
                    '塔类小计': item.塔类小计,
                    '成都移动存量': item.成都移动存量,
                    '成都移动新增': item.成都移动新增,
                    '天府移动存量': item.天府移动存量,
                    '天府移动新增': item.天府移动新增,
                    '联通存量': item.联通存量,
                    '联通新增': item.联通新增,
                    '电信存量': item.电信存量,
                    '电信新增': item.电信新增,
                }
                jsonArray.push(temp);
            });

            let xls = json2xls(jsonArray);

            fs.writeFileSync('D:/typescript/demo/accountbill/towertest2.xlsx', xls, 'binary');
            console.log('文件已经保存成功❤')
        })
    }

    main()

}



