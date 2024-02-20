
'use strict'

import installExtension, { VUEJS3_DEVTOOLS } from 'electron-devtools-installer'
import xlsx from 'node-xlsx'
const electron = require('electron');
const fs = require('fs');
const path = require('path');
const json2xls = require('json2xls')


async function transmissHandle(xlsxData) {
    //默认0和1是订单总文件和终止文件
    // for (let i = 0; i < fileData.length; i++) {
    const odtranmissorsheets = xlsxData.总订单文件表[1].data
    // console.log('🍑 ' + odtranmissorsheets)

    let odtransmisslist = []
    let odtransnum = 0

    // 传输订单文件😀😀😀
    // odtranmissorsheets = excelContent2[2].data
    // console.log(odtranmissorsheets)
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
    // console.log(odtransmisslist)
    console.log('🍌传输订单文件数（已筛选）：' + odtransnum)
    // const transmiss2 = xlsx.parse("D:/typescript/demo/accountbill/transmission.xlsx", {
    //   cellDates: true,
    // });
    //终止文件😀😀😀
    let forbidnum = 0
    let forbidensheets = xlsxData.终止文件表[0].data
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
    console.log('🍑🍑🍑')
    let allbill = sumt1 + sumt3 + sumt5 + sumt7

    return {
        forbidnum, forbidlist, yidong1, tfyidong1, liantong1, dianxing1, stocksf1a, stocksf2a, stocksf3a, stocksf4a, sumt1,
        sumt2, sumt3, sumt4, sumt5, sumt6, sumt7, sumt8, transnum
    };
}

module.export = {
    transmissHandle
}

