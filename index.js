const puppeteer = require('puppeteer');
const request = require('request');
const cheerio = require('cheerio');
const write = require('write');
const lodash = require('lodash');
const Excel = require('exceljs');

var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('My Sheet');

worksheet.columns = [
    { header: '投诉ID', key: 'id', width: 10 },
    { header: '投诉标题', key: 'title', width: 10 },
    { header: '投诉时间', key: 'ctimeStr', width: 10 },
    { header: '投诉对象', key: 'merchantname', width: 32 },
    { header: '问题类型', key: 'problemLabelListName', width: 32 },
    { header: '诉求类型', key: 'shuqiu', width: 32 },
    { header: '投诉详情', key: 'topic', width: 32 },
];

const sleep = (sec) => new Promise((resolve) => {
    setTimeout(() => {
        resolve();
    }, sec * 1000)
})

var customHeaderRequest = request.defaults({
    headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36' }
})

const reURL = 'http://ts.21cn.com/json/indexPcMorePost/order/ctime/pageNo/';
const detailPageURL = 'http://ts.21cn.com/tousu/show/id/';
const detailDataURL = 'http://ts.21cn.com/json/getPostContent/postKey/';

let pageNo = 1;

// 首页的数据可仅作为最新一条的id

(async () => {

    const browser = await puppeteer.launch({        //启动浏览器
        headless: false
    });

    const times = Array.from({ length: 3 })
    for (const time of times) {

        customHeaderRequest(reURL + pageNo, { json: true }, async (err, res, body) => {
            if (err) { return console.log(err); }
            const htmlJson = body.message;
            write.sync('./html/pageNo1.html', htmlJson, { newline: true });

            const $ = cheerio.load(htmlJson);
            const tagsEl = $('span[class=sharetag]');
            const tagArr = tagsEl.map(function (i, el) {
                return $(el).attr('tag');
            }).get();

            for (const tag of tagArr) {
                const page = await browser.newPage();       //开启浏览器新窗口
                await page.setViewport({            //配置窗口信息，具体配置的移步官方文档
                    width: 1920,
                    height: 1080
                });

                await page.goto(detailPageURL + tag);         //当前窗口加载固定 url 地址页。url 需要以 https 开头
                // const html = await page.content();   //这是返回出来的html代码

                await sleep(5);

                customHeaderRequest({
                    uri: detailPageURL + tag,
                    headers: {
                        'accept-language': 'es-ES,es;q=0.9,ru;q=0.8',
                        'accept-encoding': 'br',
                        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
                    }
                }, (err, res, body) => {

                    write.sync(`./html/${tag}.html`, body, { newline: true });

                    const $ = cheerio.load(body);
                    const postKeyValue = $('input#postKeyValue').attr('value');

                    customHeaderRequest(detailDataURL + postKeyValue, { json: true }, (err, res, body) => {
                        console.log('%c 🥝 body: ', 'font-size:20px;background-color: #465975;color:#fff;', lodash.get(body, 'post'));
                        const postData = lodash.get(body, 'post');
                        if (!postData) {
                            console.log('%c 🥤 postData: ', 'font-size:20px;background-color: #EA7E5C;color:#fff;', tag, postKeyValue, postData);
                            return;
                        }
                        const {
                            id,
                            title,
                            ctimeStr,
                            merchantname,
                            problemLabelList,
                            shuqiu,
                            topic,
                        } = postData;
                        worksheet.addRow({
                            id,
                            title,
                            ctimeStr,
                            merchantname,
                            problemLabelListName: problemLabelList.map((label) => label.name).join(','),
                            shuqiu,
                            topic,
                        });
                        workbook.xlsx.writeFile(`./generated/${new Date(ctimeStr).toLocaleDateString()}.xlsx`)
                            .then(() => {
                                console.log('csv ok');
                            });
                        if (tag === tagArr[tagArr.length - 1]) {
                            browser.close();      //关闭浏览器，对象实例销毁
                            console.log('everything is ok');
                        }
                    })
                })
            }

        });

        pageNo += 1;
    }

})();
