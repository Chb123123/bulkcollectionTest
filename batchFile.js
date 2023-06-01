let requests = require('requests') // 请求包
let fs = require('fs') // 读写文件
const request = require('request')
var xlsx = require('node-xlsx');
const ExcelJS = require('exceljs');
const axios = require('axios');
const https = require('https');

// 页号
let num = 1
// excel 文件的页号
let fileNum = 1
// 文件夹名称
let fileNameNum = 0
// 要写入的数据列表
function main() {
  let errorUrl = `https://www.vipstation.com.hk/jp/bags/ysl-saint-laurent?page${num}`
  requests(`https://www.vipstation.com.hk/jp/bags/ysl-saint-laurent?page${num}`, { encoding: 'utf8' }) // 请求路径
    .on('data', async function (chunk) {
      console.log(`当前为第${num}个页面`)
      let excelData = []
      let imgUrlList = []
      let viewList = []
      let arr = chunk.split(' var itemList =')
      let list = arr[1].split('var seriesList')[0].split('{')
      // 开始爬取页面
      for (let i = 0; i < list.length; i++) {
        let item = list[i].split(',')
        for (let j = 0; j < item.length; j++) {
          if (item[j].indexOf('ST_WEB_NAME') !== -1) {
            let urlItem = item[j].split(':')[1].replace(/"/g, '')
            viewList.push(urlItem)
          }
        }
      }
      for (let i = 0; i <= viewList.length; i++) {
        await sleep(1500)
        if (!viewList[i]) {
          break
        }
        requests(`https://www.vipstation.com.hk/jp/item/${viewList[i]}.html`).on('data', async function (data) {
          let arr1 = []
          let imgUrl = data.split('var imgList')[1].split('var videoList')[0].split('"')
          // // 图片名称
          for (let item of imgUrl) {
            if (item.indexOf('https') !== -1) {
              arr1.push(item)
            }
          }
          imgUrlList.push(arr1)
          const list = data.split(' var iteminfo')[1].split('var price')[0]
          // 截取图片数据
          let addData = ["", "", "", "", "", "", "", "", "", "", "", ""]
          // console.log(list.split('"ST_CODE":')[1].split(',')[0].replace(/"/g, ''))
          addData[0] = data.split('<title>')[1].split('</title>')[0].replace(/"/g, '')
          addData[1] = ''
          addData[2] = ''
          addData[3] = fileNum
          addData[4] = ''
          addData[5] = list.split('"ST_CODE":')[1].split(',')[0].replace(/"/g, '')
          // 商品描述
          let about = data.split('script type="application/ld+json"')[1].split(' </script>')[0].split('"description":')[1].split('"brand":')[0].split(',')
          // console.log(about)
          let aboutStr = ''
          for (let i = 0; i < about.length - 1; i++) {
            aboutStr = `${aboutStr}\n${about[i].replace(/"/g, '')}`
          }
          aboutStr = aboutStr.replace('\n', '')
          addData[6] = aboutStr
          addData[7] = ''
          addData[8] = ''
          addData[9] = ''
          addData[10] = ''
          addData[11] = ''
          console.log(`第${fileNum}条数据获取完成，开始下一条数据`)
          excelData.push(addData)
          fileNum++
        })
      }
      // 判断文件是否存在
      const fileFlag = checkFileExists('./data.xlsx')
      if (fileFlag) {
        // 文件存在，追加文件数据
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('data.xlsx');
        // 获取要追加数据的工作表
        const worksheet = workbook.getWorksheet('Sheet 1');
        console.log(worksheet.length)
        // 添加新的数据行
        for (let item of excelData) {
          worksheet.addRow(item)
        }
        // 保存工作簿，覆盖原始Excel文件
        await workbook.xlsx.writeFile('data.xlsx');
        console.log('数据已成功追加到Excel文件！');
      } else {
        // 文件不存在，新增一个文件
        const workbook = new ExcelJS.Workbook();
        // 添加一个新的工作表
        const worksheet = workbook.addWorksheet('Sheet 1');
        // 添加表头
        worksheet.addRow(["英文标题", "日文标题", "主图图片", "图片文件夹序号", '价格', '型号', "商品描述", "颜色", "尺码", "详情描述", "日元", "原价格（美金）"
        ]);
        for (let item of excelData) {
          worksheet.addRow(item)
        }
        // 保存工作簿为Excel文件
        await workbook.xlsx.writeFile('data.xlsx');

        console.log('数据已成功添加到Excel文件！');
      }
      // 循环结束
      excelData = []
      // 爬取数据完成
      console.log('当前页面爬取数据完成')
      console.log('开始爬取图片')
      for (let i = 0; i <= imgUrlList.length; i++) {
        await sleep(1000)
        if (!imgUrlList[i]) {
          break
        }
        fileNameNum++
        const folderPath = `${fileNameNum}`;
        if (!fs.existsSync(folderPath)) {
          // 创建新的文件夹
          fs.mkdirSync(folderPath);
          console.log(`成功创建文件夹 ${folderPath}`);
        } else {
          console.log(`文件夹 ${folderPath} 已存在`);
        }
        for (let j = 1; j <= imgUrlList[i].length; j++) {
          await sleep(1500)
          if (!imgUrlList[i][j]) {
            break
          }
          if (imgUrlList[i][j].indexOf('upload') !== -1) {
            let file = imgUrlList[i][j].split('.')
            let pathfile = `./${fileNameNum}/${+new Date()}.${file[file.length - 1]}`
            downloadImage(imgUrlList[i][j], pathfile).then(() => {
              console.log(`图片: ${imgUrlList[i][j]}下载完成`);
            })
              .catch((error) => {
                console.error(`图片下载失败：${error.message}`);
                console.log(`错误页面为:${errorUrl}`)
                console.log(`错误序列为: ${i}`)
                console.log(`文件夹序号为: ${fileNameNum}`)
              });
          }
        }
      }
      if (viewList.length < 20) {
        console.log('页面爬取完成，停止脚本')
        return
      } else {
        console.log(`页号：${num} 爬取完成，开始下一个页面`)
        num++
        main()
      }
    })
}
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms))
}

// 判断文件是否存在
function checkFileExists(filePath) {
  try {
    // 使用 fs.accessSync 方法检查文件是否存在
    fs.accessSync(filePath, fs.constants.F_OK);
    console.log(`文件 ${filePath} 存在`);
    return true;
  } catch (err) {
    console.error(`文件 ${filePath} 不存在`);
    console.log(`创建${filePath}文件`)
    return false;
  }
}

// 执行代码
main()


async function downloadImage(url, filePath) {
  try {
    // 创建 axios 实例，并设置 rejectUnauthorized 为 false
    const instance = axios.create({
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    });
    // 发送 HTTP GET 请求获取图片数据
    const response = await instance.get(url, { responseType: 'stream' });

    // 创建可写流，将图片数据写入文件
    const writer = fs.createWriteStream(filePath);
    response.data.pipe(writer);

    // 返回 Promise 对象，等待图片下载完成
    return new Promise((resolve, reject) => {
      writer.on('finish', resolve);
      writer.on('error', reject);
    });
  } catch (error) {
    throw new Error(`图片下载失败：${error.message}`);
  }
}
