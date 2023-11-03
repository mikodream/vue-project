<script setup lang="ts">
import { reactive, ref, onBeforeMount } from 'vue'
import ExcelJS from 'exceljs'
import { Buffer } from 'buffer'

import md5 from 'js-md5'

let currentDir = ''
// @ts-ignore
const sep = Niva.api.os.sep().then((s) => {
  console.log(sep)
  // @ts-ignore
  Niva.api.process.currentExe().then((dir) => {
    currentDir = getLeftOfFixedTextLast(dir, s)
    console.log(dir, currentDir)
  })
})

//是否需要校验
const verify = ref(true)
//卡号
const cardId = ref('')

const columns = [
  { prop: 'name', label: '商品名', width: '100' },
  { prop: 'href', label: '链接' },
  { prop: 'price', label: '价格', width: '100' },
  { prop: 'sales', label: '销量', width: '100' }
]
const tableData = reactive([{}])
let dataArrays: DataObj[] = []
var clickbutton = function () {
  // @ts-ignore
  Niva.api.clipboard.read().then((data: string) => {
    dataArrays.length = 0
    const separator = 'data-analytics-view-custom-deliverylabel="'
    //商品名字
    const shopNameSeparatorHeader = 'aria-label="'
    const shopNameSeparatorFooter = '"><div><div'
    //商品地址
    const shopHrefSeparatorHeader = ' " href="'
    const shopHrefSeparatorFooter = '" title='
    //商品销量
    const shopSalesSeparatorHeader = '<span class="msa3_z4 mgn2_12">'
    const shopSalesSeparatorFooter = 'osob'

    const html = data

    const sepaAarrays = html.split(separator)
    for (const tempText of sepaAarrays) {
      const nameAndPrice = getCenterOfFixedText(
        tempText,
        shopNameSeparatorHeader,
        shopNameSeparatorFooter
      )
      const name = getLeftOfFixedText(nameAndPrice, ', ')
      const price = getRightOfFixedText(nameAndPrice, ', ')
      const href = getCenterOfFixedText(tempText, shopHrefSeparatorHeader, shopHrefSeparatorFooter)
      const sales = getCenterOfFixedText(
        tempText,
        shopSalesSeparatorHeader,
        shopSalesSeparatorFooter
      )
      if (name !== '') {
        //封装对象
        let dataTemp: DataObj = { name: name, price: price, href: href, sales: sales }
        dataArrays.push(dataTemp)
      }
    }

    tableData.length = 0
    tableData.push(...dataArrays)
    if (tableData.length === 0) {
      // @ts-ignore
      Niva.api.dialog.showMessage('解析异常', '解析失败，解析到的结果为 0 个')
    } else {
      exportExcel().then((data) => {
        // 创建一个字节数组

        const ss = Buffer.from(data)
        const base64String = ss.toString('base64')

        const writeDri = currentDir + '/' + getDateString() + '.xlsx'
        // @ts-ignore
        Niva.api.fs.write(writeDri, base64String, 'base64')

        // // @ts-ignore
        // Niva.api.dialog.pickDir().then((dir: Promise<string>) => {

        // })
      })
      //
    }
  })
}

function getDateString() {
  // 创建一个表示当前日期和时间的Date对象
  const currentDate = new Date()

  // 获取年、月、日、时、分、秒
  const year = currentDate.getFullYear().toString().substr(-2) // 提取年份的最后两位
  const month = (currentDate.getMonth() + 1).toString().padStart(2, '0') // 月份从0开始，需要加1，然后补齐两位数
  const day = currentDate.getDate().toString().padStart(2, '0') // 补齐两位数
  const hours = currentDate.getHours().toString().padStart(2, '0') // 补齐两位数
  const minutes = currentDate.getMinutes().toString().padStart(2, '0') // 补齐两位数
  const seconds = currentDate.getSeconds().toString().padStart(2, '0') // 补齐两位数

  // 创建日期时间字符串
  const dateTimeString = `${year}${month}${day}${hours}${minutes}${seconds}`
  return dateTimeString
}

async function exportExcel() {
  // 创建一个工作簿
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('商品数据')

  // 添加表头
  const headerRow = worksheet.addRow(['名称', '价格', '链接', '销售'])
  headerRow.font = { bold: true }

  // 添加数据行
  dataArrays.forEach((item) => {
    worksheet.addRow([item.name, item.price, item.href, item.sales])
  })

  // 生成Excel文件
  const buffer = await workbook.xlsx.writeBuffer()
  return buffer
}

function getCenterOfFixedText(inputText: string, fixedTextHeader: string, fixedTextFooter: string) {
  const headerIndex = inputText.indexOf(fixedTextHeader)
  if (headerIndex !== -1) {
    const footerIndex = inputText.indexOf(fixedTextFooter, headerIndex)
    if (footerIndex !== -1) {
      const extractedText = inputText.substring(headerIndex + fixedTextHeader.length, footerIndex)
      return extractedText
    }
  }
  return ''
}

function getLeftOfFixedText(inputText: string, fixedText: string) {
  const index = inputText.indexOf(fixedText)
  if (index !== -1) {
    return inputText.substring(0, index)
  }
  return '' // 如果没有找到固定文本，可以根据需求返回 null 或其他值
}

function getLeftOfFixedTextLast(inputText: string, fixedText: string) {
  const index = inputText.lastIndexOf(fixedText)
  if (index !== -1) {
    return inputText.substring(0, index)
  }
  return '' // 如果没有找到固定文本，可以根据需求返回 null 或其他值
}

function getRightOfFixedText(inputText: string, fixedText: string) {
  const index = inputText.indexOf(fixedText)
  if (index !== -1) {
    return inputText.substring(index + fixedText.length)
  }
  return '' // 如果没有找到固定文本，可以根据需求返回 null 或其他值
}
interface DataObj {
  name: string
  href: string
  price: string
  sales: string
}

onBeforeMount(() => {
  // @ts-ignore
  Niva.api.fs.exists(currentDir+'/key.json').then(async (result) => {
    if (result) {
      verifyFile()
    }
  })
})

const onSubmit = () => {
  let machineId = generateUUID()
  let json = JSON.stringify({ machineId: machineId, cardId: cardId.value })
  console.log(json)
  // @ts-ignore
  Niva.api.fs.write(currentDir+'/key.json', json).then(() => {
    verifyFile()
  })
}

const onCancel = () => {
  // @ts-ignore
  Niva.api.process.exit()
}

function generateUUID() {
  return 'xxxxxxxxxxxx4xxxyxxxxxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0,
      v = c == 'x' ? r : (r & 0x3) | 0x8
    return v.toString(16)
  })
}

async function verifyFile() {
  // @ts-ignore
  let json = await Niva.api.fs.read(currentDir+'/key.json')
  let jsonObject = JSON.parse(json)

  let response = await verifyCard(jsonObject.machineId, jsonObject.cardId)
  console.log(response.body)
  let result = JSON.parse(response.body)
  if (result.code === 0) {
    verify.value = false
  } else {
    // @ts-ignore
    Niva.api.dialog.showMessage('程序校验失败,请联系供应商获取卡号', result.data.error_msg, 'error')
  }
}

async function verifyCard(card: string, machine_code: string) {
  let http_method = 'POST'
  let host = 'api.ssdun.cn:8520'
  let path = '/V4/card/verify'
  let user_key = '529a53ab7fc09d77a7eb9875a1d1f3c3d69be63b93985738e2941a9f1a2f8ec1'
  let app_secret = '4262'
  let timestamp = Date.now() / 1000
  let params =
    'user_key=' +
    user_key +
    '&app_secret=' +
    app_secret +
    '&card=' +
    card +
    '&machine_code=' +
    machine_code +
    '&timestamp=' +
    timestamp
  let sign = md5.md5(http_method + host + path + params + app_secret).toUpperCase()
  let body = {
    user_key: user_key,
    app_secret: app_secret,
    card: card,
    machine_code: machine_code,
    timestamp: timestamp,
    sign: sign
  }

  console.log(sign)
  // @ts-ignore
  return await Niva.api.http.request({
    method: http_method,
    url: 'http://' + host + path,
    body: body
  })
}

function niva() {
  // @ts-ignore
  return Niva
}
</script>

<template>
  <div v-if="verify" style="background-color: whitesmoke">
    <el-card>
      <el-form>
        <el-form-item label="卡号">
          <el-input v-model="cardId"></el-input>
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="onSubmit">立即创建</el-button>
          <el-button @click="onCancel">取消</el-button>
        </el-form-item>
      </el-form>
    </el-card>
  </div>
  <div v-else>
    <div>
      <el-button @click="clickbutton">复制网页源码后点此按钮</el-button>
    </div>

    <div>
      <el-table :data="tableData">
        <el-table-column
          v-for="column in columns"
          :prop="column.prop"
          :label="column.label"
          :width="column.width"
        />
      </el-table>
    </div>
  </div>
</template>

<style scoped>
header {
  line-height: 1.5;
  max-height: 100vh;
}

.logo {
  display: block;
  margin: 0 auto 2rem;
}

nav {
  width: 100%;
  font-size: 12px;
  text-align: center;
  margin-top: 2rem;
}

nav a.router-link-exact-active {
  color: var(--color-text);
}

nav a.router-link-exact-active:hover {
  background-color: transparent;
}

nav a {
  display: inline-block;
  padding: 0 1rem;
  border-left: 1px solid var(--color-border);
}

nav a:first-of-type {
  border: 0;
}

@media (min-width: 1024px) {
  header {
    display: flex;
    place-items: center;
    padding-right: calc(var(--section-gap) / 2);
  }

  .logo {
    margin: 0 2rem 0 0;
  }

  header .wrapper {
    display: flex;
    place-items: flex-start;
    flex-wrap: wrap;
  }

  nav {
    text-align: left;
    margin-left: -1rem;
    font-size: 1rem;

    padding: 1rem 0;
    margin-top: 1rem;
  }
}
</style>
