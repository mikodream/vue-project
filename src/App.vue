<script setup lang="ts">
import { reactive, ref } from 'vue'
import ExcelJS from 'exceljs'
import { collapseTextChangeRangesAcrossMultipleVersions } from 'typescript'
import { Buffer } from 'buffer'

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
    const shopSeparatorHeader = 'title="">'
    const shopSeparatorFooter = '</a></h2>'
    //商品类
    const shopTypeSeparatorHeader = 'dostawa za 17 – 24 dni" '
    const shopTypeSeparatorFooter = 'data-variants-visible="true"'
    //商品名字
    const shopNameSeparatorHeader = 'dostawa za 19 – 21 dni" aria-label="'
    const shopNameSeparatorFooter = '"><div><div'
    //商品价格
    const shopPriceSeparatorHeader = 'mgn2_30_s" aria-label="'
    const shopPriceSeparatorFooter = ' z? aktualna cena'
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
        // const byteArray = Array.from(new Uint8Array(data))
        // 使用btoa()将字节数组转换为base64编码的字符串
        // const base64String = window.btoa(String.fromCharCode.apply(null, byteArray))
        // @ts-ignore
        Niva.api.dialog.pickDir().then((dir: Promise<string>) => {
          const writeDri = dir + '/' + getDateString() + '.xlsx'
          console.log('写出的 base64 数据', base64String)
          // @ts-ignore
          Niva.api.fs.write(writeDri, base64String, 'base64')
        })
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
</script>

<template>
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
