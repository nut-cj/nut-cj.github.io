---
title: 计算机二级(office)之vlookup
date: 2019-12-05 13:49:03
categories: 学习
tags: 
- 计算机二级
- Excel
- vlookup
---

## vlookup是什么？

VLOOKUP函数是Excel中的一个纵向查找函数，它与LOOKUP函数和HLOOKUP函数属于一类函数。

## vlookup能做什么?

vlookup 在工作中都有广泛应用，例如可以用来核对数据，多个表格之间快速导入数据等函数功能。

它功能是按列查找，最终返回该列所需查询序列所对应的值；与之对应的HLOOKUP是按行查找的。

## vlookup怎么用？



```
=VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
```

| **参数**      | **简单说明**                 | 即   |**输入数据类型**       |
| ------------- | ---------------------------- | ---------------------- |---------------------- |
| lookup_value  | 要查找的值                   | 找什么 | 数值、引用或文本字符串 |
| table_array   | 要查找的区域                 | 在哪找 | 数据表区域             |
| col_index_num | 返回数据在查找区域的第几列数 | 找第几个 | 正整数                 |
| range_lookup  | 模糊匹配/精确匹配            | "仔不仔细" | TRUE（或不填）/FALSE   |

## 一些其他 Tips

想到了再说吧……

## 使用的相关文件

<p><a  href="http://nut.yixuedh.com/vlookup例表.xlsx"><img src="http://nut.yixuedh.com/下载.png"></img><center>点我下载相关文件</center></a></p>

