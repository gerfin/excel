<h3 align="center">Normal Excel Exporter & Importer By PHP Array</h3>

## ✨ Features
- **Easily export array to Excel** 极简的将数组转化为Excel
- **Easily add header or not add header**  可选择添加行标题或者不添加
- **the excel staight put to browser** excel文件直接输出到浏览器

```php
<?php

use gerfin\excel\Exporter;

class IndexController extends Controller
{
    /**
     * @param array $data 源数据 required
     * @param array $header 行标题 optional
     * @param string $excelName excel名称 optional
     * @param string $format 格式xls或xlsx optional
     * @param string $sheetName sheet名称 optional
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function index() {
        $exporter = new Exporter();
        // the first param is required, others are optional
        $exporter->export([["1","jack"],["2", "nancy"]], ["id", "name"], "excelName", "xls", "sheetName");
    }
}
```
