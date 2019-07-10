<?php
namespace gerfin\excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Exporter
{
    public function export(array $data, array $header, string $excelName = 'excel', string $sheetName = 'sheet')
    {
        $newExcel = new Spreadsheet();
        $objSheet = $newExcel->getActiveSheet();  //获取当前操作sheet的对象
        $objSheet->setTitle($sheetName);  //设置当前sheet的标题
        for ($i = 0; $i < count($header); $i++) {
            $newExcel->getActiveSheet()->getColumnDimension(chr(65 + $i))->setAutoSize(true);
            $objSheet->setCellValue(chr(65 + $i)."1", $header[$i]);
        }

        foreach ($data as $i => $value) {
           foreach ($value as $j => $v) {
               $objSheet->setCellValue(chr(65 + $j), $i+2, $v);
           }
        }

        $this->downloadExcel($newExcel, $excelName, 'Xls');
    }

    //公共文件，用来传入xls并下载
    function downloadExcel($newExcel, $filename, $format)
    {
        // $format只能为 Xlsx 或 Xls
        if ($format == 'Xlsx') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        } elseif ($format == 'Xls') {
            header('Content-Type: application/vnd.ms-excel');
        }

        header("Content-Disposition: attachment;filename="
            . $filename . date('Y-m-d') . '.' . strtolower($format));
        header('Cache-Control: max-age=0');
        $objWriter = IOFactory::createWriter($newExcel, $format);

        $objWriter->save('php://output');

        //通过php保存在本地的时候需要用到
        //$objWriter->save($dir.'/demo.xlsx');

        //以下为需要用到IE时候设置
        // If you're serving to IE 9, then the following may be needed
        //header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        //header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        //header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        //header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        //header('Pragma: public'); // HTTP/1.0
        exit;
    }
}
