<?php
namespace gerfin\excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Exporter
{
    public function export(array $data, array $header = [], string $excelName = 'excel', string $format = 'Xls', string $sheetName = 'sheet')
    {
        $newExcel = new Spreadsheet();
        $objSheet = $newExcel->getActiveSheet();
        $objSheet->setTitle($sheetName);

        $col = $header ? count($header) :
            isset($data[0]) ? count($data[0]) : 0;
        for ($i = 0; $i < $col; $i++) {
            $newExcel->getActiveSheet()->getColumnDimension(chr(65 + $i))->setAutoSize(true);
            if ($header)
                $objSheet->setCellValue(chr(65 + $i)."1", $header[$i]);
        }

        foreach ($data as $i => $value) {
           foreach ($value as $j => $v) {
               $objSheet->setCellValue(chr(65 + $j).($header ? $i + 2
                       : $i + 1), $v);
           }
        }

        $this->downloadExcel($newExcel, $excelName, $format);
    }


    function downloadExcel($newExcel, $filename, $format)
    {

        if (strtolower($format) === 'xlsx') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        }else {
            header('Content-Type: application/vnd.ms-excel');
        }

        header("Content-Disposition: attachment;filename="
            . $filename . date('Y-m-d') . '.' . strtolower($format));
        header('Cache-Control: max-age=0');
        $objWriter = IOFactory::createWriter($newExcel, $format);
        $objWriter->save('php://output');
        exit;
    }
}
