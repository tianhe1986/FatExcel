<?php

namespace FatExcel;

class FatExcel
{
    /**
     * 根据扩展名获取用于处理的write type
     * 
     * @param string $filename
     * @return string
     */
    protected function getWriteTypeFromExtension($filename)
    {
        $pathinfo = pathinfo($filename);
        if (!isset($pathinfo['extension'])) {
            return 'Xlsx';
        }

        switch (strtolower($pathinfo['extension'])) {
            case 'xlsx': // Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm': // Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return 'Xlsx';
            case 'xls': // Excel (BIFF) Spreadsheet
            case 'xlt': // Excel (BIFF) Template
                return 'Xls';
            case 'ods': // Open/Libre Offic Calc
            case 'ots': // Open/Libre Offic Calc Template
                return 'Ods';
            case 'htm':
            case 'html':
                return 'Html';
            case 'csv':
                return 'Csv';
            default:
                return null;
        }
    }

    /**
     * 将excel文件内容导出到数组
     * 
     * @param string $filename
     * @return array
     */
    public function importToArray($filename)
    {
        $reader = new \SpreadsheetReader($filename);

        $data = [];
        foreach ($reader as $row){
            $data []= $row;
        }

        return $data;
    }

    /**
     * 大数组直接导出到xlsx文件
     * 
     * @param array $data
     * @param string $filename
     * @param array $headerColumn
     */
    public function exportToFile($data, $filename, $headerColumn = [])
    {
        $writer = new \XLSXWriter();
        $writer->writeSheet($data, '', $headerColumn);
        $writer->writeToFile($filename);
    }

    /**
     * 大数组直接生成xlsx文件下载
     * 
     * @param array $data
     * @param string $filename
     * @param array $headerColumn
     */
    public function exportDownload($data, $filename = 'download.xlsx', $headerColumn = [])
    {
        $writer = new \XLSXWriter();
        $writer->writeSheet($data, '', $headerColumn);

        $this->outputDownloadHeader('xlsx', $filename);
        $writer->writeToStdOut();
        exit;
    }

    /**
     * 输出下载头信息
     * 
     * @param string $type
     */
    protected function outputDownloadHeader($type, $filename = null)
    {
        if (is_null($filename)) {
            $filename = 'download.'.strtolower($type);
        }
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');
    }
}

