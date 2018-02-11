<?php

namespace FatExcel\Tests;

use PHPUnit\Framework\TestCase;
use FatExcel\FatExcel;

class FatExcelTest extends TestCase
{
    public function testImportToArray()
    {
        $excel = new FatExcel();

        $array = $excel->importToArray(__DIR__.'/Excel/test_for_array_import.xlsx');
        $this->assertEquals(100, count($array));
        $this->assertEquals(5, count($array[0]));
        $this->assertEquals([16, 17, 18, 19, 20], $array[15]);

        $count = 0;
        $sum = 0;
        array_walk($array, 
                function ($value, $key) use (&$count, &$sum){
                    foreach ($value as $temp) {
                        if (! empty($temp)) {
                            $count++;
                            $sum += $temp;
                        }
                    }
        });
        $this->assertEquals(500, $count);
        $this->assertEquals(26250, $sum);
    }

    public function testImportToArrayWithEmpty()
    {
        $excel = new FatExcel();

        $array = $excel->importToArray(__DIR__.'/Excel/test_for_array_import_empty.xlsx');
        $this->assertEquals(100, count($array));
        $this->assertEquals(6, count($array[0]));
        $this->assertEquals(['', 16, 17, 18, '', ''], $array[15]);

        $count = 0;
        $sum = 0;
        array_walk($array, 
                function ($value, $key) use (&$count, &$sum){
                    foreach ($value as $temp) {
                        if (! empty($temp)) {
                            $count++;
                            $sum += $temp;
                        }
                    }
        });
        $this->assertEquals(385, $count);
        $this->assertEquals(20351, $sum);
    }

    public function testExportFile()
    {
        $excel = new FatExcel();
        $filename = __DIR__.'/Excel/test_export_xlsx.xlsx';

        //10行数据，每行5条
        $data = [];
        for ($i = 0; $i < 10; $i++) {
            for ($j = 0; $j < 5; $j++) {
                $data[$i][] = mt_rand(1,1000);
            }
        }

        $excel->exportToFile($data, $filename);

        $importData = $excel->importToArray($filename);

        for ($i = 0; $i < 10; $i++) {
            for ($j = 0; $j < 5; $j++) {
                $this->assertEquals($data[$i][$j], $importData[$i][$j]);
            }
        }

        unlink($filename);
    }
}