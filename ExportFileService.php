<?php

error_reporting(E_ALL);
require_once ('simplexlsx.class.php');
require_once ('Classes/PHPExcel.php');
require_once ('dump.php');


function exportSh1()                                  //специфікація схема 1 Миколаїв
	{
		$xlsx = new \SimpleXLSX('muk/main.xlsx');
		//$xlsx = new \SimpleXLSX('axsis/axsis1-4_en.xlsx');			//для англ версії
		$x = $xlsx->rows(4);
		//$x = round($y, 2);
		/*\Zend\Debug\Debug::dump($x);
        die();*/
		$objPHPExcel = new \PHPExcel();
		//$i = 2;
		for ($i = 7; $i <= 39; $i++) {
			$objPHPExcel = \PHPExcel_IOFactory::load("muk/sch1.xlsx", 'Excel2007');
			//$objPHPExcel = \PHPExcel_IOFactory::load("axsis/sh1/shab1_en.xls", 'Excel2007');          //for eng vers
			//встановлення активного листа та його назва
			$objPHPExcel->setActiveSheetIndex(0);
			$activeSheet = $objPHPExcel->getActiveSheet();
			$activeSheet->setTitle("scheme 1");
			// Orientation page
			$activeSheet->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
			$activeSheet->getCell('A3')->setValue('Адреса:' . ' ' . $x[$i][2]);
			//$activeSheet->getCell('A6')->setValue('IHS reference and address:' . ' ' . $x[$i][2]);             //for eng vers
			$activeSheet->mergeCells('A3:B3');
			//1
			$n = 10;
			$j = 27;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//3
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//2
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//4
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//26
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//5.1
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//5.2
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//6
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//28.1
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//28.2
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//29
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//7
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('DN:' . $x[$i][$j] . 'Qn:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 4]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 5]);
			//27
			$j = $j + 6;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('DN:' . $x[$i][$j] . 'Qn:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 4]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 5]);
			//8
			$j = $j + 6;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//40
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//9
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//30
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//31
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//10
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//11
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//20
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//23
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//47
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//48
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//49
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//12
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//32
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//33
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//22
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//13
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//34
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//35
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//16
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//19
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//36
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//37
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//14
			$j = 180;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//15
			$j = $j + 7;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//39
			$j = 194;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//16
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('V:' . ($x[$i][$j]) . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//18
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('V:' . ($x[$i][$j]) . 'DN:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//17 41
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//21
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//38
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue('DN:' . ($x[$i][$j]) . 'Qn:' . ($x[$i][$j + 1]));
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);
			//42
			$j = $j + 5;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//25
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//43
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//44
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//45
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//46
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//51
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//50
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//52
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//53
			$j = $j + 4;
			$n = $n + 1;
			$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
			$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
			$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
			$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);
			//загальна вартість всього
			//$activeSheet->getCell('F' . $n)->setValue($x[$i][200]);
			$activeSheet->getCell('H64')->setValue($x[$i][263]);
			//налаштування шрифта по замовчуванню
			$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')->setSize(11);
			$objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);
			//mb_internal_encoding('latin1');
			$sysFilename = iconv('UTF-8', 'CP1251', $x[$i][2]);
			$sysFilename1 = iconv('UTF-8', 'CP1251', $x[$i][3]);
			$objWriter->save("muk/sh1/" . (string) $x[$i][1] . '. ' . $sysFilename . '.xls');
		}
		exit();
	}

function exportSh2()				                  //специфікація схема 1 Миколаїв
{
	$xlsx = new \SimpleXLSX('muk/main.xlsx');
	//$xlsx = new \SimpleXLSX('axsis/axsis1-4_en.xlsx');			//для англ версії
	$x = $xlsx->rows(5);
	//$x = round($y, 2);

	dump($x);
	die();

	$objPHPExcel = new \PHPExcel();

	//$i = 2;
	for ($i = 7; $i <=72; $i++)
	{
		$objPHPExcel = \PHPExcel_IOFactory::load("muk/sch2.xlsx", 'Excel2007');
		//$objPHPExcel = \PHPExcel_IOFactory::load("axsis/sh1/shab1_en.xls", 'Excel2007');          //for eng vers

		//встановлення активного листа та його назва
		$objPHPExcel->setActiveSheetIndex(0);
		$activeSheet = $objPHPExcel->getActiveSheet();
		$activeSheet->setTitle("scheme 2");

		// Orientation page
		$activeSheet->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);

		$activeSheet->getCell('A3')->setValue('Адреса:' . ' ' . $x[$i][2]);
		//$activeSheet->getCell('A6')->setValue('IHS reference and address:' . ' ' . $x[$i][2]);             //for eng vers
		$activeSheet->mergeCells('A3:B3');

		//1
		$n = 10;
		$j = 23;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//3
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//2
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//4
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//18
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//5,1
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//5.2
		$j = $j + 4;
		$n = $n + 1;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//6
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//20,1
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//20,2
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//21
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//7
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]) . 'type:' . ($x[$i][$j + 2]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 4]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 5]);

		//19
		$n = $n + 1;
		$j = $j + 6;
		$activeSheet->getCell('E' . $n)->setValue('DN:' . $x[$i][$j] . 'Qn:' . ($x[$i][$j + 1]) . 'type:' . ($x[$i][$j + 2]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 4]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 5]);

		//32
		$n = $n + 1;
		$j = $j + 6;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//9
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//22
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//23
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//10
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//39
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//40
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//41
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//11
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//24
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//25
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//8
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//12
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//26
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//27
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);


		//17
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//28
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//29
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//13
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue('DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//14
		$n = $n + 1;
		$j = $j + 8;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//31
		$n = $n + 1;
		$j = $j + 7;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//16
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//15,33
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//30
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue('kvs:' . $x[$i][$j] . 'DN:' . ($x[$i][$j + 1]));
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 3]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 4]);

		//34
		$n = $n + 1;
		$j = $j + 5;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//35
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//36
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//37
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//38
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//43
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//42
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//44
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);

		//45
		$n = $n + 1;
		$j = $j + 4;
		$activeSheet->getCell('E' . $n)->setValue($x[$i][$j]);
		$activeSheet->getCell('D' . $n)->setValue($x[$i][$j + 1]);
		$activeSheet->getCell('G' . $n)->setValue($x[$i][$j + 2]);
		$activeSheet->getCell('H' . $n)->setValue($x[$i][$j + 3]);


		//загальна вартість всього

		//$activeSheet->getCell('F' . $n)->setValue($x[$i][200]);
		$activeSheet->getCell('H56')->setValue($x[$i][225]);

		//налаштування шрифта по замовчуванню
		$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')->setSize(11);


		$objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);

		//mb_internal_encoding('latin1');
		$sysFilename = iconv('UTF-8', 'CP1251', $x[$i][2]);
		$sysFilename1 = iconv('UTF-8', 'CP1251', $x[$i][3]);
		$objWriter->save("muk/sh2/" . (string) $x[$i][1] . '. ' . $sysFilename . '.xls');
	}
	exit();
}

function exportPKO()                                     //для "ПКО"
{
	$objPHPExcel = new \PHPExcel();

	//встановлення активного листа та його назва
	$objPHPExcel->setActiveSheetIndex(0);
	$activeSheet = $objPHPExcel->getActiveSheet();
	$activeSheet->setTitle('PKO');

	// Orientation page
	$activeSheet->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
	//$$activeSheet->getPageSetup()->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);   якщо потрібно буде для друку формат

	//Для квитанції з правої сторони
	$activeSheet->getCell('O1')->setValue('Приложение 1
к приказу Министра финансов Республики Казахстан
от 20 декабря 2012 года № 562');
	$activeSheet->mergeCells('O1:T3');

	$activeSheet->getCell('R4')->setValue('Форма КО-1');
	$activeSheet->mergeCells('R4:T4');
	$activeSheet->getRowDimension('4')->setRowHeight(40);                       //встановлення ширини для 4 строки

	$activeSheet->getCell('O6')->setValue('Организация (индивидуальный предприниматель)');
	$activeSheet->mergeCells('O6:T6');

	$activeSheet->getCell('O7')->setValue('Филиал товарищества с ограниченной ответственностью "Вираж Запчасти" в городе Актобе');   //dinamic
	$activeSheet->mergeCells('O7:T7');

	$activeSheet->getCell('O9')->setValue('КВИТАНЦИЯ');
	$activeSheet->getCell('O10')->setValue('к приходному кассовому ордеру');
	$activeSheet->getCell('O11')->setValue('№');
	$activeSheet->getCell('P11')->setValue('1500000001');					//dinamic
	$activeSheet->mergeCells('O9:T9');
	$activeSheet->mergeCells('O10:T10');
	$activeSheet->mergeCells('P11:S11');

	$activeSheet->getCell('O13')->setValue('Принято от');
	$activeSheet->getCell('O14')->setValue('Товарищество с ограниченной ответственностью Аркада Сталь Пром');   //dinamic
	$activeSheet->getCell('O17')->setValue('Основание');
	$activeSheet->getCell('O18')->setValue('Оплата от покупателя');
	$activeSheet->mergeCells('O13:Q13');
	$activeSheet->mergeCells('O14:T16');
	$activeSheet->mergeCells('O17:Q17');
	$activeSheet->mergeCells('O18:T20');

	$activeSheet->getCell('O22')->setValue('Сумма');
	$activeSheet->getCell('O23')->setValue('Десять тысяч девятьсот тенге ноль тиын');
	$activeSheet->getCell('P25')->setValue('прописью');
	$activeSheet->getCell('O26')->setValue('М.П.');
	$activeSheet->getCell('P26')->setValue('26.03.2018 г.');				//dinamic
	$activeSheet->mergeCells('O23:T24');
	$activeSheet->mergeCells('P26:T26');

	$activeSheet->getCell('O28')->setValue('Главный бухгалтер или уполномоченное лицо');
	$activeSheet->getCell('O30')->setValue('подпись');
	$activeSheet->getCell('Q29')->setValue('/');
	$activeSheet->getCell('R29')->setValue('Не предусмотрен');
	$activeSheet->getCell('R30')->setValue('расшифровка подписи');
	$activeSheet->getCell('O32')->setValue('Кассир');
	$activeSheet->getCell('P33')->setValue('подпись');
	$activeSheet->getCell('S32')->setValue('/');
	$activeSheet->getCell('T33')->setValue('расшифровка подписи');
	$activeSheet->mergeCells('O28:T28');
	$activeSheet->mergeCells('R29:T29');
	$activeSheet->mergeCells('O30:P30');
	$activeSheet->mergeCells('R30:T30');
	$activeSheet->mergeCells('P33:R33');

	//Для прихідного касового ордера з лівої сторони
	$activeSheet->getCell('B4')->setValue('Организация (индивидуальный предприниматель)');
	$activeSheet->getCell('I4')->setValue('Филиал товарищества с ограниченной ответственностью "Вираж Запчасти" в городе Актобе'); //dinamic
	$activeSheet->getCell('J6')->setValue('ИИН/БИН');
	$activeSheet->getCell('K6')->setValue('160241030045');               //dinamic
	$activeSheet->getCell('H9')->setValue('Номер документа');
	$activeSheet->getCell('K9')->setValue('Дата составления');
	$activeSheet->getCell('H10')->setValue('1500000001');             //dinamic
	$activeSheet->getCell('K10')->setValue('26.03.2016 г.');		  //dinamic
	$activeSheet->getCell('B11')->setValue('ПРИХОДНЫЙ КАССОВЫЙ ОРДЕР');
	$activeSheet->mergeCells('B4:H4');
	$activeSheet->mergeCells('I4:L4');
	$activeSheet->mergeCells('K6:L6');
	$activeSheet->mergeCells('H9:J9');
	$activeSheet->mergeCells('K9:L9');
	$activeSheet->mergeCells('H10:J10');
	$activeSheet->mergeCells('K10:L10');
	$activeSheet->mergeCells('B11:J11');

	//таблиця 1 назви колонок
	$activeSheet->getCell('B13')->setValue('Дебет');
	$activeSheet->getCell('D13')->setValue('Кредит');
	$activeSheet->getCell('D15')->setValue('корреспондирующий счет');
	$activeSheet->getCell('G13')->setValue('Сумма, в KZT');
	$activeSheet->getCell('K13')->setValue('Код целевого назначения');
	$activeSheet->mergeCells('B13:C16');
	$activeSheet->mergeCells('D13:F14');
	$activeSheet->mergeCells('D15:F16');
	$activeSheet->mergeCells('G13:J16');
	$activeSheet->mergeCells('K13:L16');
	//таблиця 1 значення колонок !!! не всі динамічні, треба перевірити
	$activeSheet->getCell('B17')->setValue('1010');
	$activeSheet->getCell('D17')->setValue('1210');
	$activeSheet->getCell('G17')->setValue('10900');
	$activeSheet->getCell('K17')->setValue('');
	$activeSheet->getCell('L17')->setValue('');
	$activeSheet->mergeCells('B13:C16');
	$activeSheet->mergeCells('D17:F17');
	$activeSheet->mergeCells('G17:J17');
	$activeSheet->mergeCells('B17:C17');

	//основна інформація під таблицею
	$activeSheet->getCell('B19')->setValue('Принято от');
	$activeSheet->getCell('D19')->setValue('Товарищество с ограниченной ответственностью Аркада Сталь Пром');   //dinamic
	$activeSheet->getCell('B21')->setValue('Основание');
	$activeSheet->getCell('D21')->setValue('Оплата от покупателя');
	$activeSheet->getCell('B23')->setValue('Сумма');
	$activeSheet->getCell('C23')->setValue('Десять тысяч девятьсот тенге ноль тиын');    						//dinamic
	$activeSheet->getCell('C24')->setValue('прописью');
	$activeSheet->getCell('B29')->setValue('Главный бухгалтер');
	$activeSheet->getCell('I29')->setValue('/');
	$activeSheet->getCell('J29')->setValue('Не предусмотрен');
	$activeSheet->getCell('B30')->setValue('или уполномоченное лицо');
	$activeSheet->getCell('F30')->setValue('подпись');
	$activeSheet->getCell('J30')->setValue('расшифровка подписи');
	$activeSheet->getCell('B32')->setValue('Получил кассир');
	$activeSheet->getCell('F33')->setValue('подпись');
	$activeSheet->getCell('I32')->setValue('/');
	$activeSheet->getCell('J33')->setValue('расшифровка подписи');
	$activeSheet->mergeCells('B19:C19');
	$activeSheet->mergeCells('D19:L19');
	$activeSheet->mergeCells('B21:C21');
	$activeSheet->mergeCells('D21:L21');
	$activeSheet->mergeCells('C23:L23');
	$activeSheet->mergeCells('C24:L24');
	$activeSheet->mergeCells('B29:E29');
	$activeSheet->mergeCells('J29:L29');
	$activeSheet->mergeCells('B30:E30');
	$activeSheet->mergeCells('F30:H30');
	$activeSheet->mergeCells('J30:L30');
	$activeSheet->mergeCells('B32:E32');
	$activeSheet->mergeCells('F33:H33');
	$activeSheet->mergeCells('J33:L33');

	//додавання лінії розрізу між документами
	$activeSheet->getCell('N1')->setValue('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - л и н и я о т р е з а - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -');
	$activeSheet->mergeCells('N1:N34');
	$objPHPExcel->getActiveSheet()->getStyle('N1')->getAlignment()->setTextRotation(90);

	$style = array(
		'borderDThick' => [
			'borders' => [
				'bottom' => [
					'style' => \PHPExcel_Style_Border::BORDER_THICK,
					'color' => array('argb' => \PHPExcel_Style_Color::COLOR_BLACK),
				],
			],
		],
		'borderDThin' => [
			'borders' => [
				'bottom' => [
					'style' => \PHPExcel_Style_Border::BORDER_THIN,
					'color' => array('argb' => \PHPExcel_Style_Color::COLOR_BLACK),
				],
			],
		],
		'borderB' => [
			'borders' => [
				'outline' => [
					'style' => \PHPExcel_Style_Border::BORDER_THIN,
					'color' => array('argb' => \PHPExcel_Style_Color::COLOR_BLACK),
				],
			],
		],
		'borderO' => [
			'borders' => [
				'outline' => [
					'style' => \PHPExcel_Style_Border::BORDER_THICK,
					'color' => array('argb' => \PHPExcel_Style_Color::COLOR_BLACK),
				],
			],
		],
		'borderAll' => [
			'borders' => [
				'allborders' => [
					'style' => \PHPExcel_Style_Border::BORDER_THIN,
					'color' => array('argb' => \PHPExcel_Style_Color::COLOR_BLACK),
				],
			],
		],
		'alignmentVH' => [
			'alignment' => [
				'wrap'		=> true,
				'vertical'	=> \PHPExcel_Style_Alignment::VERTICAL_CENTER,
				'horizontal'=> \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			],
		],
		'alignmentVC' => [
			'alignment' => [
				'wrap'		=> true,
				'vertical'	=> \PHPExcel_Style_Alignment::VERTICAL_CENTER,
				'horizontal'=> \PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
			],
		],
		'alignmentRB' => [
			'alignment' => [
				'wrap'		=> true,
				'vertical'	=> \PHPExcel_Style_Alignment::VERTICAL_BOTTOM,
				'horizontal'=> \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
			],
		],
		'alignmentLT' => [
			'alignment' => [
				'wrap'		=> true,
				'vertical'	=> \PHPExcel_Style_Alignment::VERTICAL_TOP,
				'horizontal'=> \PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
			],
		],

		'alignmentVCHR' => [
			'alignment' => [
				'wrap'		=> true,
				'vertical'	=> \PHPExcel_Style_Alignment::VERTICAL_CENTER,
				'horizontal'=> \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
			],
		],
	);

	//налаштування шрифта по замовчуванню
	$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')->setSize(8);
	//$this->_addAutoSize($activeSheet);
	$activeSheet->getStyle('D5:F6')->getFont()->setSize(7)->setBold(false);
	$activeSheet->getStyle('O1:T3')->getFont()->setSize(7)->setItalic(true);
	$activeSheet->getStyle('O9:T9')->getFont()->setSize(9)->setBold(true);
	$activeSheet->getStyle('B11:J11')->getFont()->setSize(10)->setBold(true);
	$activeSheet->getStyle('D15:F16')->getFont()->setSize(7)->setBold(false);
	$activeSheet->getStyle('F30:L30')->getFont()->setSize(6)->setItalic(true);
	$activeSheet->getStyle('F33:L33')->getFont()->setSize(6)->setItalic(true);
	$activeSheet->getStyle('O30:T30')->getFont()->setSize(6)->setItalic(true);
	$activeSheet->getStyle('O33:T33')->getFont()->setSize(6)->setItalic(true);
	$activeSheet->getStyle('O25:T25')->getFont()->setSize(6)->setItalic(true);
	$activeSheet->getStyle('I4:L4')->getFont()->setBold(true);
	$activeSheet->getStyle('K6:L6')->getFont()->setBold(true);
	$activeSheet->getStyle('O7:T7')->getFont()->setBold(true);
	$activeSheet->getStyle('O10:T10')->getFont()->setBold(true);
	$activeSheet->getStyle('P26:T26')->getFont()->setBold(true);
	$activeSheet->getStyle('N1:N34')->getFont()->setSize(6)->setBold(false);
	$activeSheet->getStyle('C24:L24')->getFont()->setSize(8)->setItalic(true);


	//вирівнювання по центру
	$activeSheet->getStyle('B4:L4')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('K6:L6')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('O9:T9')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('O10:T10')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('H9:L10')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('B13:L17')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('C24:L24')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('F30:H30')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('F33:H33')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('N1:N34')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('O7:T7')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('P26:T26')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('O30:P30')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('R30:T30')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('P33:R33')->applyFromArray($style['alignmentVH']);
	$activeSheet->getStyle('T33:T33')->applyFromArray($style['alignmentVH']);

	//вирівнювання по правому краю
	$activeSheet->getStyle('I6:J6')->applyFromArray($style['alignmentVCHR']);
	$activeSheet->getStyle('O1:T3')->applyFromArray($style['alignmentVCHR']);
	$activeSheet->getStyle('R4:T4')->applyFromArray($style['alignmentRB']);


	//вирівнювання по лівому краю
	$activeSheet->getStyle('B11:J11')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('B19:L23')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('I29:L33')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('B29:E32')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O6:T6')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O13:T16')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O18:T24')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O26:O26')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O28:T28')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('R29:T29')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O32:O32')->applyFromArray($style['alignmentVC']);
	$activeSheet->getStyle('O13:T20')->applyFromArray($style['alignmentLT']);

	//креслення таблиці
	$activeSheet->getStyle('K6:L6')->applyFromArray($style['borderB']);
	$activeSheet->getStyle('H9:L9')->applyFromArray($style['borderAll']);
	$activeSheet->getStyle('H10:J10')->applyFromArray($style['borderO']);
	$activeSheet->getStyle('K10:L10')->applyFromArray($style['borderO']);
	$activeSheet->getStyle('B13:L17')->applyFromArray($style['borderAll']);
	$activeSheet->getStyle('B17:L17')->applyFromArray($style['borderO']);


	//підкреслювання нижнє
	$activeSheet->getStyle('I4:L4')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('D19:L19')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('D21:L21')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('C23:L23')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('F29:H29')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('J29:L29')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('F32:H32')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('J32:L32')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('O7:T7')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('P11:S11')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('O14:T16')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('O18:T20')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('O23:T24')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('O29:P29')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('R29:T29')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('P32:R32')->applyFromArray($style['borderDThin']);
	$activeSheet->getStyle('T32:T32')->applyFromArray($style['borderDThin']);

	//ширина строк
	$activeSheet->getRowDimension('1')->setRowHeight(12);
	$activeSheet->getRowDimension('4')->setRowHeight(40);
	$activeSheet->getRowDimension('7')->setRowHeight(30);


	//ширина стовбців
	$styleWidth = ['A' => 1, 'B' => 7, 'C' => 6, 'D' => 6, 'E' => 6, 'F' => 6, 'G' => 8, 'H' => 5, 'I' => 3, 'J' => 12, 'K' => 11, 'L' => 10, 'M' => 0.1, 'N' => 5, 'O' => 8, 'P' => 7, 'Q' => 8, 'R' => 7, 'S' => 6, 'T' => 24];
	foreach ($styleWidth as $col => $width) {
		$activeSheet->getColumnDimension($col)->setWidth($width);
	}


	$objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);

	header('Content-Type:  application/vnd.ms-excel');
	header("Content-Disposition: attachment;filename=exportPKO.xls");
	header('Cache-Control: max-age=0');
	$objWriter->save('php://output');											//файл не буде збережений, а буде відданий браузером на скачування
	exit();


}