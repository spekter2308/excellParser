<?php

error_reporting(E_ALL);
require_once ('simplexlsx.class.php');
require_once ('Classes/PHPExcel.php');


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

	/*\Zend\Debug\Debug::dump($x);
	die();*/

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