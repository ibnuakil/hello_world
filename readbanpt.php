<?php
/**
 * PHPExcel
 * adh krn bgt
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';


if (!file_exists("instrumen_banpt.xlsx")) {
	exit("Please run 05featuredemo.php first." . EOL);
}

echo date('H:i:s') , " Load from Excel2007 file" , EOL;
$callStartTime = microtime(true);

$objPHPExcel = PHPExcel_IOFactory::load("instrumen_banpt.xlsx");
$objPHPExcel->setActiveSheetIndex(0);
echo 'Set active sheet to '.$objPHPExcel->getActiveSheetIndex(),EOL;
echo 'Get content on cell G4: '.$objPHPExcel->getActiveSheet()->getCell('G4'),EOL;
echo 'Get content on cell G10: '.$objPHPExcel->getActiveSheet()->getCell('G10'),EOL;
$objPHPExcel->setActiveSheetIndex(1);
echo 'Set active sheet to '.$objPHPExcel->getActiveSheetIndex(),EOL;
echo 'Get content on cell D5: '.$objPHPExcel->getActiveSheet()->getCell('D5'),EOL;
echo 'Get content on cell D6: '.$objPHPExcel->getActiveSheet()->getCell('D6'),EOL;
$objPHPExcel->setActiveSheetIndex(5);
echo 'Set active sheet to '.$objPHPExcel->getActiveSheetIndex(),EOL;
echo 'Get content on cell G34: '.number_format($objPHPExcel->getActiveSheet()->getCell('G34')->getCalculatedValue(),2),EOL;
echo 'Get content on cell G35: '.$objPHPExcel->getActiveSheet()->getCell('G35')->getCalculatedValue(),EOL;
echo 'Finished reading file.';