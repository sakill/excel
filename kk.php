<?php  
require './PHPExcel/PHPExcel.php';  
$PHPExcel=new PHPExcel();   //实例化一个PHPExcel类，相当于在桌面上创建一个表格  
$sheet=$PHPExcel->getActiveSheet();  //获得当前获得sheet的操作对象  
$sheet->setTitle('demo');   //设置名称  
$sheet->setCellValue("A1","姓名")->setCellValue('B1',"分数");   //填充数据  
$sheet->setCellValue("A2","张三")->setCellValue('B2',"90");   //填充数据  
$sheet->setCellValue("A3","李四")->setCellValue('B3',"80");   //填充数据  
$writer=PHPExcel_IOFactory::createWriter($PHPExcel,'Excel2007');   //按照指定格式生成Excel文件  
$writer->save('./demo.xlsx');   // 保存到指定目录下  