<?php   
    $dir=dirname(__FILE__);//查找当前脚本所在路径  
    require $dir."/db.php";//引入mysql操作类文件
    require $dir."/PHPExcel/PHPExcel.php";//引入PHPExcel    
    $config=include $dir."/dbconfig.php";  

    $db=new DB($phpexcel);  
    $objPHPExcel=new PHPExcel();//实例化PHPExcel类， 等同于在桌面上新建一个excel  
    $objSheet=$objPHPExcel->getActiveSheet();//获得当前活动单元格  
    /*设置样式*/  
    $objSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//设置excel文件默认水平垂直方向居中  
    $objSheet->getDefaultStyle()->getFont()->setSize(14)->setName("微软雅黑");//设置默认字体大小和格式  
    $objSheet->getDefaultRowDimension()->setRowHeight(30);//设置默认行高  
    $objPHPExcel->getActiveSheet()->setCellValue('A1', '序号');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', '姓名');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', '工作日值班');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', '周末值班');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', '应领金额(元)');
    $objPHPExcel->getActiveSheet()->setCellValue('F1', '实领金额(元)');
    $objPHPExcel->getActiveSheet()->setCellValue('G1', '签字确认');
    $objPHPExcel->getActiveSheet()->setCellValue('G1', '签字确认');
    $Y=2018;
    $M=6;
    $e_r=2; //初始化行数
    $x=0;
    $dutyM=$Y.'-'.$M;
    $UserInfo=$db->alluser();//查询所有在岗人员
    $duty_user=count($UserInfo); //统计值班人数
    $Leader=count($db->getAllDutyLeader());
    $total=2600;
    echo $duty_user;
    echo $Leader;
    $A=0;
    $B=0;
    $name='';
    for ($m=0; $m < 4; $m++) {
        switch ($m) {
                case '0':
                $name="张宏荣";
                break;
                case '1':
                $name="梁益学";
                break;
                case '2':
                $name="胡学英";
                break;
                case '3':
                $name="尹红梅";
                break;
            }

       for ($n=0; $n<7; $n++) { 
           $duty_Index=getCells($n);//获取值班信息所在列
            switch ($n) {
                case '0':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),(++$x));   //写入日期     
                break;
                case '1':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$name);   //姓名     
                break;
                case '2':
                  
                break;  
                case '3':
                  
                break;  
                case '4':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),'650');   //应领工资     
                break;  
                case '5':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),'650');   //实际工资     
                break;   
                case '6':
                                                                                            //签字     
                break;          
            }    
        }
        $e_r++;
    }
    for ($i=0; $i < $duty_user; $i++) {
        $A=$db->get_duty_num($UserInfo[$i],$dutyM,0); //工作日次数
        $B=$db->get_duty_num($UserInfo[$i],$dutyM,1); //周末次数
        $C=$A*50+$B*60;
       for ($k=0; $k<7; $k++) { 
           $duty_Index=getCells($k);//获取值班信息所在列
           switch ($k) {
                case '0':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),($x+$i+1));   //写入日期     
                break;
                case '1':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$UserInfo[$i]);   //姓名     
                break;
                case '2':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$A.'*50');   //工作日工资     
                break;  
                case '3':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$B.'*60');   //周末工资     
                break;  
                case '4':
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$C);   //应领工资     
                break;  
                case '5':
                if($C>650){
                    $C=650;
                }
                $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r),$C);   //实际工资 
                $total=$total+$C;    
                break;   
                case '6':
                                                                                            //签字     
                break;          
            }    
       }
       $e_r++;
    }

    $objPHPExcel->getActiveSheet()->setCellValue('A'.$e_r, '总人数');
    $objPHPExcel->getActiveSheet()->setCellValue('B'.$e_r, ($duty_user+$Leader).'人');
    $objPHPExcel->getActiveSheet()->mergeCells('C'.$e_r.':'.'E'.$e_r);
    $objPHPExcel->getActiveSheet()->setCellValue('C'.$e_r, '合计金额');
    $objPHPExcel->getActiveSheet()->mergeCells('F'.$e_r.':'.'G'.$e_r);
    $objPHPExcel->getActiveSheet()->setCellValue('F'.$e_r, $total);
    $e_r++;
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$e_r.':'.'G'.$e_r);
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$e_r, '按照渝财行【2015】72号文件精神,每人每班按照值班60元、日常值班50元的标准发放值班补助。');
    $objPHPExcel->getActiveSheet()->setTitle($Y."年".$M.'月合川报社值班补贴表');
    $objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');//生成excel文件  
    $objWriter->save($dir."/hcbs_".$M."_calc_duty.xls");//保存文件  


     /*根据下标获得单元格所在列位置*/  
    function getCells($index){  
        $arr=range('A','Z');  
        //$arr=array(A,B,C,D,E,F,G,H,I,J,K,L,M,N,....Z);  
        return $arr[$index];  
    } 

?>  