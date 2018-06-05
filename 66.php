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
    $objPHPExcel->getActiveSheet()->mergeCells('A1:A2');
    $objPHPExcel->getActiveSheet()->setCellValue('A1', '日期');
    $objPHPExcel->getActiveSheet()->mergeCells('B1:B2');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', '星期');
    $objPHPExcel->getActiveSheet()->mergeCells('C1:C2');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', '值班时段');
    $objPHPExcel->getActiveSheet()->mergeCells('D1:F1');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', '行政值班');
    $objPHPExcel->getActiveSheet()->setCellValue('D2', '值班1');
    $objPHPExcel->getActiveSheet()->setCellValue('E2', '值班2');
    $objPHPExcel->getActiveSheet()->setCellValue('F2', '值班3');
    $objPHPExcel->getActiveSheet()->mergeCells('G1:H1');
    $objPHPExcel->getActiveSheet()->setCellValue('G1', '网络值班');
    $objPHPExcel->getActiveSheet()->setCellValue('G2', '值班1');
    $objPHPExcel->getActiveSheet()->setCellValue('H2', '值班2');
    $objPHPExcel->getActiveSheet()->setCellValue('I1', '领导带班');
    $objPHPExcel->getActiveSheet()->setCellValue('I2', '带班人');
    $a=0;       //初始化周末
    $caday=3;  //初始化值班天数
    $M=6;     //初始化月份
    $Y=2018; //初始化年份
    $e_r=3; //初始化行数
    $week=""; //值班星期
    $duty_time=""; //值班时间
    $holidy=array(16,17,18); //假期时间
    $week_nor=array(); //周末正常值班
    $UserInfo=$db->getAllDutyUser();//查询所有值班人员
    $duty_user=count($UserInfo); //统计值班人数
    $LeaderInfo=$db->getAllDutyLeader();//查询所有值班领导人员
    $duty_leader=count($LeaderInfo); //统计值班人数
    $user_index=$db->getDutyIndex("user_num")[0];; //值班人员顺序
    $leader_index=$db->getDutyIndex("leader_num")[0]; //值班领导顺序
    $last_date=$db->getDutyIndex("date_record")[0]; //上月值班最后一天
    if (same_week($Y,$M,$last_date)) {
        $leader_index=$db->getDutyIndex("leader_num")[0]; //值班领导顺序
    }else{
        if(($leader_index+1)>$duty_leader){
            $leader_index=0;
        }else{
            $leader_index=$leader_index+1; 
        }
       
    }
    if ($user_index>=$duty_user) {
        $user_index=0;
    }
    $a_duty=array(); //值班天数
    $A=array(); //每个星期得开始日期
    $B=array(); //每个星期得结束日期     
    $objPHPExcel->getActiveSheet()->setTitle($Y."年".$M.'月合川报社值班表');
    $duty_day=date('t', strtotime($Y.'-'.$M));
    //程序开始
    $a_duty=dutyday($duty_day,$holidy);//排除放假日期
    AB($a_duty,$Y,$M); //星期开始结束日期
    for($i=1;$i<=$duty_day;$i++){
        if(in_array($i,$holidy)){  //排除假期值班   
            echo "假期已经排除";
            continue;
        }
        $a = date("w",strtotime($Y."-".$M."-".$i)); //判断是否是周末
        if(in_array($i,$A)){  //记录每周开始时间指针   
            $caday=$e_r;
        }
        switch ($a){ //星期几
            case 0:
            $week="星期日";
            break;
            case 1:
            $week="星期一";
            break;
            case 2:
            $week="星期二";
            break;
            case 3:
            $week="星期三";
            break;
            case 4:
            $week="星期四";
            break;
            case 5:
            $week="星期五";
            break;
            case 6:
            $week="星期六";
            break;   
        }
        if(($a =="0" || $a=="6")&&!(in_array($i,$week_nor))){ //周末值班表
            for ($j=0; $j<9; $j++) { 
                $duty_Index=getCells($j);//获取值班信息所在列
                switch ($j){
                    case 0:
                    $objPHPExcel->getActiveSheet()->mergeCells($duty_Index.$e_r.":".$duty_Index.($e_r+5)); //合并单元格
                    $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r), $M.'月'.$i.'日');   //写入日期
                    break;
                    case 1:
                    $objPHPExcel->getActiveSheet()->mergeCells($duty_Index.$e_r.":".$duty_Index.($e_r+5)); //合并单元格
                    $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r), $week);   //写入日期
                    break;
                    case 2:
                    for ($k=0; $k<6; $k++) {    //值班时间
                        switch ($k) {
                            case '0':
                            $duty_time="9:00-12:00";    
                            break;
                            case '1':
                            $duty_time="12:00-14:00";    
                            break;
                            case '2':
                            $duty_time="14:00-16:00";    
                            break;
                            case '3':
                            $duty_time="18:00-20:00";    
                            break;
                            case '4':
                            $duty_time="20:00-22:00";    
                            break;
                            case '5':
                            $duty_time="22:00-次日8:00";    
                            break;
                        }
                        $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r+$k), $duty_time);   //写入值班时间  
                    }
                    break;
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    for ($n=0; $n<6; $n++) { //值班1人员
                        if($user_index>($duty_user-1)){
                            $user_index=0;
                        }
                        $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r+$n), $UserInfo[$user_index++]);
                        $db->duty_record($UserInfo[$user_index-1],$Y.'-'.$M,1);
                    }     
                    break;
                }

            }
            $e_r=$e_r+6; //行数下调
        }else{ //平时值班
           for ($j=0; $j<9; $j++) { 
                $duty_Index=getCells($j);//获取值班信息所在列
                switch ($j){
                    case 0:
                    $objPHPExcel->getActiveSheet()->mergeCells($duty_Index.$e_r.":".$duty_Index.($e_r+2)); //合并单元格
                    $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r), $M.'月'.$i.'日');   //写入日期
                    break;
                    case 1:
                    $objPHPExcel->getActiveSheet()->mergeCells($duty_Index.$e_r.":".$duty_Index.($e_r+2)); //合并单元格
                    $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r), $week);   //写入日期
                    break;
                    case 2:
                    for ($k=0; $k<3; $k++) {
                        switch ($k) {
                            case '0':
                            $duty_time="12:00-14:00";    
                            break;
                            case '1':
                            $duty_time="18:00-20:00";    
                            break;
                            case '2':
                            $duty_time="20:00-次日8:00";    
                            break;
                        }
                        $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r+$k), $duty_time);   //写入值班时间      
                    }
                    break;
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    for ($n=0; $n<3; $n++) { //值班1人员
                        if($user_index>($duty_user-1)){
                            $user_index=0;
                        }
                        $objPHPExcel->getActiveSheet()->setCellValue($duty_Index.($e_r+$n), $UserInfo[$user_index++]);
                        $db->duty_record($UserInfo[$user_index-1],$Y.'-'.$M,0);  
                    }     
                    break;      
                }
            }
            $e_r=$e_r+3; //行数下调
        }


        //合并值班领导单元格
        if($leader_index>3){
            $leader_index=0;
        }
        // $a = date("w",strtotime($Y."-".$M."-".($i+1)); //判断周末被放掉
        // if ($a==0) {
        //     $objPHPExcel->getActiveSheet()->mergeCells('I'.$caday.":".'I'.($e_r-1)); //合并单元格
        //     $objPHPExcel->getActiveSheet()->setCellValue('I'.$caday, $LeaderInfo[$leader_index++]);  //值班领导 
        // }else if($i>=$duty_day){
        //     $objPHPExcel->getActiveSheet()->mergeCells('I'.$caday.":".'I'.($e_r-1)); //合并单元格
        //     $objPHPExcel->getActiveSheet()->setCellValue('I'.$caday, $LeaderInfo[$leader_index++]);  //值班领导 
        // } 
        // 
        if(in_array($i,$B)){  //结束指针
            $objPHPExcel->getActiveSheet()->mergeCells('I'.$caday.":".'I'.($e_r-1)); //合并单元格
            $objPHPExcel->getActiveSheet()->setCellValue('I'.$caday, $LeaderInfo[$leader_index++]);  //值班领导 
        }
    }

    //生成表格
    $db->setDutyIndex(($user_index),($leader_index-1),end($a_duty));//记录下月次数
    $objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');//生成excel文件  
    $objWriter->save($dir."/hcbs_".$M."_duty.xls");//保存文件  


      /*根据下标获得单元格所在列位置*/  
    function getCells($index){  
        $arr=range('A','Z');  
        //$arr=array(A,B,C,D,E,F,G,H,I,J,K,L,M,N,....Z);  
        return $arr[$index];  
    } 

     /*统计天数放到个数组里面*/  
    function dutyday($duty_day,$holidy){  
        for($i=1;$i<=$duty_day;$i++){
            if(in_array($i,$holidy)){  //排除假期值班   
                continue;
            }
            $a_duty[]=$i;
        }
        return $a_duty;
    }

    /*形成一个星期得开始日期和结束日期*/  
    function AB($a_duty,$Y,$M){  
        Global $A,$B;
        $num=count($a_duty);
        $a[]=$a_duty[0];
        for ($i=0; $i <$num ; $i++) { 
            $now = date("W",strtotime($Y."-".$M."-".$a_duty[$i])); //判断是否是现在是第几周
            $nex = date("W",strtotime($Y."-".$M."-".$a_duty[$i+1])); //判断是否是现在是第几周
            if (($nex-$now)>0) {
                $a[]=$a_duty[$i+1];
                $b[]=$a_duty[$i];
            }
        }
        $b[]=$a_duty[$num-1];
        $A=$a;
        $B=$b;
    }

    /*判断下个月得第一个星期和上个月最后一个星期是否在一周内*/  
    function same_week($Y,$M,$last_date){  
        $now = date("W",strtotime($Y."-".$M."-"."1")); //这个月第一天
        echo $now;
        $pre = date("W",strtotime($Y."-".($M-1)."-".$last_date)); //上个月最后一天
        echo $pre;
        if(($now-$pre)>0){
            return 0;
        }else{
            return 1;
        }
    }   
?>  