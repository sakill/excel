<?php 
	require dirname(__FILE__)."/dbconfig.php";//引入配置文件

	class db{
		public $conn=null;

		public function __construct($config){//构造方法 实例化类时自动调用 
				$this->conn=mysql_connect($config['host'],$config['username'],$config['password']) or die(mysql_error());//连接数据库
				mysql_select_db($config['database'],$this->conn) or die(mysql_error());//选择数据库
				mysql_query("set names ".$config['charset']) or die(mysql_error());//设定mysql编码
		}
		/**
		**根据传入sql语句 查询mysql结果集
		**/
		public function getResult($sql){
			$resource=mysql_query($sql,$this->conn) or die(mysql_error());//查询sql语句
			$res=array();
			while(($row=mysql_fetch_row($resource))!=false){
				foreach($row as $x=>$x_value) {
					$res[] = $x_value;
				}
			}
			return $res;
		}

		/**
		** 查询所有的值班人员
		**/
		public function getAllDutyUser(){
			$sql="select user from duty_people  where state=1";
			$res=$this->getResult($sql);
			return $res;
		}

		/**
		** 查询所有的值班领导
		**/
		public function getAllDutyLeader(){
			$sql="select user from duty_people  where state=2";
			$res=$this->getResult($sql);
			return $res;
		}


		/**
		** 记录每月最后值班领导和值班人员
		**/
		public function setDutyIndex($user_num,$leader_num,$date){
			$sql="UPDATE duty_index SET user_num = $user_num,leader_num = $leader_num,date_record = $date";
			$resource=mysql_query($sql,$this->conn);
		}

		/**
		** 获取值班顺序
		**/
		public function getDutyIndex($index){
			$sql="select $index from duty_index";
			$resource=mysql_query($sql,$this->conn);
			return mysql_fetch_row($resource);
		}

		/**
		** 记录值班记录
		**/
		public function duty_record($index,$month,$is){
			$sql="INSERT INTO duty_record (user,duty_month,is_week) VALUES ('$index','$month',$is)";
			$resource=mysql_query($sql,$this->conn);
		}


		/**
		** 所有在岗值班人员
		**/
		public function alluser(){
			$sql="select user from duty_people where state=0 || state=1";
			$res=$this->getResult($sql);
			return $res;
		}

		/**
		** 值班次数
		**/
		public function get_duty_num($user,$date,$is){
			$sql="select user from duty_record where user='$user' and duty_month='$date' and is_week=$is";
			$res=$this->getResult($sql);
			return count($res);

		}
	}
?>

