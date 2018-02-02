## 依赖：
* phpexcel 1.8

## 使用方法：

```` php
//1.实例化
/* 
 $file 要导入的.xlsx 文件
 $index 默认为0，可传入sheet的索引值，或者名字
*/
$my=new \myexcel\myexcel($file,$index);

//2.get_data
$my->get_data($row1=false,$row2=false,$col1=false,$col2=false);
/* 
  $row1,$row2 开始和结束的行数
  $col1,$col2 开始和结束的列名
  也可以传两个参数("A:1","D:60")，即从第一行A到第60行的D
  不传参的话，获取所有
*/

//3.get_row 或 get_col 获取指定行或指定列

$data=$my->get_row(5,"A","E");
$data=$my->get_col("D",3);

//4.write 往excel文件中写入数据，并保存为 filename.xlsx 文件
/* 
$data=array(
    array(),  //第一行
    array(),  //第二行
);
$data 必须是个二维数组
*/
$my->write($data,"filename");



```````