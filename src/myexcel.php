<?php
namespace tansuo1989;

class myexcel{
    public $allData=array();
    public $excel="";
    public $sheet="";

    public function __construct($file=false,$index=0){
        if($file==false){return;}
        $this->excel=\PHPExcel_IOFactory::load($file);
        if(is_numeric($index)){
            $this->sheet=$this->excel->getSheet($index);
        }else{
            $this->sheet=$this->excel->getSheetByName($index);
        }
        if(!$this->sheet){exit($index." is no exist !");}
        $this->allData=$this->sheet->toArray(null, true, true, true);
    }

    public function get_data($row1=false,$row2=false,$col1=false,$col2=false){
        if($row1==false){
            return $this->allData;
        }
        if(strpos($row1,":")){
             $start=explode(":",$row1);
             $end=explode(":",$row2);
             $row1=$start[1];
             $row2=$end[1];
             $col1=$start[0];
             $col2=$end[0];
        }
        $col1=strtoupper($col1);
        $col2=strtoupper($col2);
        $re=array();
        foreach($this->allData as $k=>$v){
            if($k>=$row1 && $k<=$row2){
                foreach($v as $kk=>$vv){
                    if($this->compare($kk,$col1)>=0&&$this->compare($kk,$col2)<=0){
                        $re[$k][$kk]=$vv;
                    }
                }
            }
        }
        return $re;
    }

    public function get_row($row=1,$col1=false,$col2=false){
        if($col1==false){
            return $this->allData[$row];
        }
        return $this->get_data($row,$row,$col1,$col2); 
    }

    public function get_col($col="A",$row1=1,$row2=false){
        $col=strtoupper($col);
        $data=$row2==false?$this->allData:$this->get_data($row1,$row2,$col,$col);
        $re=array();
        foreach($data as $v){
          $re[]=$v[$col];
        }
        return$re;
    }

    public function write($arr,$filename=false){
        $excel=new \PHPExcel();
        $excel->getSheet()->fromArray($arr);
        $filename=$filename?:time();
        $objWriter = \PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
        if($objWriter->save($filename.".xlsx")){return true;}
        return false;
    }

    protected function compare($a,$b){
        $alen=strlen($a);
        $blen=strlen($b);
        if($alen==$blen){
           return strcmp($a,$b);
        }else{
            return $alen-$blen;
        }
    }





}//endclass