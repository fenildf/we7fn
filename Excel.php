<?php
//没有命名空间类加载方式 调用类用\
require_once IA_ROOT . '/framework/library/phpexcel/PHPExcel.php';
require_once IA_ROOT . '/framework/library/phpexcel/PHPExcel/IOFactory.php';

class Excel{

   
    private $xls='.xls';//保存文件后缀
     private $xlsx='.xlsx';//保存文件后缀
    private $excelPath;//文件保存的绝对位置
    private $letter=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
    //Excel5和Excel2007
    private $ExcelVersion=['Excel2007','Excel5'];

    private $sheetNum=0;
    //phpExcel实例化对象
    private $phpExcel;
    private $phpWriter;
    private $xlsReader;
    private $phpSheet;

    public function __construct()
    {
        //  实例化PHPExcel类
        $this->phpExcel = new PHPExcel();
       
    }

    /**
     * 创建新的Sheet 支持链式操作
     * @param string $sheet_title
     * @param array  $data       导出数据内容
     * @param array  $excelHeader导出表头
     * @return $this
     * @throws \Exception
     * @throws \PHPExcel_Exception
     */
    public function createSheet($sheet_title='Sheet1',$data=[],$excelHeader=[])
    {
        if ( empty($excelHeader)||!is_array($excelHeader)){
            throw new Exception("Parameter is incorrect");
            return $this;
        }
        $sheet_num = $this->getNewSheetNum();
        $objPHPExcel=$this->phpExcel;
        $objPHPExcel->createSheet($sheet_num);
        //设置当前的sheet
        
        $objPHPExcel->setActiveSheetIndex($sheet_num);
        //设置sheet的name
        $objPHPExcel->getActiveSheet()->setTitle($sheet_title);
        $sheet=$objPHPExcel->getActiveSheet();
       //表头设置
        $excelHeader=array_values($excelHeader);
        foreach($excelHeader as $item=>$value){
            $sheet->setCellValue($this->letter[$item]."1",$value);
        }
       //表内容设置
        foreach($data as $item=>$value ){
            $value=array_values($value);
            foreach($value as $i=>$v)
            //$sheet->setCellValue($this->letter[$i].($item+2),$value[$i]);
            $sheet->setCellValueExplicit($this->letter[$i].($item+2),$value[$i], PHPExcel_Cell_DataType::TYPE_STRING);
        }
        return $this;
    }


    /**
     * 导出下载
     * @return output  
     */ 
    public function downFile($excelName='',$ver='xls')
    {
        ob_start();
        if(empty($excelName)){
            $excelName = 'Excel'.date("Ymdhis");
        }
        try{
            if($ver=='xls'){
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[1]);
              $ext = $this->xls;
            }else{
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[0]);
              $ext = $this->xlsx;
            }   
        }catch(Exception $e){
            throw new Exception("Export failed");
        }
        header('Content-Type: application/vnd.ms-excel; charset=utf-8');
        header("Content-Disposition: attachment;filename=".$excelName.$ext);
        header('Cache-Control: max-age=0');
        $this->phpWriter->save('php://output');  
        ob_end_flush();
        die();
    }
    /**
     * 导出保存服务器上
     * @param  String  $path 保存路径
     * @param  boolean $activate 自定义保存路径需要将此处设置为true
     * @return Object  
     */ 
    public function saveFile($excelName='',$filepath='',$ver='xls')
    {
        
        if (empty($filepath)) {
            $filepath  =  IA_ROOT.'/data/Excel/';
        }
        if(empty($excelName)){
            $excelName = 'Excel'.date("Ymdhis");
        }
        if(!$this->checkPath($filepath)){
            throw new Exception("The current directory is not writable");
        } else{
            try{
                 if($ver=='xls'){
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[1]);
              $ext = $this->xls;
            }else{
              $this->phpWriter = PHPExcel_IOFactory::createWriter($this->phpExcel,$this->ExcelVersion[0]);
              $ext = $this->xlsx;
            }  
            }catch(Exception $e){
                throw new Exception("Export failed");
            }
          
            $this->excelPath=$filepath.$excelName.$ext;
            $this->phpWriter->save($this->excelPath); 
            return $this;
        }

    }
    /**
     * 导入基本设置
     * @param  String  $path 保存路径
     * @param  boolean $activate 自定义保存路径需要将此处设置为true
     * @return Object  
     * @throws Exception
     * @throws \PHPExcel_Exception
     */ 
    public function loadExcel($filepath)
    {
      if(!is_file($filepath)){
        throw new Exception("File does not exist");
      }
       try{$type = strtolower( pathinfo($filepath, PATHINFO_EXTENSION) );
            if($type=='xlsx'){
                $xlsReader =  PHPExcel_IOFactory::createReader($this->ExcelVersion[0]);
                $xlsReader->setReadDataOnly(true); 
                $xlsReader->setLoadSheetsOnly(true);
                $this->xlsReader=$xlsReader->load($filepath);
            }elseif('xls'==$type){
                $xlsReader =  PHPExcel_IOFactory::createReader($this->ExcelVersion[1]);
                $xlsReader->setReadDataOnly(true); 
                $xlsReader->setLoadSheetsOnly(true);
                $this->xlsReader=$xlsReader->load($filepath);
            }elseif('csv'==$type){
              $handle = fopen($filepath, 'r');
        $dataArray = array();
        $row = 0;
        while ($data = fgetcsv($handle)) {
            $num = count($data);

            for ($i = 0; $i < $num; $i++) {
                $dataArray[$row][$i] = mb_convert_encoding($data[$i], "utf-8", 'GBK');
            }
            $row++;

        }

        $this->xlsReader= $dataArray;
            }
       }catch(Exception $e){
            throw new Exception("Reading failed");
       }
      return $this->xlsReader;
    }
    /**
     * 获取新的Sheet编号
     * @return int
     */
    protected function getNewSheetNum(){
        $sheet_num=$this->sheetNum;
        $this->sheetNum=$sheet_num+1;
        return $sheet_num;
    }

    /**
     * 检查目录是否可写
     * @param  string   $path    目录
     * @return boolean
     */
    protected function checkPath($path)
    {
        if (is_dir($path)) {
            return true;
        }
        if (mkdir($path, 0755, true)) {
            return true;
        } else {
            return false;
        }
    }
    /**
     * 返回数组的维度
     * @param  Array   $arr 任意数组
     * @return number  数组维度
     */
    protected function array_depth($arr)
    {
        if(!is_array($arr)) return 0;
        $max_depth = 0;
        foreach($arr as $item1)
        {
            $t1 = $this->array_depth($item1);
            if( $t1 > $max_depth) $max_depth = $t1;
        }
        return $max_depth + 1;
    }
}
