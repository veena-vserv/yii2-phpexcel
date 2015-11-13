<?php
/**
 * Created by PhpStorm.
 * User: yantze.yang
 * Date: 2015/11/12
 * Time: 15:49
 */

namespace yantze\helper;
use yii\base\Exception;

/**
 * Class Excel
 *
 * @property \PHPExcel _objPHPExcel
 * @property string _savePath
 * @property array _suffix
 * @package app\helpers
 */
class Excel
{

    private $_objPHPExcel; //excel 句柄

    private $_savePath = "attach"; //保存目录

    private $_properties = [
        'creator' => "dataPlatform",
        'lastModifiedBy' => "dataPlatform",
        'title' => "Office 2007 XLSX Document",
        'subject' => "Office 2007 XLSX Document",
        'description' => "Document for Office 2007 XLSX, generated using PHP classes.",
        'keywords' => "office 2007 openxml php",
        'category' => "Reporting Forms",
    ];

    function __construct () {

        // set max executor time
        set_time_limit(300);

        // Create new PHPExcel object
        $this->_objPHPExcel = new \PHPExcel();
        $this->_objPHPExcel->setActiveSheetIndex(0);

        $this->_activeSheet = $this->_objPHPExcel->getActiveSheet();
    }

    /**
     * Set document properties
     * @param array $options document options
     */
    function setProperties($options) {
        $this->_properties = array_merge($this->_properties, $options);

        $this->_objPHPExcel->getProperties()->setCreator($this->_properties['creator'])
            ->setLastModifiedBy($this->_properties['lastModifiedBy'])
            ->setTitle($this->_properties['title'])
            ->setSubject($this->_properties['subject'])
            ->setDescription($this->_properties['description'])
            ->setKeywords($this->_properties['keywords'])
            ->setCategory($this->_properties['category']);
    }

    function setTitle ( $title ) {
        // Rename worksheet
        $this->_objPHPExcel->getActiveSheet()->setTitle ( $title );
    }

    public function addHead($head, $startRow = 1) {
        $pos = 'A' . $startRow;

        $this->_activeSheet->fromArray(
            $head,      // The data to set
            NULL,         // Array values with this value will not be set
            $pos         // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );
    }

    /**
     * 让 'A' 这种字母可以加数字而不变成数字
     *  eg.
     *
     *  $a = 'A';
     *  var_dump(++$a)
     *  string => 'B'
     *
     * @param $a
     * @param $b
     * @return bool
     */
    private function add($a, $b) {
        if (is_null($a) || !is_numeric($b)) {
            return false;
        }

        if ($b > -1) {
            for ($i = 0; $i < $b; ++$i) {
                ++$a;
            }
        } else {
            $b = -1 * $b;
            for ($i = 0; $i < $b; ++$i) {
                --$a;
            }
        }

        return $a;
    }

    /**
     * 添加 高级表头， 支持合并单元格和居中
     * @param $advanceHead
     * @param int $startRow
     * @param int $startCol
     */
    function addAdvancedMenu( $advanceHead, $startRow = 1, $startCol = 0) {
        foreach ( $advanceHead as $k => $cell ) {
            //$v=iconv("utf-8","gb2312",$v);
            if (isset($cell['colspan']) && $cell['colspan'] > 1)
            {
                // 设置 表格值
                $this->_activeSheet->getCellByColumnAndRow($startCol, $startRow)->setValue($cell['name']);

                // 设置 结束位置
                $cols = $cell['colspan']-1;

                // 合并单元格
                $this->_activeSheet->mergeCellsByColumnAndRow($startCol, $startRow, $startCol + $cols, $startRow);

                // 设置表格居中
                $this->_activeSheet->getCellByColumnAndRow($startCol, $startRow)
                    ->getStyle()
                    ->getAlignment()
                    ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                $startCol += $cell['colspan'];
            }

        }
    }

    /**
     * 设置主要的数据
     * @param $arrayData array
     * @param $startRow int
     */
    public function setData($arrayData, $startRow) {

        // demo data
        // $arrayData = array(
        //     array(NULL, 2010, 2011, 2012),
        //     array('Q1',   12,   15,   21),
        //     array('Q2',   56,   73,   86),
        //     array('Q3',   52,   61,   69),
        //     array('Q4',   30,   32,    0),
        // );

        // $activeSheet->freezePane('A2');  // 冻结窗口
        // $this->_activeSheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 1);

        $pos = 'A' . $startRow;
        $this->_activeSheet->fromArray(
            $arrayData,
            NULL,
            $pos
        );

    }

    function save($type = "xlsx", $filename = "") {

        if ( ! $filename ) {
            $filename = time ();
        }

        switch ($type) {
            case "htm":
            case "html" :

                // not test
                $objWriter = \PHPExcel_IOFactory::createWriter ( $this->_objPHPExcel, "HTML" );

                break;
            case "xls" :

                $objWriter = \PHPExcel_IOFactory::createWriter ( $this->_objPHPExcel, "Excel5" );

                break;
            case "xlsx" :

                $objWriter = \PHPExcel_IOFactory::createWriter ( $this->_objPHPExcel, "Excel2007" );

                break;
            case "csv" :

                // not test
                $objWriter = \PHPExcel_IOFactory::createWriter ( $this->_objPHPExcel, 'CSV' )->setDelimiter ( ',' )->setEnclosure ( '' )->setLineEnding ( "\r\n" )->setSheetIndex ( 0 );

                break;
            default :

                print( "sorry！".$type." not support" );

        }

        $this->setAttachDir(__DIR__); // $_savePath = __DIR__
        $path = $this->_savePath . "/" . $filename . "." . $type;

        $objWriter->save($path);

//        $EOL = '\n';
//        $callStartTime = microtime(true);
//        $callEndTime = microtime(true);
//        $callTime = $callEndTime - $callStartTime;
//
//        echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , $EOL;
//        echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , $EOL;
//        // Echo memory usage
//        echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , $EOL;

//        // Echo memory peak usage
//        echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , $EOL;
    }

    function output($suffix = "xlsx", $filename = null) {
        if ( ! $filename ) {
            $filename = time ();
        }

        switch ($suffix) {
            case "xls" :
                header ( 'Content-Type: application/vnd.ms-excel' );
                header ( 'Content-Disposition: attachment;filename="' . $filename . '.' . $suffix . '"' );
                header ( 'Cache-Control: max-age=0' );
                $_type = 'Excel5';
                break;
            case "xlsx" :
                // Redirect output to a client’s web browser (Excel2007)
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header ( 'Content-Disposition: attachment;filename="' . $filename . '.' . $suffix . '"' );
                header('Cache-Control: max-age=0');

                // If you're serving to IE over SSL, then the following may be needed
//                header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
//                header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
//                header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
//                header ('Pragma: public'); // HTTP/1.0
                $_type = 'Excel2007';
                break;
            case "html" :
                header ( "Content-Type:HTML text data" );
                header ( 'Content-Disposition: attachment;filename="' . $filename . '.' . $suffix . '"' );
                header ( 'Cache-Control: max-age=0' );
                $_type = 'HTML';
                break;
            case "csv" :
                //header ( "Content-type:text/csv" );
                Header('Content-Type: application/msexcel;charset=gbk');
                header ( 'Content-Disposition:attachment;filename="' . $filename . '.' . $suffix . '"' );
                //header ( 'Cache-Control:must-revalidate,post-check=0,pre-check=0' );
                //header ( 'Expires:0' );
                //header ( 'Pragma:public' );
                $_type = 'CSV';
                break;
            default :
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header ( 'Content-Disposition: attachment;filename="' . $filename . '.' . $suffix . '"' );
                header('Cache-Control: max-age=0');

                $_type = 'Excel2007';
                break;
        }

        $objWriter = \PHPExcel_IOFactory::createWriter ( $this->_objPHPExcel, $_type );
        $objWriter->save ( 'php://output' );

        // 如果缺少了 exit 这行命令，excel 会提示文件损坏，需要修复
        exit;
    }

    function setAttachDir ( $fullPath ) {

        if ( ! is_dir ( $fullPath ) ) {

            mkdir ( $fullPath, 755 );
        }
        if(!is_writable($fullPath)){
            chmod($fullPath, 755);
        }
        $this->_savePath = $fullPath;

    }
}