<?php
/**
 * Created by PhpStorm.
 * User: shayvmo
 * Date: 2019-9-23
 * Time: 15:28
 * Use for: phpExcel封装
 */

namespace shayvmo;

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

error_reporting(E_ALL);
ini_set('display_errors', false);
ini_set('memory_limit', '128M');

class phpExcel
{
    public $Creator;
    public $LastModifiedBy;
    public $Title;
    public $Subject;
    public $Description;
    public $Keywords;
    public $Category;
    public $data;
    public $index_key = [
        'A','B','C','D','E','F','G','H','I','J',
        'K','L','M','N','O','P','Q','R','S','T',
        'U','V','W','X','Y','Z',

    ];

    public function __construct($data)
    {
        $this->Creator = !empty($data['properties']['Creator'])?$data['properties']['Creator']:'Eric';//文件创建者
        $this->LastModifiedBy = !empty($data['properties']['LastModifiedBy'])?$data['properties']['LastModifiedBy']:'';//最后更新
        $this->Title = !empty($data['properties']['Title'])?$data['properties']['Title']:'Eric';//标题
        $this->Subject = !empty($data['properties']['Subject'])?$data['properties']['Subject']:'导出文档';//主题
        $this->Description = !empty($data['properties']['Description'])?$data['properties']['Description']:'';//描述
        $this->Keywords = !empty($data['properties']['Keywords'])?$data['properties']['Keywords']:'';//关键词
        $this->Category = !empty($data['properties']['Category'])?$data['properties']['Category']:'';//分类
        $this->data = $data;//数据
    }

    /**
     * 导出操作
     * @param string $file_name 文件名
     * @param string $save_path 保存路径，默认不填是下载表格
     */
    public function exportExcel($file_name='',$save_path='php://output')
    {

        try {
            if(empty($file_name)) {
                $file_name = 'test_file'.time().'.xls';
            }
            $file_arr = pathinfo(strtolower($file_name));
            $file_type = $file_arr['extension'];
            $spreadsheet = new Spreadsheet();

            $spreadsheet->getProperties()->setCreator($this->Creator)
                ->setLastModifiedBy($this->LastModifiedBy)
                ->setTitle($this->Title)
                ->setSubject($this->Subject)
                ->setDescription($this->Description)
                ->setKeywords($this->Keywords)
                ->setCategory($this->Category);


            $activeSheet = $spreadsheet->getActiveSheet();

            if (!empty($this->data['options']['print']) && is_bool($this->data['options']['print']) && $this->data['options']['print'] === true) {
                $activeSheet->getPageSetup()->setPaperSize(PageSetup:: PAPERSIZE_A4);
                /* 设置打印时边距 */
                $pValue = 1 / 2.54;
                $activeSheet->getPageMargins()->setTop($pValue / 2);
                $activeSheet->getPageMargins()->setBottom($pValue * 2);
                $activeSheet->getPageMargins()->setLeft($pValue / 2);
                $activeSheet->getPageMargins()->setRight($pValue / 2);
            }


            //列宽
            if(!empty($this->data['options']['setWidth']))
            {
                foreach ($this->data['options']['setWidth'] as $key=>$value)
                {
                    $activeSheet->getColumnDimension($key)->setWidth($value);
                }
            }

            //合并
            if(!empty($this->data['options']['mergeCells']))
            {
                foreach ($this->data['options']['mergeCells'] as $value)
                {
                    $activeSheet->mergeCells($value);
                }
            }

            //字体加粗
            if(!empty($this->data['options']['bold']))
            {
                foreach ($this->data['options']['bold'] as $value)
                {
                    $activeSheet->getStyle($value)->getFont()->setBold(true);
                }

            }

            //设置背景色
            if(!empty($this->data['options']['setARGB']))
            {
                foreach ($this->data['options']['setARGB'] as $key=>$value)
                {
                    $activeSheet->getStyle($key)
                        ->getFill()->setFillType(Fill::FILL_SOLID)
                        ->getStartColor()->setARGB($value);
                }

            }





            //公式
            if(!empty($this->data['options']['formula'])) {
                foreach ( $this->data['options']['formula'] as $k=>$v)
                {
                    $activeSheet->setCellValue($k,$v);
                }
            }

            // Add some data
            if (!empty($this->data['data'])) {
                foreach ($this->data['data'] as $key => $value) {
                    foreach ($value as $k => $v) {
                        if(in_array($k,$this->index_key,true)) {
                            $activeSheet->setCellValueExplicit($k . ($key+1), $v,DataType::TYPE_STRING);
                        } elseif(isset($this->index_key[$k])) {
                            $activeSheet->setCellValueExplicit($this->index_key[$k] . ($key+1), $v,DataType::TYPE_STRING);
                        }
                    }
                }
            }

            $highest_row = $activeSheet->getHighestRow();//最大行数

            //字体
            if(!empty($this->data['options']['font'])) {

                foreach ($this->data['options']['font'] as $k=>$v)
                {
                    $array = [
                        'font' => [
                            'name' => isset($v['name'])?$v['name']:'Arial',
                            'size' => isset($v['size'])?$v['size']:11,
                            'bold' => isset($v['bold'])?$v['bold']:false,
                            'italic' => isset($v['italic'])?$v['italic']:false,
                            'underline' => Font::UNDERLINE_NONE,
                            'strikethrough' => isset($v['strikethrough'])?$v['strikethrough']:false,
                            'color' => ['rgb' => isset($v['color'])?$v['color']:'000000']
                        ]
                    ];
                    if (preg_match('/^[A-Z]$/',$k)) {
                        $activeSheet->getStyle($k.'1:'.$k.$highest_row)->applyFromArray($array);
                    } else {
                        $activeSheet->getStyle($k)->applyFromArray($array);
                    }
                }

            }


            //设置居中样式
            if(!empty($this->data['options']['alignment']))
            {
                //水平
                $horizontal = [
                    'left'=>Alignment::HORIZONTAL_LEFT,
                    'right'=>Alignment::HORIZONTAL_RIGHT,
                    'center'=>Alignment::HORIZONTAL_CENTER,
                ];
                //垂直
                $vertical = [
                    'top'=>Alignment::VERTICAL_TOP,
                    'bottom'=>Alignment::VERTICAL_BOTTOM,
                    'center'=>Alignment::VERTICAL_CENTER,
                ];

                foreach ($this->data['options']['alignment'] as $key=>$value)
                {
                    $alignment = [
                        'alignment' => [
                            'horizontal' => isset($value[0])?$horizontal[$value[0]]:Alignment::HORIZONTAL_LEFT,
                            'vertical' => isset($value[1])?$vertical[$value[1]]:Alignment::VERTICAL_TOP,
                            'wrapText' => true,
                        ]
                    ];
                    $pCoordinate = strtoupper($key);
                    //匹配
                    if (preg_match('/^([A-Z]+\d+)([:]([A-Z]+\d+))?$/',$pCoordinate)) {
                        $activeSheet->getStyle($pCoordinate)->applyFromArray($alignment);

                    } else if (preg_match('/^([A-Z])([:]([A-Z]))?$/',$pCoordinate)) {

                        if (strpos($pCoordinate,':') === false) {
                            $activeSheet->getStyle($pCoordinate.'1:'.$pCoordinate.$highest_row)->applyFromArray($alignment);
                        } else {
                            list($a,$b) = explode(':',$pCoordinate);
                            $activeSheet->getStyle($a.'1:'.$b.$highest_row)->applyFromArray($alignment);

                        }
                        unset($alignment,$pCoordinate);
                    } else {
                        unset($alignment,$pCoordinate);
                        continue;
                    }
                }
            }


            //设置单元格边框
            if(!empty($this->data['options']['setBorder']))
            {
                foreach ($this->data['options']['setBorder'] as $key=>$value)
                {
                    $border = [
                        'borders'=>[
                            'allBorders' => [
                                'borderStyle' => Border::BORDER_THIN,
                                'color' => [ 'rgb' => $value ]
                            ]
                        ]
                    ];
                    $pCoordinate = strtoupper($key);
                    //匹配
                    if (preg_match('/^([A-Z]+\d+)([:]([A-Z]+\d+))?$/',$pCoordinate)) {
                        $activeSheet->getStyle($pCoordinate)->applyFromArray($border);

                    } else if (preg_match('/^([A-Z])([:]([A-Z]))?$/',$pCoordinate)) {
                        if (strpos($pCoordinate,':') === false) {
                            $activeSheet->getStyle($pCoordinate.'1:'.$pCoordinate.$highest_row)->applyFromArray($border);
                        } else {
                            list($a,$b) = explode(':',$pCoordinate);
                            $highest_column = $activeSheet->getHighestColumn();//最大列
                            $activeSheet->getStyle($a.'1:'.$b.$highest_row)->applyFromArray($border);
                        }
                        unset($alignment,$pCoordinate);
                    } else {
                        unset($alignment,$pCoordinate);
                        continue;
                    }
                }
            }

            //行高
            if(!empty($this->data['options']['lineHeight'])) {
                foreach ($this->data['options']['lineHeight'] as $k => $v) {
                    if( ( isset($v[0]) && is_array($v[0]) ) && ( isset($v[1])&&is_numeric($v[1])&&$v[1]>0 ) ) {

                        foreach ($v[0] as $key=>$value) {
                            $activeSheet->getRowDimension($value)->setRowHeight($v[1]);
                        }
                    }
                }
            }



            // Rename worksheet
            $Sheet_index = $spreadsheet->getActiveSheetIndex();
            $Sheet_Title = isset($this->data['worksheet'][$Sheet_index]['Title'])?$this->data['worksheet'][$Sheet_index]['Title']:'Sheet1';
            $activeSheet->setTitle($Sheet_Title);

            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $spreadsheet->setActiveSheetIndex(0);

            switch ($file_type)
            {
                case 'xls':
                    $Content_type = 'application/vnd.ms-excel';
                    break;
                case 'xlsx':
                    $Content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                    break;
                default:
                    $Content_type = 'application/vnd.ms-excel';
                    break;
            }
            ob_clean();
            ob_start();

            if($save_path != 'php://output') {
                if(!empty($save_path) && !is_dir($save_path)) {
                    mkdir($save_path);
                }
                $save_path = $save_path.'\\'.$file_name;
            } else {

                // Redirect output to a client’s web browser (Xls)
                header('Content-Type: '.$Content_type);
                header('Content-Disposition: attachment;filename="' . $file_name . '"');
                header('Cache-Control: max-age=0');
                // If you're serving to IE 9, then the following may be needed
                header('Cache-Control: max-age=1');

                // If you're serving to IE over SSL, then the following may be needed
                header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
                header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
                header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
                header('Pragma: public'); // HTTP/1.0

            }

            $writer = IOFactory::createWriter($spreadsheet, ucfirst($file_type));
            $writer->save($save_path);


            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            ob_end_flush();
            exit('saved at : '.$save_path);

        } catch (\Exception $e) {

            exit($e->getMessage());
        }

    }
}