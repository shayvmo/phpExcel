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
ini_set('display_errors', true);
ini_set('memory_limit', '128M');

class phpExcel_beta
{
    public $FileName;
    public $Creator;
    public $LastModifiedBy;
    public $Title;
    public $Subject;
    public $Description;
    public $Keywords;
    public $Category;
    public $alignment_style='';
    public $data;
    public $data_key = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
        'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
        'U', 'V', 'W', 'X', 'Y', 'Z'
    ];

    public function __construct($data)
    {
        $this->FileName = !empty($data['filename'])?$data['filename']:'Eric';
        $this->Creator = !empty($data['properties']['Creator'])?$data['properties']['Creator']:'Eric';//文件创建者
        $this->LastModifiedBy = !empty($data['properties']['LastModifiedBy'])?$data['properties']['LastModifiedBy']:'';//最后更新
        $this->Title = !empty($data['properties']['Title'])?$data['properties']['Title']:'Eric';//标题
        $this->Subject = !empty($data['properties']['Subject'])?$data['properties']['Subject']:'导出文档';//主题
        $this->Description = !empty($data['properties']['Description'])?$data['properties']['Description']:'';//描述
        $this->Keywords = !empty($data['properties']['Keywords'])?$data['properties']['Keywords']:'';//关键词
        $this->Category = !empty($data['properties']['Category'])?$data['properties']['Category']:'';//分类
        $this->data = $data;//数据
    }

    public function example_xls()
    {
        // Create new Spreadsheet object
        $spreadsheet = new Spreadsheet();
//        $spreadsheet->setActiveSheetIndex(1);//需要先创建Sheet

        // Set document properties
        $spreadsheet->getProperties()->setCreator($this->Creator)
            ->setLastModifiedBy($this->LastModifiedBy)
            ->setTitle($this->Title)
            ->setSubject($this->Subject)
            ->setDescription($this->Description)
            ->setKeywords($this->Keywords)
            ->setCategory($this->Category);

        // Add some data
        if (!empty($this->data['data'])) {
//            $spreadsheet->getActiveSheet()->fromArray($this->data['data'],NULL,$this->data['startCell']);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
            foreach ($this->data['data'] as $key => $value) {
                $new_value = array_values($value);
                foreach ($new_value as $k => $v) {
                    $spreadsheet->getActiveSheet()->setCellValueExplicit($this->data_key[$k] . ($key + 1), $v,DataType::TYPE_STRING);
                }
                unset($new_value);
            }
        }

        // Rename worksheet
        $Sheet_index = $spreadsheet->getActiveSheetIndex();
        $spreadsheet->getActiveSheet()->setTitle($this->data['worksheet'][$Sheet_index]['Title']);

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);
//        self::exportAction($spreadsheet,$this->FileName);
        self::exportLocal($spreadsheet,$this->FileName);


    }

    public function exportExcel()
    {
        $spreadsheet = new Spreadsheet();
//        $spreadsheet->setActiveSheetIndex(1);//需要先创建Sheet

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

        //设置单元格边框
        if(!empty($this->data['options']['setBorder']))
        {
            $border = [
                'borders'=>[
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => [ 'rgb' => $this->data['options']['setBorder'][1] ]
                    ]
                ]
            ];
            $activeSheet->getStyle($this->data['options']['setBorder'][0])->applyFromArray($border);

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

            $alignment_arr = [];//含有居中样式的单元格
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
//                    exit('success');
                    if (strpos($pCoordinate,':') === false) {
                        if(!isset($alignment_arr[$pCoordinate])) {
                            $alignment_arr[$pCoordinate] = $alignment;
                        }
                    } else {
                        list($a,$b) = explode(':',$pCoordinate);
                        if(ord($a)<ord($b)) {
                            $merge_arr = range(ord($a),ord($b));
                            foreach ($merge_arr as $value )
                            {
                                if(!isset($alignment_arr[chr($value)])) {
                                    $alignment_arr[chr($value)] = $alignment;
                                }
                            }
                        } elseif (ord($a) == ord($b)) {
                            if(!isset($alignment_arr[$a])) {
                                $alignment_arr[$a] = $alignment;
                            }
                        }

                    }
                    unset($alignment,$pCoordinate);
                } else {
                    unset($alignment,$pCoordinate);
                    continue;
                }
            }
        }

        // Add some data
        if (!empty($this->data['data'])) {
            foreach ($this->data['data'] as $key => $value) {
                foreach ($value as $k => $v) {
                    $activeSheet->setCellValueExplicit($k . ($key+1), $v,DataType::TYPE_STRING);
                    if (!empty($alignment_arr) && isset($alignment_arr[$k])) {
                        $activeSheet->getStyle($k . ($key+1))->applyFromArray($alignment_arr[$k]);
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
//        self::exportAction($spreadsheet,$this->FileName);
        self::exportLocal($spreadsheet,$this->FileName);
    }

    public function example_types()
    {
        $spreadsheet = new Spreadsheet();

        // Set document properties
        $spreadsheet->getProperties()->setCreator($this->Creator)
            ->setLastModifiedBy($this->LastModifiedBy)
            ->setTitle($this->Title)
            ->setSubject($this->Subject)
            ->setDescription($this->Description)
            ->setKeywords($this->Keywords)
            ->setCategory($this->Category);

        $spreadsheet->getDefaultStyle()
            ->getFont()
            ->setName('Arial')
            ->setSize(10);

// Add some data, resembling some different data types
        $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'String')
            ->setCellValue('B1', 'Simple')
            ->setCellValue('C1', 'PhpSpreadsheet');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A2', 'String')
            ->setCellValue('B2', 'Symbols')
            ->setCellValue('C2', '!+&=()~§±æþ');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A3', 'String')
            ->setCellValue('B3', 'UTF-8')
            ->setCellValue('C3', 'Создать MS Excel Книги из PHP скриптов');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A4', 'Number')
            ->setCellValue('B4', 'Integer')
            ->setCellValue('C4', 12);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A5', 'Number')
            ->setCellValue('B5', 'Float')
            ->setCellValue('C5', 34.56);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A6', 'Number')
            ->setCellValue('B6', 'Negative')
            ->setCellValue('C6', -7.89);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A7', 'Boolean')
            ->setCellValue('B7', 'True')
            ->setCellValue('C7', true);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A8', 'Boolean')
            ->setCellValue('B8', 'False')
            ->setCellValue('C8', false);

        $dateTimeNow = time();
        $spreadsheet->getActiveSheet()
            ->setCellValue('A9', 'Date/Time')
            ->setCellValue('B9', 'Date')
            ->setCellValue('C9', Date::PHPToExcel($dateTimeNow));
        $spreadsheet->getActiveSheet()
            ->getStyle('C9')
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_DATE_YYYYMMDD2);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A10', 'Date/Time')
            ->setCellValue('B10', 'Time')
            ->setCellValue('C10', Date::PHPToExcel($dateTimeNow));
        $spreadsheet->getActiveSheet()
            ->getStyle('C10')
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_DATE_TIME4);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A11', 'Date/Time')
            ->setCellValue('B11', 'Date and Time')
            ->setCellValue('C11', Date::PHPToExcel($dateTimeNow));
        $spreadsheet->getActiveSheet()
            ->getStyle('C11')
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_DATE_DATETIME);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A12', 'NULL')
            ->setCellValue('C12', null);

        $richText = new RichText();
        $richText->createText('你好 ');

        $payable = $richText->createTextRun('你 好 吗？');
        $payable->getFont()->setBold(true);
        $payable->getFont()->setItalic(true);
        $payable->getFont()->setColor(new Color(Color::COLOR_DARKGREEN));

        $richText->createText(', unless specified otherwise on the invoice.');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A13', 'Rich Text')
            ->setCellValue('C13', $richText);

        $richText2 = new RichText();
        $richText2->createText("black text\n");

        $red = $richText2->createTextRun('red text');
        $red->getFont()->setColor(new Color(Color::COLOR_RED));

        $spreadsheet->getActiveSheet()
            ->getCell('C14')
            ->setValue($richText2);
        $spreadsheet->getActiveSheet()
            ->getStyle('C14')
            ->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue('A17', 'Hyperlink');

        $spreadsheet->getActiveSheet()
            ->setCellValue('C17', 'PhpSpreadsheet Web Site');
        $spreadsheet->getActiveSheet()
            ->getCell('C17')
            ->getHyperlink()
            ->setUrl('https://github.com/PHPOffice/PhpSpreadsheet')
            ->setTooltip('Navigate to PhpSpreadsheet website');

        $spreadsheet->getActiveSheet()
            ->setCellValue('C18', '=HYPERLINK("mailto:abc@def.com","abc@def.com")');

        $spreadsheet->getActiveSheet()
            ->getColumnDimension('B')
            ->setAutoSize(true);
        $spreadsheet->getActiveSheet()
            ->getColumnDimension('C')
            ->setAutoSize(true);

// Rename worksheet
        $spreadsheet->getActiveSheet()->setTitle('Datatypes');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Xls)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $this->FileName . '.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($spreadsheet, 'Xls');
        $writer->save('php://output');
        exit;
    }

    public function example_style()
    {
        $spreadsheet = new Spreadsheet();
        $helper = new Sample();
        // Set document properties
        $spreadsheet->getProperties()->setCreator($this->Creator)
            ->setLastModifiedBy($this->LastModifiedBy)
            ->setTitle($this->Title)
            ->setSubject($this->Subject)
            ->setDescription($this->Description)
            ->setKeywords($this->Keywords)
            ->setCategory($this->Category);

        // Add some data
        if (!empty($this->data)) {

            foreach ($this->data as $key => $value) {
                foreach ($value as $k => $v) {
                    $spreadsheet->setActiveSheetIndex(0)
                        ->setCellValue($this->data_key[$key] . ($k + 1), $v);
                }
            }
        }

        $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray([
            'font' => [
                'name' => 'Arial',
                'bold' => true,
                'italic' => false,
                'underline' => Font::UNDERLINE_DOUBLE,
                'strikethrough' => false,
                'color' => ['rgb' => '808080']],
            'borders' => [
                'bottom' => [
                    'borderStyle' => Border::BORDER_DASHDOT,
                    'color' => ['rgb' => '808080']
                ],
                'top' => [
                    'borderStyle' => Border::BORDER_DASHDOT,
                    'color' => ['rgb' => '808080']
                ]
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'quotePrefix' => true
        ]);

//        $helper->write($spreadsheet, __FILE__);
        self::exportAction($spreadsheet,$this->FileName);
    }





    /**
     * 实际导出操作( xls 和 xlsx )
     * @param $spreadsheet
     * @param string $fileName
     * @param string $type
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function exportAction($spreadsheet,$fileName='excel',$type='Xls')
    {

        switch ($type)
        {
            case 'Xls':
                $type = 'Xls';
                $Content_type = 'application/vnd.ms-excel';
                $extension = '.xls';
                break;
            case 'Xlsx':
                $type = 'Xlsx';
                $Content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                $extension = '.xlsx';
                break;
            default:
                $type = 'Xls';
                $Content_type = 'application/vnd.ms-excel';
                $extension = '.xls';
                break;
        }

        // Redirect output to a client’s web browser (Xls)
        header('Content-Type: '.$Content_type);
        header('Content-Disposition: attachment;filename="' . $fileName . $extension.'"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($spreadsheet, $type);
        $writer->save('php://output');
        exit;
    }

    /**
     * 保存在服务器本地
     * @param $spreadsheet
     * @param string $fileName
     * @param string $savePath
     * @param string $type
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function exportLocal($spreadsheet,$fileName='excel',$savePath='',$type='Xls')
    {

        switch ($type)
        {
            case 'Xls':
                $type = 'Xls';
                break;
            case 'Xlsx':
                $type = 'Xlsx';
                break;
            default:
                $type = 'Xls';
                break;
        }

        if(!empty($savePath) && !is_dir($savePath)) {
            mkdir($savePath);
        }
        $writer = IOFactory::createWriter($spreadsheet, $type);
        $writer->save($savePath.$fileName.'.'.strtolower($type));
        exit;
    }
}