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

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('memory_limit', '128M');

class phpExcel
{
    public $FileName;
    public $Creator;
    public $LastModifiedBy;
    public $Title;
    public $Subject;
    public $Description;
    public $Keywords;
    public $Category;
    public $data;
    public $data_key = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
        'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
        'U', 'V', 'W', 'X', 'Y', 'Z',
    ];

    public function __construct($data)
    {
        $this->FileName = $data['filename'];
        $this->Creator = $data['properties']['Creator'];//文件创建者
        $this->LastModifiedBy = $data['properties']['LastModifiedBy'];//最后更新
        $this->Title = $data['properties']['Title'];//标题
        $this->Subject = $data['properties']['Subject'];//主题
        $this->Description = $data['properties']['Description'];//描述
        $this->Keywords = $data['properties']['Keywords'];//关键词
        $this->Category = $data['properties']['Category'];//分类
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
            $spreadsheet->getActiveSheet()->fromArray($this->data['data'],NULL,$this->data['startCell']);
            /*foreach ($this->data['data'] as $key => $value) {
                foreach ($value as $k => $v) {
                    $spreadsheet->setActiveSheetIndex(0)
                        ->setCellValue($this->data_key[$key] . ($k + 1), $v);
                }
            }*/
        }

        // Miscellaneous glyphs, UTF-8
        /*$spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A4', 'Miscellaneous glyphs')
            ->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');*/

        // Rename worksheet
        $Sheet_index = $spreadsheet->getActiveSheetIndex();
        $spreadsheet->getActiveSheet()->setTitle($this->data['worksheet'][$Sheet_index]['Title']);

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);
        self::exportAction($spreadsheet,$this->FileName);


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
}