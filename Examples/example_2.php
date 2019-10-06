<?php
/**
 * Created by PhpStorm.
 * User: shayvmo
 * Date: 2019-9-23
 * Time: 16:41
 * Use for: php导出例子1
 */

require_once '../vendor/autoload.php';

include_once '../Core/phpExcel.php';

$excel = new \shayvmo\phpExcel([
    'filename'=>'test_file'.time(),
    'properties'=>[
        'Creator'=>'Eric',//文件创建者
        'LastModifiedBy'=>'Eric',//最后更新
        'Title'=>'导出文档',//标题
        'Subject'=>'导出文档',//主题
        'Description'=>'php导出文档',//描述
        'Keywords'=>'',//关键词
        'Category'=>'',//分类
    ],
    'worksheet'=>[
        ['Title'=>'sheet01']
    ],
    'startCell'=>'A1',
    'options'=>[
        'print'=>[],//设置打印格式

        'freezePane'=>[],//锁定行数，例如表头为第一行，则锁定表头输入A2

        'setARGB'=>[
            'A1'=>'FFFFFF00',
            'F2'=>'FFFFFF00'
        ],//设置背景色

        'setWidth'=>[
            'A'=>10,
            'B'=>15,
            'C'=>20
        ],//设置宽度

        'setBorder'=>['A1:I3','000000'],//设置单元格边框：位置，颜色

        'mergeCells'=>['A1:B1','B2:B3'],//设置合并单元格

        'formula'=>[],//设置公式，例如['F2' => '=IF(D2>0,E42/D2,0)']

        'format'=>['A'=>''],//设置格式，整列设置，例如['A' => 'General']

        //设置居中样式
        'alignment'=>[
//            'A1'=>['left','top'],//水平，垂直
//            'D1'=>['center','center'],//水平，垂直
//            'G1'=>['right','bottom'],//水平，垂直
            'A:I'=>['center','center'],//水平，垂直
        ],

        'bold'=>['A1:I1','E2'],//设置加粗样式，例如['A1', 'A2']
    ],
    'data'=>[
        [
            'A'=>'供应商',
            'B'=>'商品名称',
            'C'=>'系统编号',
            'D'=>'条形码',
            'E'=>'规格',
            'F'=>'单位',
            'G'=>'采购价',
            'H'=>'零售价',
            'I'=>'库存',
        ],
        [
            'A'=>'1',
            'B'=>'2',
            'C'=>'3',
            'D'=>'4',
            'E'=>'5',
            'F'=>'6',
            'G'=>'7',
            'H'=>'8',
            'I'=>'9',
        ],
        /*'A'=>[
            1=>'Hello',
            2=>'World',
            3=>'!',
        ],
        'B'=>[
            2=>'Excited',
            3=>'For',
            4=>'You!'
        ],
        'C'=>['供应商','商品名称','系统编号','条形码','规格','单位','采购价','零售价','库存']*/
    ]
]);

$excel->example_xls_2();
