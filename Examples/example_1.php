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
        'print'=>'',//设置打印格式
        'freezePane'=>'',//锁定行数，例如表头为第一行，则锁定表头输入A2
        'setARGB'=>'',//设置背景色，例如['A1', 'C1']
        'setWidth'=>'',//设置宽度，例如['A' => 30, 'C' => 20]
        'setBorder'=>'',//设置单元格边框
        'mergeCells'=>'',//设置合并单元格，例如['A1:J1' => 'A1:J1']
        'formula'=>'',//设置公式，例如['F2' => '=IF(D2>0,E42/D2,0)']
        'format'=>'',//设置格式，整列设置，例如['A' => 'General']
        'alignCenter'=>'',//设置居中样式，例如['A1', 'A2']
        'bold'=>'',//设置加粗样式，例如['A1', 'A2']
    ],
    'data'=>[
        ['Hello','World','!'],
        ['Excited','For','You!'],
    ]
]);

$excel->example_xls();
