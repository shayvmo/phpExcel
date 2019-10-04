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
    'filename'=>'test_file',
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
    'data'=>[
        ['Hello','World','!'],
        ['Excited','For','You!'],
    ]
]);

$excel->example_xls();
