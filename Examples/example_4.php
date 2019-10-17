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
    'options'=>[
        'print'=>true,//设置打印格式

        //锁定行数，例如表头为第一行，则锁定表头输入A2
//        'freezePane'=>[],

        //设置背景色 单元格=>颜色RGB
        'setARGB'=>[
            'A1' => 'FFFFFF00',
            'F2' => 'FFFFFF00'
        ],

        //设置宽度
        'setWidth'=>[
            'A'=>15,
            'B'=>15,
            'C'=>20,
            'G'=>15,
        ],

        //设置单元格边框：位置=>颜色
        'setBorder'=>[
            'A1:G8'=>'000000'
        ],

        //设置合并单元格
        'mergeCells'=>['A1:G1','A2:G2','A3:B3','C3:D3','E3:G3','A4:B4','C4:D4','E4:G4'],

        //设置公式，例如['F2' => '=IF(D2>0,E42/D2,0)']
        'formula'=>['G9' => '=G7+G8'],

        //设置格式，整列设置，例如['A' => 'General']
//        'format'=>['A'=>''],

        //设置居中样式
        'alignment'=>[
//            'A1'=>['left','top'],//水平，垂直
//            'D1'=>['center','center'],//水平，垂直
//            'G1'=>['right','bottom'],//水平，垂直
//            'A1:I2'=>['center','center'],//水平，垂直
            'A:I'=>['center','center'],//水平，垂直
        ],

        //字体
        'font'=>[
            'A1'=>[
                'name' => '仿宋',//字体,选填
                'size' => 15,//字体大小,选填
                'bold' => true,//是否加粗,选填
                'italic' => false,//斜体,选填
                'strikethrough' => true,//删除线,选填
                'color' => '000000'//颜色,选填
            ],
            'B'=>[
                'name' => 'Arial',//字体,选填
                'size' => 11,//字体大小,选填
                'bold' => true,//是否加粗,选填
                'italic' => true,//斜体,选填
                'strikethrough' => true,//删除线,选填
                'color' => '808080'//颜色,选填
            ],
        ],

        //设置加粗样式，例如['A1', 'A2']
        'bold'=>['A5:H5'],

        //设置行高【行数组=>对应行高】
        'lineHeight'=> [
            [[1,2,3],30],
            [[4],20],
        ],

    ],
    'data'=>[
        [
            '开心仓'
        ],
        [
            'A'=>'入库验收单'
        ],
        [
            'A'=>'由快乐仓发往开心仓',
            'C'=>'操作日期：'.date('Y-m-d H:i:s'),
            'E'=>'进货单号：'.rand(1000,9999).rand(1000,9999).rand(1000,9999),
        ],
        [
            'A'=>'订单号：'.rand(1000,9999).rand(1000,9999).rand(1000,9999),
            'C'=>'操作员：Eric',
            'E'=>'经办人：Eric',
        ],
        [
            'A'=>'货号',
            'B'=>'商品名称',
            'C'=>'条形码',
            'D'=>'单位',
            'E'=>'数量',
            'F'=>'零售价',
            'G'=>'零售金额',
        ],
        [
            'A'=>'0001200023',
            'B'=>'商品1',
            'C'=>'7610700601068',
            'D'=>'1',
            'E'=>'5',
            'F'=>'6.05',
            'G'=>5*6.05,
        ],
        [
            'A'=>'0001200024',
            'B'=>'商品2',
            'C'=>'7610700601068,6911011010091,6911011010095',
            'D'=>'瓶',
            'E'=>'7',
            'F'=>'6.5',
            'G'=>7*6.5,
        ],
        [
            'A'=>'0001200025',
            'B'=>'上品3',
            'C'=>'6911011010091,6911011010095',
            'D'=>'条',
            'E'=>'5',
            'F'=>'6',
            'G'=>5*6,
        ]
    ]
]);

/**
 * 1、文件名称。例如：1.xls
 * 自动根据文件名后缀，导出xls 或者xlsx
 *
 *2、保存路径
 * 默认是导出Excel下载，填写路径可以保存在服务器本地
 */
$excel->exportExcel(time().'.xls',__DIR__);
