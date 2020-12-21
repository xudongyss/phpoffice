# PHPOffice

[TOC]

## 安装

```
composer require xudongyss/phpoffice
```
## 快速使用

### Excel

#### 初始化
```php
require_once 'vendor/autoload.php';

use XuDongYss\PhpOffice\Excel;

$excel = new Excel();
```

#### 导入

```php
$_data = $excel->import('001.xlsx');
//$_data 单元格数据，数组
```

```php
Array
(
    [1] => Array
        (
            [A] => 订单号
            [B] => 店铺
            [C] => 订单状态
            [D] => 核销方式
            [E] => 订单创建时间
            [F] => 买家付款时间
            [G] => 付款方式
            [H] => 支付流水号
            [I] => 商品金额合计
            [J] => 优惠合计
            [K] => 应收订单金额
            [L] => 订单实付金额
            [M] => 全部商品
            [N] => 买家手机号
        )

    [2] => Array
        (
            [A] => 
            [B] => 
            [C] => 交易成功
            [D] => 自提
            [E] => 2020-12-18 11:03:31
            [F] => 2020-12-18 11:03:49
            [G] => 微信支付
            [H] => 
            [I] => 10
            [J] => 0
            [K] => 10
            [L] => 10
            [M] => 长寿花(数量：1)
            [N] => 
        )
)
```

#### 导出

```php
$filename = '导出测试';
$_data = [];
$column_field = [
    ['column'=> '用户ID', 'field'=> 'id', 'data_type'=> 's'],
    ['column'=> '手机号', 'field'=> 'mobile', 'data_type'=> 's'],
    ['column'=> '昵称', 'field'=> 'nickname'],
    ['column'=> '余额', 'field'=> 'avail_money'],
    ['column'=> '已提现金额', 'field'=> 'cash_money'],
    ['column'=> '总余额', 'field'=> 'all_money'],
    ['column'=> 'VIP', 'field'=> 'vip_title'],
    ['column'=> '银行卡号', 'field'=> 'bank_card', 'data_type'=> 's'],
    ['column'=> '开户银行', 'field'=> 'bank_name'],
    ['column'=> '持卡人姓名', 'field'=> 'bank_realname'],
    ['column'=> '注册时间', 'field'=> 'register_time'],
];
$excel->export($filename, $_data, $column_field);
```

```php
//$_data
Array
(
    [0] => Array
        (
            [id] => 100000012
            [mobile] => 13500000021
            [nickname] => 13500****21
            [all_money] => 1600000.00
            [avail_money] => 1600000.00
            [cash_money] => 0.00
            [bank_name] => Ngân hàng Ngoại thương Việt Nam
            [bank_card] => 363636636636
            [bank_realname] => 啦啦啦肯
            [vip_title] => 普通会员
        )
)
```

