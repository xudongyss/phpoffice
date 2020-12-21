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

