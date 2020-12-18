<?php
require_once 'vendor/autoload.php';

use XuDongYss\PhpOffice\Excel;

$_data = Excel::import('001.xlsx');

echo '<pre>';print_r($_data);