<?php
namespace XuDongYss\PhpOffice;

use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel{
	/**
	 * 导入
	 * @param file	$file	文件或者上传的文件流
	 */
	/**
	 * 
	 * @param file	$file	文件或者上传的文件流
	 * @return array|mixed  出错会抛出异常
	 */
	public static function import($file) {
		$_data = [];
		
		/* 加载文件 */
		$spreadsheet = IOFactory::load($file);
		
		$_data = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
		
		return $_data;
	}
}