<?php
namespace XuDongYss\PhpOffice;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Excel{
	private Spreadsheet $spreadsheet;
	
	/**
	 * 初始化
	 */
	public function __construct() {
		
	}
	
	/**
	 * 获取操作对象
	 */
	public function getSpreadsheet() {
		return $this->spreadsheet;
	}
	
	/**
	 * 返回 excel 表格数据
	 * @param string		$file	文件或者上传的文件流
	 * @return array|mixed  		数组|出错会抛出异常
	 */
	public function import($file) {
		$_data = [];
		
		/* 加载文件 */
		$this->spreadsheet = IOFactory::load($file);
		/* 设置当前单元格 */
// 		$spreadsheet->setActiveSheetIndexByName('线下集市订单');
		/* 将当前表格单元格数据转化成数组 */
		$_data = $this->spreadsheet->getActiveSheet()->toArray(null, true, true, true);
		
		return $_data;
	}
	
	/**
	 * 导出到 Excel
	 */
	public static function export() {
		
	}
}