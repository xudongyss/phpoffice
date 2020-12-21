<?php
namespace XuDongYss\PhpOffice;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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
	 * 导出设置
	 * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
	 */
	protected function exportConfig() {
		$this->spreadsheet = new Spreadsheet();
		$this->spreadsheet->getProperties()
						  ->setCreator('Unknown')
						  ->setLastModifiedBy('Unknown')
						  ->setKeywords('Unknown')
						  ->setCategory('zstyle');
	}
	
	/**
	 * 导出到 Excel
	 * @param string 	$filename
	 * @param [] 		$_data
	 * @param [] 		$column_field	列名称,字段名,单元格数据格式,
	 * 									数据示例：[['column'=> '标题','field'=> '字段','data_type'=>'单元格数据格式']]
	 * 									@param $dataType	\PhpOffice\PhpSpreadsheet\Cell\DataType		data_type 值参考
	 * @param string 	$sheet_title
	 * @param string 	$last_row
	 * @return string
	 */
	public function export($filename, $_data, $column_field, $sheet_title = '', $last_row = '') {
		/* 初始化 */
		$this->exportConfig();
		/* 新建工作表 */
		$this->spreadsheet->setActiveSheetIndex(0);
		/* 设置工作表名称 */
		$sheet_title = $sheet_title ? $sheet_title : $filename;
		$this->spreadsheet->getActiveSheet()->setTitle($sheet_title);
		
		$_abc = $this->_abc();
		/* 设置列 */
		$i = 1;
		foreach($column_field as $k=> $v) {
			$this->spreadsheet->getActiveSheet()->setCellValue($_abc[$k].$i, $v['column']);
		}
		/* 设置行 */
		$i++;
		foreach($_data as $item) {
			foreach($column_field as $_k=> $_v) {
				if(isset($_v['data_type'])) {
					$this->spreadsheet->getActiveSheet()->setCellValueExplicit($_abc[$_k].$i, $item[$_v['field']], $_v['data_type']);
				}else {
					$this->spreadsheet->getActiveSheet()->setCellValue($_abc[$_k].$i, $item[$_v['field']]);
				}
			}
			$i++;
		}
		/* 最后一行 */
		if($last_row) {
			foreach($last_row as $k=> $v) {
				switch($v['key']) {
					case 'fun':
						switch($v['value']) {
							case 'SUM':
								$pValue = '=SUM('.$_abc[$k].'2:'.$_abc[$k].($i-1).')';
								break;
							default:
								break;
						}
						break;
					default:
						$pValue = $v['value'];
						break;
				}
				
				$this->spreadsheet->getActiveSheet()->setCellValue($_abc[$k].$i, $pValue);
			}
		}
		
		return $this->excelOut($filename.'.'.date('YmdHis'));
	}
	
	/**
	 * 输出Excel表格
	 * @param  [object] $objPHPExcel PHPExcel对象，即数据集
	 * @param  [string] $fileName    导出的文件名
	 * @return
	 */
	protected function excelOut($filename, $path = 'Uploads/Office') {
		$filename = $filename.'.xlsx';
		$writer = new Xlsx($this->spreadsheet);
		$writer->save($filename);
		
		$file = trim($path, '/').'/'.$filename;
		rename($filename, $file);
		
		return $file;
	}
	
	/**
	 * 导出
	 */
	protected function exportColumn($columnNumber = 52) {
		$_data = [];
		
		for($i = 65; $i <= 90; $i++) {
			if($columnNumber <= 0) return $_data;
			
			$_data[] = chr($i);
			
			$columnNumber--;
		}
		
		for($i = 65; $i <= 90; $i++) {
			for($j = 65; $j <= 90; $j++) {
				if($columnNumber <= 0) return $_data;
				$_data[] = chr($i).chr($j);
				
				$columnNumber--;
			}
		}
		
		return $_data;
	}
}