<?php
/**
 * EExcel Class
 * EExcel is a Yii wrapper for PHPExcel Class
 * 
 * @author Alfa Adhitya <alfa2159@gmail.com>
 */
class EExcel extends CApplicationComponent
{

	/**
	 * @var string Path to PHPExcel Lib
	 */
	public $phpExcelPath = 'application.vendor.phpexcel.PHPExcel';

	/**
	 * @var PHPExcel The currently PHPExcel instance Loaded
	 */
	private $_excel;

	/**
	 * @var boolean
	 */
	private $_isInitialized = false;

	/**
	 * @var string
	 */
	private $_titleCell = null;

	/**
	 * Header format
	 * @var array header format
	 */
	private $_headerFormat = array();

	/**
	 * Title Format
	 * @var array title format
	 */
	private $_titleFormat = array(
		'font' => array(
			'bold' => true,
			'size' => '14',
		)
	);

	/**
	 * Initialize Component
	 */
	public function init()
	{
		if(!file_exists(Yii::getPathOfAlias($this->phpExcelPath)))
			throw new CHttpException(Yii::t('EExcel', 'PHPExcel cannot be loaded'), 500);
		
		if(!$this->_isInitialized) {
			spl_autoload_unregister(array('YiiBase','autoload'));
			Yii::import($this->phpExcelPath, true);

			$this->_excel = new PHPExcel();
			spl_autoload_register(array('YiiBase','autoload'));
			$this->_isInitialized = true;
		}

		// default param
		$this->_headerFormat = array(
			'font' => array(
		        'bold' => true
			),
			'fill' => array(
	            'type' => PHPExcel_Style_Fill::FILL_SOLID,
	            'color' => array('rgb' => 'FF0000')
	        )
		);
		parent::init();
	}

	/**
	 * Set title header
	 * @param string $title title
	 * @param string $cell Insert title on this cell address as the top left coordinate
	 * @return current EExcel class object
	 */
	public function setTitle($title = null, $cell = 'A1')
	{
		if($title!==null) {
			$this->_titleCell = $cell;
			$this->_excel->getActiveSheet()->setCellValue((string) $this->_titleCell, $title);
		}
		return $this;
	}

	/**
	 * Get header format value
	 * @return array The header format value
	 */
	public function getheaderFormat()
	{
		return $this->_headerFormat;
	}

	/**
	 * Set header format
	 * @param array $value header style
	 * @return current EExcel class object
	 */
	public function setheaderFormat($value = array())
	{
		$this->_headerFormat = CMap::mergeArray($this->_headerFormat, $value);
		return $this;
	}

	/**
	 * Get title format value
	 * @return array The title format value
	 */
	public function gettitleFormat()
	{
		return $this->_titleFormat;
	}

	/**
	 * Set title format
	 * @param array $value title style
	 * @return current EExcel class object
	 */
	public function settitleFormat($value = array())
	{
		$this->_titleFormat = CMap::mergeArray($this->_titleFormat, $value);
		return $this;
	}

	/**
	 * Apply header format
	 * @param string $cellRange cell range to apply header format
	 * @return current EExcel class object
	 */
	public function applyHeaderFormat($cellRange = null)
	{
		$this->_excel->getActiveSheet()->getStyle($cellRange)->applyFromArray($this->headerFormat);
		return $this;
	}

	/**
	 * Split cell into array, separate between letter and number of cell
	 * @param string $cell cell value
	 * @return array array of string, separate between letter and number of cell
	 */
	private function splitCell($cell)
	{
		$analyze = preg_split('/(?<=\d)(?=[a-z])|(?<=[a-z])(?=\d)/i', $cell);
		return $analyze;
	}

	/**
	 * Set data source
	 * @param array  $source Source array
	 * @param string $startCell Insert array starting from this cell address as the top left coordinate
	 */
	public function setData($source = array(), $startCell = 'A3')
	{
		$analyze = $this->splitCell($startCell);
		foreach($source as $key => $row ) {
			$currentRow = $analyze[1]++;
			$this->_excel->getActiveSheet()->fromArray($row, NULL, $analyze[0] . $currentRow);
		}

		/**
		 * Styling data content
		 */
		$this->_excel->getActiveSheet()->getStyle(
		    $startCell . ':' .
		    $this->_excel->getActiveSheet()->getHighestColumn() . 
		    $this->_excel->getActiveSheet()->getHighestRow()
		)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

		/**
		 * Styling Title
		 */
		if($this->_titleCell!==null) {
			$analyze = $this->splitCell($this->_titleCell);
			$this->_excel->getActiveSheet()->getRowDimension($analyze[1])->setRowHeight(25);
			
			$headerCellRange = (string)$this->_titleCell . ':' . $this->_excel->getActiveSheet()->getHighestColumn().$analyze[1];
			$this->_excel->getActiveSheet()->getStyle($headerCellRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$this->_excel->getActiveSheet()->getStyle($headerCellRange)->applyFromArray($this->titleFormat);
			$this->_excel->setActiveSheetIndex(0)->mergeCells($headerCellRange);
		}

		return $this;
	}

	/**
	 * Save excel file
	 * @param string $filepath path to save file
	 */
	public function save($filepath = null)
	{
		$ext = pathinfo($filepath, PATHINFO_EXTENSION);
		switch($ext) {
			case 'xlsx':
				$writerType = 'Excel2007';
			break;
			case 'xls':
			default:
				$writerType = 'Excel5';
			break;
		}
		$objWriter = PHPExcel_IOFactory::createWriter($this->_excel, 'Excel2007');
		$objWriter->save($filepath);
	}

	/**
	 * Send the excel document in browser with a specific name
	 * @param string $filename specific filename of downloaded file
	 */
	public function download($filename = null)
	{
		ob_start();
		$ext = pathinfo($filename, PATHINFO_EXTENSION);
		switch($ext) {
			case 'xlsx':
				$writerType = 'Excel2007';
				$mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
			break;
			case 'xls':
			default:
				$writerType = 'Excel5';
				$mime = 'application/vnd.ms-excel';
			break;
		}

		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-type: ' . $mime);
		header('Content-Disposition: attachment; filename="' . $filename . '"');
		header('Cache-Control: max-age=0');
		
		$objWriter = PHPExcel_IOFactory::createWriter($this->_excel, $writerType);
		ob_clean();
		$objWriter->save('php://output');
	}

}