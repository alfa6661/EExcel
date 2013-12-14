## EExcel - a Yii wrapper for PHPExcel Class


### Installation

1. Copy EExcel file to protected/components.
2. Download [PHPExcel](http://phpexcel.codeplex.com/releases/view/96183).
3. Create a phpexcel directory on protected/vendor.
4. In your protected/config/main.php, add the following :

```php
//..
'components'=>array(
	//..
    'excel' => array(
		'class' => 'EExcel'
	),
    //..	
)
```


### Usage

```php
public function actionTest()
{
    $data = array(
		array('Hello', 'World', '!!!'),
		array('X', 'Y', 'Z')
	);
    $excel = Yii::app()->excel
    	->setTitle('Laporan', 'B1')
    	->setData($data, 'B3')
    	->setheaderFormat(array(
			'font' => array(
		        'bold' => true
			),
			'fill' => array(
	            'color' => array('rgb' => 'FF0000')
	        )
		))
    	->applyHeaderFormat('B3:D3')
    	->download('file.xlsx');        	
}
```

### Save file to disk

```php
->save('/path/filename.xls');
```