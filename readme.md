# PHPExcel

## Feature
- Run Limit Time 300s
- Simple
- Flexible
- High Performance

## Installation

The preferred way to install this extension is through composer.

Either run
```
composer require yantze/yii2-phpexcel
```
or add below line to the require section of your application's composer.json file.
```
"yantze/yii2-phpexcel" : "*"
```

## Usage
```
$head = [
	'fhead' => [
		[
			'colspan' => 2,
			'name' => '',
		],[
			'colspan' => 2,
			'name' => 'second',
		],
	],
	'head' => [
		[
			'sort' => 'tb_col_name1',
			'name' => 'col1',
		],[
			'sort' => 'tb_col_name2',
			'name' => 'col2',
		],[
			'sort' => 'tb_col_name3',
			'name' => 'col3',
		],[
			'sort' => 'tb_col_name4',
			'name' => 'col4',
		],
	]
];

$arrayData = [
	[NULL, 2010, 2011, 2012],
	['Q1',   12,   15,   21],
	['Q2',   56,   73,   86],
	['Q3',   52,   61,   69],
	['Q4',   30,   32,    0],
];

$name = 'filename';

$excel = new \yantze\helper\Excel();
$curRow = 1;

if (count($head['fhead']) > 1 || $head['fhead'][0]['name'] != '') {
	$excel->addAdvancedMenu($head['fhead'], $curRow++);
}

$head1 = ArrayHelper::getColumn($head['head'], 'name');
$excel->addHead($head1, $curRow++);

$excel->setData($arrayData, $curRow);
$excel->output("xlsx", $name);
```

## Useful PHPExcel Articles
[PHPExcel Offical Accessing Cells](https://github.com/PHPOffice/PHPExcel/blob/develop/Documentation/markdown/Overview/07-Accessing-Cells.md)

## License
MIT
