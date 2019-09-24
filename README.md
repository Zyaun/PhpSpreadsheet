本文为基础教程，主要用PHP实现对EXCEL文件的读取和写入。已封装成类，非常方便使用。

[PHPExcel](https://github.com/PHPOffice/PHPExcel)已经被官方弃用，并推荐换用[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet)来代替。我们的顺序是先读取，然后写入，最后下载EXCEL文件。

### 环境依赖

- PHP >= 5.6

- PHP_ZIP拓展

- PHP_XML拓展

- PHP_GD拓展



#### 已存在项目引入

```
composer require phpoffice/phpspreadsheet
```
#### 或者在composer.json文件中引入

```
"require": {
    "phpoffice/phpspreadsheet": "^1.3"
}
```
然后执行composer install

### 头部加载要用到的类

```
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
```
### EXCEL文件导入

```
/**
     * 导入文件插入数据库
     * @param string $tableName 表名
     * @param array $filed 字段名 [ 'create_id', 'create_time' ]
     * @param string $filePath 文件路径
     * @return bool|int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \yii\db\Exception
     */
    public function importExcel($tableName='', $filed=[], $filePath='')
    {
        $file = request()->file('excel');
        if(!$file){
            return $this->error("请先选择文件");
        }
        $info = $file->validate(['size'=>10240,'ext'=>'xlsx,xls,csv'])->move(ROOT_PATH . 'public' . DS . 'excel');
        if(!$info){     // 上传失败获取错误信息
            return $this->error($file->getError());
        }
        $exclePath = $info->getSaveName();  //获取文件名
        $filePath = ROOT_PATH . 'public' . DS . 'excel' . DS . $exclePath;    //上传文件的地址
        $params = $this->excelToArray($filed,$filePath);       // 得到数组
        foreach ($params as $key => $value) {
            $params[$key]['add_time'] = time();
        }
        if(model($tableName)->allowField(true)->saveAll($params) !==false){
            return $this->success('导入成功');
        }
        return $this->error('导入失败');
    }
 
    /**
     * excel文件 转 Array
     * @param string $filed      数组字段 
     * @param string $filePath 文件路径
     * @return array
     */
    public static function excelToArray($filed,$filePath)
    {
        $spreadsheet = IOFactory::load($filePath);// 载入excel表格
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow(); // 总行数
        $highestColumn = $worksheet->getHighestColumn(); // 总列数
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
        $data = [];
        for ($row = 2; $row <= $highestRow; ++$row) { // 从第二行开始
             $i = 0;
            $row_data = [];
            for ($column = 1; $column <= $highestColumnIndex; $column++) {
                $row_data[$filed[$i]] = $worksheet->getCellByColumnAndRow($column, $row)->getValue();
                 $i++;
            }
            $data[] = $row_data;
        }
        return $data;
    }

```
如何调用？ 封装后so easy， 英文有插入数据库操作，只需填入表名和字段即可

```
// 导入测试
  public function test(){
    return $this->importExcel('links',['link_name','links']);      
  }
```

### EXCEL文件导出

```
 /**
  * 导出excel表
  * $data：要导出excel表的数据，接受一个二维数组
  * $head：excel表的表头，接受一个一维数组
  * $key：$data中对应表头的键的数组，接受一个一维数组
  * $name：excel表的表名
  * 备注：此函数缺点是，表头（对应列数）不能超过26；循环不够灵活，一个单元格中不方便存放两个数据库字段的值
 */
public function exportExcel($data=[], $head=[], $keys=[] , $name='测试表')
{
   $count = count($head);  //计算表头数量
   $spreadsheet = new Spreadsheet();
   $sheet = $spreadsheet->getActiveSheet();
   for ($i = 65; $i < $count + 65; $i++) {     //数字转字母从65开始，循环设置表头：
    $sheet->setCellValue(strtoupper(chr($i)) . '1', $head[$i - 65]);
   }
   /*--------------开始从数据库提取信息插入Excel表中------------------*/
   foreach ($data as $key => $item) {             //循环设置单元格：
     //$key+2,因为第一行是表头，所以写到表格时   从第二行开始写
     for ($i = 65; $i < $count + 65; $i++) {     //数字转字母从65开始：
       $sheet->setCellValue(strtoupper(chr($i)) . ($key + 2), $item[$keys[$i - 65]]);
       $spreadsheet->getActiveSheet()->getColumnDimension(strtoupper(chr($i)))->setWidth(20); //固定列宽
     }
   }
   header('Content-Type: application/vnd.ms-excel');
   header('Content-Disposition: attachment;filename="' . $name . '.xlsx"');
   header('Cache-Control: max-age=0');
   $writer = new Xlsx($spreadsheet);
   $writer->save('php://output');
   //删除清空：
   $spreadsheet->disconnectWorksheets();
   unset($spreadsheet);
   exit;
}
```

如何调用？ easy too.
```
// 导出测试
public function test2(){     
    $head = ['网站名称','网站地址'];      //设置表头
    $keys = ['link_name','links'];      //数据中对应的字段，用于读取相应数据
    $data = model('links')->field('link_name,links')->select();      // 数据
    $this->exportExcel($data, $head, $keys, '友情链接');
}
```

html代码
```
<body>
    <a href="{:url('test2')}" >导出</a>
    <form action="{:url('test')}"  method="post" enctype="multipart/form-data" >
        <input name="excel" type="file">
        <button  >导入</button>
    </form>
</body>
```