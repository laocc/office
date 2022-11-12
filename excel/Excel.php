<?php
declare(strict_types=1);

namespace esp\helper\excel;

use esp\error\Error;
use PHPExcel_Exception;
use PHPExcel_Reader_Exception;

/**
 * 基于phpoffice/phpexcel，不过这个性能比不上phpoffice/phpspreadsheet
 *
 * Class Excel
 * @package library
 */
class Excel
{
    private $objPHPExcel;
    private $tables;
    private $sheets;

    /**
     * Excel constructor.
     * @param string $filename
     * @throws PHPExcel_Reader_Exception
     * @throws PHPExcel_Exception
     */
    public function __construct(string $filename)
    {
        /**
         * 屏蔽警告：Deprecated: Array and string offset access syntax with curly braces is deprecated in
         * 很多地方调用数组用的是 $value{0}这种形式，或全部手要改成$value[0]
         * 太多了，改不完，所以只有屏蔽
         */
        ini_set('error_reporting', strval(E_ALL & ~E_DEPRECATED));

        if (!is_readable($filename)) {
            if (is_readable(_ROOT . $filename)) {
                $filename = _ROOT . $filename;
            } else {
                throw new Error("{$filename}无法读取");
            }
        }
        $this->objPHPExcel = \PHPExcel_IOFactory::load($filename);
    }

    /**
     * 读取整个excel内容
     *
     * @throws PHPExcel_Exception
     */
    public function getData()
    {
        $sheetCount = $this->objPHPExcel->getSheetCount();//获取sheet表格数目
        for ($i = 0; $i < $sheetCount; $i++) {
            $this->tables[$i] = $this->readSheet($i);
        }
        return $this->tables;
    }

    /**
     * 读取mysql设置表，返回数组内容
     *
     * @param string $keyColumn
     * @param int $rows
     * @return mixed
     * @throws PHPExcel_Exception
     */
    public function getMysql(string $keyColumn = 'B', int $rows = 6)
    {
        $sheetCount = $this->objPHPExcel->getSheetCount();//获取sheet表格数目
        for ($i = 0; $i < $sheetCount; $i++) {
            $this->tables[$i] = $this->readMysqlTale($i, $keyColumn, $rows);
        }
        return $this->tables;
    }

    /**
     * 直接读取mysql设置表并生成sql
     *
     * @param string|null $engine Innodb,MyISAM
     * @param string|null $charset utf8mb4,utf8,gb2312
     * @return array
     * @throws PHPExcel_Exception
     */
    public function buildMysql(string $engine = null, string $charset = null)
    {
        $tables = $this->getMysql();
        $sqlAll = [];
        if (!$engine) $engine = 'Innodb';//MyISAM
        if (!$charset) $charset = 'utf8mb4';//utf8,gb2312
        foreach ($tables as $sheet) {
            foreach ($sheet as $table) {
                switch ($table['action']) {
                    case 'skip':
                    case 'lock':
                        continue;
                        break;
                    case 'update':

                        break;
                    default:
                        $drop = "TABLE IF EXISTS `{$table['table']}`;";

                        $fields = [];
                        foreach ($table['fields'] as $fid) {
                            $fields[] = "{$fid['name']} {$fid['type']} COMMENT '{$fid['label']} {$fid['notes']}'";
                        }
                        $fields[] = "primary key({$table['primary']})";
                        foreach ($table['keys'] as $key) {
                            $fields[] = "key {$key} ({$key})";
                        }
                        foreach ($table['spatial'] as $key) {
                            $fields[] = "Spatial Index {$key} ({$key})";
                        }
                        $build = "TABLE IF NOT EXISTS `{$table['table']}` (%s) ENGINE={$engine} DEFAULT CHARSET={$charset} COMMENT='{$table['label']}';";
                        $build = sprintf($build, implode(', ', $fields));
                        $sqlAll[] = "DROP {$drop} CREATE {$build}";
                }

            }
        }
        return $sqlAll;
    }


    /**
     * @param int $index
     * @param string $key
     * @param int $rows
     * @return array
     * @throws PHPExcel_Exception
     */
    private function readMysqlTale(int $index, string $key, int $rows)
    {
        $this->objPHPExcel->setActiveSheetIndex($index);
        $sheet = $this->objPHPExcel->getActiveSheet();

        $rowCount = $sheet->getHighestRow();//获取表格行数
        $columnCount = $sheet->getHighestColumn();//获取表格列数
        $this->sheets[$index] = [
            'label' => $sheet->getTitle(),
            'row' => $rowCount,
            'column' => $columnCount,
        ];

        $dataArr = array();
        $table = [];
        $tabRow = 0;
        for ($row = 1; $row <= $rowCount; $row++) {
            $rowIndex = 0;
            for ($column = ord($key); $column <= (ord($key) + $rows); $column++) {
                $rs = $sheet->getCell(chr($column) . $row)->getValue();
                if ($rowIndex === 1 and empty($rs)) break;
                $rs = trim(strval($rs));

                if ($rs && substr($rs, 0, 3) === 'tab') {
                    if (!empty($table) and isset($table['table'])) {
                        foreach ($table['fields'] as $ri => $rr) {
                            if (!isset($rr['name']) or empty($rr['name'])) {
                                unset($table['fields'][$ri]);
                            }
                        }
                        $dataArr[] = $table;
                        $tabRow = 0;
                    }
                    $table = [
                        'table' => $rs,
                        'label' => $sheet->getCell(chr($column + 3) . $row)->getValue(),
                        'action' => 'create',
                        'fields' => [],
                        'keys' => [],
                        'spatial' => [],
                    ];
                    break;
                }

                switch ($rowIndex) {
                    case 0:
                        $table['fields'][$tabRow]['key'] = $rs;
                        if (in_array($rs, ['skip', 'update', 'lock'])) {
                            $table['action'] = $rs;
                        }
                        break;
                    case 1:
                        $table['fields'][$tabRow]['name'] = $rs;
                        if (empty($rs)) {
                            unset($table['fields'][$tabRow]);
                            break;
                        }
                        if ($table['fields'][$tabRow]['key'] === 'key') {
                            $table['keys'][] = $rs;
                        } else if ($table['fields'][$tabRow]['key'] === 'index') {
                            $table['spatial'][] = $rs;
                        }
                        unset($table['fields'][$tabRow]['key']);
                        break;
                    case 2:
                        $table['fields'][$tabRow]['type'] = $rs;
                        if (stripos($rs, 'AUTO_INCREMENT')) {
                            $table['primary'] = $table['fields'][$tabRow]['name'];
                        }
                        break;
                    case 4:
                        $table['fields'][$tabRow]['label'] = $rs ?: $table['fields'][$tabRow]['name'];
                        break;
                    case 5:
                        $table['fields'][$tabRow]['notes'] = $rs;
                        break;
                    default:
                }
                $rowIndex++;
            }
            $tabRow++;
        }


        if (!empty($table) and isset($table['table'])) {
            foreach ($table['fields'] as $ri => $rr) {
                if (!isset($rr['name']) or empty($rr['name'])) {
                    unset($table['fields'][$ri]);
                }
            }
            $dataArr[] = $table;
            $tabRow = 0;
        }
        return $dataArr;
    }

    /**
     * @param int $index
     * @return array
     * @throws PHPExcel_Exception
     */
    private function readSheet(int $index)
    {
        $this->objPHPExcel->setActiveSheetIndex($index);
        $sheet = $this->objPHPExcel->getActiveSheet();

        $rowCount = $sheet->getHighestRow();//获取表格行数
        $columnCount = $sheet->getHighestColumn();//获取表格列数
        $this->sheets[$index] = [
            'label' => $sheet->getTitle(),
            'row' => $rowCount,
            'column' => $columnCount,
        ];

        $dataArr = array();

        for ($row = 1; $row <= $rowCount; $row++) {
            $dataArr[$row] = [];
            for ($column = 'A'; $column <= $columnCount; $column++) {
                $dataArr[$row][] = $this->objPHPExcel->getActiveSheet()->getCell($column . $row)->getValue();
            }
        }

        return $dataArr;
    }


}