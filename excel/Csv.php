<?php
declare(strict_types=1);

namespace esp\helper\excel;

use Generator;
use function esp\helper\mk_dir;

/**
 * 将文本导入 Excel Csv 并下载
 *
 * Class Csv
 * @package esp\helper\excel
 */
final class Csv
{
    public $filename;
    private $header;
    private $writeHead;
    private $space = ",";//或,
    private $debug = false;

    /**
     * Csv constructor.
     * @param string|null $filename 文件保存路径
     * @param bool $writeHead 是否要写入header头
     */
    public function __construct(string $filename = null, bool $writeHead = true)
    {
        if (is_null($filename)) {
            $rand = mt_rand() . time();
            $filename = _RUNTIME . "/{$rand}.csv";
        }
        if (substr($filename, -4) !== '.csv') $filename .= ".csv";
        mk_dir($filename);

        $this->filename = $filename;
        $this->writeHead = $writeHead;
    }

    /**
     * 设置第一行
     * 为防止数字被转换，在行头字段名后加n，该列数字前将被加上`号
     * 比如：手机号n，
     *
     * @param array $header
     */
    public function setHeader(array $header)
    {
        $this->header = $header;
        if (!$this->writeHead) return;

        $col = array_map(function ($lab) {
            if (substr($lab, -1) === 'n') $lab = substr($lab, 0, -1);
            if (!$this->debug) return iconv("UTF-8", 'gbk', $lab);
            return $lab;
        }, array_values($header));

        file_put_contents($this->filename, implode($this->space, $col) . "\n");//FILE_APPEND
    }

    /**
     * 导入内容，可多次导入
     *
     * @param array $data 导入的数据字段数和header要相同
     * @return $this
     */
    public function assign(array $data)
    {
        foreach ($data as $i => $rs) {
            $line = [];
            foreach ($this->header as $key => $label) {
                $n = '';
                if (substr($label, -1) === 'n') $n = '`';
                if (!$this->debug) {
                    $line[] = iconv("UTF-8", 'gbk//TRANSLIT//IGNORE', $n . ($rs[$key] ?? ''));
                } else {
                    $line[] = $n . ($rs[$key] ?? '');
                }
            }
            file_put_contents($this->filename, implode($this->space, $line) . "\n", FILE_APPEND);
        }
        return $this;
    }

    /**
     * 下载文件
     * @param string $downName 最终下载得到的文件名，与__construct中指定的文件名无关，那个是保存的临时文件，这儿是最终下载的文件名
     * @param bool $unlink 删除临时文件
     * @return bool|string
     */
    public function downFile(string $downName, bool $unlink = false)
    {
        file_put_contents($this->filename, "END\n", FILE_APPEND);
        if ($this->debug) {
            return file_get_contents($this->filename);
        }
        if (substr($downName, -4) !== '.csv') $downName .= ".csv";
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=\"{$downName}\"");
        header('Cache-Control: max-age=0');

        foreach ($this->readExcelData() as $value) echo $value;
        if ($unlink) unlink($this->filename);

        exit();
    }

    /**
     * 迭代方式读取
     *
     * @return Generator
     */
    private function readExcelData(): Generator
    {
        $handle = fopen($this->filename, 'rb');
        while (feof($handle) === false) {
            yield fgets($handle);
        }
        fclose($handle);
    }

}