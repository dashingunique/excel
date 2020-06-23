<?php

namespace dashingunique\excel\traits;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Reader\SheetInterface;
use Box\Spout\Writer\WriterInterface;
use dashingunique\library\Collection;

/**
 * 导入Excel（csv、ods）文件
 * Trait ImportData
 * @package dashingunique\excel\traits
 * @property bool $withHeader
 */
trait ImportData
{
    /**
     * @var int 图表编号
     */
    protected $sheetNumber = 1;

    /**
     * 获取导入文件类型
     * @param string $path 文件路径
     * @return mixed
     */
    abstract protected function getType($path);

    /**
     * 设置读写信息
     * @param ReaderInterface|WriterInterface $readerOrWriter
     * @return mixed
     */
    abstract protected function setOptions(&$readerOrWriter);

    /**
     * 导入文件
     * @param $path
     * @param callable|null $callback
     * @return Collection
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function import($path, ?callable $callback = null)
    {
        $reader = $this->reader($path);
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheetNumber != $key) {
                continue;
            }
            $collection = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return uniqueCollection($collection ?? []);
    }

    /**
     * 导入表格信息
     * @param string $path
     * @param callable|null $callback
     *
     * @return Collection
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     *
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function importSheets($path, callable $callback = null)
    {
        $reader = $this->reader($path);

        $collections = [];
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            $collections[] = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return new Collection($collections);
    }

    /**
     * 读取文件信息
     * @param $path
     *
     * @return \Box\Spout\Reader\ReaderInterface
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     *
     * @throws \Box\Spout\Common\Exception\IOException
     */
    private function reader($path)
    {
        $reader = ReaderFactory::create($this->getType($path));
        $this->setOptions($reader);
        /* @var \Box\Spout\Reader\ReaderInterface $reader */
        $reader->open($path);

        return $reader;
    }

    /**
     * 导入单元格
     * @param SheetInterface $sheet
     * @param callable|null $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, callable $callback = null)
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        if ($this->withHeader) {
            foreach ($sheet->getRowIterator() as $k => $row) {
                if ($k == 1) {
                    $headers = $this->toStrings($row);
                    $count_header = count($headers);
                    continue;
                }
                if ($count_header > $count_row = count($row)) {
                    $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                } elseif ($count_header < $count_row = count($row)) {
                    $row = array_slice($row, 0, $count_header);
                }
                if ($callback) {
                    if ($result = $callback(array_combine($headers, $row))) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = array_combine($headers, $row);
                }
            }
        } else {
            foreach ($sheet->getRowIterator() as $row) {
                if ($callback) {
                    if ($result = $callback($row)) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = $row;
                }
            }
        }

        return $collection;
    }

    /**
     * 字符串转换
     * @param array $values
     *
     * @return array
     */
    private function toStrings($values)
    {
        foreach ($values as &$value) {
            if ($value instanceof \Datetime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string)$value;
            }
        }

        return $values;
    }
}