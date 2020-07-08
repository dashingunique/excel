<?php


namespace dashingUnique\excel\traits;


use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\WriterInterface;
use dashingUnique\library\Collection;
use Generator;
use Box\Spout\Writer\XLSX\Writer as XLSXWriter;
use Box\Spout\Writer\ODS\Writer as ODSWriter;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use InvalidArgumentException;

/**
 * 导出Excel（csv,ods）
 * Trait Exportable
 * @package dashingUnique\excel\traits
 * @property bool $withHeader
 * @property Collection $data
 */
trait Exportable
{
    /**
     * @var Style 头部样式
     */
    private $headerStyle;

    /**
     * @var array 备注信息 => ['备注:备注']
     */
    protected $remark;

    /**
     * 获取文件类型
     * @param string $path
     *
     * @return string
     */
    abstract protected function getType($path);

    /**
     * 设置读写信息
     * @param ReaderInterface|WriterInterface $readerOrWriter
     * @return mixed
     */
    abstract protected function setOptions(&$readerOrWriter);

    /**
     * 导出Excel（csv、oda）信息
     * @param string $path
     * @param callable|null $callback
     *
     * @return string
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException*@throws \Box\Spout\Common\Exception\SpoutException
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\SpoutException
     */
    public function export($path, callable $callback = null)
    {
        self::exportOrDownload($path, 'openToFile', $callback);
        return realpath($path) ?: $path;
    }

    /**
     * 下载内容
     * @param $path
     * @param callable|null $callback
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\SpoutException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function download($path, callable $callback = null)
    {
        self::exportOrDownload($path, 'openToBrowser', $callback);
        return '';
    }

    /**
     * 设置备注信息
     * @param string $remark
     * @return $this
     */
    public function remark(string $remark)
    {
        $this->remark = $remark;
        return $this;
    }


    /**
     * 导出并下载Excel（csv,odb）
     * @param $path
     * @param string $function
     * @param callable|null $callback
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Common\Exception\SpoutException
     */
    private function exportOrDownload($path, $function, callable $callback = null)
    {
        $writer = WriterFactory::create($this->getType($path));
        $this->setOptions($writer);
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer->$function($path);

        $hasSheets = ($writer instanceof XLSXWriter || $writer instanceof ODSWriter);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->data instanceof Collection ? $this->data : uniqueCollection([$this->data]);

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                $this->writeRowsFromCollection($writer, $collection, $callback);
            } elseif ($collection instanceof Generator) {
                $this->writeRowsFromGenerator($writer, $collection, $callback);
            } elseif (is_array($collection)) {
                $this->writeRowsFromArray($writer, $collection, $callback);
            } else {
                throw new InvalidArgumentException('Unsupported type for $data');
            }
            if (is_string($key)) {
                $writer->getCurrentSheet()->setName($key);
            }
            if ($hasSheets && $data->keys()->last() !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $this->writeRemark($writer);
        $writer->close();
    }

    /**
     * 写入表格信息（通过集合）
     * @param WriterInterface $writer
     * @param Collection $collection
     * @param callable|null $callback
     */
    private function writeRowsFromCollection($writer, Collection $collection, ?callable $callback = null)
    {
        // Apply callback
        if ($callback) {
            $collection->transform(function ($value) use ($callback) {
                return $callback($value);
            });
        }
        // Prepare collection (i.e remove non-string)
        $this->prepareCollection($collection);
        // Add header row.
        if ($this->withHeader) {
            $this->writeHeader($writer, $collection->first());
        }
        // Write all rows
        $writer->addRows($collection->toArray());
    }

    /**
     * 写入备注信息
     * @param $writer
     */
    private function writeRemark($writer)
    {
        if (!empty($this->remark)) {
            // Write all rows
            $writer->addRow($this->remark);
        }
    }

    /**
     * 写入数据（通过生成器）
     * @param $writer
     * @param Generator $generator
     * @param callable|null $callback
     */
    private function writeRowsFromGenerator($writer, Generator $generator, ?callable $callback = null)
    {
        foreach ($generator as $key => $item) {
            // Apply callback
            if ($callback) {
                $item = $callback($item);
            }

            // Prepare row (i.e remove non-string)
            $item = $this->transformRow($item);

            // Add header row.
            if ($this->withHeader && $key === 0) {
                $this->writeHeader($writer, $item);
            }
            // Write rows (one by one).
            $writer->addRow($item->toArray());
        }
    }



    /**
     * 写入信息（通过数组）
     * @param $writer
     * @param array $array
     * @param callable|null $callback
     */
    private function writeRowsFromArray($writer, array $array, ?callable $callback = null)
    {
        $collection = uniqueCollection($array);

        if (is_object($collection->first()) || is_array($collection->first())) {
            // provided $array was valid and could be converted to a collection
            $this->writeRowsFromCollection($writer, $collection, $callback);
        }
    }



    /**
     * 写入header
     * @param $writer
     * @param $first_row
     */
    private function writeHeader($writer, $first_row)
    {
        if ($first_row === null) {
            return;
        }
        $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
        if ($this->headerStyle) {
            $writer->addRowWithStyle($keys, $this->headerStyle);
        } else {
            $writer->addRow($keys);
        }
    }

    /**
     * 如果需要，通过除去非字符串来准备收集。
     * @param Collection $collection
     */
    protected function prepareCollection(Collection $collection)
    {
        $need_conversion = false;
        $first_row = $collection->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->transform($collection);
        }
    }

    /**
     * 转换集合
     * @param Collection $collection
     */
    private function transform(Collection $collection)
    {
        $collection->transform(function ($data) {
            return $this->transformRow($data);
        });
    }

    /**
     * 转换一行（即删除非字符串）
     * @param $data
     * @return Collection
     */
    private function transformRow($data)
    {
        return uniqueCollection($data)->map(function ($value) {
            return is_null($value) ? (string)$value : $value;
        })->filter(function ($value) {
            return is_string($value) || is_int($value) || is_float($value);
        });
    }

    /**
     * 设置header 头样式
     * @param Style $style
     *
     * @return Exportable
     */
    public function headerStyle(Style $style)
    {
        $this->headerStyle = $style;
        return $this;
    }
}