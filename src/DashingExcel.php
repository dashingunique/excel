<?php

namespace dashingUnique\excel;

use Box\Spout\Common\Type;
use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use Box\Spout\Writer\WriterInterface;
use dashingUnique\excel\traits\Exportable;
use dashingUnique\excel\traits\ImportData;
use dashingUnique\library\Collection;
use dashingUnique\library\Str;

/**
 * Excel操作类
 * Class DashingExcel
 * @package dashingUnique\excel
 */
class DashingExcel
{
    use ImportData;
    use Exportable;

    /**
     * @var Collection
     */
    protected $data;

    /**
     * @var bool
     */
    protected $withHeader = true;

    /**
     * @var
     */
    private $csvConfig = [
        'delimiter' => ',',
        'enclosure' => '"',
        'eol' => "\n",
        'encoding' => 'UTF-8',
        'bom' => true,
    ];

    /**
     * @var callable
     */
    protected $readerConfigurator = null;

    /**
     * @var callable
     */
    protected $writerConfigurator = null;

    /**
     * DashingExcel constructor.
     * @param Collection $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
    }

    /**
     * 设置数据信息
     * @param $data
     * @return $this
     */
    public function data($data)
    {
        $this->data = $data;
        return $this;
    }

    /**
     * 获取文件后缀名
     * @param $path
     * @return string
     */
    protected function getType($path)
    {
        if (Str::endsWith($path, Type::CSV)) {
            return Type::CSV;
        } elseif (Str::endsWith($path, Type::ODS)) {
            return Type::ODS;
        } else {
            return Type::XLSX;
        }
    }

    /**
     * 设置单元格数量
     * @param $sheetNumber
     * @return $this
     */
    public function sheet($sheetNumber)
    {
        $this->sheetNumber = $sheetNumber;

        return $this;
    }

    /**
     * 不需要header信息
     * @return $this
     */
    public function withoutHeaders()
    {
        $this->withHeader = false;
        return $this;
    }

    /**
     * 配置CSV信息
     * @param string $delimiter
     * @param string $enclosure
     * @param string $eol
     * @param string $encoding
     * @param bool $bom
     *
     * @return $this
     */
    public function configureCsv($delimiter = ',', $enclosure = '"', $eol = "\n", $encoding = 'UTF-8', $bom = false)
    {
        $this->csvConfig = compact('delimiter', 'enclosure', 'eol', 'encoding', 'bom');

        return $this;
    }

    /**
     * 使用回调配置基础的读取阅读器。
     *
     * @param callable $callback
     *
     * @return $this
     */
    public function configureReaderUsing(?callable $callback = null)
    {
        $this->readerConfigurator = $callback;

        return $this;
    }

    /**
     * 使用回调配置基础的写入阅读器。
     * @param callable $callback
     *
     * @return $this
     */
    public function configureWriterUsing(?callable $callback = null)
    {
        $this->writerConfigurator = $callback;
        return $this;
    }

    /**
     * 设置读写信息
     * @param ReaderInterface|WriterInterface $readerOrWriter
     */
    protected function setOptions(&$readerOrWriter)
    {
        if ($readerOrWriter instanceof CSVReader || $readerOrWriter instanceof CSVWriter) {
            $readerOrWriter->setFieldDelimiter($this->csvConfig['delimiter']);
            $readerOrWriter->setFieldEnclosure($this->csvConfig['enclosure']);
            if ($readerOrWriter instanceof CSVReader) {
                $readerOrWriter->setEndOfLineCharacter($this->csvConfig['eol']);
                $readerOrWriter->setEncoding($this->csvConfig['encoding']);
            }
            if ($readerOrWriter instanceof CSVWriter) {
                $readerOrWriter->setShouldAddBOM($this->csvConfig['bom']);
            }
        }

        if ($readerOrWriter instanceof ReaderInterface && is_callable($this->readerConfigurator)) {
            call_user_func(
                $this->readerConfigurator,
                $readerOrWriter
            );
        } elseif ($readerOrWriter instanceof WriterInterface && is_callable($this->writerConfigurator)) {
            call_user_func(
                $this->writerConfigurator,
                $readerOrWriter
            );
        }
    }
}