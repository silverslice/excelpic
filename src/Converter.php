<?php

namespace Silverslice\ExcelPic;

class Converter
{
    /** @var Zip */
    protected $zip;

    /**
     * @var string Temporary directory for unzipped files
     */
    protected $tempDir;

    public function __construct()
    {
        $this->zip = new Zip();
    }

    /**
     * Opens xlsx file
     *
     * @param $file
     *
     * @return $this
     *
     * @throws \Exception
     */
    public function open($file)
    {
        if (!is_file($file)) {
            throw new \Exception("File $file not found");
        }

        $dir = $this->getTempDir();
        $this->zip->extract($file, $dir);

        return $this;
    }

    /**
     * Converts all pictures to movable and resizable
     *
     * @return $this
     */
    public function convertImagesToResizable()
    {
        $file = $this->tempDir . '/xl/drawings/drawing1.xml';

        $dom = new \DOMDocument();
        $dom->load($file);
        $documentElement = $dom->documentElement;

        $delete = [];
        $list = $dom->getElementsByTagName('oneCellAnchor');
        foreach ($list as $oneCellAnchor) {
            // check if it is picture
            $pic = $oneCellAnchor->getElementsByTagName('pic');
            if ($pic->length) {
                $this->processDrawing($oneCellAnchor);
                $delete[] = $oneCellAnchor;
            }
        }

        foreach ($delete as $node) {
            $documentElement->removeChild($node);
        }

        $dom->save($file);

        return $this;
    }

    /**
     * Saves file
     *
     * @param $filename
     *
     * @return bool
     */
    public function save($filename)
    {
        $res = $this->zip->archive($this->tempDir, $filename);
        (new FileHelper())->removeDirectory($this->tempDir);

        return $res;
    }

    /**
     * Sets temporary directory
     *
     * @param string $dir
     *
     * @return $this
     *
     * @throws \Exception
     */
    public function setTempDirectory($dir)
    {
        if (!is_dir($dir)) {
            throw new \Exception("Directory $dir not found");
        }

        $this->tempDir = $dir;

        return $this;
    }

    protected function getTempDir()
    {
        if (!isset($this->tempDir)) {
            $this->tempDir = sys_get_temp_dir() . '/' . uniqid('xlsx');
        }

        return $this->tempDir;
    }

    protected function copyChildNodes(\DOMNode $from, \DOMNode $to)
    {
        foreach ($from->childNodes as $child) {
            $node = $child->cloneNode(true);
            $to->appendChild($node);
        }
        foreach ($from->attributes as $name => $node) {
            $to->setAttribute($name, $node);
        }

        return true;
    }

    /**
     * Replace oneCellAnchor element to twoCellAnchor element
     *
     * @link http://officeopenxml.com/drwPicInSpread-twoCell.php
     *
     * @param \DOMElement $oneCellAnchor
     */
    protected function processDrawing(\DOMElement $oneCellAnchor)
    {
        $dom = $oneCellAnchor->ownerDocument;

        $twoCellAnchor = $dom->createElement('xdr:twoCellAnchor');
        $this->copyChildNodes($oneCellAnchor, $twoCellAnchor);
        $this->createToElement($twoCellAnchor);

        $dom->documentElement->appendChild($twoCellAnchor);
    }

    protected function createToElement(\DOMElement $node)
    {
        $from = $node->getElementsByTagName('from')->item(0);
        $ext  = $node->getElementsByTagName('ext')->item(0);
        $dom = $node->ownerDocument;

        $data = [
            'col'    => $from->childNodes->item(0)->nodeValue,
            'colOff' => $ext->getAttribute('cx'),
            'row'    => $from->childNodes->item(2)->nodeValue,
            'rowOff' => $ext->getAttribute('cy'),
        ];

        $to = $dom->createElement('xdr:to');
        $to->appendChild($dom->createElement('xdr:col', $data['col']));
        $to->appendChild($dom->createElement('xdr:colOff', $data['colOff']));
        $to->appendChild($dom->createElement('xdr:row', $data['row']));
        $to->appendChild($dom->createElement('xdr:rowOff', $data['rowOff']));

        $node->insertBefore($to, $from->nextSibling);
        $node->removeChild($ext);
    }
}
