<?php
/**
 * Created by PhpStorm.
 * User: mrShes
 * Date: 30.01.2018
 * Time: 14:58
 */

namespace App\Http\Controllers;

use PhpOffice\PhpWord\IOFactory;

class PhpWordExtend
{

    protected $phpWord;
    protected $path;
    protected $extension = [
        'docx',
        'doc'
    ];

    public function __construct($path)
    {
        $this->path = $path;
    }

    /**
     * @return mixed
     */
    public function getPhpWord()
    {
        $file = pathinfo($this->path);
        $extension = $file['extension'];
        if (in_array($extension, $this->extension)) {
            return $this->phpWord = IOFactory::load($this->path);
        }
    }


    public function get()
    {
        $phpWord = $this->getPhpWord();
        if ($phpWord) {
            $response = '';
            $sections = $phpWord->getSections();
            foreach ($sections as $key => $section) {
                $sectionElements = $section->getElements();
                foreach ($sectionElements as $elementKey => $element) {
                    $response .= $this->getTextFromElement($element);
                }
            }
            return $response;
        }
    }


    protected function shift()
    {
        return '<br>\n';
    }

    protected function getClassName($class)
    {
        $class_name = get_class($class);
        $pattern = '/(\w*$)/';
        $get_name = preg_match($pattern, $class_name, $name);
        return $name[1];
    }


    function getTextFromElement($element)
    {
        $text = '';
        $name = $this->getClassName($element);
        switch ($name) {
            case 'Text':
                $text .= $element->getText();
                break;
            case 'TextRun':
                $textRunElements = $element->getElements();
                if (!empty($textRunElements)) {
                    foreach ($textRunElements as $textRunElement) {
                        $text .= $this->getTextFromElement($textRunElement);
                        $text .= $this->shift();
                    }
                }
                break;
            case 'Table':
                $rows = $element->getRows();
                foreach ($rows as $row) {
                    $cells = $row->getCells();
                    foreach ($cells as $cell) {
                        $cell_elements = $cell->getElements();
                        foreach ($cell_elements as $cell_element) {
                            $cell_text = $this->getTextFromElement($cell_element);
                            $text .= $cell_text;
                        }
                    }
                    if (!empty($cell_text)) {
                        $text .= $this->shift();
                    }
                }
                break;
            case 'TextBreak':
//                //
                break;
        }
        return $text;
    }
}