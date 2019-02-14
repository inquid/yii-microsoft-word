<?php

namespace inquid\word;

use PhpOffice\PhpWord\TemplateProcessor;
use yii\base\Component;

/**
 * This is just an example.
 */
class TemplateExecuter extends Component
{
    public $templateFolder;
    public $templateName;
    public $saveFolder;

    /**
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     */
    public function execute($baseArray)
    {
        if ($this->templateFolder == null) {
            $this->templateFolder = \Yii::getAlias('@vendor/inquid/yii-microsoft-word/templates');
        }
        if ($this->templateName == null) {
            $this->templateName = 'Default.docx';
        }
        $templateProcessor = new TemplateProcessor("{$this->templateFolder}/{$this->templateName}");
        foreach ($baseArray as $item){
            if($this->array_depth($item)>1){
                $templateProcessor->setValue($item[0][0],$item[1][0]);
                $templateProcessor->setValue($item[0][1],$item[1][1]);
            }
            $templateProcessor->setValue($item[0],$item[1]);
        }
        return $templateProcessor;
    }

    /**
     * @param  $templateProcessor
     * @param $documentNamme
     * @return bool|string
     */
    public function saveDocument($templateProcessor, $documentNamme)
    {
        if ($this->saveFolder == null) {
            $this->saveFolder = \Yii::getAlias('files');
        }
        return \Yii::getAlias($templateProcessor->saveAs("{$this->saveFolder}/{$documentNamme}"));
    }

    function array_depth($array) {
        $max_indentation = 1;

        $array_str = print_r($array, true);
        $lines = explode("\n", $array_str);

        foreach ($lines as $line) {
            $indentation = (strlen($line) - strlen(ltrim($line))) / 4;

            if ($indentation > $max_indentation) {
                $max_indentation = $indentation;
            }
        }

        return ceil(($max_indentation - 1) / 2) + 1;
    }
}
