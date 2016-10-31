<?php


namespace lvincesl\query;

/**
 * Classe de gestion d'exportation de requêtes SQL vers Excel
 *
 * @category Template
 * @package  Lvincesl
 * @author   Lionel Vinceslas <lionel.vinceslas@lapsote.net>
 * @license  CECILL http://www.cecill.info/licences/Licence_CeCILL-C_V1-fr.txt
 * @link     lionel-vinceslas.eurower.net
 *
 * @date 25/10/2016
 * @return string
 */
class Query_to_excel
{
    var $pdo;
    var $query;
    var $csv_ouput_filename     = 'output.csv';
    var $excel_ouput_filename   = 'output.xls';
    

    public function __construct($pdo, $query=null)
    {
        $this->pdo = $pdo;
        if (!is_null($this->query)) {
            $this->query = $query;
        }
    }

    public function toString()
    {
        return null;
    }

    public function setQuery($query)
    {
        $this->query = $query;
    }

    public function setCsvOutputFilename($csv_output_filename)
    {
        $this->csv_output_filename = $csv_output_filename;
    }

    public function setExcelOutputFilename($excel_output_filename)
    {
        $this->excel_output_filename = $excel_output_filename;
    }

    public function getQuery()
    {
        return $this->query;
    }

    public function getCsvOutputFilename()
    {
        return $this->csv_output_filename;
    }

    public function getExcelOutputFilename()
    {
        return $this->excel_output_filename;
    }

    public function toExcel()
    {
        return null;
    }

    public function toCsv()
    {
        $result = $this->pdo->query($query);
        $sep    = ";";
      

        header("Content-Type: text/csv; charset=UTF-8");
        header("Content-Disposition: attachment; filename=".$this->csv_output_filename);
        header("Pragma: no-cache");
        header("Expires: 0");

        for ($i = 0; $i <  $result->columnCount(); $i++) {
            echo utf8_decode($result->getColumnMeta($i)['name']) . $sep;
        }

        print("\n");

        while($row = $result->fetch())
        {
            $schema_insert = "";
            for($j=0; $j<$result->columnCount();$j++)
            {
                if(!isset($row[$j]))
                    $schema_insert .= "".$sep;
                elseif ($row[$j] != "")
                    $schema_insert .= str_replace(';', ' , ', html_entity_decode($row[$j])).$sep;
                else
                    $schema_insert .= "".$sep;

            }

            $schema_insert = str_replace($sep."$", "", $schema_insert);

            $schema_insert = preg_replace("/\r\n|\n\r|\n|\r/", " ", $schema_insert);

            $schema_insert .= "\t";
            print(utf8_decode(trim($schema_insert)));
            print "\n";
        }

        exit(0);
        }
    }
}

?>
