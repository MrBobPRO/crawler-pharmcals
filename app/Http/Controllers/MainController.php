<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MainController extends Controller
{
    public function index()
    {
        $fromPage = 301;
        $toPage = 400;

        for ($i = $fromPage; $i <= $toPage; $i++) {
            $response = Http::withBasicAuth('Ortos', 'Ortos2023')->get('http://pharm.cals.am/pharm/report/get_data.php', [
                'prand' => (float)rand() / (float)getrandmax(),
                'pbtn' => 'search',
                'pdate1' => '01-01-2018',
                'pdate2' => '08-03-2023',
                'pname' => '',
                'pgeneric' => '',
                'pdosform' => '',
                'pcountry' => '',
                'pmanuf' => '',
                'ptype' => '1',
                'ppage' => $i,
                'psid' => '1822168846',
            ]);

            if ($response->successful()) {
                $result = json_decode($response->body());

                // preapare excel reader/writer
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(public_path('results.xlsx'));
                $sheet = $spreadsheet->getActiveSheet();

                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                $reader->setReadDataOnly(true);

                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
                $writer->setPreCalculateFormulas(false);

                for ($j = 0; $j < count($result->items); $j++) {
                    $drug = $this->validateDrug($result->items[$j]->caption);
                    $count = $result->items[$j]->count;

                    $highestRow = $sheet->getHighestRow() + 1;
                    $sheet->setCellValue('A' . $highestRow, $drug);
                    $sheet->setCellValue('B' . $highestRow, $count);
                }

                // Save generated file
                $writer->save(public_path('results.xlsx'));
            } else {
                dd('Error while parsing page' . ($i + 1));
            }
        }
        return 'Success!';
    }

    private function validateDrug($string): string
    {
        // replace <br> with blanks
        $string = str_replace('<br>', ' ', $string);

        // strip tags
        $string = strip_tags($string);

        // remove whitespaces
        $string = Str::squish($string);

        return $string;
    }
}
