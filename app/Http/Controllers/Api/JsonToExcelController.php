<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class JsonToExcelController extends Controller
{
    //
    public function test_page(Request $request)
    {
        $data = $request->all();
        // var_dump($data);
        // var_dump(json_decode($data['data']));
        // var_dump(json_decode($data['data1']));
        // die;

        $arr1 = json_decode($data['data']);
        $arr2 = json_decode($data['data1']);

        $arrarge = array_merge($arr1, $arr2);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        // Header

        $sheet->setCellValue('A1', 'Kondisi');
        $sheet->setCellValue('B1', 'Januari');
        $sheet->setCellValue('C1', 'Februari');
        $sheet->setCellValue('D1', 'Maret');
        $sheet->setCellValue('E1', 'April');
        $sheet->setCellValue('F1', 'Mei');
        $sheet->setCellValue('G1', 'Juni');
        $sheet->setCellValue('H1', 'Juli');
        $sheet->setCellValue('I1', 'Agustus');
        $sheet->setCellValue('J1', 'September');
        $sheet->setCellValue('K1', 'Oktober');
        $sheet->setCellValue('L1', 'November');
        $sheet->setCellValue('M1', 'Desember');

        $count = 2;

        foreach ($arrarge as $value) {
            // dd($value);
            // foreach ($value as $x) {
            $sheet->setCellValue('A' . $count, $value->condition);
            $sheet->setCellValue('B' . $count, $value->Januari);
            $sheet->setCellValue('C' . $count, $value->Februari);
            $sheet->setCellValue('D' . $count, $value->Maret);
            $sheet->setCellValue('E' . $count, $value->April);
            $sheet->setCellValue('F' . $count, $value->Mei);
            $sheet->setCellValue('G' . $count, $value->Juni);
            $sheet->setCellValue('H' . $count, $value->Juli);
            $sheet->setCellValue('I' . $count, $value->Agustus);
            $sheet->setCellValue('J' . $count, $value->September);
            $sheet->setCellValue('K' . $count, $value->Oktober);
            $sheet->setCellValue('L' . $count, $value->November);
            $sheet->setCellValue('M' . $count, $value->Desember);
            // }

            $count = $count + 1;
        }

        // $sheet->setCellValue('A1', 'Hello World !');

        $fileName = "cek.xlsx";

        $writer = new Xlsx($spreadsheet);
        $writer->save('hello_world.xlsx');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode($fileName) . '"');
        $writer->save('php://output');
    }
}
