<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use setasign\Fpdi\Fpdi;

require_once 'vendor/autoload.php';

//Initialization of paths
$diploma_path = "diploma/DiplomaMedici.pdf";
$excel_path = "excel/ListaMediciPrezenta.xlsx";
$output_folder = "PDF_Output";

// Calea către fișierul PDF existent
$existing_pdf_path = $output_folder . "/diploma.pdf";

// Verifică dacă există un fișier PDF existent și șterge-l
if (file_exists($existing_pdf_path)) {
    unlink($existing_pdf_path);
}

// Încarcă fișierul Excel
$spreadsheet = IOFactory::load($excel_path);
$sheet = $spreadsheet->getActiveSheet();

// Obține numele persoanelor din coloana A
$nameColumn = 'B';
$emailColumn = 'E';
$highestRow = $sheet->getHighestRow();

for ($row = 2; $row <= $highestRow; $row++) {
    $emailValue= $sheet->getCell($emailColumn . $row)->getValue();
    $nameValue = $sheet->getCell($nameColumn . $row)->getValue();
    if (!empty($nameValue)) {

        //Replace diacritics
        $nameValue = str_replace('Ă', 'a', $nameValue);
        $nameValue = str_replace('Â', 'a', $nameValue);
        $nameValue = str_replace('Î', 'i', $nameValue);
        $nameValue = str_replace('Ș', 's', $nameValue);
        $nameValue = str_replace('Ț', 't', $nameValue);
        $nameValue = str_replace('ă', 'a', $nameValue);
        $nameValue = str_replace('â', 'a', $nameValue);
        $nameValue = str_replace('î', 'i', $nameValue);
        $nameValue = str_replace('ș', 's', $nameValue);
        $nameValue = str_replace('ț', 't', $nameValue);

        //Create new pdf
        $pdf = new Fpdi();

        //Copy pdf template
        $pdf->setSourceFile($diploma_path);
        $templateId = $pdf->importPage(1);

        $pdf->AddPage();

        $pdf->useTemplate($templateId, ['adjustPageSize' => true]);

        //index for diploma number
        $index = $row - 1;

        //Edit font and add text
        //Name
        $pdf->SetFont('Times');
        $pdf->SetFontSize(14);
        $pdf->SetTextColor(0, 0, 0);
        $pdf->SetXY( 77, 88.5);
        $pdf->Write(0, $nameValue);

        //Diploma Number
        //$pdf->SetXY(260, 8.6);
       // $pdf->Write(0, $index);

        //testing on browser
        //$pdf->Output();

        //save pdf in path end close the pdf
        if(!empty($emailValue))
        {
            $pdf_path = $output_folder . "/" . $emailValue . ".pdf";
        } else {
            $pdf_path = $output_folder . "/" . $nameValue . ".pdf";
        }

        $pdf->Output($pdf_path, 'F');
        //$pdf->Output();
        $pdf->close();
        echo $index;
        echo "\n";

    } else {
        break;
    }
}


exit();



