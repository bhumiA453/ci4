<?php
namespace App\Controllers;

use CodeIgniter\Controller;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
// use PhpOffice\PhpPresentation\Shape\Chart;
// use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;

class ExcelToPptController extends Controller
{
    public function index()
    {
        return view('upload_excel_to_ppt');
    }

    public function convert()
    {
        $excelFile = $this->request->getFile('excel_file');

        if ($excelFile->isValid() && $excelFile->getClientMimeType() === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            
            // Load the Excel file
            $spreadsheet = IOFactory::load($excelFile->getTempName());
            $worksheet = $spreadsheet->getActiveSheet();

            // Create a PowerPoint presentation
            $presentation = new PhpPresentation();
            
            foreach ($worksheet->getRowIterator() as $row) {
                $slide = $presentation->createSlide();

                // Create a text box on the slide
                $shape = $slide->createRichTextShape();
                $shape->setHeight(300);
                $shape->setWidth(600);

                // Set text alignment to left
                $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

                $textRun = $shape->createTextRun('');

                foreach ($row->getCellIterator() as $cell) {
                    $cellValue = $cell->getValue();

                    // Add cell value to the text box
                    $textRun = $shape->createTextRun($cellValue . "\t");

                    // Customize text formatting (if needed)
                    $textRun->getFont()->setColor(new Color('000000')); // Text color (black)
                    // Add more formatting as needed

                    // Add cell value to the text box
                    // $textRun->addText($cellValue . "\t", ['bold' => false]);
                }
            }
            /*// Create a new spreadsheet object and save it
            $newSpreadsheet = new Spreadsheet();
            $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx');
            $writer->save('public/converted_excel_file.xlsx');

            // Save the PowerPoint presentation
            $outputFile = 'public/presentation.pptx';
            $pptWriter = IOFactory::createWriter($presentation, 'PowerPoint2007');
            $pptWriter->save($outputFile);*/

            // Save the PowerPoint presentation
            $outputFile = 'public/presentation.pptx';
            \PhpOffice\PhpPresentation\IOFactory::createWriter($presentation, 'PowerPoint2007')->save($outputFile);

            /*$pptWriter = new PhpOffice\PhpPresentation\Writer\PowerPoint2007($presentation);
            $pptWriter->save($outputFile);*/


            echo "Conversion complete. PowerPoint presentation saved as: $outputFile";
        } else {
            echo 'Invalid file format. Please upload a valid Excel file.';
        }
    }
}
