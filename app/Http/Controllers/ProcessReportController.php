<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Config;
use Illuminate\Support\Facades\Schema;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class ProcessReportController extends Controller
{
    public function processReport($tableName)
    {

        $query = DB::connection('sqlsrv')->table('ALLINOFFERSKULIST as a')
            ->select(
                'a.VENDOR_CODE',
                'a.DESCRIPTION',
                'a.PART_NUMBER',
                'a.ITEM_DESC',
                'a.TIRESIZE_WIDTH',
                'a.TIRESIZE_RATIO',
                'a.TIRESIZE_WHEEL',
                'a.LEGACY_PART_REFERENCE',
                'a.WEIGHT',
                'a.TIRE_CATALOGSIZE',
                'a.FET',
                DB::raw('a.TIRESIZE_WIDTH + a.TIRESIZE_RATIO + a.TIRESIZE_WHEEL as sfullsize'),
                DB::raw('COALESCE(v.QTY_AVAIL, 0) as Vernon_AVAIL'),
                DB::raw('COALESCE(l.QTY_AVAIL, 0) as Latrobe_AVAIL'),
                DB::raw('COALESCE(x.QTY_AVAIL, 0) as KENTUCKY_AVAIL')
            )
            ->leftJoin('VERNON_AVAIL as v', function ($join) {
                $join->on('a.vendor_code', '=', 'v.vendor_code')
                    ->on('a.part_number', '=', 'v.part_number');
            })
            ->leftJoin('LATROBE_AVAIL as l', function ($join) {
                $join->on('a.vendor_code', '=', 'l.vendor_code')
                    ->on('a.part_number', '=', 'l.part_number');
            })
            ->leftJoin('KENTUCKY_AVAIL as x', function ($join) {
                $join->on('a.vendor_code', '=', 'x.vendor_code')
                    ->on('a.part_number', '=', 'x.part_number');
            })
            ->orderBy('a.TIRESIZE_WHEEL', 'asc')
            ->orderBy('a.TIRESIZE_WIDTH', 'asc')
            ->orderBy('a.TIRESIZE_RATIO', 'asc')
            ->orderBy('Vernon_AVAIL', 'desc')
            ->get();

        // dd($query);

        $query2 = DB::table('ALLINOFFERSKULIST as a')
            ->select(
                'a.VENDOR_CODE',
                'a.DESCRIPTION',
                'a.PART_NUMBER',
                'a.ITEM_DESC',
                'a.TIRESIZE_WIDTH',
                'a.TIRESIZE_RATIO',
                'a.TIRESIZE_WHEEL',
                'a.LEGACY_PART_REFERENCE',
                'a.WEIGHT',
                'a.TIRE_CATALOGSIZE',
                'a.FET',
                DB::raw('a.TIRESIZE_WIDTH + a.TIRESIZE_RATIO + a.TIRESIZE_WHEEL as sfullsize'),
                DB::raw('COALESCE(v.QTY_AVAIL, 0) as Vernon_AVAIL'),
                DB::raw('COALESCE(l.QTY_AVAIL, 0) as Latrobe_AVAIL'),
                DB::raw('COALESCE(x.QTY_AVAIL, 0) as KENTUCKY_AVAIL')
            )
            ->leftJoin('VERNON_AVAIL as v', function ($join) {
                $join->on('a.vendor_code', '=', 'v.vendor_code')
                    ->on('a.part_number', '=', 'v.part_number');
            })
            ->leftJoin('LATROBE_AVAIL as l', function ($join) {
                $join->on('a.vendor_code', '=', 'l.vendor_code')
                    ->on('a.part_number', '=', 'l.part_number');
            })
            ->leftJoin('KENTUCKY_AVAIL as x', function ($join) {
                $join->on('a.vendor_code', '=', 'x.vendor_code')
                    ->on('a.part_number', '=', 'x.part_number');
            })
            ->orderBy('a.VENDOR_CODE', 'asc')
            ->orderBy('a.TIRESIZE_WHEEL', 'asc')
            ->orderBy('a.TIRESIZE_WIDTH', 'asc')
            ->orderBy('a.TIRESIZE_RATIO', 'asc')
            ->orderBy('Vernon_AVAIL', 'desc')
            ->get();

        ///EXCEL

        $spreadsheet = new Spreadsheet();
        $border_thick = [
            'font' => [
                'bold' => true,
            ],

            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                    'color' => ['argb' => '000000'],

                ],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],



        ];
        $border_thin = [

            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '333333'],

                ],
            ]

        ];
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('CA, PA & KY BySize');
        $sheet->getPageSetup()->setScale(65);
        $spreadsheet->getActiveSheet()->getSheetView()->setZoomScale(65);

        $sheet->getPageMargins()->setTop(0);
        $sheet->getPageMargins()->setRight(0);
        $sheet->getPageMargins()->setLeft(0);
        $sheet->getPageMargins()->setBottom(0);

        $sheet->mergeCells('A1:A3');
        $sheet->mergeCells('B1:D1');
        $sheet->mergeCells('C2:D2');
        $sheet->mergeCells('C3:D3');
        $sheet->mergeCells('C4:D4');
        $sheet->mergeCells('E2:I2');
        $sheet->mergeCells('E3:I3');
        $sheet->mergeCells('E4:I4');
        $sheet->mergeCells('I1:J1');
        $sheet->mergeCells('E1:H1');
        $sheet->freezePane('A6');
        $sheet->getColumnDimension('W')->setWidth(0);

        $sheet->mergeCells('N1:O1');
        $sheet->mergeCells('Q2:T2');
        $sheet->mergeCells('Q3:T3');
        $sheet->mergeCells('Q4:T4');
        $sheet->mergeCells('T1:U1');
        $sheet->mergeCells('Q1:S1');

        // KENTUCKY
        $sheet->mergeCells('Z1:AA1');
        $sheet->mergeCells('AC2:AF2');
        $sheet->mergeCells('AC3:AF3');
        $sheet->mergeCells('AC4:AF4');
        $sheet->mergeCells('AF1:AG1');
        $sheet->mergeCells('AC1:AE1');



        // $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        // $drawing->setName('WTD - Wholesale Tire Distributors');
        // $drawing->setDescription('Wholesale Tire Distributors');
        // // $drawing->setPath('wtd-logo.jpg'); // put your path and image here
        // $drawing->setCoordinates('A1');
        // $drawing->setResizeProportional(true);
        // $drawing->setWidth(120);
        // $drawing->setHeight(120);
        // $drawing->setOffsetX(10);
        // $drawing->getShadow()->setVisible(false);
        // $drawing->setWorksheet($sheet);

        $sheet
            ->getStyle('B2:J4')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FFFFFF');
        $sheet
            ->getStyle('B1:I1')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FFFFFF');

        $styleArray = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  13,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000']
                ]
            ]
        ];
        $styleArrayTopNum = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
        ];
        $styleArray2 = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  13,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000']
                ]
            ]
        ];
        $styleArrayTopTotal = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
        ];
        $styleHeader = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => array('argb' => 'b8cce4')
            ]
        ];

        $styleHeader_wtd = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => array('argb' => 'b8cce4')
            ]
        ];

        $styleTransparent = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  10,
                'color' => array('rgb' => 'FF0000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ]
        ];

        $sheet->getStyle('A1:A3')->applyFromArray($styleHeader_wtd);
        $sheet->getStyle('B2:D4')->applyFromArray($styleArray2);
        $sheet->getStyle('E2:J2')->applyFromArray($styleArray);
        $sheet->getStyle('E3:J3')->applyFromArray($styleArray);
        $sheet->getStyle('E4:J4')->applyFromArray($styleArray);
        $sheet->getStyle('B1:C1')->applyFromArray($styleArrayTopNum);
        $sheet->getStyle('E1:J1')->applyFromArray($styleArrayTopTotal);
        $sheet->getStyle('A5:M5')->applyFromArray($styleHeader);
        $sheet->getStyle('K2')->applyFromArray($styleTransparent);
        $sheet->getStyle('K3')->applyFromArray($styleTransparent);
        $sheet->getStyle("I1:J1")->getFont()->setSize(18);

        $sheet->getStyle('O1:J1')->applyFromArray($styleArrayTopTotal);
        $sheet->getStyle('N2:O4')->applyFromArray($styleArray2);
        $sheet->getStyle('P2:U2')->applyFromArray($styleArray);
        $sheet->getStyle('P3:U3')->applyFromArray($styleArray);
        $sheet->getStyle('P4:U4')->applyFromArray($styleArray);
        $sheet->getStyle('N1:U1')->applyFromArray($styleArrayTopNum);
        $sheet->getStyle('V2')->applyFromArray($styleTransparent);
        $sheet->getStyle('V3')->applyFromArray($styleTransparent);
        $sheet->getStyle("T1:U1")->getFont()->setSize(18);

        // KENTUCKY
        $sheet->getStyle('Z1:AA1')->applyFromArray($styleArrayTopTotal);
        $sheet->getStyle('Z2:AA4')->applyFromArray($styleArray2);
        $sheet->getStyle('AB2:AG2')->applyFromArray($styleArray);
        $sheet->getStyle('AB3:AG3')->applyFromArray($styleArray);
        $sheet->getStyle('AB4:AG4')->applyFromArray($styleArray);
        $sheet->getStyle('Y1:AG1')->applyFromArray($styleArrayTopNum);
        $sheet->getStyle('AH2')->applyFromArray($styleTransparent);
        $sheet->getStyle('AH3')->applyFromArray($styleTransparent);
        $sheet->getStyle("AF1:AG1")->getFont()->setSize(18);


        $sheet->getStyle('P5:V5')->applyFromArray($styleHeader);

        // KENTUCKY
        $sheet->getStyle('Z5:AH5')->applyFromArray($styleHeader);


        $sheet->getRowDimension('1')->setRowHeight(23.25);
        $sheet->getRowDimension('2')->setRowHeight(43.5);
        $sheet->getRowDimension('3')->setRowHeight(43.5);
        $sheet->getRowDimension('4')->setRowHeight(23.25);
        $sheet->getRowDimension('5')->setRowHeight(45);

        $sheet->getColumnDimension('A')->setWidth(21.86);
        $sheet->getColumnDimension('B')->setWidth(21.43);
        $sheet->getColumnDimension('C')->setWidth(21.43);
        $sheet->getColumnDimension('D')->setWidth(15.86);
        $sheet->getColumnDimension('E')->setWidth(30.86);
        $sheet->getColumnDimension('F')->setWidth(14.43);
        $sheet->getColumnDimension('G')->setWidth(11.71);
        $sheet->getColumnDimension('H')->setWidth(20.43);
        $sheet->getColumnDimension('I')->setAutoSize(true);
        $sheet->getColumnDimension('J')->setWidth(15.86);
        $sheet->getColumnDimension('K')->setWidth(26.23);
        $sheet->getColumnDimension('L')->setWidth(11.71);
        $sheet->getColumnDimension('M')->setWidth(15.71);


        // $sheet->getColumnDimension('N')->setWidth(0);
        $sheet->getColumnDimension('O')->setWidth(0);
        $sheet->getColumnDimension('P')->setWidth(0);
        $sheet->getColumnDimension('Q')->setWidth(12.86);
        $sheet->getColumnDimension('R')->setWidth(12.43);
        $sheet->getColumnDimension('S')->setWidth(11.71);
        $sheet->getColumnDimension('T')->setWidth(12.43);
        $sheet->getColumnDimension('U')->setAutoSize(true);
        $sheet->getColumnDimension('V')->setWidth(26.23);
        $sheet->getColumnDimension('W')->setWidth(0);
        $sheet->getColumnDimension('X')->setWidth(0);

        // KENTUCKY
        $sheet->getColumnDimension('Z')->setWidth(0);
        $sheet->getColumnDimension('AA')->setWidth(0);
        $sheet->getColumnDimension('AB')->setWidth(0);
        $sheet->getColumnDimension('AC')->setWidth(12.86);
        $sheet->getColumnDimension('AD')->setWidth(12.43);
        $sheet->getColumnDimension('AE')->setWidth(11.71);
        $sheet->getColumnDimension('AF')->setWidth(12.43);
        $sheet->getColumnDimension('AG')->setAutoSize(true);
        $sheet->getColumnDimension('AH')->setWidth(26.23);
        $sheet->getColumnDimension('AI')->setWidth(0);
        $sheet->getColumnDimension('AJ')->setWidth(0);

        $sheet->setCellValue('A1', 'WTD Wholesale Tire Distributors');
        $sheet->setCellValue('B1', 'California TOLL FREE: (877)449-4335');
        $sheet->setCellValue('B2', 'Customer Name:');
        $sheet->setCellValue('B3', 'Account #:');
        $sheet->setCellValue('B4', 'PO #:');
        $sheet->setCellValue('E1', 'California TOTAL AMOUNT LOAD:');
        $sheet->setCellValue('E2', 'California Total Weight (full load attained at 32k lbs):');
        $sheet->setCellValue('E3', 'California Total Piece Count:');
        $sheet->setCellValue('E4', 'Date:');
        $sheet->setCellValue('J4', date('M j, Y'));

        // $sheet->setCellValue('N1', 'Pennsylvania TOLL FREE: (877)449-4335');
        // $sheet->setCellValue('N2', 'Customer Name:');
        // $sheet->setCellValue('N3', 'Account #:');
        // $sheet->setCellValue('N4', 'PO #:');
        $sheet->setCellValue('Q1', 'Pennsylvania TOTAL AMOUNT LOAD:');
        $sheet->setCellValue('Q2', 'Pennsylvania Total Weight (full load attained at 32k lbs):');
        $sheet->setCellValue('Q3', 'Pennsylvania Total Piece Count:');
        // $sheet->setCellValue('P4', 'Date:');
        // $sheet->setCellValue('U4', date('M j, Y'));

        $sheet->setCellValue('Z1', 'Kentucky TOLL FREE: (877)449-4335');
        $sheet->setCellValue('Z2', 'Customer Name:');
        $sheet->setCellValue('Z3', 'Account #:');
        $sheet->setCellValue('Z4', 'PO #:');
        $sheet->setCellValue('AC1', 'Kentucky TOTAL AMOUNT LOAD:');
        $sheet->setCellValue('AC2', 'Kentucky Total Weight (full load attained at 32k lbs):');
        $sheet->setCellValue('AC3', 'Kentucky Total Piece Count:');


        $sheet->setCellValue('A5', 'Brand');
        $sheet->setCellValue('B5', 'Product Code');
        $sheet->setCellValue('C5', 'Size');
        $sheet->setCellValue('D5', 'Load Speed');
        $sheet->setCellValue('E5', "Pattern");
        $sheet->setCellValue('F5', 'Position');
        $sheet->setCellValue('G5', 'Ply');
        $sheet->setCellValue('H5', "California On\nHand Inventory");
        $sheet->setCellValue('I5', 'Price/Tire');
        $sheet->setCellValue('J5', 'Order QTY');
        $sheet->setCellValue('K5', 'Subtotal');
        $sheet->setCellValue('L5', 'F.E.T.');
        $sheet->setCellValue('M5', 'Cut/Fill');

        // $sheet->setCellValue('N5', 'Brand');
        $sheet->setCellValue('O5', 'Product Code');
        $sheet->setCellValue('P5', 'Item Desc');
        $sheet->setCellValue('Q5', "Latrobe On\nHand Inventory");
        $sheet->setCellValue('R5', 'Price/Tire');
        $sheet->setCellValue('S5', 'Order QTY');
        $sheet->setCellValue('T5', 'SubTotal');
        $sheet->setCellValue('U5', 'F.E.T.');
        $sheet->setCellValue('V5', 'Cut/Fill');
        //$sheet->setCellValue('M5', 'Weight');

        $sheet->setCellValue('Z5', 'Brand');
        $sheet->setCellValue('AA5', 'Product Code');
        $sheet->setCellValue('AB5', 'Item Desc');
        $sheet->setCellValue('AC5', "Kentucky On\nHand Inventory");
        $sheet->setCellValue('AD5', 'Price/Tire');
        $sheet->setCellValue('AE5', 'Order QTY');
        $sheet->setCellValue('AF5', 'SubTotal');
        $sheet->setCellValue('AG5', 'F.E.T.');
        $sheet->setCellValue('AH5', 'Cut/Fill');

        $sheet->getStyle('E5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('Q5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('AC5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('A1')->getAlignment()->setWrapText(true);



        $allSize = [];
        $i = 6;
        foreach ($query as $row) {
            $unq_id = $row->VENDOR_CODE . $row->PART_NUMBER;
            $nPart_num = $row->PART_NUMBER;
            $n_brand = $row->DESCRIPTION;

            $sql3 = DB::connection('sqlsrv2')->table($tableName)
                ->select('Part_Number as nPartNumber', 'offer as nOffer', 'LATMWF as LATOffer', 'KENMWF as KENOffer', 'Vendor_Code as nVCode')
                ->where('Part_Number', $nPart_num)
                ->where('offer', '<>', 0)
                ->first();

            $sql4 = DB::connection('sqlsrv')->table('MWF_MASTERTABLE')
                ->select('BRAND as mBrand', 'PRODUCT_CODE as mProdCode', 'SIZE as mSize', 'LOAD_SPEED as mLoadSpeed', 'PATTERN as mPattern', 'POSITION as mPosition', 'PLY as mPly')
                ->where('PRODUCT_CODE', $nPart_num)
                ->first();

            $sheet->getRowDimension($i)->setRowHeight(18.75);

            $cellStyles = [
                'alignment' => [
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                ],
                'fill' => array(
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'startColor' => ($i % 2 == 0) ? array('argb' => 'ffff99') : array('argb' => 'FFFFFF')
                ),
            ];

            $sheet->getStyle('A' . $i . ':M' . $i)->applyFromArray($cellStyles);
            $sheet->getStyle('P' . $i . ':V' . $i)->applyFromArray($cellStyles);
            $sheet->getStyle('Z' . $i . ':AH' . $i)->applyFromArray($cellStyles);

            if ($sql3 !== null && $row->PART_NUMBER == $sql3->nPartNumber && strtoupper($row->VENDOR_CODE) == strtoupper($sql3->nVCode)) {
                $_manufacturer = ($row->DESCRIPTION == 'MISC TIRE') ? $row->LEGACY_PART_REFERENCE . '-' : $row->DESCRIPTION;

                $sheet->setCellValue('A' . $i, $sql4->mBrand ?? '');
                $sheet->setCellValue('B' . $i, $sql4->mProdCode ?? '');
                $sheet->setCellValue('C' . $i, $sql4->mSize ?? '');
                $sheet->setCellValue('D' . $i, $sql4->mLoadSpeed ?? '');
                $sheet->setCellValue('E' . $i, $sql4->mPattern ?? '');
                $sheet->setCellValue('F' . $i, $sql4->mPosition ?? '');
                $sheet->setCellValue('G' . $i, ($sql4 !== null && is_numeric($sql4->mPly)) ? $sql4->mPly : '');

                $weight = ($row->WEIGHT == 0) ? ($size[$row->TIRESIZE_WIDTH . $row->TIRESIZE_RATIO . $row->TIRESIZE_WHEEL]['final_weight'] ?? 0) : $row->WEIGHT;
                $sheet->setCellValue('BE' . $i, number_format($weight, 2));

                $onhand_qty = ($row->Vernon_AVAIL > 200) ? 200 : $row->Vernon_AVAIL;
                $sheet->setCellValue('H' . $i, $onhand_qty);
                $sheet->setCellValue('I' . $i, $sql3->nOffer);
                $sheet->setCellValue('L' . $i, $row->FET);

                $onhand_qty_lat = ($row->Latrobe_AVAIL > 200) ? 200 : $row->Latrobe_AVAIL;
                $sheet->setCellValue('O' . $i, $row->PART_NUMBER);
                $sheet->setCellValue('P' . $i, $row->ITEM_DESC);
                $sheet->setCellValue('Q' . $i, $onhand_qty_lat);
                $sheet->setCellValue('R' . $i, $sql3->LATOffer);
                $sheet->setCellValue('U' . $i, $row->FET);

                $onhand_qty_lat = ($row->KENTUCKY_AVAIL > 200) ? 200 : $row->KENTUCKY_AVAIL;
                $sheet->setCellValue('Z' . $i, $_manufacturer);
                $sheet->setCellValue('AA' . $i, $row->PART_NUMBER);
                $sheet->setCellValue('AB' . $i, $row->ITEM_DESC);
                $sheet->setCellValue('AC' . $i, $onhand_qty_lat);
                $sheet->setCellValue('AD' . $i, $sql3->KENOffer);
                $sheet->setCellValue('AG' . $i, $row->FET);

                $sheet->setCellValue('K' . $i, '=SUM(I' . $i . '*J' . $i . ')+(L' . $i . '*J' . $i . ')');
                $sheet->setCellValue('T' . $i, '=SUM(R' . $i . '*S' . $i . ')+(U' . $i . '*S' . $i . ')');
                $sheet->setCellValue('AF' . $i, '=SUM(AD' . $i . '*AE' . $i . ')+(AG' . $i . '*AE' . $i . ')');
                $sheet->setCellValue('BD' . $i, '=SUM(BE' . $i . '*J' . $i . ')');
                $sheet->setCellValue('X' . $i, '=SUM(BE' . $i . '*S' . $i . ')');
                $sheet->setCellValue('AI' . $i, '=SUM(BE' . $i . '*AE' . $i . ')');

                $sheet->getStyle('J' . $i . ':J' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');
                $sheet->getStyle('S' . $i . ':S' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');
                $sheet->getStyle('AE' . $i . ':AE' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');

                $sheet->getStyle('K' . $i . ':K' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
                $sheet->getStyle('L' . $i . ':L' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                $sheet->getStyle('M' . $i . ':M' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);

                $sheet->getStyle('C' . $i . ':C' . $i)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
                $allSizes = $row->TIRESIZE_WIDTH . $row->TIRESIZE_RATIO . $row->TIRESIZE_WHEEL;
                $allSize[] = $allSizes;

                //$sheet->setAutoFilter('C6:C'.$i);

                $i++;
            }
        }

        $vals = array_count_values($allSize);
        $j = 6;
        foreach ($vals as $val => $key) {
            $addRow = $key + $j;
            $sheet->insertNewRowBefore($addRow);

            $ranges = ['A', 'P', 'Z'];
            foreach ($ranges as $range) {
                $sheet->getStyle($range . $addRow . ':M' . $addRow)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('D9D9D9');
            }

            $j += $key;
        }

        $numberFormatRanges = [
            'I' => '6:I',
            'J' => '6:J',
            'K' => '6:K',
            'L' => '6:L',
            'T' => '6:T',
            'U' => '6:U',
            'AD' => '6:AD',
            'AF' => '6:AF',
            'AG' => '6:AG',
        ];

        foreach ($numberFormatRanges as $col => $range) {
            $sheet->getStyle($col . $range . $j)->getNumberFormat()->setFormatCode('[$$-en-US]* #,##0.00');
        }

        $sumFormulas = [
            'J3' => 'J6:J',
            'J2' => 'BD6:BD',
            'I1' => 'K6:K',
            'U3' => 'S6:S',
            'U2' => 'X6:X',
            'T1' => 'T6:T',
            'AG3' => 'AE6:AE',
            'AG2' => 'AI6:AI',
            'AF1' => 'AF6:AF',
        ];

        foreach ($sumFormulas as $cell => $range) {
            $sheet->setCellValue($cell, '=SUM(' . $range . $j . ')');
        }

        $alignmentRanges = [
            'A1:AH' => 1,
            'J2:J3' => true,
            'K2:K3' => true,
            'V2:V3' => true,
            'AH2:AH3' => true,
        ];

        foreach ($alignmentRanges as $range => $indent) {
            $sheet->getStyle($range . $j)->getAlignment()->setIndent($indent);
        }

        $conditional1 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
        $conditional1->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
        $conditional1->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHANOREQUAL);
        $conditional1->addCondition('32001');
        $conditional1->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\fill::FILL_SOLID);
        $conditional1->getStyle()->getFill()->getEndColor()->setARGB('FF0000');

        $conditionalStyles = $sheet->getStyle('J2')->getConditionalStyles();
        $conditionalStyles[] = $conditional1;
        $sheet->getStyle('J2')->setConditionalStyles($conditionalStyles);

        $formulaTexts = [
            'K2' => 'IF(J2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")',
            'K3' => 'IF(J2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")',
            'V2' => 'IF(U2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")',
            'V3' => 'IF(U2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")',
            'AH2' => 'IF(AG2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")',
            'AH3' => 'IF(AG2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")',
        ];

        foreach ($formulaTexts as $cell => $formula) {
            $sheet->setCellValue($cell, '=' . $formula);
            $sheet->getStyle($cell)->getAlignment()->setWrapText(true);
        }

        $styleArrayAll = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '595959']
                ]
            ]
        ];
        $sheet->getStyle('A1:AH' . $j)->applyFromArray($styleArrayAll);

        /* TAB 2 */
        $spreadsheet->createSheet();
        $sheet2 = $spreadsheet->setActiveSheetIndex(1);
        $sheet2->setTitle('CA, PA & KY ByBrand');
        $sheet2->getPageSetup()->setScale(65);
        $spreadsheet->getActiveSheet()->getSheetView()->setZoomScale(65);

        $sheet2->getPageMargins()->setTop(0);
        $sheet2->getPageMargins()->setRight(0);
        $sheet2->getPageMargins()->setLeft(0);
        $sheet2->getPageMargins()->setBottom(0);

        $sheet2->mergeCells('A1:A3');
        $sheet2->mergeCells('B1:D1');
        $sheet2->mergeCells('C2:D2');
        $sheet2->mergeCells('C3:D3');
        $sheet2->mergeCells('C4:D4');
        $sheet2->mergeCells('E2:I2');
        $sheet2->mergeCells('E3:I3');
        $sheet2->mergeCells('E4:I4');
        $sheet2->mergeCells('I1:J1');
        $sheet2->mergeCells('E1:H1');
        $sheet2->freezePane('A6');
        $sheet2->getColumnDimension('W')->setWidth(0);

        $sheet2->mergeCells('N1:O1');
        $sheet2->mergeCells('Q2:T2');
        $sheet2->mergeCells('Q3:T3');
        $sheet2->mergeCells('Q4:T4');
        $sheet2->mergeCells('T1:U1');
        $sheet2->mergeCells('Q1:S1');

        // KENTUCKY
        $sheet2->mergeCells('Z1:AA1');
        $sheet2->mergeCells('AC2:AF2');
        $sheet2->mergeCells('AC3:AF3');
        $sheet2->mergeCells('AC4:AF4');
        $sheet2->mergeCells('AF1:AG1');
        $sheet2->mergeCells('AC1:AE1');



        $sheet2
            ->getStyle('B2:J4')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FFFFFF');
        $sheet2
            ->getStyle('B1:I1')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FFFFFF');

        $styleArray = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  13,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000']
                ]
            ]
        ];
        $styleArrayTopNum = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
        ];
        $styleArray2 = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  13,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000']
                ]
            ]
        ];
        $styleArrayTopTotal = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
        ];
        $styleHeader = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => array('argb' => 'b8cce4')
            ]
        ];

        $styleHeader_wtd = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  12,
                'color' => array('rgb' => '000000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => array('argb' => 'b8cce4')
            ]
        ];

        $styleTransparent = [
            'font' => [
                'bold'  =>  true,
                'size'  =>  10,
                'color' => array('rgb' => 'FF0000'),
                'name'  =>  'Calibri'
            ],
            'alignment' => [
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            ]
        ];

        $sheet2->getStyle('A1:A3')->applyFromArray($styleHeader_wtd);
        $sheet2->getStyle('B2:D4')->applyFromArray($styleArray2);
        $sheet2->getStyle('E2:J2')->applyFromArray($styleArray);
        $sheet2->getStyle('E3:J3')->applyFromArray($styleArray);
        $sheet2->getStyle('E4:J4')->applyFromArray($styleArray);
        $sheet2->getStyle('B1:C1')->applyFromArray($styleArrayTopNum);
        $sheet2->getStyle('E1:J1')->applyFromArray($styleArrayTopTotal);
        $sheet2->getStyle('A5:M5')->applyFromArray($styleHeader);
        $sheet2->getStyle('K2')->applyFromArray($styleTransparent);
        $sheet2->getStyle('K3')->applyFromArray($styleTransparent);
        $sheet2->getStyle("I1:J1")->getFont()->setSize(18);

        $sheet2->getStyle('O1:J1')->applyFromArray($styleArrayTopTotal);
        $sheet2->getStyle('N2:O4')->applyFromArray($styleArray2);
        $sheet2->getStyle('P2:U2')->applyFromArray($styleArray);
        $sheet2->getStyle('P3:U3')->applyFromArray($styleArray);
        $sheet2->getStyle('P4:U4')->applyFromArray($styleArray);
        $sheet2->getStyle('N1:U1')->applyFromArray($styleArrayTopNum);
        $sheet2->getStyle('V2')->applyFromArray($styleTransparent);
        $sheet2->getStyle('V3')->applyFromArray($styleTransparent);
        $sheet2->getStyle("T1:U1")->getFont()->setSize(18);

        // KENTUCKY
        $sheet2->getStyle('Z1:AA1')->applyFromArray($styleArrayTopTotal);
        $sheet2->getStyle('Z2:AA4')->applyFromArray($styleArray2);
        $sheet2->getStyle('AB2:AG2')->applyFromArray($styleArray);
        $sheet2->getStyle('AB3:AG3')->applyFromArray($styleArray);
        $sheet2->getStyle('AB4:AG4')->applyFromArray($styleArray);
        $sheet2->getStyle('Y1:AG1')->applyFromArray($styleArrayTopNum);
        $sheet2->getStyle('AH2')->applyFromArray($styleTransparent);
        $sheet2->getStyle('AH3')->applyFromArray($styleTransparent);
        $sheet2->getStyle("AF1:AG1")->getFont()->setSize(18);


        $sheet2->getStyle('P5:V5')->applyFromArray($styleHeader);

        // KENTUCKY
        $sheet2->getStyle('Z5:AH5')->applyFromArray($styleHeader);




        $sheet2->getRowDimension('1')->setRowHeight(23.25);
        $sheet2->getRowDimension('2')->setRowHeight(43.5);
        $sheet2->getRowDimension('3')->setRowHeight(43.5);
        $sheet2->getRowDimension('4')->setRowHeight(23.25);
        $sheet2->getRowDimension('5')->setRowHeight(45);

        $sheet2->getColumnDimension('A')->setWidth(21.86);
        $sheet2->getColumnDimension('B')->setWidth(21.43);
        $sheet2->getColumnDimension('C')->setWidth(21.43);
        $sheet2->getColumnDimension('D')->setWidth(15.86);
        $sheet2->getColumnDimension('E')->setWidth(30.86);
        $sheet2->getColumnDimension('F')->setWidth(14.43);
        $sheet2->getColumnDimension('G')->setWidth(11.71);
        $sheet2->getColumnDimension('H')->setWidth(20.43);
        $sheet2->getColumnDimension('I')->setAutoSize(true);
        $sheet2->getColumnDimension('J')->setWidth(15.86);
        $sheet2->getColumnDimension('K')->setWidth(26.23);
        $sheet2->getColumnDimension('L')->setWidth(11.71);
        $sheet2->getColumnDimension('M')->setWidth(15.71);


        // $sheet2->getColumnDimension('N')->setWidth(0);
        $sheet2->getColumnDimension('O')->setWidth(0);
        $sheet2->getColumnDimension('P')->setWidth(0);
        $sheet2->getColumnDimension('Q')->setWidth(12.86);
        $sheet2->getColumnDimension('R')->setWidth(12.43);
        $sheet2->getColumnDimension('S')->setWidth(11.71);
        $sheet2->getColumnDimension('T')->setWidth(12.43);
        $sheet2->getColumnDimension('U')->setAutoSize(true);
        $sheet2->getColumnDimension('V')->setWidth(26.23);
        $sheet2->getColumnDimension('W')->setWidth(0);
        $sheet2->getColumnDimension('X')->setWidth(0);

        // KENTUCKY
        $sheet2->getColumnDimension('Z')->setWidth(0);
        $sheet2->getColumnDimension('AA')->setWidth(0);
        $sheet2->getColumnDimension('AB')->setWidth(0);
        $sheet2->getColumnDimension('AC')->setWidth(12.86);
        $sheet2->getColumnDimension('AD')->setWidth(12.43);
        $sheet2->getColumnDimension('AE')->setWidth(11.71);
        $sheet2->getColumnDimension('AF')->setWidth(12.43);
        $sheet2->getColumnDimension('AG')->setAutoSize(true);
        $sheet2->getColumnDimension('AH')->setWidth(26.23);
        $sheet2->getColumnDimension('AI')->setWidth(0);
        $sheet2->getColumnDimension('AJ')->setWidth(0);


        $sheet2->setCellValue('A1', 'WTD Wholesale Tire Distributors');
        $sheet2->setCellValue('B1', 'California TOLL FREE: (877)449-4335');
        $sheet2->setCellValue('B2', 'Customer Name:');
        $sheet2->setCellValue('B3', 'Account #:');
        $sheet2->setCellValue('B4', 'PO #:');
        $sheet2->setCellValue('E1', 'California TOTAL AMOUNT LOAD:');
        $sheet2->setCellValue('E2', 'California Total Weight (full load attained at 32k lbs):');
        $sheet2->setCellValue('E3', 'California Total Piece Count:');
        $sheet2->setCellValue('E4', 'Date:');
        $sheet2->setCellValue('J4', date('M j, Y'));

        // $sheet2->setCellValue('N1', 'Pennsylvania TOLL FREE: (877)449-4335');
        // $sheet2->setCellValue('N2', 'Customer Name:');
        // $sheet2->setCellValue('N3', 'Account #:');
        // $sheet2->setCellValue('N4', 'PO #:');
        $sheet2->setCellValue('Q1', 'Pennsylvania TOTAL AMOUNT LOAD:');
        $sheet2->setCellValue('Q2', 'Pennsylvania Total Weight (full load attained at 32k lbs):');
        $sheet2->setCellValue('Q3', 'Pennsylvania Total Piece Count:');
        // $sheet2->setCellValue('P4', 'Date:');
        // $sheet2->setCellValue('U4', date('M j, Y'));

        $sheet2->setCellValue('Z1', 'Kentucky TOLL FREE: (877)449-4335');
        $sheet2->setCellValue('Z2', 'Customer Name:');
        $sheet2->setCellValue('Z3', 'Account #:');
        $sheet2->setCellValue('Z4', 'PO #:');
        $sheet2->setCellValue('AC1', 'Kentucky TOTAL AMOUNT LOAD:');
        $sheet2->setCellValue('AC2', 'Kentucky Total Weight (full load attained at 32k lbs):');
        $sheet2->setCellValue('AC3', 'Kentucky Total Piece Count:');


        $sheet2->setCellValue('A5', 'Brand');
        $sheet2->setCellValue('B5', 'Product Code');
        $sheet2->setCellValue('C5', 'Size');
        $sheet2->setCellValue('D5', 'Load Speed');
        $sheet2->setCellValue('E5', "Pattern");
        $sheet2->setCellValue('F5', 'Position');
        $sheet2->setCellValue('G5', 'Ply');
        $sheet2->setCellValue('H5', "California On\nHand Inventory");
        $sheet2->setCellValue('I5', 'Price/Tire');
        $sheet2->setCellValue('J5', 'Order QTY');
        $sheet2->setCellValue('K5', 'Subtotal');
        $sheet2->setCellValue('L5', 'F.E.T.');
        $sheet2->setCellValue('M5', 'Cut/Fill');

        // $sheet2->setCellValue('N5', 'Brand');
        $sheet2->setCellValue('O5', 'Product Code');
        $sheet2->setCellValue('P5', 'Item Desc');
        $sheet2->setCellValue('Q5', "Latrobe On\nHand Inventory");
        $sheet2->setCellValue('R5', 'Price/Tire');
        $sheet2->setCellValue('S5', 'Order QTY');
        $sheet2->setCellValue('T5', 'SubTotal');
        $sheet2->setCellValue('U5', 'F.E.T.');
        $sheet2->setCellValue('V5', 'Cut/Fill');
        //$sheet2->setCellValue('M5', 'Weight');

        $sheet2->setCellValue('Z5', 'Brand');
        $sheet2->setCellValue('AA5', 'Product Code');
        $sheet2->setCellValue('AB5', 'Item Desc');
        $sheet2->setCellValue('AC5', "Kentucky On\nHand Inventory");
        $sheet2->setCellValue('AD5', 'Price/Tire');
        $sheet2->setCellValue('AE5', 'Order QTY');
        $sheet2->setCellValue('AF5', 'SubTotal');
        $sheet2->setCellValue('AG5', 'F.E.T.');
        $sheet2->setCellValue('AH5', 'Cut/Fill');

        $sheet2->getStyle('E5')->getAlignment()->setWrapText(true);
        $sheet2->getStyle('Q5')->getAlignment()->setWrapText(true);
        $sheet2->getStyle('AC5')->getAlignment()->setWrapText(true);
        $sheet2->getStyle('A1')->getAlignment()->setWrapText(true);



        $allSize = [];
        $i = 6;

        foreach ($query2 as $row) {

            $unq_id = $row->VENDOR_CODE . $row->PART_NUMBER;
            $nPart_num = $row->PART_NUMBER;

            $sql3 = DB::connection('sqlsrv2')->table($tableName)
                ->select('Part_Number as nPartNumber', 'offer as nOffer', 'LATMWF as LATOffer', 'KENMWF as KENOffer', 'Vendor_Code as nVCode')
                ->where('Part_Number', $nPart_num)
                ->where('offer', '<>', 0)
                ->first();
            // dd($sql3);

            $sql4 = DB::connection('sqlsrv')->table('MWF_MASTERTABLE')
                ->select('BRAND as mBrand', 'PRODUCT_CODE as mProdCode', 'SIZE as mSize', 'LOAD_SPEED as mLoadSpeed', 'PATTERN as mPattern', 'POSITION as mPosition', 'PLY as mPly')
                ->where('PRODUCT_CODE', $nPart_num)
                ->first();

            // $sheet2->getRowDimension($i)->setRowHeight(18.75);

            if ($i % 2 == 0) {

                $sheet2->getStyle('A' . $i . ':M' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'ffff99')

                        ),
                    )
                );
                $sheet2->getStyle('P' . $i . ':V' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'ffff99')

                        ),
                    )
                );

                $sheet2->getStyle('Z' . $i . ':AH' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'ffff99')

                        ),
                    )
                );
            } else {
                $sheet2->getStyle('A' . $i . ':J' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'FFFFFF')

                        ),
                    )
                );
                $sheet2->getStyle('P' . $i . ':V' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'FFFFFF')

                        ),
                    )
                );
                $sheet2->getStyle('Z' . $i . ':AH' . $i)->applyFromArray(
                    array(
                        'alignment' => [
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        ],
                        'fill' => array(
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => array('argb' => 'FFFFFF')

                        ),
                    )
                );
            }

            if ($sql3 !== null && $row->PART_NUMBER == $sql3->nPartNumber && strtoupper($row->VENDOR_CODE) == strtoupper($sql3->nVCode)) {
                if (strtoupper($row->DESCRIPTION) == 'MISC TIRE') {
                    $_manufacturer = $row->LEGACY_PART_REFERENCE . '-';
                } else {
                    $_manufacturer = $row->DESCRIPTION;
                }

                // $sheet2->setCellValue('A'.$i, $_manufacturer);
                $sheet2->setCellValue('A' . $i, $sql4->mBrand ?? '');
                $sheet2->setCellValue('B' . $i, $sql4->mProdCode ?? '');


                $sheet2->setCellValue('C' . $i, $sql4->mSize ?? '');
                $sheet2->setCellValue('D' . $i, $sql4->mLoadSpeed ?? '');

                $sheet2->setCellValue('E' . $i, $sql4->mPattern ?? '');
                // if($sql4['mPosition'] == 0){
                //     $sheet2->setCellValue('F'.$i, '');
                // }else{
                //     $sheet2->setCellValue('F'.$i, $sql4['mPosition']);
                // }

                $sheet2->setCellValue('F' . $i, $sql4->mPosition ?? '');
                $sheet2->setCellValue('G' . $i, $sql4 !== null && is_numeric($sql4->mPly) ? $sql4->mPly : '');
                // $sheet2->setCellValue('G'.$i, $sql4->mPly);

                if ($row->WEIGHT == 0) {
                    if (isset($size[$row->TIRESIZE_WIDTH . $row->TIRESIZE_RATIO . $row->TIRESIZE_WHEEL]['final_weight'])) {
                        $weight = $size[$row->TIRESIZE_WIDTH . $row->TIRESIZE_RATIO . $row->TIRESIZE_WHEEL]['final_weight'];
                    } else {
                        $weight = 0;
                    }
                } else {
                    $weight = $row->WEIGHT;
                }
                $sheet2->setCellValue('BE' . $i, number_format($weight, 2));

                if ($row->Vernon_AVAIL > 200) {
                    $onhand_qty = 200;
                } else {
                    $onhand_qty = $row->Vernon_AVAIL;
                }
                $sheet2->setCellValue('H' . $i, $onhand_qty);
                $sheet2->setCellValue('I' . $i, $sql3->nOffer);
                $sheet2->setCellValue('L' . $i, $row->FET);


                if ($row->Latrobe_AVAIL > 200) {
                    $onhand_qty_lat = 200;
                } else {
                    $onhand_qty_lat = $row->Latrobe_AVAIL;
                }
                // $sheet2->setCellValue('N'.$i, $_manufacturer);
                $sheet2->setCellValue('O' . $i, $row->PART_NUMBER);
                $sheet2->setCellValue('P' . $i, $row->ITEM_DESC);
                $sheet2->setCellValue('Q' . $i, $onhand_qty_lat);
                $sheet2->setCellValue('R' . $i, $sql3->LATOffer);
                $sheet2->setCellValue('U' . $i, $row->FET);

                // KENTUCKY
                if ($row->KENTUCKY_AVAIL > 200) {
                    $onhand_qty_lat = 200;
                } else {
                    $onhand_qty_lat = $row->KENTUCKY_AVAIL;
                }
                $sheet2->setCellValue('Z' . $i, $_manufacturer);
                $sheet2->setCellValue('AA' . $i, $row->PART_NUMBER);
                $sheet2->setCellValue('AB' . $i, $row->ITEM_DESC);
                $sheet2->setCellValue('AC' . $i, $onhand_qty_lat);
                $sheet2->setCellValue('AD' . $i, $sql3->KENOffer);
                $sheet2->setCellValue('AG' . $i, $row->FET);


                $sheet2->setCellValue(
                    'K' . $i,
                    '=SUM(I' . $i . '*J' . $i . ')+(L' . $i . '*J' . $i . ')'
                );

                $sheet2->setCellValue(
                    'T' . $i,
                    '=SUM(R' . $i . '*S' . $i . ')+(U' . $i . '*S' . $i . ')'
                );

                $sheet2->setCellValue(
                    'AF' . $i,
                    '=SUM(AD' . $i . '*AE' . $i . ')+(AG' . $i . '*AE' . $i . ')'
                );

                $sheet2->setCellValue(
                    'BD' . $i,
                    '=SUM(BE' . $i . '*J' . $i . ')'
                );

                $sheet2->setCellValue(
                    'X' . $i,
                    '=SUM(BE' . $i . '*S' . $i . ')'
                );

                $sheet2->setCellValue(
                    'AI' . $i,
                    '=SUM(BE' . $i . '*AE' . $i . ')'
                );

                $sheet2->getStyle('J' . $i . ':J' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');
                $sheet2->getStyle('S' . $i . ':S' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');
                $sheet2->getStyle('AE' . $i . ':AE' . $i)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('86CB6B');

                $sheet2->getStyle('K' . $i . ':K' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
                $sheet2->getStyle('L' . $i . ':L' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                $sheet2->getStyle('M' . $i . ':M' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);

                $sheet2->getStyle('C' . $i . ':C' . $i)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
                $allSizes = $row->TIRESIZE_WIDTH . $row->TIRESIZE_RATIO . $row->TIRESIZE_WHEEL;
                $allSize[] = $allSizes;

                //$sheet2->setAutoFilter('C6:C'.$i);

                $i++;
            }
        }



        $vals = array_count_values($allSize);
        $j = 6;
        foreach ($vals as $val => $key) {
            $addRow = $key + $j;
            $sheet2->insertNewRowBefore($addRow);
            $sheet2->getStyle('A' . $addRow . ':M' . $addRow)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('D9D9D9');
            $sheet2->getStyle('P' . $addRow . ':W' . $addRow)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('D9D9D9');
            $sheet2->getStyle('Z' . $addRow . ':AJ' . $addRow)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('D9D9D9');
            $j++;
            $j = $key + $j;
        }

        $sheet2->getStyle('I6:I' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('L6:L' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('K6:K' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('I1:J1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

        $sheet2->getStyle('R6:R' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('T6:T' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('U6:U' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('T1:U1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        // KENTUCKY
        $sheet2->getStyle('AD6:AD' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AF6:AF' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AG6:AG' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AF1:AG1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

        // $sheet2->getStyle('L5:L'.$row_count)->getNumberFormat()->setFormatCode(PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->setCellValue(
            'J3',
            '=SUM(J6:J' . $j . ')'
        );
        // $sheet2->getStyle('L4')->getNumberFormat()->setFormatCode(PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->setCellValue(
            'J2',
            '=SUM(BD6:BD' . $j . ')'
        );
        $sheet2->setCellValue(
            'I1',
            '=SUM(K6:K' . $j . ')'
        );

        $sheet2->setCellValue(
            'U3',
            '=SUM(S6:S' . $j . ')'
        );
        $sheet2->setCellValue(
            'U2',
            '=SUM(X6:X' . $j . ')'
        );
        $sheet2->setCellValue(
            'T1',
            '=SUM(T6:T' . $j . ')'
        );

        $sheet2->setCellValue(
            'AG3',
            '=SUM(AE6:AE' . $j . ')'
        );
        $sheet2->setCellValue(
            'AG2',
            '=SUM(AI6:AI' . $j . ')'
        );
        $sheet2->setCellValue(
            'AF1',
            '=SUM(AF6:AF' . $j . ')'
        );

        $sheet2->getStyle('A1:AH' . $j)->getAlignment()->setIndent(1);


        $conditional1 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
        $conditional1->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
        $conditional1->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHANOREQUAL);
        $conditional1->addCondition('32001');
        $conditional1->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\fill::FILL_SOLID);
        $conditional1->getStyle()->getFill()->getEndColor()->setARGB('FF0000');

        $conditionalStyles = $sheet2->getStyle('J2')->getConditionalStyles();
        $conditionalStyles[] = $conditional1;

        $sheet2->getStyle('J2')->setConditionalStyles($conditionalStyles);


        $sheet2->setCellValue('K2', '=IF(J2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('K2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('K3', '=IF(J2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('K3')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('V2', '=IF(U2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('V2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('V3', '=IF(U2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('V3')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('AH2', '=IF(AG2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('AH2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('AH3', '=IF(AG2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('AH3')->getAlignment()->setWrapText(true);


        $styleArrayAll = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '595959']
                ]
            ]
        ];
        $sheet2->getStyle('A1:AH' . $j)->applyFromArray($styleArrayAll);
        $spreadsheet->setActiveSheetIndex(0);
        $sheet2->getStyle('I6:I' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('L6:L' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('K6:K' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('I1:J1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

        $sheet2->getStyle('R6:R' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('T6:T' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('U6:U' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('T1:U1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        // KENTUCKY
        $sheet2->getStyle('AD6:AD' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AF6:AF' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AG6:AG' . $j)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->getStyle('AF1:AG1')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

        // $sheet2->getStyle('L5:L'.$row_count)->getNumberFormat()->setFormatCode(PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->setCellValue(
            'J3',
            '=SUM(J6:J' . $j . ')'
        );
        // $sheet2->getStyle('L4')->getNumberFormat()->setFormatCode(PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        $sheet2->setCellValue(
            'J2',
            '=SUM(BD6:BD' . $j . ')'
        );
        $sheet2->setCellValue(
            'I1',
            '=SUM(K6:K' . $j . ')'
        );

        $sheet2->setCellValue(
            'U3',
            '=SUM(S6:S' . $j . ')'
        );
        $sheet2->setCellValue(
            'U2',
            '=SUM(X6:X' . $j . ')'
        );
        $sheet2->setCellValue(
            'T1',
            '=SUM(T6:T' . $j . ')'
        );

        $sheet2->setCellValue(
            'AG3',
            '=SUM(AE6:AE' . $j . ')'
        );
        $sheet2->setCellValue(
            'AG2',
            '=SUM(AI6:AI' . $j . ')'
        );
        $sheet2->setCellValue(
            'AF1',
            '=SUM(AF6:AF' . $j . ')'
        );

        $sheet2->getStyle('A1:AH' . $j)->getAlignment()->setIndent(1);


        $conditional1 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
        $conditional1->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
        $conditional1->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHANOREQUAL);
        $conditional1->addCondition('32001');
        $conditional1->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\fill::FILL_SOLID);
        $conditional1->getStyle()->getFill()->getEndColor()->setARGB('FF0000');

        $conditionalStyles = $sheet2->getStyle('J2')->getConditionalStyles();
        $conditionalStyles[] = $conditional1;

        $sheet2->getStyle('J2')->setConditionalStyles($conditionalStyles);


        $sheet2->setCellValue('K2', '=IF(J2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('K2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('K3', '=IF(J2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('K3')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('V2', '=IF(U2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('V2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('V3', '=IF(U2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('V3')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('AH2', '=IF(AG2>32000,"WEIGHT EXCEEDS LIMIT, PLEASE ASSIGN CUT ITEM","")');
        $sheet2->getStyle('AH2')->getAlignment()->setWrapText(true);

        $sheet2->setCellValue('AH3', '=IF(AG2<32000,"WEIGHT UNDER LIMIT, PLEASE ASSIGN FILL ITEM","")');
        $sheet2->getStyle('AH3')->getAlignment()->setWrapText(true);


        $styleArrayAll = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '595959']
                ]
            ]
        ];
        $sheet2->getStyle('A1:AH' . $j)->applyFromArray($styleArrayAll);
        $spreadsheet->setActiveSheetIndex(0);


        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $fileName = $tableName . date('Y-m-d') . '.xlsx';
        $directory = 'Zoe Report';
        $filePath = public_path($directory . '/' . $fileName);

        $writer->save($filePath);
    }
}
