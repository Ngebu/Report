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
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use App\Http\Controllers\ProcessReportControllerVernon;


class ReportController extends Controller
{
    protected $processReportController;
    protected $processReportControllerVernon;

    public function __construct(ProcessReportController $processReportController, ProcessReportControllerVernon $processReportControllerVernon)
    {
        $this->processReportController = $processReportController;
        $this->processReportControllerVernon = $processReportControllerVernon;
    }


    public function tablesVLK()
    {
        
        // all tables (VERNON ONLY)
        $tableNamesVLK = [
            'ACETIREMIAMI',
            'ALFATIRES',
            'ALLAMERICANFLEET',
            'ALMATIRES',
            'ALPHATIRES',
            'ARMENMWF_FASHIONWHEELSWHOLESALE',
            'ARMENMWF_H_AND_SWHEELDISTRIBUTORS',
            'ARMENMWF_USAWHEELANDTIRE',
            'ARROWSHOP',
            'BAUERBUILT',
            'BEACONTIRE',
            'BESTONE',
            'BILLWILLIAMSTIRE',
            'BLACKSTIRESERVICE',
            'CATIREWHOLESALE',
            'CLARKTIREWHOLESALE',
            'COMPASSTIRE',
            'COMPETITIONWHEELS',
            'CONNECTICUTTIRE',
        ];

        foreach ($tableNamesVLK as $tableName) {
            $this->processReportController->processReport($tableName);
        }
    }

    // Call this from the frontend if you only want to retrieve the value for VernonOnly.
    public function tablesV()
    {
        // all tables (VERNON ONLY)
        $tableNamesV = [
            'ALEXSTIREINC',
            'ATRCORPATDCOVINA'

        ];

        foreach ($tableNamesV as $tableNameV) {
            $this->processReportControllerVernon->processReportVernonOnly($tableNameV);
        }
    }
}
