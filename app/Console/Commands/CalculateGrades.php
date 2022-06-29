<?php

namespace App\Console\Commands;

use App\Exports\GradesExport;
use Carbon\Carbon;
use Illuminate\Console\Command;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Excel;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

class CalculateGrades extends Command
{
    private Excel $excel;

    public function __construct(Excel $excel, Carbon $carbon)
    {
        parent::__construct();
        $this->excel = $excel;
        $this->carbon = $carbon;
    }

    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'hoffelijk:calculate-grades {filename?}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Calculate the grades from a XLSX file.';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $storageFilepath = $this->getStorageFilepath();

        if (!$storageFilepath) {
            $this->info('File not found.');
        }

        $gradesCollection = $this->excel->toCollection(collect(), $storageFilepath, null, Excel::XLSX)->first();

        // [1] is the max question score in the document
        $maxQuestionScoreInt = $this->maxQuestionScore($gradesCollection[1]);

        $studentTotalScoreCollection = $this->calculateStudentScores($gradesCollection, $maxQuestionScoreInt);

        $exportedFile = $this->calculatedScoresExportToExcel($studentTotalScoreCollection);

        $this->info("The calculated file is stored inside storage/app/public/$exportedFile.xlsx");

        return true;
    }

    private function getFilename(): string
    {
        $filename = $this->argument('filename');

        if (!$filename) {
            $filename = $this->ask('What is the filename?');
        }

        return $filename;
    }

    private function maxQuestionScore($maxQuestionScore): int
    {
        $maxQuestionScoreInt = 0;

        // forget text from the array
        $maxQuestionScore->forget(0);

        foreach ($maxQuestionScore as $score) {
            $maxQuestionScoreInt += $score;
        }

        return $maxQuestionScoreInt;
    }

    private function getStorageFilepath(): string
    {
        $filenameWithXlsxExtension = $this->getFilename() . '.xlsx';
        return storage_path() . '/gradelists/' . $filenameWithXlsxExtension;
    }

    private function calculateStudentScores($gradesCollection, $maxQuestionScoreInt): Collection
    {
        $studentTotalScoreCollection = collect();

        // skip the first two lines with $i=2
        for ($i = 2; $i < count($gradesCollection); $i++) {
            $studentTotalScore = 0;
            $studentCode = $gradesCollection[$i][0];
            $gradesCollection[$i]->forget(0);

            // adds the question grade to the studentTotalScore
            foreach ($gradesCollection[$i] as $grade) {
                $studentTotalScore += $grade;
            }

            $studentTotalScoreCollection->put($studentCode, ['studentcode' => $studentCode, 'student total score' => $studentTotalScore, 'student calculated grade' => $this->studentCalculatedGrade($studentTotalScore, $maxQuestionScoreInt)]);
        }

        return $studentTotalScoreCollection;
    }

    private function studentCalculatedGrade(int $studentTotalScore, int $maxQuestionScoreInt): float
    {
        $studentScorePercentage = ($studentTotalScore / $maxQuestionScoreInt) * 100;
        $grade = 0;

        if ($studentScorePercentage <= 20) {
            $grade = 1.0;
        }

        if ($studentScorePercentage > 20 && $studentScorePercentage < 70) {
            $grade = 1 + ($studentTotalScore * round(4.5 / ($maxQuestionScoreInt / 100 * 70), 2));
        }

        if ($studentScorePercentage >= 70) {
            $grade = 5.5 + ($studentTotalScore - ($maxQuestionScoreInt / 100 * 70)) * round(4.5 / ($maxQuestionScoreInt / 100 * 30), 2);
        }

        if ($studentScorePercentage >= 100) {
            $grade = 10.0;
        }

        return number_format($grade, 1, '.', '');
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    private function calculatedScoresExportToExcel(Collection $studentTotalScoreCollection): string
    {
        $filename = $this->carbon->now()->format('YmdHs');
        $this->excel->store(new GradesExport($studentTotalScoreCollection), "$filename.xlsx", 'public', Excel::XLSX);
        return $filename;
    }
}
