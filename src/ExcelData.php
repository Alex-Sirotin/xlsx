<?php

namespace ExcelCompare;

use Faker\Factory as FakeData;

class ExcelData
{
    private $fake;
//    /** @var int  */
//    private int $rowCount;
//    /** @var array  */
//    private array $data = [];

    public function __construct()
    {
//        $this->rowCount = $rowCount;
        //$this->fake = new FakeData();
        $this->fake = FakeData::create();
    }

    public function getText(int $length = 200): string
    {
        return $this->fake->text($length);
    }

    public function getNumber(): string
    {
        return $this->fake->unique()->randomNumber();
    }

    public function getFloat(): float
    {
        return $this->fake->unique()->randomFloat();
    }

    public function getDateTime(): \DateTime
    {
        return $this->fake->dateTime();
    }

    public function getUnixTime(): int
    {
        return $this->fake->unixTime();
    }

    public function getTime(): string
    {
        return $this->fake->time();
    }

    public function getDate(): string
    {
        return $this->fake->date();
    }

    public function getLongNumber(): string
    {
        return $this->fake->creditCardNumber(null, false, '');
    }

    public function getName(string $gender = null): string
    {
        return $this->fake->name($gender);
    }

//    private function addRow(): array
//    {
//        return [
//            $this->fake->text(),
//            $this->fake->dateTime(),
//            $this->fake->randomNumber(),
//            $this->fake->randomFloat(),
//            $this->fake->unixTime(),
//            $this->fake->time(),
//            $this->fake->date(),
//            $this->fake->creditCardNumber(),
//            $this->fake->name(),

//            $this->fake->title(),
//            $this->fake->dayOfWeek(),
//            $this->fake->monthName(),
//            $this->fake->year(),
//            $this->fake->password(),
//        ];
//    }
//
//    /**
//     * Генерация массива данных
//     *
//     * @param int|null $rowCount Количесвто строк, если не задано - будет использовано из конструктора
//     * @return void
//     */
//    public function generate(int $rowCount = null)
//    {
//        if ($rowCount) {
//            $this->$rowCount = $rowCount;
//        }
//        $this->data = [];
//        foreach(range(1, $this->rowCount) as $i) {
//            $this->data[] = array_merge($i, $this->addRow());
//        }
//    }
//    public function getData(): array
//    {
//        return $this->data;
//    }
}
