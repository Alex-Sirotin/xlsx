<?php

namespace ExcelCompare;

use Faker\Generator as FakeData;
class ExcelData
{
    /** @var FakeData */
    private $fake;
    /** @var int  */
    private $rowCount;
    /** @var array  */
    private $data = [];

    public function __construct(int $rowCount)
    {
        $this->rowCount = $rowCount;
        $this->fake = new FakeData();
    }
    private function addRow(): array
    {
        return [
            $this->fake->name(),
            $this->fake->title(),
            $this->fake->creditCardNumber(),
            $this->fake->date(),
            $this->fake->dayOfWeek(),
            $this->fake->monthName(),
            $this->fake->year(),
            $this->fake->time(),
            $this->fake->dateTime(),
            $this->fake->randomNumber(),
            $this->fake->randomFloat(),
            $this->fake->text(),
            $this->fake->password(),
            $this->fake->unixTime(),
        ];
    }
    /**
     * Генерация массива данных
     *
     * @param int $rowCount Количесвто строк, если не задано - будет использовано из конструктора
     * @return void
     */
    public function generate(int $rowCount = null)
    {
        if ($rowCount) {
            $this->$rowCount = $rowCount;
        }
        $this->data = [];
        foreach(range(1, $this->rowCount) as $i) {
            $this->data[] = array_merge($i, $this->addRow());
        }
    }
    public function getData(): array
    {
        return $this->data;
    }
}
