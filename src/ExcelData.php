<?php

namespace ExcelCompare;

use Faker\Factory as FakeData;

class ExcelData
{
    private $fake;

    public function __construct()
    {
        $this->fake = FakeData::create();
    }

    public function getText(int $length = 200): string
    {
        return $this->fake->text($length);
    }

    public function getNumber(): int
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
}
