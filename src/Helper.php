<?php

use dashingUnique\excel\DashingExcel;

if (!function_exists('DashingExcel')) {
    /**
     * @param null $data
     * @return DashingExcel
     */
    function DashingExcel($data = null)
    {
        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }
        return blank($data) ? new DashingExcel() : new DashingExcel($data);
    }
}